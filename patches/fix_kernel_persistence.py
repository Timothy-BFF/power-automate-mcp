"""
Patch: Persistent Kernel Manager with Session-Aware Routing
============================================================

File to modify: main.py (kernel/session management section)
Branch: fix/kernel-persistence-and-preload
Issue: #3 — Sandbox kernel state not persisting across execute_code calls

PROBLEM:
    Each execute_code call currently gets a fresh Python kernel,
    losing all imports, variables, and state from previous calls.
    This breaks any multi-step workflow (PDF processing, data analysis, etc.)

FIX:
    Implement a persistent kernel manager that:
    1. Maps session_id -> kernel instance (keeps kernel alive between calls)
    2. Health-checks kernels before reuse (restarts dead ones)
    3. Auto-cleans idle kernels after configurable timeout
    4. Pre-loads common libraries on kernel creation (via kernel_startup.py)
    5. Thread-safe access for concurrent requests
"""

import threading
import time
import logging
from typing import Dict, Optional, Any
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)


# =============================================================================
# CONFIGURATION
# =============================================================================

KERNEL_CONFIG = {
    # How long a kernel can sit idle before being cleaned up
    "idle_timeout_minutes": 30,

    # Maximum number of concurrent kernels (memory protection)
    "max_kernels": 6,

    # How often the cleanup thread checks for idle kernels
    "cleanup_interval_seconds": 60,

    # Whether to pre-load common libraries on kernel creation
    "preload_libraries": True,

    # Maximum consecutive errors before force-recycling a kernel
    "max_consecutive_errors": 5,
}


# =============================================================================
# IMPORT STARTUP CODE FROM COMPANION MODULE
# =============================================================================

try:
    from patches.kernel_startup import KERNEL_STARTUP_CODE
except ImportError:
    # Fallback minimal startup if companion module not found
    KERNEL_STARTUP_CODE = "import os, sys, json, base64, pathlib"
    logger.warning(
        "kernel_startup.py not found — using minimal startup code. "
        "Install patches/kernel_startup.py for full library pre-loading."
    )


# =============================================================================
# KERNEL WRAPPER
# =============================================================================

class ManagedKernel:
    """
    Wraps a sandbox kernel instance with metadata for lifecycle management.
    
    Tracks:
    - Creation time and last-used time (for idle detection)
    - Execution count (for monitoring)
    - Consecutive error count (for health assessment)
    - Pre-load status
    """

    def __init__(self, session_id: str, kernel: Any):
        self.session_id = session_id
        self.kernel = kernel
        self.created_at = datetime.utcnow()
        self.last_used_at = datetime.utcnow()
        self.execution_count = 0
        self.consecutive_errors = 0
        self.is_preloaded = False

    def touch(self):
        """Update last-used timestamp."""
        self.last_used_at = datetime.utcnow()
        self.execution_count += 1

    def record_success(self):
        """Reset error counter on successful execution."""
        self.consecutive_errors = 0

    def record_error(self):
        """Increment error counter."""
        self.consecutive_errors += 1

    def is_idle(self, timeout_minutes: int) -> bool:
        """Check if kernel has been idle longer than timeout."""
        cutoff = datetime.utcnow() - timedelta(minutes=timeout_minutes)
        return self.last_used_at < cutoff

    def is_healthy(self) -> bool:
        """
        Check if the kernel is still alive and hasn't exceeded error threshold.
        
        INTEGRATION NOTE:
            Replace `self.kernel.is_alive()` with whatever health check
            method your actual kernel implementation provides. Common options:
                - self.kernel.is_alive()
                - self.kernel.client.is_alive()
                - self.kernel.poll() is None  (for subprocess-based kernels)
        """
        try:
            if self.consecutive_errors >= KERNEL_CONFIG["max_consecutive_errors"]:
                logger.warning(
                    f"Kernel {self.session_id} exceeded max consecutive errors "
                    f"({self.consecutive_errors}), marking unhealthy"
                )
                return False

            # ============================================================
            # INTEGRATION POINT: Replace with your kernel's health check
            # ============================================================
            alive = self.kernel.is_alive()
            return alive

        except Exception as e:
            logger.error(f"Kernel health check failed for {self.session_id}: {e}")
            return False

    def execute(self, code: str) -> Any:
        """
        Execute code on this kernel.
        
        INTEGRATION NOTE:
            Replace `self.kernel.execute(code)` with your actual
            kernel execution method.
        """
        self.touch()
        try:
            # ============================================================
            # INTEGRATION POINT: Replace with your kernel's execute method
            # ============================================================
            result = self.kernel.execute(code)
            self.record_success()
            return result
        except Exception as e:
            self.record_error()
            raise

    def shutdown(self):
        """
        Gracefully shut down this kernel.
        
        INTEGRATION NOTE:
            Replace with your kernel's shutdown method.
        """
        try:
            # ============================================================
            # INTEGRATION POINT: Replace with your kernel's shutdown method
            # ============================================================
            self.kernel.shutdown()
            logger.info(
                f"Kernel {self.session_id} shut down "
                f"(lived {(datetime.utcnow() - self.created_at).total_seconds():.0f}s, "
                f"{self.execution_count} executions)"
            )
        except Exception as e:
            logger.error(f"Error shutting down kernel {self.session_id}: {e}")


# =============================================================================
# PERSISTENT KERNEL MANAGER
# =============================================================================

class PersistentKernelManager:
    """
    Manages sandbox kernel lifecycle with session-aware persistence.

    Key behaviors:
    - Same session_id always gets the same kernel (state persists!)
    - Dead kernels are automatically replaced
    - Idle kernels are cleaned up after configurable timeout
    - New kernels are pre-loaded with common libraries
    - Thread-safe for concurrent MCP requests
    - Respects max kernel limit (memory protection)

    Usage in main.py:

        # At module level:
        kernel_manager = PersistentKernelManager(
            create_kernel_fn=your_create_kernel_function
        )

        # In handle_tool_call():
        if tool_name == "execute_code":
            session_id = args.get("session_id", "default")
            managed = kernel_manager.get_or_create(session_id)
            result = managed.execute(code)
    """

    def __init__(self, create_kernel_fn, config: dict = None):
        """
        Args:
            create_kernel_fn: Callable that creates a new raw kernel instance.
                              This is YOUR existing kernel creation function.
            config: Optional config overrides (defaults to KERNEL_CONFIG).
        """
        self._create_kernel_fn = create_kernel_fn
        self._config = config or KERNEL_CONFIG
        self._kernels: Dict[str, ManagedKernel] = {}
        self._lock = threading.Lock()
        self._cleanup_thread = None
        self._running = False

        # Start background cleanup
        self._start_cleanup_thread()

        logger.info(
            f"PersistentKernelManager initialized "
            f"(max_kernels={self._config['max_kernels']}, "
            f"idle_timeout={self._config['idle_timeout_minutes']}min, "
            f"preload={self._config['preload_libraries']})"
        )

    # =========================================================================
    # MAIN ENTRY POINT
    # =========================================================================

    def get_or_create(self, session_id: str = "default") -> ManagedKernel:
        """
        Get an existing kernel for this session, or create a new one.
        
        THIS IS THE KEY METHOD. Call this from handle_tool_call().
        Same session_id = same kernel = state persists.

        Args:
            session_id: Session identifier. Defaults to "default".
                        All calls with the same session_id share state.

        Returns:
            ManagedKernel instance ready for code execution.

        Raises:
            RuntimeError: If max kernel limit reached and no idle kernels
                          can be evicted.
        """
        with self._lock:
            # ----- Case 1: Existing healthy kernel -----
            if session_id in self._kernels:
                managed = self._kernels[session_id]
                if managed.is_healthy():
                    logger.info(
                        f"Reusing kernel for session '{session_id}' "
                        f"(exec #{managed.execution_count + 1}, "
                        f"alive {(datetime.utcnow() - managed.created_at).total_seconds():.0f}s)"
                    )
                    return managed
                else:
                    logger.warning(
                        f"Kernel for session '{session_id}' is unhealthy, replacing"
                    )
                    managed.shutdown()
                    del self._kernels[session_id]

            # ----- Case 2: Need new kernel, check capacity -----
            if len(self._kernels) >= self._config["max_kernels"]:
                evicted = self._evict_idle_kernel()
                if not evicted:
                    raise RuntimeError(
                        f"Max kernel limit reached ({self._config['max_kernels']}). "
                        f"No idle kernels available for eviction. "
                        f"Active sessions: {list(self._kernels.keys())}"
                    )

            # ----- Case 3: Create new kernel -----
            return self._create_new_kernel(session_id)

    # =========================================================================
    # KERNEL CREATION
    # =========================================================================

    def _create_new_kernel(self, session_id: str) -> ManagedKernel:
        """Create a new kernel, pre-load libraries, and register it."""
        logger.info(f"Creating new kernel for session '{session_id}'")

        # Create the raw kernel using the provided factory function
        raw_kernel = self._create_kernel_fn()

        # Wrap it in our managed wrapper
        managed = ManagedKernel(session_id=session_id, kernel=raw_kernel)

        # Pre-load common libraries if configured
        if self._config["preload_libraries"]:
            try:
                managed.kernel.execute(KERNEL_STARTUP_CODE)
                managed.is_preloaded = True
                logger.info(f"Pre-loaded libraries for session '{session_id}'")
            except Exception as e:
                logger.warning(
                    f"Library pre-load failed for session '{session_id}': {e}. "
                    f"Kernel is still usable, libraries will need manual import."
                )

        # Register in our map
        self._kernels[session_id] = managed

        logger.info(
            f"Kernel created for session '{session_id}' "
            f"(total active: {len(self._kernels)}/{self._config['max_kernels']})"
        )

        return managed

    # =========================================================================
    # EVICTION
    # =========================================================================

    def _evict_idle_kernel(self) -> bool:
        """
        Evict the oldest idle kernel to make room for a new one.
        Returns True if a kernel was evicted, False if none were idle.
        Falls back to LRU eviction if no kernels are past idle timeout.
        """
        # First try: evict kernels past idle timeout
        oldest_idle = None
        oldest_time = datetime.utcnow()

        for sid, managed in self._kernels.items():
            if managed.is_idle(self._config["idle_timeout_minutes"]):
                if managed.last_used_at < oldest_time:
                    oldest_idle = sid
                    oldest_time = managed.last_used_at

        if oldest_idle:
            logger.info(f"Evicting idle kernel for session '{oldest_idle}'")
            self._kernels[oldest_idle].shutdown()
            del self._kernels[oldest_idle]
            return True

        # Fallback: force-evict least recently used
        if self._kernels:
            lru_sid = min(
                self._kernels,
                key=lambda s: self._kernels[s].last_used_at
            )
            logger.warning(
                f"No idle kernels — force-evicting LRU session '{lru_sid}' "
                f"(last used {self._kernels[lru_sid].last_used_at.isoformat()})"
            )
            self._kernels[lru_sid].shutdown()
            del self._kernels[lru_sid]
            return True

        return False

    # =========================================================================
    # BACKGROUND CLEANUP
    # =========================================================================

    def _start_cleanup_thread(self):
        """Start background thread that periodically cleans up idle kernels."""
        self._running = True
        self._cleanup_thread = threading.Thread(
            target=self._cleanup_loop,
            daemon=True,
            name="kernel-cleanup"
        )
        self._cleanup_thread.start()

    def _cleanup_loop(self):
        """Periodically check for and remove idle kernels."""
        while self._running:
            time.sleep(self._config["cleanup_interval_seconds"])
            self._cleanup_idle_kernels()

    def _cleanup_idle_kernels(self):
        """Remove all kernels that have been idle beyond the timeout."""
        with self._lock:
            to_remove = []
            for sid, managed in self._kernels.items():
                if managed.is_idle(self._config["idle_timeout_minutes"]):
                    to_remove.append(sid)

            for sid in to_remove:
                logger.info(
                    f"Cleaning up idle kernel for session '{sid}' "
                    f"(idle for >{self._config['idle_timeout_minutes']}min)"
                )
                self._kernels[sid].shutdown()
                del self._kernels[sid]

            if to_remove:
                logger.info(
                    f"Cleaned up {len(to_remove)} idle kernel(s). "
                    f"Active: {len(self._kernels)}/{self._config['max_kernels']}"
                )

    # =========================================================================
    # STATUS & MONITORING
    # =========================================================================

    def get_status(self) -> dict:
        """
        Get current kernel manager status for monitoring/debugging.
        
        Can be exposed via an API endpoint:
            @app.get("/kernels/status")
            def kernel_status():
                return kernel_manager.get_status()
        """
        with self._lock:
            return {
                "active_kernels": len(self._kernels),
                "max_kernels": self._config["max_kernels"],
                "sessions": {
                    sid: {
                        "created_at": m.created_at.isoformat(),
                        "last_used_at": m.last_used_at.isoformat(),
                        "execution_count": m.execution_count,
                        "consecutive_errors": m.consecutive_errors,
                        "is_preloaded": m.is_preloaded,
                        "is_healthy": m.is_healthy(),
                        "idle_seconds": (
                            datetime.utcnow() - m.last_used_at
                        ).total_seconds(),
                    }
                    for sid, m in self._kernels.items()
                }
            }

    # =========================================================================
    # SHUTDOWN
    # =========================================================================

    def shutdown_all(self):
        """Gracefully shut down all kernels. Call on server shutdown."""
        self._running = False
        with self._lock:
            logger.info(f"Shutting down all {len(self._kernels)} kernel(s)")
            for sid, managed in self._kernels.items():
                managed.shutdown()
            self._kernels.clear()
            logger.info("All kernels shut down")


# =============================================================================
# INTEGRATION GUIDE FOR main.py
# =============================================================================
#
# Step 1 — Import at module level (after existing imports):
#
#     from patches.fix_kernel_persistence import PersistentKernelManager
#
#
# Step 2 — Initialize at module level (replace existing kernel setup):
#
#     # Replace 'your_create_kernel_function' with whatever function
#     # currently creates a new sandbox kernel in your codebase.
#     kernel_manager = PersistentKernelManager(
#         create_kernel_fn=your_create_kernel_function
#     )
#
#
# Step 3 — In handle_tool_call() where execute_code is processed (~line 500):
#
#     if tool_name == "execute_code":
#         code = args.get("code")
#         if not code:
#             # From Issue #1 fix (empty args mitigation)
#             return build_empty_args_error_response(msg_id)
#
#         session_id = args.get("session_id", "default")
#
#         try:
#             managed_kernel = kernel_manager.get_or_create(session_id)
#             result = managed_kernel.execute(code)
#             return {
#                 "jsonrpc": "2.0",
#                 "id": msg_id,
#                 "result": {
#                     "content": [{"type": "text", "text": str(result)}]
#                 }
#             }
#         except RuntimeError as e:
#             # Max kernels reached
#             return {
#                 "jsonrpc": "2.0",
#                 "id": msg_id,
#                 "result": {
#                     "isError": True,
#                     "content": [{"type": "text", "text": str(e)}]
#                 }
#             }
#
#
# Step 4 — On server shutdown (in your shutdown/cleanup handler):
#
#     kernel_manager.shutdown_all()
#
#
# Step 5 (Optional) — Expose monitoring endpoint:
#
#     @app.get("/kernels/status")
#     def kernel_status():
#         return kernel_manager.get_status()
#
# =============================================================================
