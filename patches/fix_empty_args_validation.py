"""
Patch: Improved execute_code Empty Arguments Validation
=======================================================

File to modify: main.py (around line 507)
Branch: fix/empty-args-validation-and-retry

This patch improves the error response when execute_code is called
with empty arguments (args={}). The current implementation returns
a plain error string that agents often fail to recover from.

The improved version returns an MCP-compliant structured error with
explicit retry instructions, significantly improving agent recovery rate.

Bug Reference:
- 5 occurrences observed (March 3-4, 2026)
- Always 122-byte payload (JSON envelope only)
- Triggered by follow-up calls after large tool responses
- Affects Claude; not observed on GPT-5.2
"""


# =============================================================================
# CURRENT CODE (main.py ~line 500-515)
# =============================================================================
#
# Replace the existing empty args validation block with the improved version.
#
# BEFORE:
# -------
#
#     if tool_name == "execute_code":
#         code = args.get("code")
#         if not code:
#             logger.warning(
#                 f"MCP direct: execute_code argument validation failed: "
#                 f"Missing required parameter(s) for 'execute_code': code. "
#                 f"Please provide: code"
#             )
#             # Return error to agent
#             response = {
#                 "jsonrpc": "2.0",
#                 "id": msg_id,
#                 "result": {
#                     "content": [
#                         {
#                             "type": "text",
#                             "text": (
#                                 "Missing required parameter(s) for "
#                                 "'execute_code': code. Please provide: code"
#                             )
#                         }
#                     ]
#                 }
#             }
#             return response
#
#
# =============================================================================
# IMPROVED CODE — REPLACE THE ABOVE WITH THIS:
# =============================================================================

def build_empty_args_error_response(msg_id: int, tool_name: str = "execute_code") -> dict:
    """
    Build an MCP-compliant structured error response for when execute_code
    is called with empty arguments.

    This is a known issue where the agent's code parameter is dropped during
    serialization, typically after large tool responses that consume most of
    the context window.

    The response is designed to:
    1. Clearly signal this is an error (isError: True)
    2. Explain what happened in terms the agent can understand
    3. Provide an explicit example of the correct format
    4. Suggest the agent retry with a simpler/shorter code block

    Args:
        msg_id: The JSON-RPC message ID from the original request
        tool_name: The tool that was called (default: execute_code)

    Returns:
        MCP-compliant JSON-RPC error response dict
    """
    error_message = (
        f"ERROR: {tool_name} was called with empty arguments — "
        f"the required 'code' parameter was not provided. "
        f"This is a known serialization issue that can occur after "
        f"large tool responses. "
        f"\n\n"
        f"TO FIX: Please retry this tool call and ensure the 'code' "
        f"parameter contains your Python code as a string value. "
        f"\n\n"
        f"CORRECT FORMAT EXAMPLE:\n"
        f'{{"code": "print(\'hello world\')"}}'
        f"\n\n"
        f"TIP: If your code is very long, try breaking it into smaller "
        f"steps across multiple execute_code calls."
    )

    return {
        "jsonrpc": "2.0",
        "id": msg_id,
        "result": {
            "isError": True,
            "content": [
                {
                    "type": "text",
                    "text": error_message
                }
            ]
        }
    }


# =============================================================================
# INTEGRATION POINT — In main.py handle_tool_call() around line 500:
# =============================================================================
#
#     if tool_name == "execute_code":
#         code = args.get("code")
#         if not code:
#             logger.warning(
#                 f"MCP direct: {tool_name} argument validation failed: "
#                 f"Empty arguments received (payload ~122 bytes). "
#                 f"Likely context window truncation after large response. "
#                 f"Returning structured retry hint to agent."
#             )
#             return build_empty_args_error_response(msg_id, tool_name)
#
#         # ... rest of execute_code handling ...
#


# =============================================================================
# OPTIONAL: Add metrics tracking for this bug pattern
# =============================================================================

class EmptyArgsTracker:
    """
    Track empty args occurrences for monitoring and alerting.
    Useful for measuring whether the bug frequency changes over time
    or correlates with specific models/payload sizes.
    """

    def __init__(self):
        self.occurrences = []

    def record(self, timestamp: str, tool_name: str, msg_id: int, 
               source_ip: str = None, model: str = None):
        self.occurrences.append({
            "timestamp": timestamp,
            "tool_name": tool_name,
            "msg_id": msg_id,
            "source_ip": source_ip,
            "model": model,
            "payload_bytes": 122,  # Always 122 bytes observed
        })

    def get_count(self) -> int:
        return len(self.occurrences)

    def get_frequency(self, window_minutes: int = 60) -> float:
        """Get occurrences per hour within the specified window."""
        if not self.occurrences:
            return 0.0
        from datetime import datetime, timedelta
        now = datetime.utcnow()
        cutoff = now - timedelta(minutes=window_minutes)
        recent = [o for o in self.occurrences 
                  if datetime.fromisoformat(o["timestamp"]) > cutoff]
        return len(recent) * (60 / window_minutes)


# Global tracker instance
empty_args_tracker = EmptyArgsTracker()
