# ═══════════════════════════════════════════════════════
# Power Automate MCP Server — Railway-Optimized Dockerfile
# ═══════════════════════════════════════════════════════

# Stage 1: Build
FROM node:18-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci --ignore-scripts

COPY tsconfig.json ./
COPY src/ ./src/

RUN npm run build

# Stage 2: Production
FROM node:18-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev --ignore-scripts

COPY --from=builder /app/dist ./dist

EXPOSE ${PORT:-8080}

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://localhost:${PORT:-8080}/health || exit 1

CMD ["node", "dist/index.js"]
