# === Stage 1: Build ===
FROM node:20-slim AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .

ENV DOCKER_BUILD=true
RUN npm run build

# === Stage 2: Production ===
FROM node:20-slim AS runner

WORKDIR /app

# Install Python + MarkItDown + hwp5html
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip \
    && rm -rf /var/lib/apt/lists/*

# Install all Python tools via pip3
RUN pip3 install --break-system-packages \
    markitdown "markitdown[pdf]" "markitdown[docx,pptx,xlsx]" \
    pyhwp six olefile lxml
ENV HOME="/root"
ENV NODE_ENV=production
ENV NEXT_TELEMETRY_DISABLED=1

# Copy built app
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static
COPY --from=builder /app/public ./public

EXPOSE 3100

ENV PORT=3100
ENV HOSTNAME="0.0.0.0"

CMD ["node", "server.js"]
