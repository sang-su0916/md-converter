# === Stage 1: Build ===
FROM node:20-slim AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .
RUN npm run build

# === Stage 2: Production ===
FROM node:20-slim AS runner

WORKDIR /app

# Install Python + MarkItDown + hwp5html
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip python3-venv pipx \
    && rm -rf /var/lib/apt/lists/*

# Install MarkItDown with all format support
RUN pipx install markitdown && \
    pipx inject markitdown "markitdown[pdf]" --force && \
    pipx inject markitdown "markitdown[docx,pptx,xlsx]" --force

# Install hwp5html for HWP support
RUN pipx install pyhwp

# Add pipx bin to PATH
ENV PATH="/root/.local/bin:${PATH}"
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
