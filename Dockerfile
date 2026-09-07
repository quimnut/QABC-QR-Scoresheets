# ─── Stage 1: Build Astro static site ────────────────────────────────────────
FROM node:24-slim AS builder

WORKDIR /build

# Copy package files first for layer caching
COPY frontend/package.json frontend/package-lock.json* ./

RUN npm ci

# Copy source
COPY frontend/ .

# Build static output to /build/dist
RUN npm run build

# ─── Stage 2: Serve with nginx ───────────────────────────────────────────────
FROM nginx:1.27-alpine AS runtime

# Remove the default nginx config
RUN rm /etc/nginx/conf.d/default.conf

# Copy our nginx config
COPY nginx.conf /etc/nginx/conf.d/app.conf

# Copy built static files
COPY --from=builder /build/dist /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
