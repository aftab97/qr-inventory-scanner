# ---- Frontend builder (Vite React) ----
FROM node:18-alpine AS frontend-builder
WORKDIR /app

COPY package.json package-lock.json* pnpm-lock.yaml* yarn.lock* ./
RUN set -eux; \
  if [ -f pnpm-lock.yaml ]; then \
    npm i -g pnpm@8; \
    pnpm install --frozen-lockfile; \
  elif [ -f yarn.lock ]; then \
    yarn install --frozen-lockfile; \
  elif [ -f package-lock.json ]; then \
    npm ci; \
  else \
    npm install; \
  fi

COPY . .
RUN npx vite build --base=./

# ---- Backend builder ----
FROM node:18-alpine AS backend-builder
WORKDIR /srv/backend

COPY backend/package.json backend/package-lock.json* backend/pnpm-lock.yaml* backend/yarn.lock* ./
RUN set -eux; \
  if [ -f pnpm-lock.yaml ]; then \
    npm i -g pnpm@8; \
    pnpm install --prod --frozen-lockfile; \
  elif [ -f yarn.lock ]; then \
    yarn install --production --frozen-lockfile; \
  elif [ -f package-lock.json ]; then \
    npm ci --omit=dev; \
  else \
    npm install --omit=dev; \
  fi

COPY backend/ ./

# ---- Final runtime (nginx + node via supervisord) ----
FROM node:18-alpine

RUN apk add --no-cache nginx supervisor
RUN mkdir -p /var/log/supervisor /var/www/html /etc/nginx/conf.d

COPY --from=frontend-builder /app/dist /var/www/html

WORKDIR /srv/backend
COPY --from=backend-builder /srv/backend /srv/backend

ENV NODE_ENV=production \
    PORT=3000 \
    APP_PORT=8080

COPY nginx.conf /etc/nginx/nginx.conf
COPY supervisord.conf /etc/supervisord.conf

EXPOSE 8080

RUN ln -sf /dev/stdout /var/log/nginx/access.log && \
    ln -sf /dev/stderr /var/log/nginx/error.log

CMD ["/usr/bin/supervisord", "-c", "/etc/supervisord.conf"]