FROM node:18-alpine AS frontend-builder
WORKDIR /app
COPY package.json package-lock.json* pnpm-lock.yaml* yarn.lock* ./
RUN if [ -f pnpm-lock.yaml ]; then npm i -g pnpm && pnpm install --frozen-lockfile; \
    elif [ -f yarn.lock ]; then yarn install --frozen-lockfile; \
    else npm ci; fi
COPY . .
RUN npm run build

FROM node:18-alpine AS backend-builder
WORKDIR /srv/backend
COPY backend/package.json backend/package-lock.json* backend/pnpm-lock.yaml* backend/yarn.lock* ./
RUN if [ -f pnpm-lock.yaml ]; then npm i -g pnpm && pnpm install --frozen-lockfile; \
    elif [ -f yarn.lock ]; then yarn install --frozen-lockfile; \
    else npm ci; fi
COPY backend/ ./

FROM node:18-alpine
RUN apk add --no-cache nginx supervisor
RUN mkdir -p /var/log/supervisor /var/www/html /etc/nginx/conf.d
COPY --from=frontend-builder /app/dist /var/www/html
WORKDIR /srv/backend
COPY --from=backend-builder /srv/backend /srv/backend
ENV NODE_ENV=production PORT=3000 APP_PORT=8080
COPY nginx.conf /etc/nginx/nginx.conf
COPY supervisord.conf /etc/supervisord.conf
EXPOSE 8080
RUN ln -sf /dev/stdout /var/log/nginx/access.log && ln -sf /dev/stderr /var/log/nginx/error.log
CMD ["/usr/bin/supervisord", "-c", "/etc/supervisord.conf"]