# syntax=docker/dockerfile:1

FROM node:24-bookworm-slim AS build
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY nest-cli.json tsconfig.json tsconfig.build.json ./
COPY src ./src
RUN npm run build && npm prune --omit=dev

FROM node:24-bookworm-slim AS runtime
ENV NODE_ENV=production
ENV PORT=3000
WORKDIR /app

RUN groupadd --system feira-nova \
    && useradd --system --gid feira-nova --create-home feira-nova

COPY --chown=feira-nova:feira-nova package.json package-lock.json ./
COPY --from=build --chown=feira-nova:feira-nova /app/node_modules ./node_modules
COPY --from=build --chown=feira-nova:feira-nova /app/dist ./dist
COPY --chown=feira-nova:feira-nova data ./data
COPY --chown=feira-nova:feira-nova template ./template
COPY --chown=feira-nova:feira-nova scripts ./scripts
COPY --chown=feira-nova:feira-nova src/index.js ./src/index.js

USER feira-nova
EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:' + process.env.PORT + '/health').then(r => { if (!r.ok) process.exit(1) }).catch(() => process.exit(1))"

CMD ["node", "dist/main.js"]
