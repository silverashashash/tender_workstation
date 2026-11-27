FROM node:22-alpine AS builder
WORKDIR /app

COPY package.json ./
COPY server/package.json ./server/
COPY client/package.json ./client/
RUN npm install && npm --prefix server install && npm --prefix client install

COPY server ./server
COPY client ./client
RUN npm run build

FROM node:22-alpine
WORKDIR /app
ENV NODE_ENV=production
ENV PORT=3001
ENV DATA_DIR=/data

COPY --from=builder /app/server/package.json ./server/package.json
COPY --from=builder /app/server/node_modules ./server/node_modules
COPY --from=builder /app/server/dist ./server/dist
COPY --from=builder /app/client/dist ./client/dist

RUN mkdir -p /data/uploads
VOLUME ["/data"]
EXPOSE 3001

CMD ["node", "server/dist/index.js"]
