# Optional: run anywhere that supports Docker (Railway, Fly.io, a VM...).
FROM node:22-alpine
WORKDIR /app
COPY package.json server.js index.html ./
ENV NODE_ENV=production PORT=3000 DATA_FILE=/data/data.json
RUN mkdir -p /data && chown node:node /data
USER node
VOLUME ["/data"]
EXPOSE 3000
HEALTHCHECK CMD wget -qO- http://127.0.0.1:3000/health || exit 1
CMD ["node", "server.js"]
