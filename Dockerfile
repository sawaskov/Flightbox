# Build from repo root: docker build -t flightbox .
FROM node:20-alpine
WORKDIR /app
COPY backend/package.json ./
RUN npm install --omit=dev
COPY backend/ ./
ENV NODE_ENV=production
ENV LISTEN_HOST=0.0.0.0
EXPOSE 3000
CMD ["node", "server.js"]
