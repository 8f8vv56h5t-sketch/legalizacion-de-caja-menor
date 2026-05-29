FROM node:20-alpine

WORKDIR /app

COPY app ./app

ENV NODE_ENV=production
ENV PORT=3000
ENV DATA_DIR=/data

EXPOSE 3000

CMD ["node", "app/server.js"]
