FROM node:20-alpine

WORKDIR /app

RUN apk add --no-cache zip unzip

COPY package*.json ./
RUN npm install --omit=dev

COPY app ./app

ENV NODE_ENV=production
ENV PORT=3000
ENV DATA_DIR=/data

EXPOSE 3000

CMD ["node", "app/server.js"]
