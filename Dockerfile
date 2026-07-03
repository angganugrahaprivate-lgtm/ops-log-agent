FROM node:20-slim

WORKDIR /app

# Copy manifest dulu biar layer cache npm install optimal
COPY package*.json ./

RUN npm install --omit=dev

# Copy sisa source code
COPY . .

ENV NODE_ENV=production

EXPOSE 3000

CMD ["node", "server.js"]
