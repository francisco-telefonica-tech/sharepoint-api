FROM node:10-alpine
RUN mkdir -p /data/node_modules && chown -R node:node /data
WORKDIR /data
COPY package*.json ./
USER node
RUN npm install
COPY --chown=node:node . .
EXPOSE 3000
CMD [ "node", "app.js" ]