FROM node:24-alpine

EXPOSE 80
EXPOSE 443
WORKDIR /var/web/
ENTRYPOINT ["npx", "--offline", "serve", "-p", "80", "./public"]

RUN echo {}>/var/web/package.json
RUN npm install serve@14
