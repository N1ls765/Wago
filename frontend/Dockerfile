# WAGO EPD Explorer als fertiges Image
#   Stufe 1: mit Node bauen (npm run build -> dist/)
#   Stufe 2: dist/ mit nginx ausliefern, /aas/ an den AAS-Server weiterreichen

FROM node:22-alpine AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --no-audit --no-fund
COPY . .
RUN npm run build

FROM nginx:1.27-alpine
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/dist /usr/share/nginx/html
EXPOSE 80
