FROM node:22-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
# Public values inlined into the bundle and manifest at build time: never secrets.
ARG VITE_ENTRA_CLIENT_ID
ARG VITE_API_BASE_URL
ARG ADDIN_PROD_URL
ARG VITE_MAX_EMAILS=300
ARG VITE_MAX_BODY_CHARS=6000
RUN npm run build

FROM nginx:1.27-alpine
COPY --from=build /app/dist /usr/share/nginx/html
