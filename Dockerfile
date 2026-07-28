# =========================
# 🏗️  Etapa 1: Build
# =========================
FROM node:22-alpine AS builder

WORKDIR /app

# Copia os arquivos essenciais primeiro para aproveitar o cache de camadas
COPY package*.json ./
COPY tsconfig*.json ./

# Instala todas as dependências (incluindo devDependencies para compilar o TS)
RUN npm ci

# Copia o código-fonte
COPY . .

# Compila o TypeScript
RUN npm run build

# =========================
# 🚀 Etapa 2: Produção
# =========================
FROM node:22-alpine

WORKDIR /app

# Copia os manifestos para instalar apenas prod-deps
COPY package*.json ./

# Instala somente dependências de produção de forma rápida
RUN npm ci --omit=dev && npm cache clean --force

# Copia apenas o código compilado da etapa de build
COPY --from=builder /app/dist ./dist

# Define variáveis de ambiente
ENV NODE_ENV=production
ENV PORT=3000

# Expõe a porta do servidor
EXPOSE 3000

# Comando de inicialização
CMD ["node", "dist/server.js"]