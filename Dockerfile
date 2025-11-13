# =========================
# 🏗️  Etapa 1: Build
# =========================
FROM node:22 AS builder

WORKDIR /app

# Copia os arquivos essenciais primeiro (melhor para cache)
COPY package*.json ./
COPY tsconfig*.json ./

# Instala dependências
RUN npm install

# Copia o código-fonte
COPY . .

# Compila o TypeScript
RUN npm run build

# =========================
# 🚀 Etapa 2: Produção
# =========================
FROM node:22-alpine

WORKDIR /app

# Copia apenas os arquivos necessários da build anterior
COPY --from=builder /app/dist ./dist
COPY package*.json ./

# Instala apenas as dependências de produção
RUN npm install --omit=dev

# Define variáveis padrão
ENV NODE_ENV=production
ENV PORT=3000

# Expõe a porta da aplicação
EXPOSE 3000

# Comando de inicialização
CMD ["node", "dist/server.js"]
