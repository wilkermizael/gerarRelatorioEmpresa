import express from "express";
import cors from "cors";
import * as path from "path";

// ✅ Ajuste nas importações com fallback para CommonJS / ES Modules
const relatoriosRouter = require("./routes/relatorios").default || require("./routes/relatorios");
const relatorioBoletosRoutes = require("./routes/relatoriosBoletosEmpresa").default || require("./routes/relatoriosBoletosEmpresa");
const relatoriosGerais = require("./routes/relatoriosGerais").default || require("./routes/relatoriosGerais");
const votacao = require("./routes/votacao").default || require("./routes/votacao");
const relatoriosDespesas = require("./routes/relatoriosDespesas").default || require("./routes/relatoriosDespesas");

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true, parameterLimit: 50000 }));

// Usa as rotas
app.use("/uploads", express.static(path.join(__dirname, "../uploads")));
app.use("/relatorios", relatoriosRouter);
app.use("/relatorios", relatorioBoletosRoutes);
app.use("/relatorios", relatoriosGerais);
app.use("/relatorios", relatoriosDespesas);
app.use("/relatorios", votacao); 

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));