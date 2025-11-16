import express from "express";
import cors from "cors";
import * as path from "path";
import relatoriosRouter from "./routes/relatorios";
import relatorioBoletosRoutes from "./routes/relatoriosBoletosEmpresa";
import relatoriosGerais from "./routes/relatoriosGerais";
const app = express();
app.use(cors());
app.use(express.json());

// Usa as rotas
app.use("/uploads", express.static(path.join(__dirname, "../uploads")));
app.use("/relatorios", relatoriosRouter);
app.use("/relatorios", relatorioBoletosRoutes); // ← adiciona aqui
app.use("/relatorios", relatoriosGerais); // ← adiciona aqui
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));

