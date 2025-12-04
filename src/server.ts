import express from "express";
import cors from "cors";
import * as path from "path";
import relatoriosRouter from "./routes/relatorios";
import relatorioBoletosRoutes from "./routes/relatoriosBoletosEmpresa";
import relatoriosGerais from "./routes/relatoriosGerais";
import votacao from "./routes/votacao";
import relatoriosDespesas from "./routes/relatoriosDespesas";
const app = express();
app.use(cors());
app.use(express.json());

// Usa as rotas
app.use("/uploads", express.static(path.join(__dirname, "../uploads")));
app.use("/relatorios", relatoriosRouter);
app.use("/relatorios", relatorioBoletosRoutes);
app.use("/relatorios", relatoriosGerais);
app.use("/relatorios", relatoriosDespesas)
app.use("/relatorios", votacao); 
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));

