import express from "express";
import cors from "cors";
import relatoriosRouter from "./routes/relatorios";

const app = express();
app.use(cors());
app.use(express.json());

// Usa as rotas
app.use("/relatorios", relatoriosRouter);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
