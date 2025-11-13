"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = require("express");
const exceljs_1 = __importDefault(require("exceljs"));
const pdf_lib_1 = require("pdf-lib");
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const router = (0, express_1.Router)();
router.get("/mensagem", (req, res) => {
    res.status(200).json({
        sucesso: true,
        mensagem: "✅ Backend de relatórios funcionando corretamente!",
    });
});
router.post("/empresas", async (req, res) => {
    try {
        const { empresas, relatorioEmpresa, formato } = req.body;
        if (!empresas || !Array.isArray(empresas)) {
            return res.status(400).json({ error: "Empresas inválidas." });
        }
        if (!relatorioEmpresa) {
            return res.status(400).json({ error: "Configuração de relatório ausente." });
        }
        if (!["xlsx", "pdf"].includes(formato)) {
            return res.status(400).json({ error: "Formato inválido. Use 'xlsx' ou 'pdf'." });
        }
        const colunasSelecionadas = Object.keys(relatorioEmpresa).filter((coluna) => relatorioEmpresa[coluna] === true);
        if (colunasSelecionadas.length === 0) {
            return res.status(400).json({ error: "Nenhuma coluna selecionada." });
        }
        const outputDir = path_1.default.join(__dirname, "../../uploads");
        if (!fs_1.default.existsSync(outputDir))
            fs_1.default.mkdirSync(outputDir);
        const fileName = `relatorio_empresas_${Date.now()}.${formato}`;
        const filePath = path_1.default.join(outputDir, fileName);
        if (formato === "xlsx") {
            const workbook = new exceljs_1.default.Workbook();
            const sheet = workbook.addWorksheet("Empresas");
            sheet.addRow(colunasSelecionadas);
            empresas.forEach((empresa) => {
                const linha = colunasSelecionadas.map((col) => empresa[col] ?? "");
                sheet.addRow(linha);
            });
            await workbook.xlsx.writeFile(filePath);
        }
        else {
            const pdfDoc = await pdf_lib_1.PDFDocument.create();
            const page = pdfDoc.addPage([595, 842]);
            const font = await pdfDoc.embedFont(pdf_lib_1.StandardFonts.Helvetica);
            const { height } = page.getSize();
            let y = height - 50;
            page.drawText("Relatório de Empresas", { x: 50, y, size: 16, font });
            y -= 30;
            empresas.forEach((empresa) => {
                colunasSelecionadas.forEach((col) => {
                    page.drawText(`${col}: ${empresa[col] ?? ""}`, { x: 50, y, size: 10, font });
                    y -= 15;
                });
                y -= 10;
                if (y < 50)
                    y = height - 50;
            });
            const pdfBytes = await pdfDoc.save();
            fs_1.default.writeFileSync(filePath, pdfBytes);
        }
        return res.download(filePath, fileName, (err) => {
            if (err)
                console.error("Erro ao enviar o arquivo:", err);
            fs_1.default.unlinkSync(filePath);
        });
    }
    catch (error) {
        console.error(error);
        return res.status(500).json({ error: "Erro ao gerar relatório." });
    }
});
exports.default = router;
