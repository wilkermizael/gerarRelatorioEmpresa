import { Router, Request, Response } from "express";
import ExcelJS from "exceljs";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import fs from "fs";
import path from "path";
import axios from "axios";
const router = Router();

router.get("/mensagem", (req: Request, res: Response) => {
  res.status(200).json({
    sucesso: true,
    mensagem: "✅ Backend de relatórios funcionando corretamente!",
  });
});

router.post("/empresas", async (req: Request, res: Response) => {
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

    const colunasSelecionadas = Object.keys(relatorioEmpresa).filter(
      (coluna) => relatorioEmpresa[coluna] === true
    );

    if (colunasSelecionadas.length === 0) {
      return res.status(400).json({ error: "Nenhuma coluna selecionada." });
    }

    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fileName = `relatorio_empresas_${Date.now()}.${formato}`;
    const filePath = path.join(outputDir, fileName);

    // 🟩 Geração XLSX
    if (formato === "xlsx") {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Empresas");

      sheet.addRow(colunasSelecionadas);
      empresas.forEach((empresa: any) => {
        const linha = colunasSelecionadas.map((col) => empresa[col] ?? "");
        sheet.addRow(linha);
      });

      await workbook.xlsx.writeFile(filePath);
    }

else {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595, 842]); // A4
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const { width, height } = page.getSize();

  // === Cabeçalho ===
  const logoUrl =
    "https://jsimrqytfiwiayxbdiro.supabase.co/storage/v1/object/public/senalbabucket/Logo/logoSenalba.jpeg";

  try {
    const logoResponse = await axios.get(logoUrl, { responseType: "arraybuffer" });
    const logoBytes = logoResponse.data;
    const logoImage = await pdfDoc.embedJpg(logoBytes);

    const logoWidth = 70;
    const logoHeight = 70 * (logoImage.height / logoImage.width);
    page.drawImage(logoImage, {
      x: 50,
      y: height - logoHeight - 40,
      width: logoWidth,
      height: logoHeight,
    });
  } catch {
    console.warn("⚠️ Logo não carregada — continuando sem imagem.");
  }

  page.drawText("Senalba MG", {
    x: 140,
    y: height - 60,
    size: 16,
    font: boldFont,
    color: rgb(0, 0.2, 0.6),
  });

  page.drawText("Relatório de Empresas", {
    x: 140,
    y: height - 80,
    size: 12,
    font,
    color: rgb(0, 0, 0),
  });

  let y = height - 120;
  const marginX = 50;
  const colWidth = (width - marginX * 2) / colunasSelecionadas.length;
  const rowHeight = 20;

  const drawCell = (x: number, y: number, text: string, bold = false) => {
    const f = bold ? boldFont : font;
    page.drawText(text, {
      x: x + 2,
      y: y - 14,
      size: 9,
      font: f,
    });
  };

  // === Cabeçalho da tabela ===
  colunasSelecionadas.forEach((col, i) => {
    const x = marginX + i * colWidth;
    page.drawRectangle({
      x,
      y: y - rowHeight,
      width: colWidth,
      height: rowHeight,
      borderColor: rgb(0.7, 0.7, 0.7),
      borderWidth: 0.8,
    });
    drawCell(x, y - 4, col.toUpperCase(), true);
  });
  y -= rowHeight;

  // === Linhas de dados ===
  empresas.forEach((empresa: any, index: number) => {
    if (y < 70) {
      // quebra de página
      const newPage = pdfDoc.addPage([595, 842]);
      page.drawText(`Página ${index + 1}`, {
        x: width - 100,
        y: 40,
        size: 8,
        font,
      });
      y = height - 100;
    }

    colunasSelecionadas.forEach((col, i) => {
      const x = marginX + i * colWidth;
      page.drawRectangle({
        x,
        y: y - rowHeight,
        width: colWidth,
        height: rowHeight,
        borderColor: rgb(0.85, 0.85, 0.85),
        borderWidth: 0.6,
      });
      const texto = String(empresa[col] ?? "");
      drawCell(x, y - 2, texto);
    });

    y -= rowHeight;
  });

  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(filePath, pdfBytes);
}



    return res.download(filePath, fileName, (err) => {
      if (err) console.error("Erro ao enviar o arquivo:", err);
      fs.unlinkSync(filePath);
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});

export default router;
