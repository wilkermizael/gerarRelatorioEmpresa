import { Router, Request, Response } from "express";
import axios from "axios";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import dotenv from "dotenv";
dotenv.config();
const router = Router();

// Função utilitária para converter YYYY-MM-DD → DD/MM/YYYY
function formatarDataBrasileira(data: string | null) {
  if (!data) return "";
  const d = new Date(data);
  if (isNaN(d.getTime())) return data;
  return d.toLocaleDateString("pt-BR");
}

// Função para converter datetime
function formatarDataHoraBrasileira(data: string | null) {
  if (!data) return "";
  const d = new Date(data);
  if (isNaN(d.getTime())) return data;
  return d.toLocaleString("pt-BR", { hour12: false });
}

router.post("/boletos/geral", async (req: Request, res: Response) => {
  try {
    const { dataInicial, dataFinal, application } = req.body;

    if (!dataInicial || !dataFinal) {
      return res.status(400).json({
        error: "dataInicial e dataFinal são obrigatórios."
      });
    }

    // ============================
    // 1. MONTAR A URL SAFE2PAY
    // ============================
    let url = `https://api.safe2pay.com.br/v2/transaction/list`;
    url += `?PageNumber=1`;
    url += `&RowsPerPage=1000`;
    url += `&CreatedDateInitial=${dataInicial}`;
    url += `&CreatedDateEnd=${dataFinal}`;

    if (application && application.trim() !== "") {
      url += `&Object.Application=${encodeURIComponent(application)}`;
    }

    // ============================
    // 2. CONSULTAR SAFE2PAY
    // ============================
    const resposta = await axios.get(url, {
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        "X-API-KEY": process.env.SAFE2PAY_KEY as string,
      },
    });

    const dados = resposta.data?.ResponseDetail?.Objects ?? [];

    if (!Array.isArray(dados) || dados.length === 0) {
      return res.status(200).json({ aviso: "Nenhum boleto encontrado." });
    }

    // ============================
    // 3. CONFIGURAR PLANILHA
    // ============================
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Boletos");

    const header = [
      "Empresa",
      "CNPJ",
      "Email",
      "Telefone",
      "Status",              // <-- (Mensagem virou Status)
      "Tipo",
      "Data Criação",
      "Data Pagamento",
      "Valor",               // <-- Valor (Split) virou Valor
      "Taxa",
      "Vencimento"
    ];

    const headerRow = sheet.addRow(header);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };

    // Ativar filtros automáticos
    sheet.autoFilter = {
      from: 'A1',
      to: 'K1'
    };

    // ============================
    // 4. SOMATÓRIOS
    // ============================
    let totalRecebido = 0;
    let totalAReceber = 0;

    // ============================
    // 5. ADICIONAR LINHAS
    // ============================
    dados.forEach((t: any) => {
      const empresa = t.Customer ?? {};
      const boleto = t.PaymentObject ?? {};
      const split = t.Splits?.[0] ?? null;

      const valor = Number(split?.Amount ?? 0);

      // Status simplificado
      let statusFinanceiro = "";
      if (t.Message === "Liberado") statusFinanceiro = "A Receber";
      else if (t.Message === "Autorizado") statusFinanceiro = "Pago";
      else statusFinanceiro = t.Message;

      if (statusFinanceiro === "Pago") totalRecebido += valor;
      if (statusFinanceiro === "A Receber") totalAReceber += valor;

      const row = sheet.addRow([
        empresa.Name ?? "",
        empresa.Identity ?? "",
        empresa.Email ?? "",
        empresa.Phone ?? "",
        statusFinanceiro,               // <-- Coluna Status
        t.Application ?? "",
        formatarDataBrasileira(t.CreatedDate),
        formatarDataHoraBrasileira(t.CreatedDateTime),
        valor,                          // <-- Valor (número real)
        t.TaxValue ?? "",
        formatarDataBrasileira(boleto.DueDate)
      ]);

      // 🎯 FORMATAÇÃO MONETÁRIA
      row.getCell(9).numFmt = "R$ #,##0.00";   // Valor
      row.getCell(10).numFmt = "R$ #,##0.00";  // Taxa
    });

    // ============================
    // 6. RESUMO
    // ============================
    const totalRecebidoRow = sheet.addRow(["TOTAL RECEBIDO (Pago)", totalRecebido]);
    totalRecebidoRow.getCell(2).numFmt = "R$ #,##0.00";

    const totalAReceberRow = sheet.addRow(["TOTAL A RECEBER (Liberado)", totalAReceber]);
    totalAReceberRow.getCell(2).numFmt = "R$ #,##0.00";

    sheet.columns.forEach((col) => (col.width = 22));

    // ============================
    // 7. SALVAR ARQUIVO
    // ============================
    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fileName = `relatorio_boletos_${Date.now()}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    await workbook.xlsx.writeFile(filePath);

    return res.download(filePath, fileName, () => {
      fs.unlinkSync(filePath);
    });

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});

export default router;
