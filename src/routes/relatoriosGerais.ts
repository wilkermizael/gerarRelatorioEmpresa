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

// Converte qualquer entrada em número válido independente da estrutura
function extrairValor(valorPrincipal: any, valorFallback: any = 0): number {
  let val = valorPrincipal !== undefined && valorPrincipal !== null && valorPrincipal !== "" ? valorPrincipal : valorFallback;
  
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  if (typeof val === "string") {
    const limpo = val.replace(/\./g, "").replace(",", ".");
    const num = parseFloat(limpo);
    return isNaN(num) ? 0 : num;
  }
  return 0;
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
      "Status",
      "Tipo",
      "Data Criação",
      "Data Pagamento",
      "Valor Bruto",        // Coluna 9
      "Valor Líquido",      // Coluna 10
      "Taxa",               // Coluna 11
      "Vencimento"          // Coluna 12
    ];

    const headerRow = sheet.addRow(header);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };

    sheet.autoFilter = {
      from: 'A1',
      to: 'L1'
    };

    // ============================
    // 4. SOMATÓRIOS
    // ============================
    let totalRecebidoLiquido = 0;
    let totalAReceberBruto = 0;

    // ============================
    // 5. ADICIONAR LINHAS
    // ============================
    dados.forEach((t: any) => {
      const empresa = t.Customer ?? {};
      const boleto = t.PaymentObject ?? {};

      // Extração direta sem condicional entre Amount e TaxValue
      const valorBruto = extrairValor(t.Amount, boleto.Amount);
      const taxa = extrairValor(t.TaxValue, boleto.TaxValue);
      
      // Se NetValue não vier explicitamente preenchido na transação Negocial, calcula Bruto - Taxa
      let valorLiquido = extrairValor(t.NetValue, boleto.NetValue);
      if (valorLiquido === 0 && valorBruto > 0) {
        valorLiquido = valorBruto - taxa;
      }

      const tipoAplicacao = t.Application ?? t.PaymentMethod?.Name ?? "N/A";

      // Status Financeiro
      let statusFinanceiro = "";
      if (t.Message === "Processamento") {
        statusFinanceiro = "A Receber";
      } else if (t.Message === "Liberado") {
        statusFinanceiro = "Liberado";
      } else if (t.Message === "Autorizado") {
        statusFinanceiro = "Pago";
      } else {
        statusFinanceiro = t.Message ?? "N/A";
      }

      // Somatórios
      if (statusFinanceiro === "Pago") totalRecebidoLiquido += valorLiquido;
      if (statusFinanceiro === "A Receber") totalAReceberBruto += valorBruto;

      const row = sheet.addRow([
        empresa.Name ?? "",
        empresa.Identity ?? "",
        empresa.Email ?? "",
        empresa.Phone ?? "",
        statusFinanceiro,
        tipoAplicacao,
        formatarDataBrasileira(t.CreatedDate),
        formatarDataHoraBrasileira(t.CreatedDateTime),
        valorBruto,                          // Coluna 9
        valorLiquido,                        // Coluna 10
        taxa,                                // Coluna 11
        formatarDataBrasileira(boleto.DueDate)
      ]);

      // Formatação Monetária nas células
      row.getCell(9).numFmt = "R$ #,##0.00";
      row.getCell(10).numFmt = "R$ #,##0.00";
      row.getCell(11).numFmt = "R$ #,##0.00";
    });

    // ============================
    // 6. RESUMO
    // ============================
    const totalRecebidoRow = sheet.addRow(["TOTAL RECEBIDO (Líquido / Pago)", totalRecebidoLiquido]);
    totalRecebidoRow.getCell(2).numFmt = "R$ #,##0.00";

    const totalAReceberRow = sheet.addRow(["TOTAL A RECEBER (Bruto / Processamento)", totalAReceberBruto]);
    totalAReceberRow.getCell(2).numFmt = "R$ #,##0.00";

    sheet.columns.forEach((col) => (col.width = 22));

    // ============================
    // 7. SALVAR ARQUIVO
    // ============================
    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

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