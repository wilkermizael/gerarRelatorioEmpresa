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

router.post("/boletos", async (req: Request, res: Response) => {
  try {
    const { dataInicial, dataFinal, identity, application } = req.body;

    if (!dataInicial || !dataFinal || !identity) {
      return res.status(400).json({
        error: "dataInicial, dataFinal e identity são obrigatórios."
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
    url += `&Object.Customer.Identity=${identity}`;

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

    // Dados da Empresa (extraídos da primeira ocorrência do retorno)
    const primeiraEmpresa = dados[0]?.Customer ?? {};
    const nomeEmpresa = primeiraEmpresa.Name || "EMPRESA NÃO IDENTIFICADA";
    const cnpjEmpresa = primeiraEmpresa.Identity || identity;

    // ============================
    // 3. CONFIGURAR PLANILHA
    // ============================
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Boletos");

    // --- CABEÇALHO INSTITUCIONAL ---
    // Linha 1: Título do Sindicato
    sheet.mergeCells("A1:L1");
    const titleCell = sheet.getCell("A1");
    titleCell.value = "SENALBA MG - SINDICATO DOS EMPREGADOS";
    titleCell.font = { name: "Arial", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4B39EF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 30;

    // Linha 2: Razão Social e CNPJ
    sheet.mergeCells("A2:L2");
    const subTitleCell = sheet.getCell("A2");
    subTitleCell.value = `RELATÓRIO DE CONTRIBUIÇÕES: ${nomeEmpresa.toUpperCase()} | CNPJ: ${cnpjEmpresa}`;
    subTitleCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF1E293B" } };
    subTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
    subTitleCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(2).height = 22;

    // Linha 3: Informações de Período e Geração
    sheet.mergeCells("A3:L3");
    const infoCell = sheet.getCell("A3");
    infoCell.value = `Período: ${formatarDataBrasileira(dataInicial)} até ${formatarDataBrasileira(dataFinal)} ${
      application ? `| Filtro: ${application}` : "| Todos os Acordos"
    } | Gerado em: ${new Date().toLocaleString("pt-BR")}`;
    infoCell.font = { name: "Arial", size: 9, italic: true, color: { argb: "FF64748B" } };
    infoCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(3).height = 18;

    // Linha 4: Espaço em branco
    sheet.addRow([]);

    // --- COLUNAS DA TABELA (Linha 5) ---
    const header = [
      "Empresa",
      "CNPJ",
      "Email",
      "Telefone",
      "Status",
      "Tipo de Contribuição",
      "Data Criação",
      "Data Pagamento",
      "Valor Bruto",
      "Taxa",
      "Valor Líquido",
      "Vencimento"
    ];

    const headerRow = sheet.addRow(header);
    headerRow.height = 24;
    headerRow.eachCell((cell) => {
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Filtros automáticos na linha 5
    sheet.autoFilter = {
      from: "A5",
      to: "L5"
    };

    // ============================
    // 4. SOMATÓRIOS
    // ============================
    let totalBruto = 0;
    let totalTaxas = 0;
    let totalLiquido = 0;

    // ============================
    // 5. ADICIONAR LINHAS DE DADOS
    // ============================
    // ============================
    // 5. ADICIONAR LINHAS DE DADOS
    // ============================
    dados.forEach((t: any) => {
      const empresa = t.Customer ?? {};
      const boleto = t.PaymentObject ?? {};

      // Pega o valor bruto direto do campo Amount (com fallback para boleto.Amount ou 0)
      const valorBruto = Number(t.Amount ?? boleto.Amount ?? 0);
      const taxa = Number(t.TaxValue ?? 0);
      const valorLiquido = valorBruto - taxa; // Valor Líquido = Bruto - Taxa

      // Status financeiro
      let statusFinanceiro = "";
      if (t.Message === "Liberado") statusFinanceiro = "A Receber";
      else if (t.Message === "Autorizado") statusFinanceiro = "Pago";
      else statusFinanceiro = t.Message || "-";

      totalBruto += valorBruto;
      totalTaxas += taxa;
      totalLiquido += valorLiquido;

      const row = sheet.addRow([
        empresa.Name ?? "",
        empresa.Identity ?? "",
        empresa.Email ?? "",
        empresa.Phone ?? "",
        statusFinanceiro,
        t.Application ?? "",
        formatarDataBrasileira(t.CreatedDate),
        formatarDataHoraBrasileira(t.CreatedDateTime),
        valorBruto,     // Coluna I (9): Amount
        taxa,           // Coluna J (10): TaxValue
        valorLiquido,   // Coluna K (11): Valor Líquido
        formatarDataBrasileira(boleto.DueDate)
      ]);

      // Formatação monetária R$
      row.getCell(9).numFmt = "R$ #,##0.00";
      row.getCell(10).numFmt = "R$ #,##0.00";
      row.getCell(11).numFmt = "R$ #,##0.00";
      row.alignment = { vertical: "middle" };
    });

    // ============================
    // 6. LINHAS DE RESUMO / TOTAIS
    // ============================
    sheet.addRow([]); // Espaço

    const totalRow = sheet.addRow([
      "TOTAIS ACUMULADOS", "", "", "", "", "", "", "",
      totalBruto,
      totalTaxas,
      totalLiquido,
      ""
    ]);

    totalRow.height = 22;
    sheet.mergeCells(`A${totalRow.number}:H${totalRow.number}`);

    totalRow.getCell(1).font = { name: "Arial", size: 10, bold: true };
    totalRow.getCell(1).alignment = { horizontal: "right", vertical: "middle" };

    [9, 10, 11].forEach((colIdx) => {
      const cell = totalRow.getCell(colIdx);
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF059669" } };
      cell.numFmt = "R$ #,##0.00";
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECFDF5" } };
      cell.alignment = { vertical: "middle" };
    });

    // Ajuste de largura das colunas
    sheet.columns.forEach((col) => {
      col.width = 24;
    });

    // ============================
    // 7. SALVAR E RETORNAR
    // ============================
    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    const fileName = `relatorio_boletos_${Date.now()}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    await workbook.xlsx.writeFile(filePath);

    return res.download(filePath, fileName, () => {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    });

  } catch (error) {
    console.error("Erro na rota /boletos:", error);
    return res.status(500).json({ error: "Erro ao gerar relatório de boletos." });
  }
});

export default router;