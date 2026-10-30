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
    sheet.mergeCells("A1:M1");
    const titleCell = sheet.getCell("A1");
    titleCell.value = "SENALBA MG - SINDICATO DOS EMPREGADOS";
    titleCell.font = { name: "Arial", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4B39EF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 30;

    sheet.mergeCells("A2:M2");
    const subTitleCell = sheet.getCell("A2");
    subTitleCell.value = `RELATÓRIO DE CONTRIBUIÇÕES: ${nomeEmpresa.toUpperCase()} | CNPJ: ${cnpjEmpresa}`;
    subTitleCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF1E293B" } };
    subTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
    subTitleCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(2).height = 22;

    sheet.mergeCells("A3:M3");
    const infoCell = sheet.getCell("A3");
    infoCell.value = `Período: ${formatarDataBrasileira(dataInicial)} até ${formatarDataBrasileira(dataFinal)} ${
      application ? `| Filtro: ${application}` : "| Todos os Acordos"
    } | Gerado em: ${new Date().toLocaleString("pt-BR")}`;
    infoCell.font = { name: "Arial", size: 9, italic: true, color: { argb: "FF64748B" } };
    infoCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(3).height = 18;

    // Linha de espaçamento
    sheet.addRow([]);

    // --- COLUNAS DA TABELA (Linha 5) ---
    const header = [
      "Empresa",
      "CNPJ",
      "Email",
      "Telefone",
      "Status Transação",   // Mapeado com a tabela Safe2Pay
      "Tipo Contribuição",
      "Data Criação",
      "Data Pagamento",
      "Valor Cobrado",       // Amount (Valor nominal do boleto)
      "Valor Pago",          // Preenchido apenas se Pago (Status 3)
      "Taxa",                // TaxValue
      "Valor Líquido",       // (Valor Pago - Taxa) ou 0 se não pago
      "Vencimento"
    ];

    const headerRow = sheet.addRow(header);
    headerRow.height = 24;
    headerRow.eachCell((cell) => {
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    sheet.autoFilter = {
      from: "A5",
      to: "M5"
    };

    // MAPA DE STATUS TRANSACIONAL SAFE2PAY
    const MAPA_STATUS_SAFE2PAY: Record<number | string, string> = {
      1: "Pendente",
      2: "Processamento",
      3: "Autorizado (Pago)",
      5: "Em disputa",
      6: "Estornado",
      7: "Baixado",
      8: "Recusado",
      11: "Liberado (Aguardando Pagamento)",
      12: "Em cancelamento",
      13: "Chargeback",
      14: "Pré-Autorizado",
      15: "Devolução de Contestação"
    };

    // ============================
    // 4. SOMATÓRIOS SEPARADOS
    // ============================
    let totalCobrado = 0;
    let totalEfetivamentePago = 0;
    let totalTaxas = 0;
    let totalLiquido = 0;

    // ============================
    // 5. ADICIONAR LINHAS DE DADOS
    // ============================
    dados.forEach((t: any) => {
      const empresa = t.Customer ?? {};
      const boleto = t.PaymentObject ?? {};

      // Código do status transacional vindo do Safe2Pay
      const statusCode = Number(t.TransactionStatus?.Code ?? t.TransactionStatusId ?? t.TransactionStatus ?? 0);
      const statusDescricao = MAPA_STATUS_SAFE2PAY[statusCode] || t.Message || "Outro";

      // Amount = Valor nominal do boleto emitido
      const valorCobrado = Number(t.Amount ?? boleto.Amount ?? 0);
      const taxa = Number(t.TaxValue ?? 0);

      // REGRA DE NEGÓCIO: Só é considerado "Valor Pago" se o status for 3 (Autorizado)
      const foiPago = statusCode === 3 || t.Message === "Autorizado";
      const valorPago = foiPago ? valorCobrado : 0;
      const valorLiquido = foiPago ? (valorPago - taxa) : 0;

      // Acumula somatórios
      totalCobrado += valorCobrado;
      totalEfetivamentePago += valorPago;
      totalTaxas += foiPago ? taxa : 0;
      totalLiquido += valorLiquido;

      const row = sheet.addRow([
        empresa.Name ?? "",
        empresa.Identity ?? "",
        empresa.Email ?? "",
        empresa.Phone ?? "",
        statusDescricao,
        t.Application ?? "",
        formatarDataBrasileira(t.CreatedDate),
        foiPago ? formatarDataHoraBrasileira(t.PaymentDate || t.CreatedDateTime) : "-",
        valorCobrado,       // Coluna I (9): Amount
        valorPago,          // Coluna J (10): Valor Efetivamente Pago
        taxa,               // Coluna K (11): Taxa
        valorLiquido,       // Coluna L (12): Valor Líquido Recebido
        formatarDataBrasileira(boleto.DueDate)
      ]);

      // Formatação monetária R$ para as colunas numéricas (I, J, K, L)
      [9, 10, 11, 12].forEach((colIdx) => {
        row.getCell(colIdx).numFmt = "R$ #,##0.00";
      });

      // Destaque visual suave para boletos pagos vs pendentes/baixados
      if (foiPago) {
        row.getCell(5).font = { bold: true, color: { argb: "FF059669" } }; // Verde
      } else if (statusCode === 7 || statusCode === 6) {
        row.getCell(5).font = { color: { argb: "FFDC2626" } }; // Vermelho para Baixado/Estornado
      }

      row.alignment = { vertical: "middle" };
    });

    // ============================
    // 6. RESUMO DE TOTAIS NO RODAPÉ
    // ============================
    sheet.addRow([]); // Espaçamento

    const totalRow = sheet.addRow([
      "TOTAIS DO PERÍODO", "", "", "", "", "", "", "",
      totalCobrado,
      totalEfetivamentePago,
      totalTaxas,
      totalLiquido,
      ""
    ]);

    totalRow.height = 24;
    sheet.mergeCells(`A${totalRow.number}:H${totalRow.number}`);

    totalRow.getCell(1).font = { name: "Arial", size: 10, bold: true };
    totalRow.getCell(1).alignment = { horizontal: "right", vertical: "middle" };

    // Estilização das colunas financeiras de resumo
    [9, 10, 11, 12].forEach((colIdx) => {
      const cell = totalRow.getCell(colIdx);
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0F172A" } };
      cell.numFmt = "R$ #,##0.00";
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
      cell.alignment = { vertical: "middle" };
    });

    // Destaque verde especificamente para o Total Líquido Entrado
    totalRow.getCell(12).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF059669" } };
    totalRow.getCell(12).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECFDF5" } };

    // Ajuste da largura das colunas
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