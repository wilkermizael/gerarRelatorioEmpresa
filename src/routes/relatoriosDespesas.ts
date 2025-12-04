import { Router, Request, Response } from "express";
import ExcelJS from "exceljs";
import dotenv from "dotenv";
dotenv.config();
const router = Router();


router.post("/despesa-mensal", async (req: Request, res: Response) => {
  try {
    const { dados, data_inicio } = req.body;

    if (!data_inicio) {
      return res.status(400).json({ error: "data_inicio é obrigatória." });
    }

    if (!Array.isArray(dados) || dados.length === 0) {
      return res.status(400).json({
        error: "dados deve ser uma lista e não pode estar vazia.",
      });
    }

    // ────────────────────────────────────────────
    // 🎯 FASE DE FILTRAGEM (Mantido o código que está funcionando)
    // ────────────────────────────────────────────

    // 1. EXTRAIR APENAS MÊS E ANO DE data_inicio
    const dataISO = data_inicio.substring(0, 10);
    const partesFiltro = dataISO.split("-").map(Number);

    const anoFiltro = partesFiltro[0];
    const mesFiltro = partesFiltro[1]; // 1–12
    
    // Pegar o nome do mês para o título.
    const dataParaNomeMes = new Date(Date.UTC(anoFiltro, mesFiltro - 1, 1));
    const nomeMes = dataParaNomeMes.toLocaleString("pt-BR", { month: "long" });


    // 2. FILTRAR DADOS: Compara APENAS Ano e Mês
    const despesasFiltradas = dados.filter((item) => {
      if (!item.date_despesa) return false;
      
      const partesDespesa = item.date_despesa.split("-");
      if (partesDespesa.length !== 3) return false;

      const [anoDesp, mesDesp] = partesDespesa.map(Number);
      
      return anoDesp === anoFiltro && mesDesp === mesFiltro;
    });


    // ────────────────────────────────────────────
    // 📊 CRIAR EXCEL (CORRIGIDO A INSERÇÃO DE CABEÇALHOS)
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Despesas");

    // Definição das colunas/chaves
    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Descrição", key: "descricao", width: 40 },
      { header: "Valor (R$)", key: "valor", width: 15 }, 
    ];

    sheet.columns = headers; // Define a estrutura do ExcelJS

    // 1. TÍTULO (Linha 1)
    sheet.mergeCells("A1:C1");
    const titulo = sheet.getCell("A1");
    // Ajustado para o formato solicitado: Despesas Mensais Dezembro 2025
    titulo.value = `Despesas Mensais ${nomeMes} ${anoFiltro}`;
    titulo.font = { bold: true, size: 18 };
    titulo.alignment = { horizontal: "center" };


    // 2. CABEÇALHOS (Linha 2)
    // 🟢 CORREÇÃO: Forçar a inserção dos headers na Linha 2
    const headerRow = sheet.getRow(2);
    headerRow.values = headers.map(h => h.header);

    // Formatação da Linha 2 (agora aplicada à linha que contém os headers)
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDDDDD" },
      };
    });

    // 3. INSERIR DADOS (A partir da Linha 3)
    // sheet.addRow() agora começará na linha 3, após o cabeçalho
    despesasFiltradas.forEach((item) => {
      // Cria a data em UTC a partir das partes para evitar o deslocamento do fuso.
      const [year, month, day] = item.date_despesa.split("-").map(Number);
      const dataUTC = new Date(Date.UTC(year, month - 1, day));
      
      const dataFormatada = dataUTC.toLocaleDateString("pt-BR");

      sheet.addRow({
        data: dataFormatada,
        descricao: item.descricao || "",
        valor: item.valor || 0,
      });
    });

    sheet.getColumn("valor").numFmt = "R$ #,##0.00";

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=despesas_${mesFiltro}_${anoFiltro}.xlsx`
    );

    return res.send(buffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório de despesas." });
  }
});

export default router;

