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
router.post("/empresas/filtrar", async (req: Request, res: Response) => {
  try {
    const { tipoAcordo, tipoConvencao, empresas, relatorioEmpresa, formato } = req.body;

    // ============================
    // 1. VALIDAÇÕES
    // ============================
    if (!empresas || !Array.isArray(empresas)) {
      return res.status(400).json({ error: "Lista de empresas inválida." });
    }

    if (!relatorioEmpresa) {
      return res.status(400).json({ error: "Configuração de relatório ausente." });
    }

    if (!["xlsx", "pdf"].includes(formato)) {
      return res.status(400).json({ error: "Formato inválido. Use 'xlsx' ou 'pdf'." });
    }

    // ============================
    // 2. PREPARA OS FILTROS
    // ============================

    const mapaAcordo = {
      mensalidade: "acordo_mensalidade",
      sindical: "acordo_sindical",
      negocial: "acordo_negocial",
    };

    let empresasFiltradas = empresas;

    // --- FILTRO POR ACORDO (OPCIONAL)
    if (tipoAcordo && mapaAcordo[tipoAcordo]) {
      const colunaAcordo = mapaAcordo[tipoAcordo];
      empresasFiltradas = empresasFiltradas.filter((e: any) => e[colunaAcordo] === true);
    }

    // --- FILTRO POR CONVENÇÃO (OPCIONAL)
    if (tipoConvencao && tipoConvencao.trim() !== "") {
      empresasFiltradas = empresasFiltradas.filter(
        (e: any) =>
          String(e.tipo_convencao).trim().toUpperCase() ===
          tipoConvencao.trim().toUpperCase()
      );
    }

    // --- Caso nenhum filtro, mantemos todas as empresas
    if (empresasFiltradas.length === 0) {
      return res.status(200).json({ aviso: "Nenhuma empresa encontrada com os filtros." });
    }

    // ============================
    // 3. SELEÇÃO DE COLUNAS
    // ============================
    const colunasSelecionadas = Object.keys(relatorioEmpresa).filter(
      (coluna) => relatorioEmpresa[coluna] === true
    );

    if (colunasSelecionadas.length === 0) {
      return res.status(400).json({ error: "Nenhuma coluna selecionada." });
    }

    // ============================
    // 4. ARQUIVO TEMPORÁRIO
    // ============================
    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fileName = `relatorio_empresas_filtradas_${Date.now()}.${formato}`;
    const filePath = path.join(outputDir, fileName);

    // ============================
    // 5. FORMATADOR DE CABEÇALHOS
    // ============================
    const formatHeader = (key: string) =>
      key
        .replace(/_/g, " ")
        .toLowerCase()
        .replace(/\b\w/g, (l) => l.toUpperCase());

    // ======================================================================
    // ============================= XLSX ===================================
    // ======================================================================
    if (formato === "xlsx") {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Empresas");

      const headerFormatted = colunasSelecionadas.map((c) => formatHeader(c));
      const headerRow = sheet.addRow(headerFormatted);
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.height = 22;

      headerRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "4472C4" },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });

      empresasFiltradas.forEach((empresa: any) => {
        const linha = colunasSelecionadas.map((col) => {
          let valor = empresa[col] ?? "";

          if (valor === true) valor = "Ativa";
          if (valor === false) valor = "Inativa";

          return String(valor);
        });

        sheet.addRow(linha);
      });

      colunasSelecionadas.forEach((col, i) => {
        const maxLength = Math.max(
          formatHeader(col).length,
          ...empresasFiltradas.map((e: any) => String(e[col] ?? "").length)
        );
        sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 40);
      });

      await workbook.xlsx.writeFile(filePath);
    }

    // ======================================================================
    // ============================= PDF ====================================
    // ======================================================================
    else {
      const pdfDoc = await PDFDocument.create();
      let page = pdfDoc.addPage([595, 842]);
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const { width, height } = page.getSize();

      page.drawText("SENALBA MG - Relatório de Empresas", {
        x: 40,
        y: height - 50,
        size: 16,
        font: boldFont,
        color: rgb(0, 0.3, 0.6),
      });

      const startX = 40;
      const startY = height - 90;
      const rowHeight = 22;
      const tableWidth = width - 80;
      const colWidth = tableWidth / colunasSelecionadas.length;
      let y = startY;

      page.drawRectangle({
        x: startX,
        y: y - rowHeight + 5,
        width: tableWidth,
        height: rowHeight,
        color: rgb(0.85, 0.9, 0.98),
      });

      let x = startX;
      colunasSelecionadas.forEach((col) => {
        page.drawText(formatHeader(col), {
          x: x + 4,
          y: y - 8,
          size: 10,
          font: boldFont,
          color: rgb(0, 0.2, 0.5),
        });
        x += colWidth;
      });

      y -= rowHeight;

      for (const empresa of empresasFiltradas) {
        x = startX;

        for (const col of colunasSelecionadas) {
          let valor = empresa[col] ?? "";

          if (valor === true) valor = "Ativa";
          if (valor === false) valor = "Inativa";

          let texto = String(valor);
          if (texto.length > 30) texto = texto.slice(0, 27) + "...";

          page.drawText(texto, {
            x: x + 4,
            y: y - 8,
            size: 9,
            font,
          });

          page.drawRectangle({
            x,
            y: y - rowHeight + 5,
            width: colWidth,
            height: rowHeight,
            borderWidth: 0.5,
            borderColor: rgb(0.8, 0.8, 0.8),
          });

          x += colWidth;
        }

        y -= rowHeight;

        if (y < 60) {
          page = pdfDoc.addPage([595, 842]);
          y = height - 90;
        }
      }

      const pdfBytes = await pdfDoc.save();
      fs.writeFileSync(filePath, pdfBytes);
    }

    return res.download(filePath, fileName, () => {
      fs.unlinkSync(filePath);
    });

  } catch (error) {
    console.error("🔥 ERRO back-end:", error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
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
  // 🟩 Geração XLSX (profissional)
if (formato === "xlsx") {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Empresas");

  // === Função de formatação dos nomes de colunas ===
  const formatHeader = (key: string) =>
    key
      .replace(/_/g, " ")
      .toLowerCase()
      .replace(/\b\w/g, (l) => l.toUpperCase());

  // === Cabeçalho formatado ===
  const headerRow = sheet.addRow(colunasSelecionadas.map((c) => formatHeader(c)));
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  headerRow.height = 22;

  // Fundo azul claro
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "4472C4" },
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FFCCCCCC" } },
      bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
  });

  // === Linhas de dados ===
  empresas.forEach((empresa: any) => {
    const linha = colunasSelecionadas.map((col) => {
      let valor = empresa[col] ?? "";

      // ✅ Substitui valores booleanos
      if (valor === true) valor = "Sim";
      if (valor === false) valor = "Não";

      // ✅ Garante texto simples e sem quebra
      return String(valor);
    });

    const dataRow = sheet.addRow(linha);

    dataRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFD9D9D9" } },
        bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
        left: { style: "thin", color: { argb: "FFD9D9D9" } },
        right: { style: "thin", color: { argb: "FFD9D9D9" } },
      };
      cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    });
  });

  // === Ajusta automaticamente a largura das colunas ===
  colunasSelecionadas.forEach((col, i) => {
    const maxLength = Math.max(
      formatHeader(col).length,
      ...empresas.map((empresa: any) => String(empresa[col] ?? "").length)
    );
    sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 35);
  });

  // === Título e resumo ===
  sheet.insertRow(1, ["SENALBA MG - Relatório de Empresas"]);
  const titleRow = sheet.getRow(1);
  titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
  sheet.mergeCells(1, 1, 1, colunasSelecionadas.length);
  titleRow.alignment = { horizontal: "center" };

  // === Gera arquivo ===
  await workbook.xlsx.writeFile(filePath);
}


// 🟥 Geração PDF (profissional)
else {
  const pdfDoc = await PDFDocument.create();
  let page = pdfDoc.addPage([595, 842]); // A4
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const { width, height } = page.getSize();

  // --- Cabeçalho ---
  page.drawText("SENALBA MG", {
    x: 50,
    y: height - 50,
    size: 18,
    font: boldFont,
    color: rgb(0, 0.3, 0.6),
  });

  page.drawText("Relatório de Empresas", {
    x: 50,
    y: height - 70,
    size: 13,
    font,
    color: rgb(0, 0, 0),
  });

  // --- Tabela ---
  const startX = 40;
  const startY = height - 100;
  const rowHeight = 22;
  const tableWidth = width - 80;
  const colCount = colunasSelecionadas.length;
  const colWidth = tableWidth / colCount;

  let y = startY;

  // Função para formatar cabeçalhos
  const formatHeader = (key: string) =>
    key
      .replace(/_/g, " ")
      .toLowerCase()
      .replace(/\b\w/g, (l) => l.toUpperCase());

  // Cabeçalho visual
  page.drawRectangle({
    x: startX,
    y: y - rowHeight + 5,
    width: tableWidth,
    height: rowHeight,
    color: rgb(0.85, 0.9, 0.98),
  });

  let x = startX;
  colunasSelecionadas.forEach((col) => {
    page.drawText(formatHeader(col), {
      x: x + 4,
      y: y - 8,
      size: 10,
      font: boldFont,
      color: rgb(0, 0.2, 0.5),
    });
    x += colWidth;
  });

  y -= rowHeight;

  // Linhas da tabela
  for (const empresa of empresas) {
    x = startX;
    colunasSelecionadas.forEach((col) => {
      let valor = empresa[col] ?? "";

      // ✅ Substitui valores booleanos
      if (valor === true) valor = "Sim";
      if (valor === false) valor = "Não";

      // ✅ Quebra textos longos
      let texto = String(valor);
      if (texto.length > 30) texto = texto.slice(0, 27) + "...";

      page.drawText(texto, {
        x: x + 4,
        y: y - 8,
        size: 9,
        font,
        color: rgb(0, 0, 0),
        maxWidth: colWidth - 8,
      });

      // ✅ Desenha a borda da célula
      page.drawRectangle({
        x,
        y: y - rowHeight + 5,
        width: colWidth,
        height: rowHeight,
        borderColor: rgb(0.8, 0.8, 0.8),
        borderWidth: 0.5,
      });

      x += colWidth;
    });

    y -= rowHeight;

    // Nova página quando chegar no fim
    if (y < 60) {
      page = pdfDoc.addPage([595, 842]);
      y = height - 100;
    }
  }

  // Salva o PDF
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

router.post("/sindicalizados", async (req: Request, res: Response) => {
  try {
    const { sindicalizados, escolhaColunaTabela } = req.body;

    if (!Array.isArray(sindicalizados)) {
      return res.status(400).json({ error: "Lista de sindicalizados inválida." });
    }

    if (!escolhaColunaTabela || typeof escolhaColunaTabela !== "object") {
      return res.status(400).json({ error: "Configuração de colunas inválida." });
    }

    // Seleciona apenas colunas onde o usuário marcou TRUE
    const colunasSelecionadas = Object.keys(escolhaColunaTabela).filter(
      (col) => escolhaColunaTabela[col] === true
    );

    if (colunasSelecionadas.length === 0) {
      return res.status(400).json({ error: "Nenhuma coluna selecionada." });
    }

    //
    // ──────────────────────────────────────────
    //   PREPARAÇÃO DO ARQUIVO
    // ──────────────────────────────────────────
    //

    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fileName = `relatorio_sindicalizados_${Date.now()}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    //
    // ──────────────────────────────────────────
    //   GERAR XLSX (MESMO PADRÃO PROFISSIONAL)
    // ──────────────────────────────────────────
    //

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sindicalizados");

    // Função para formatar os headers
    const formatHeader = (key: string) => {
  if (key === "status") return "Status (Ativo / Inativo)";

  return key
    .replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, (l) => l.toUpperCase());
};


    //
    // CABEÇALHO DA TABELA
    //
    const headerRow = sheet.addRow(colunasSelecionadas.map((c) => formatHeader(c)));
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 22;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "4472C4" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };
    });

    //
    // LINHAS DE DADOS
    //
    sindicalizados.forEach((item: any) => {
      const linha = colunasSelecionadas.map((col) => {
        let valor = item[col] ?? "";

        if (valor === true) valor = "Sim";
        if (valor === false) valor = "Não";

        return String(valor);
      });

      const row = sheet.addRow(linha);

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFD9D9D9" } },
          bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
          left: { style: "thin", color: { argb: "FFD9D9D9" } },
          right: { style: "thin", color: { argb: "FFD9D9D9" } },
        };
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      });
    });

    //
    // AJUSTAR LARGURA DAS COLUNAS
    //
    colunasSelecionadas.forEach((col, i) => {
      const maxLength = Math.max(
        formatHeader(col).length,
        ...sindicalizados.map((i: any) => String(i[col] ?? "").length)
      );

      sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 35);
    });

    //
    // TÍTULO NO TOPO
    //
    sheet.insertRow(1, ["SENALBA MG - Relatório de Sindicalizados"]);
    const titleRow = sheet.getRow(1);

    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    sheet.mergeCells(1, 1, 1, colunasSelecionadas.length);
    titleRow.alignment = { horizontal: "center" };

    //
    // GERAR ARQUIVO
    //
    await workbook.xlsx.writeFile(filePath);

    //
    // ENVIAR ARQUIVO
    //
    return res.download(filePath, fileName, (err) => {
      if (err) console.error("Erro ao enviar:", err);
      fs.unlinkSync(filePath);
    });

  } catch (error) {
    console.error("Erro no relatório:", error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});
router.post("/sindicalizados/filtro", async (req: Request, res: Response) => {
  try {
    const { sindicalizados, escolhaColunaTabela, filtros } = req.body;

    console.log("BODY COMPLETO RECEBIDO:", JSON.stringify(req.body, null, 2));

    if (!Array.isArray(sindicalizados)) {
      return res.status(400).json({ error: "Lista de sindicalizados inválida." });
    }

    if (!escolhaColunaTabela || typeof escolhaColunaTabela !== "object") {
      return res.status(400).json({ error: "Configuração de colunas inválida." });
    }

    //
    // ──────────────────────────────────────────
    //   APLICAR FILTROS (SE EXISTIREM)
    // ──────────────────────────────────────────
    //

    //
// ──────────────────────────────────────────
//   APLICAR FILTROS (COM id_empresa)
// ──────────────────────────────────────────

let listaFiltrada = [...sindicalizados];

if (filtros && typeof filtros === "object") {

  const normalizar = (v: any) =>
    String(v ?? "").trim().toLowerCase();

  listaFiltrada = listaFiltrada.filter((item) => {
    const itemStatus = normalizar(item.status);
    const itemEmpresaId = String(item.id_empresa ?? "");
    const itemUnidade = normalizar(item.unidade);
    const itemTipo = normalizar(item.tipo_desconto);

    // STATUS → ativo / inativo
    if (filtros.status) {
      const filtroStatus = normalizar(filtros.status);
      if (itemStatus !== filtroStatus) return false;
    }

    // FILTRAR POR ID DA EMPRESA
    if (filtros.id_empresa) {
      const filtroEmpresaId = String(filtros.id_empresa);
      if (itemEmpresaId !== filtroEmpresaId) return false;
    }

    // UNIDADE
    if (filtros.unidade) {
      if (itemUnidade !== normalizar(filtros.unidade)) return false;
    }

    // TIPO DE DESCONTO
    if (filtros.tipo_desconto) {
      if (itemTipo !== normalizar(filtros.tipo_desconto)) return false;
    }

    return true;
  });
}

//
// Se não encontrar nada → retorna erro
//
if (listaFiltrada.length === 0) {
  return res.status(404).json({
    error: "Nenhum resultado encontrado após aplicar filtros."
  });
}

    //
    // ──────────────────────────────────────────
    //   SELEÇÃO DAS COLUNAS
    // ──────────────────────────────────────────
    //

    const colunasSelecionadas = Object.keys(escolhaColunaTabela).filter(
      (col) => escolhaColunaTabela[col] === true
    );

    if (colunasSelecionadas.length === 0) {
      return res.status(400).json({ error: "Nenhuma coluna selecionada." });
    }

    //
    // ──────────────────────────────────────────
    //   PREPARAÇÃO DO XLSX
    // ──────────────────────────────────────────
    //

    const outputDir = path.join(__dirname, "../../uploads");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fileName = `relatorio_sindicalizados_${Date.now()}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sindicalizados");

    const formatHeader = (key: string) =>
      key
        .replace(/_/g, " ")
        .toLowerCase()
        .replace(/\b\w/g, (l) => l.toUpperCase());

    //
    // CABEÇALHO
    //
    const headerRow = sheet.addRow(colunasSelecionadas.map((c) => formatHeader(c)));
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 22;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "4472C4" },
      };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    });

    //
    // DADOS
    //
    listaFiltrada.forEach((item: any) => {
  const linha = colunasSelecionadas.map((col) => {
    if (col === "status") {
      // status booleano → string bonita
      return item.status === true ? "Ativo" : "Inativo";
    }

    return String(item[col] ?? "");
  });

  const row = sheet.addRow(linha);

  row.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    cell.alignment = {
      vertical: "middle",
      horizontal: "left",
      wrapText: true
    };
  });
});

    //
    // AJUSTE DE LARGURA
    //
    colunasSelecionadas.forEach((col, i) => {
      const maxLength = Math.max(
        formatHeader(col).length,
        ...listaFiltrada.map((i: any) => String(i[col] ?? "").length)
      );
      sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 35);
    });

    //
    // TÍTULO
    //
    sheet.insertRow(1, ["SENALBA MG - Relatório de Sindicalizados"]);
    const titleRow = sheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    sheet.mergeCells(1, 1, 1, colunasSelecionadas.length);
    titleRow.alignment = { horizontal: "center" };

    //
    // GERAR E ENVIAR
    //
    await workbook.xlsx.writeFile(filePath);

    return res.download(filePath, fileName, (err) => {
      if (err) console.error("Erro ao enviar:", err);
      fs.unlinkSync(filePath);
    });

  } catch (error) {
    console.error("Erro no relatório:", error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});


export default router;

