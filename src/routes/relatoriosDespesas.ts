import { Router, Request, Response } from "express";
import ExcelJS from "exceljs";
import dotenv from "dotenv";
dotenv.config();
const router = Router();

router.post("/despesa-diaria", async (req: Request, res: Response) => {
  try {
    const { dados } = req.body; // data_inicio foi removido, pois a lista já vem filtrada

    if (!Array.isArray(dados) || dados.length === 0) {
      return res.status(400).json({
        error: "dados deve ser uma lista e não pode estar vazia.",
      });
    }

    // ────────────────────────────────────────────
    // 🎯 FASE DE PREPARAÇÃO DA DATA
    // ────────────────────────────────────────────
    // Pega a data da primeira despesa para usar no título e no nome do arquivo
    const dataReferencia = dados[0].date_despesa; // Ex: "2026-08-14"
    let dataFormatadaTitulo = "";
    let nomeArquivo = "data_indefinida";

    if (dataReferencia) {
      const [ano, mes, dia] = dataReferencia.split("-").map(Number);
      const dataUTC = new Date(Date.UTC(ano, mes - 1, dia));
      
      dataFormatadaTitulo = dataUTC.toLocaleDateString("pt-BR"); // Fica "14/08/2026"
      nomeArquivo = `${dia}_${mes}_${ano}`;
    }

    // ────────────────────────────────────────────
    // 📊 CRIAR EXCEL 
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Despesas Diárias");

    // Mantida a mesma estrutura de 5 colunas
    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Fornecedor", key: "fornecedor", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Pagamento", key: "categoria_pagamento", width: 20 },
      { header: "Valor (R$)", key: "valor", width: 15 }, 
    ];

    sheet.columns = headers;

    // 1. TÍTULO (Linha 1)
    sheet.mergeCells("A1:E1");
    const titulo = sheet.getCell("A1");
    titulo.value = `Despesas do Dia ${dataFormatadaTitulo}`;
    titulo.font = { bold: true, size: 18 };
    titulo.alignment = { horizontal: "center" };

    // 2. CABEÇALHOS (Linha 2)
    const headerRow = sheet.getRow(2);
    headerRow.values = headers.map(h => h.header);

    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDDDDD" },
      };
    });

    // 3. INSERIR DADOS (A partir da Linha 3)
    let totalValor = 0; 

    // Como já está filtrado, percorremos todos os dados recebidos
    dados.forEach((item) => {
      let dataFormatada = "-";
      
      if (item.date_despesa) {
        const [year, month, day] = item.date_despesa.split("-").map(Number);
        const dataUTC = new Date(Date.UTC(year, month - 1, day));
        dataFormatada = dataUTC.toLocaleDateString("pt-BR");
      }
      
      const valorItem = Number(item.valor) || 0;
      totalValor += valorItem; 

      sheet.addRow({
        data: dataFormatada,
        descricao: item.descricao || "-",
        valor: valorItem,
        categoria_pagamento: item.categoria_pagamento || "-",
        fornecedor: item.fornecedor || "-",
      });
    });

    // 4. INSERIR A LINHA DE TOTAL AO FINAL
    const totalRow = sheet.addRow({
      data: "",
      descricao: "TOTAL",
      valor: totalValor,
      categoria_pagamento: "",
      fornecedor: "",
    });

    // Formata a linha de total em negrito
    totalRow.font = { bold: true };

    // Formata a coluna de valores como moeda
    sheet.getColumn("valor").numFmt = "R$ #,##0.00";

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=despesas_diarias_${nomeArquivo}.xlsx`
    );

    return res.send(buffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório de despesas diário." });
  }
});
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
    // 🎯 FASE DE FILTRAGEM 
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
    // 📊 CRIAR EXCEL 
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Despesas");

    // 🟢 REMOVIDAS as colunas solicitadas (Ficaram apenas 5)
    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Fornecedor", key: "fornecedor", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Pagamento", key: "categoria_pagamento", width: 20 },
      { header: "Valor (R$)", key: "valor", width: 15 }, 
      
      
    ];

    sheet.columns = headers;

    // 🟢 AJUSTADO para mesclar apenas as 5 colunas (A até E)
    sheet.mergeCells("A1:E1");
    const titulo = sheet.getCell("A1");
    titulo.value = `Despesas Mensais ${nomeMes} ${anoFiltro}`;
    titulo.font = { bold: true, size: 18 };
    titulo.alignment = { horizontal: "center" };

    // 2. CABEÇALHOS (Linha 2)
    const headerRow = sheet.getRow(2);
    headerRow.values = headers.map(h => h.header);

    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDDDDD" },
      };
    });

    // 3. INSERIR DADOS (A partir da Linha 3)
    let totalValor = 0; 

    despesasFiltradas.forEach((item) => {
      const [year, month, day] = item.date_despesa.split("-").map(Number);
      const dataUTC = new Date(Date.UTC(year, month - 1, day));
      
      const dataFormatada = dataUTC.toLocaleDateString("pt-BR");
      const valorItem = Number(item.valor) || 0;

      totalValor += valorItem; 

      // 🟢 AJUSTADO: Removidos os campos extras na inserção
      sheet.addRow({
        data: dataFormatada,
        descricao: item.descricao || "-",
        valor: valorItem,
        categoria_pagamento: item.categoria_pagamento || "-",
        fornecedor: item.fornecedor || "-",
      });
    });

    // 4. INSERIR A LINHA DE TOTAL AO FINAL
    // 🟢 AJUSTADO: Removidos os campos extras do total
    const totalRow = sheet.addRow({
      data: "",
      descricao: "TOTAL",
      valor: totalValor,
      categoria_pagamento: "",
      fornecedor: "",
    });

    // Formata a linha de total em negrito
    totalRow.font = { bold: true };

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
router.post("/despesa-mensal-classificado", async (req: Request, res: Response) => {
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

    // Extrair ano e mês de data_inicio (formato esperado: YYYY-MM-DD)
    const dataISO = data_inicio.substring(0, 10);
    const partesFiltro = dataISO.split("-").map(Number);

    const anoFiltro = partesFiltro[0];
    const mesFiltro = partesFiltro[1]; // 1–12

    // Nome do mês para uso em títulos/nomes de arquivo
    const dataParaNomeMes = new Date(Date.UTC(anoFiltro, mesFiltro - 1, 1));
    const nomeMes = dataParaNomeMes.toLocaleString("pt-BR", { month: "long" });

    const despesasFiltradas = dados.filter((item) => {
      if (!item.date_despesa) return false;
      const partesDespesa = item.date_despesa.split("-");
      if (partesDespesa.length !== 3) return false;
      const [anoDesp, mesDesp] = partesDespesa.map(Number);
      return anoDesp === anoFiltro && mesDesp === mesFiltro;
    });

    const workbook = new ExcelJS.Workbook();
    
    // ─────────────────────────────────────────────────────────────────
    // 📊 ABA 1: DESPESAS DETALHADAS
    // ─────────────────────────────────────────────────────────────────
    const sheet1 = workbook.addWorksheet("Detalhamento");
    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Valor (R$)", key: "valor", width: 15 },
      { header: "Categoria", key: "categoria", width: 20 },
      { header: "Subcategoria", key: "subcategoria", width: 20 },
    ];
    sheet1.columns = headers;

    sheet1.mergeCells("A1:E1");
    sheet1.getCell("A1").value = `Detalhamento de Despesas Classificado - ${nomeMes}/${anoFiltro}`;
    sheet1.getCell("A1").font = { bold: true, size: 14 };
    sheet1.getCell("A1").alignment = { horizontal: "center" };

    // Cabeçalhos (Linha 2)
    sheet1.getRow(2).values = headers.map(h => h.header);
    sheet1.getRow(2).eachCell(cell => {
      cell.font = { bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDDDDD" } };
    });

    let somaTotalGeral = 0;
    const resumoCategorias = {}; // Objeto para acumular somas da Aba 2

    // Preencher linhas e acumular totais
    despesasFiltradas.forEach((item) => {
      const valor = parseFloat(item.valor) || 0;
      const cat = item.despesa_categoria || "Sem Categoria";
      
      somaTotalGeral += valor;
      resumoCategorias[cat] = (resumoCategorias[cat] || 0) + valor;

      const [y, m, d] = item.date_despesa.split("-").map(Number);
      const dataFormatada = new Date(Date.UTC(y, m - 1, d)).toLocaleDateString("pt-BR");

      sheet1.addRow({
        data: dataFormatada,
        descricao: item.descricao || "",
        valor: valor,
        categoria: cat,
        subcategoria: item.despesa_sub_categoria || "",
      });
    });

    // 🟢 1. LINHA DE TOTAL GERAL (Aba 1)
    const rowTotal = sheet1.addRow({
      descricao: "TOTAL GERAL",
      valor: somaTotalGeral
    });
    rowTotal.font = { bold: true };
    sheet1.getColumn("valor").numFmt = "R$ #,##0.00";

    // ─────────────────────────────────────────────────────────────────
    // 📈 ABA 2: RESUMO POR CATEGORIA
    // ─────────────────────────────────────────────────────────────────
    const sheet2 = workbook.addWorksheet("Resumo por Categoria");
    
    sheet2.columns = [
      { header: "Categoria", key: "cat", width: 30 },
      { header: "Total Acumulado (R$)", key: "total", width: 25 }
    ];

    sheet2.getRow(1).font = { bold: true };

    // Inserir os dados acumulados
    Object.keys(resumoCategorias).forEach(cat => {
      sheet2.addRow({
        cat: cat,
        total: resumoCategorias[cat]
      });
    });

    sheet2.getColumn("total").numFmt = "R$ #,##0.00";

    // Exportação
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=Relatorio_${nomeMes}.xlsx`);
    return res.send(buffer);

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório completo." });
  }
});
router.post("/despesa-anual", async (req: Request, res: Response) => {
  try {
    const { listaDespesas, listaArrecadacao, dataFiltro } = req.body;

    // --- 1. Validação e Preparação do Filtro ---
    if (!dataFiltro) {
      return res.status(400).json({ error: "dataFiltro é obrigatória." });
    }

    // Garante que as listas sejam arrays, mesmo que venham nulas (para evitar .forEach() em null)
    const despesasBase = Array.isArray(listaDespesas) ? listaDespesas : [];
    const arrecadacoesBase = Array.isArray(listaArrecadacao) ? listaArrecadacao : [];
    
    const dataISO = dataFiltro.substring(0, 4);
    const anoFiltro = Number(dataISO);
    const tituloRelatorio = `Prestação de Contas Anual - ${anoFiltro}`;

    // Variáveis de acumulação ANUAL e MENSAL TOTAL
    const resumoDespesasPorCategoria = {};    // Total Despesas por Categoria (Anual)
    const resumoArrecadacaoPorTipo = {};      // Total Arrecadação por Tipo (Anual)
    const totalDespesasMes = {};             // Total Despesas por Mês (Global)
    const totalArrecadacaoMes = {};          // Total Arrecadação por Mês (Global)
    let totalDespesasAnual = 0;
    let totalArrecadacaoAnual = 0;

    // NOVAS ESTRUTURAS: Agregação por MÊS e por Categoria/Tipo
    const resumoDespesasPorCategoriaMes = {};
    const resumoArrecadacaoPorTipoMes = {};

    // ────────────────────────────────────────────
    // 2. FILTRAGEM E AGREGAÇÃO POR ANO E POR MÊS
    // ────────────────────────────────────────────

    // A. Filtragem e Agregação de Despesas
    despesasBase.forEach((item) => {
      if (!item.date_despesa) return;
      const partesDespesa = item.date_despesa.split("-");
      const anoDesp = Number(partesDespesa[0]);
      const mesDesp = Number(partesDespesa[1]);
      
      if (anoDesp === anoFiltro) {
        const valor = parseFloat(item.valor) || 0;
        const categoria = item.despesa_categoria || "Sem Categoria";
        
        // Acumulação Anual
        totalDespesasAnual += valor; 
        
        // Acumulação por Categoria (Anual)
        resumoDespesasPorCategoria[categoria] = (resumoDespesasPorCategoria[categoria] || 0) + valor;
        
        // Acumulação por Mês (Mensal Total)
        totalDespesasMes[mesDesp] = (totalDespesasMes[mesDesp] || 0) + valor;

        // NOVO: Acumulação por Categoria e por Mês
        if (!resumoDespesasPorCategoriaMes[mesDesp]) {
            resumoDespesasPorCategoriaMes[mesDesp] = {};
        }
        resumoDespesasPorCategoriaMes[mesDesp][categoria] = 
            (resumoDespesasPorCategoriaMes[mesDesp][categoria] || 0) + valor;
      }
    });

    // B. Filtragem e Agregação de Arrecadações
    arrecadacoesBase.forEach((item) => {
      if (!item.data_recebido) return;
      const partesArrec = item.data_recebido.split("-");
      const anoArr = Number(partesArrec[0]);
      const mesArr = Number(partesArrec[1]);
      
      if (anoArr === anoFiltro) {
        const valor = parseFloat(item.valor) || 0;
        const tipo = item.tipo_arrecadacao || "Sem Tipo";
        
        // Acumulação Anual
        totalArrecadacaoAnual += valor; 
        
        // Acumulação por Tipo (Anual)
        resumoArrecadacaoPorTipo[tipo] = (resumoArrecadacaoPorTipo[tipo] || 0) + valor;

        // Acumulação por Mês (Mensal Total)
        totalArrecadacaoMes[mesArr] = (totalArrecadacaoMes[mesArr] || 0) + valor;

        // NOVO: Acumulação por Tipo e por Mês
        if (!resumoArrecadacaoPorTipoMes[mesArr]) {
            resumoArrecadacaoPorTipoMes[mesArr] = {};
        }
        resumoArrecadacaoPorTipoMes[mesArr][tipo] = 
            (resumoArrecadacaoPorTipoMes[mesArr][tipo] || 0) + valor;
      }
    });

    const balancoAnual = totalArrecadacaoAnual - totalDespesasAnual;

    // ────────────────────────────────────────────
    // 3. GERAÇÃO DO EXCEL
    // ────────────────────────────────────────────
    
    const workbook = new ExcelJS.Workbook();
    const mesesNomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

    const headerStyle = {
      font: { bold: true },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDDDDD" } }
    };
    const currencyFormat = "R$ #,##0.00";

    // --- 3.1 ABA 1: Balanço Anual (Total por Mês) ---
    const sheet1 = workbook.addWorksheet("Balanço Anual");
    sheet1.mergeCells("A1:D1");
    sheet1.getCell("A1").value = tituloRelatorio;
    sheet1.getCell("A1").font = { bold: true, size: 16 };
    sheet1.getCell("A1").alignment = { horizontal: "center" };

    sheet1.columns = [
      { header: "Mês", key: "item", width: 15 },
      { header: "Total Arrecadação (R$)", key: "arr", width: 20 },
      { header: "Total Despesas (R$)", key: "desp", width: 20 },
      { header: "Saldo (R$)", key: "saldo", width: 15 },
    ];

    sheet1.getRow(3).values = ["Mês", "Total Arrecadação (R$)", "Total Despesas (R$)", "Saldo (R$)"];
    sheet1.getRow(3).eachCell(cell => Object.assign(cell, headerStyle));

    for (let mes = 1; mes <= 12; mes++) {
      const nomeMesCurto = mesesNomes[mes - 1]; // Usa a lista de meses
      const arr = totalArrecadacaoMes[mes] || 0;
      const desp = totalDespesasMes[mes] || 0;
      const saldo = arr - desp;

      sheet1.addRow({ item: nomeMesCurto, arr: arr, desp: desp, saldo: saldo });
    }

    const finalRow1 = sheet1.addRow({
        item: "TOTAL ANUAL",
        arr: totalArrecadacaoAnual,
        desp: totalDespesasAnual,
        saldo: balancoAnual
    });
    finalRow1.font = { bold: true };
    sheet1.getColumn("arr").numFmt = currencyFormat;
    sheet1.getColumn("desp").numFmt = currencyFormat;
    sheet1.getColumn("saldo").numFmt = currencyFormat;

    // --- 3.2 ABA 2: Despesas por Categoria (Anual) ---
    const sheet2 = workbook.addWorksheet("Despesas Categoria (Anual)");
    sheet2.columns = [
      { header: "Categoria", key: "categoria", width: 30 },
      { header: "Total (R$)", key: "total", width: 20 }
    ];
    sheet2.getRow(1).eachCell(cell => Object.assign(cell, headerStyle));
    Object.keys(resumoDespesasPorCategoria).forEach(cat => {
      sheet2.addRow({ categoria: cat, total: resumoDespesasPorCategoria[cat] });
    });
    sheet2.addRow({ categoria: "TOTAL ANUAL", total: totalDespesasAnual }).font = { bold: true };
    sheet2.getColumn("total").numFmt = currencyFormat;

    // --- 3.3 ABA 3: Arrecadação por Tipo (Anual) ---
    const sheet3 = workbook.addWorksheet("Arrecadação Tipo (Anual)");
    sheet3.columns = [
      { header: "Tipo", key: "tipo", width: 30 },
      { header: "Total (R$)", key: "total", width: 20 }
    ];
    sheet3.getRow(1).eachCell(cell => Object.assign(cell, headerStyle));
    Object.keys(resumoArrecadacaoPorTipo).forEach(tipo => {
      sheet3.addRow({ tipo: tipo, total: resumoArrecadacaoPorTipo[tipo] });
    });
    sheet3.addRow({ tipo: "TOTAL ANUAL", total: totalArrecadacaoAnual }).font = { bold: true };
    sheet3.getColumn("total").numFmt = currencyFormat;

    // ────────────────────────────────────────────
    // 🟢 3.4 ABA 4: Despesas por Mês e Categoria (CORRIGIDO NOME)
    // ────────────────────────────────────────────
    const sheet4 = workbook.addWorksheet("Despesas Mês-Categoria"); // <<<<<< CORREÇÃO AQUI
    sheet4.mergeCells("A1:N1");
    sheet4.getCell("A1").value = `Despesas por Categoria - Análise Mensal (${anoFiltro})`;
    sheet4.getCell("A1").font = { bold: true, size: 14 };
    
    const todasCategorias = new Set(Object.keys(resumoDespesasPorCategoria));
    
    const columns4 = [{ header: "Categoria", key: "categoria", width: 25 }];
    mesesNomes.forEach((m, index) => {
        columns4.push({ header: m, key: `m${index + 1}`, width: 12 });
    });
    columns4.push({ header: "Total Anual (R$)", key: "totalAnual", width: 20 });
    sheet4.columns = columns4;
    
    // Define o cabeçalho na linha 3
    const headerRow4 = sheet4.getRow(3);
    headerRow4.values = columns4.map(c => c.header);
    headerRow4.eachCell(cell => Object.assign(cell, headerStyle));

    // Popula as linhas
    todasCategorias.forEach(categoria => {
        // Tipagem corrigida para aceitar as chaves dinâmicas
        const row: Record<string, string | number> = { 
          categoria: categoria, 
          totalAnual: resumoDespesasPorCategoria[categoria] 
        };
        for (let m = 1; m <= 12; m++) {
            row[`m${m}`] = resumoDespesasPorCategoriaMes[m]?.[categoria] || 0;
        }
        sheet4.addRow(row);
    });

    // Linha de Total Mensal
    const totalRowData4: Record<string, string | number> = { categoria: "TOTAL POR MÊS" };
    for (let m = 1; m <= 12; m++) {
        totalRowData4[`m${m}`] = totalDespesasMes[m] || 0;
    }
    totalRowData4.totalAnual = totalDespesasAnual;
    const totalRow4 = sheet4.addRow(totalRowData4);
    totalRow4.font = { bold: true };

    // Aplica formatação monetária nas colunas de mês e total
    for (let i = 2; i <= 14; i++) { // Colunas 2 (m1) a 14 (Total Anual)
        sheet4.getColumn(i).numFmt = currencyFormat;
    }


    // ────────────────────────────────────────────
    // 🟢 3.5 ABA 5: Arrecadação por Mês e Tipo (CORRIGIDO NOME)
    // ────────────────────────────────────────────
    const sheet5 = workbook.addWorksheet("Arrecadação Mês-Tipo"); // <<<<<< CORREÇÃO AQUI
    sheet5.mergeCells("A1:N1");
    sheet5.getCell("A1").value = `Arrecadação por Tipo - Análise Mensal (${anoFiltro})`;
    sheet5.getCell("A1").font = { bold: true, size: 14 };
    
    const todosTipos = new Set(Object.keys(resumoArrecadacaoPorTipo));
    
    const columns5 = [{ header: "Tipo", key: "tipo", width: 25 }];
    mesesNomes.forEach((m, index) => {
        columns5.push({ header: m, key: `m${index + 1}`, width: 12 });
    });
    columns5.push({ header: "Total Anual (R$)", key: "totalAnual", width: 20 });
    sheet5.columns = columns5;
    
    // Define o cabeçalho na linha 3
    const headerRow5 = sheet5.getRow(3);
    headerRow5.values = columns5.map(c => c.header);
    headerRow5.eachCell(cell => Object.assign(cell, headerStyle));

    // Popula as linhas
    todosTipos.forEach(tipo => {
        // Tipagem corrigida para aceitar as chaves dinâmicas
        const row: Record<string, string | number> = { 
          tipo: tipo, 
          totalAnual: resumoArrecadacaoPorTipo[tipo] 
        };
        for (let m = 1; m <= 12; m++) {
            row[`m${m}`] = resumoArrecadacaoPorTipoMes[m]?.[tipo] || 0;
        }
        sheet5.addRow(row);
    });

    // Linha de Total Mensal
    const totalRowData5: Record<string, string | number> = { tipo: "TOTAL POR MÊS" };
    for (let m = 1; m <= 12; m++) {
        totalRowData5[`m${m}`] = totalArrecadacaoMes[m] || 0;
    }
    totalRowData5.totalAnual = totalArrecadacaoAnual;
    const totalRow5 = sheet5.addRow(totalRowData5);
    totalRow5.font = { bold: true };

    // Aplica formatação monetária
    for (let i = 2; i <= 14; i++) { // Colunas 2 (m1) a 14 (Total Anual)
        sheet5.getColumn(i).numFmt = currencyFormat;
    }

    // ────────────────────────────────────────────
    // 4. EXPORTAÇÃO
    // ────────────────────────────────────────────
    
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename=Prestacao_Anual_${anoFiltro}.xlsx`);
    return res.send(buffer);

  } catch (error) {
    console.error('❌ ERRO CRÍTICO NA GERAÇÃO DO RELATÓRIO ANUAL:', error);
    return res.status(500).json({ error: "Erro interno do servidor ao gerar relatório anual." });
  }
});


export default router;

