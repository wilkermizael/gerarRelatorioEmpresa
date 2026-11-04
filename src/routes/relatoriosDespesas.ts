import { Router, Request, Response } from "express";
import ExcelJS from "exceljs";
import dotenv from "dotenv";
dotenv.config();
const router = Router();

/*router.post("/despesa-diaria", async (req: Request, res: Response) => {
  try {
    const { dados } = req.body;

    if (!Array.isArray(dados) || dados.length === 0) {
      return res.status(400).json({
        error: "dados deve ser uma lista e não pode estar vazia.",
      });
    }

    // ────────────────────────────────────────────
    // 🎯 PREPARAÇÃO DA DATA
    // ────────────────────────────────────────────
    const dataReferencia = dados[0].date_despesa;
    let dataFormatadaTitulo = "";
    let nomeArquivo = "data_indefinida";

    if (dataReferencia) {
      const [ano, mes, dia] = dataReferencia.split("-").map(Number);
      const dataUTC = new Date(Date.UTC(ano, mes - 1, dia));
      
      dataFormatadaTitulo = dataUTC.toLocaleDateString("pt-BR");
      nomeArquivo = `${dia}_${mes}_${ano}`;
    }

    // ────────────────────────────────────────────
    // 📊 CRIAR EXCEL 
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Despesas Diárias");

    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Fornecedor", key: "fornecedor", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Categoria", key: "categoria_pagamento", width: 25 },
      { header: "Banco", key: "banco", width: 20 },
      { header: "Forma de Pagamento", key: "forma_pagamento", width: 22 },
      { header: "Valor (R$)", key: "valor", width: 15 }, 
    ];

    sheet.columns = headers;

    // 1. TÍTULO (Linha 1) - Ajustado para A1:G1 para cobrir as 7 colunas
    sheet.mergeCells("A1:G1");
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

    // 3. INSERIR DADOS
    let totalValor = 0; 

    dados.forEach((item) => {
      let dataFormatada = "-";
      
      if (item.date_despesa) {
        const [year, month, day] = item.date_despesa.split("-").map(Number);
        const dataUTC = new Date(Date.UTC(year, month - 1, day));
        dataFormatada = dataUTC.toLocaleDateString("pt-BR");
      }
      
      const valorItem = Number(item.valor) || 0;
      totalValor += valorItem; 

      // Busca flexível dos campos
      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || item.Forma_Pagamento || "-";

      sheet.addRow({
        data: dataFormatada,
        fornecedor: item.fornecedor || "-",
        descricao: item.descricao || "-",
        categoria_pagamento: item.categoria_pagamento || "-",
        banco: bancoValor,
        forma_pagamento: formaPagamentoValor,
        valor: valorItem,
      });
    });

    // 4. LINHA DE TOTAL
    const totalRow = sheet.addRow({
      data: "",
      fornecedor: "",
      descricao: "TOTAL",
      categoria_pagamento: "",
      banco: "",
      forma_pagamento: "",
      valor: totalValor,
    });

    totalRow.font = { bold: true };
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
});*/

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
    const dataISO = data_inicio.substring(0, 10);
    const partesFiltro = dataISO.split("-").map(Number);

    const anoFiltro = partesFiltro[0];
    const mesFiltro = partesFiltro[1]; // 1–12
    
    const dataParaNomeMes = new Date(Date.UTC(anoFiltro, mesFiltro - 1, 1));
    const nomeMes = dataParaNomeMes.toLocaleString("pt-BR", { month: "long" });

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

    // Estrutura expandida para 7 colunas (incluindo Banco e Forma de Pagamento)
    const headers = [
      { header: "Data", key: "data", width: 15 },
      { header: "Fornecedor", key: "fornecedor", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Categoria", key: "categoria_pagamento", width: 25 },
      { header: "Banco", key: "banco", width: 20 },
      { header: "Forma de Pagamento", key: "forma_pagamento", width: 22 },
      { header: "Valor (R$)", key: "valor", width: 15 }, 
    ];

    sheet.columns = headers;

    // 1. TÍTULO (Linha 1) - Ajustado para A1:G1 para cobrir as 7 colunas
    sheet.mergeCells("A1:G1");
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

      // Busca flexível dos campos banco e forma_pagamento
      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || item.Forma_Pagamento || "-";

      sheet.addRow({
        data: dataFormatada,
        fornecedor: item.fornecedor || "-",
        descricao: item.descricao || "-",
        categoria_pagamento: item.categoria_pagamento || "-",
        banco: bancoValor,
        forma_pagamento: formaPagamentoValor,
        valor: valorItem,
      });
    });

    // 4. INSERIR A LINHA DE TOTAL AO FINAL
    const totalRow = sheet.addRow({
      data: "",
      fornecedor: "",
      descricao: "TOTAL",
      categoria_pagamento: "",
      banco: "",
      forma_pagamento: "",
      valor: totalValor,
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
      { header: "Categoria", key: "categoria", width: 20 },
      { header: "Subcategoria", key: "subcategoria", width: 20 },
      { header: "Banco", key: "banco", width: 20 },
      { header: "Forma de Pagamento", key: "forma_pagamento", width: 22 },
      { header: "Valor (R$)", key: "valor", width: 15 },
    ];
    sheet1.columns = headers;

    // Mesclagem ajustada de A1 até G1 para cobrir as 7 colunas
    sheet1.mergeCells("A1:G1");
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

      // Busca flexível dos novos campos
      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || item.Forma_Pagamento || "-";

      sheet1.addRow({
        data: dataFormatada,
        descricao: item.descricao || "-",
        categoria: cat,
        subcategoria: item.despesa_sub_categoria || "-",
        banco: bancoValor,
        forma_pagamento: formaPagamentoValor,
        valor: valor,
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

    const despesasBase = Array.isArray(listaDespesas) ? listaDespesas : [];
    const arrecadacoesBase = Array.isArray(listaArrecadacao) ? listaArrecadacao : [];
    
    const dataISO = String(dataFiltro).substring(0, 4);
    const anoFiltro = Number(dataISO);
    const tituloRelatorio = `Prestação de Contas Anual - ${anoFiltro}`;

    // Helper para buscar a string de data disponível no objeto
    const obterDataBruta = (item: any, tipo: 'despesa' | 'receita') => {
      if (tipo === 'despesa') {
        return item.date_despesa || item.data || item.created_at || "";
      } else {
        return item.data_recebido || item.date_receita || item.date_arrecadacao || item.data || item.created_at || "";
      }
    };

    // Helper robusto para converter qualquer formato de data em Timestamp UTC para ordenação
    const extrairTimestamp = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return 0;

      // 1. Tenta formato ISO (Ex: "2026-09-16" ou "2026-09-16T14:30:00")
      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        const y = parseInt(matchISO[1], 10);
        const m = parseInt(matchISO[2], 10);
        const d = parseInt(matchISO[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      // 2. Tenta formato BR (Ex: "16/09/2026")
      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        const d = parseInt(matchBR[1], 10);
        const m = parseInt(matchBR[2], 10);
        const y = parseInt(matchBR[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      return 0;
    };

    // Helper para formatar a data na exibição do Excel (DD/MM/YYYY)
    const formatarData = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return "-";

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        return `${matchISO[3]}/${matchISO[2]}/${matchISO[1]}`;
      }

      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        return `${matchBR[1]}/${matchBR[2]}/${matchBR[3]}`;
      }

      return "-";
    };

    // Helper para extrair o Ano e Mês numéricos de um item
    const extrairAnoEMes = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return { ano: 0, mes: 0 };

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        return { ano: parseInt(matchISO[1], 10), mes: parseInt(matchISO[2], 10) };
      }

      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        return { ano: parseInt(matchBR[3], 10), mes: parseInt(matchBR[2], 10) };
      }

      return { ano: 0, mes: 0 };
    };

    // Variáveis de acumulação ANUAL e MENSAL TOTAL
    const resumoDespesasPorCategoria: Record<string, number> = {};
    const resumoDespesasPorBanco: Record<string, number> = {};
    const resumoDespesasPorFormaPagamento: Record<string, number> = {};
    const resumoArrecadacaoPorTipo: Record<string, number> = {};
    
    const totalDespesasMes: Record<number, number> = {};
    const totalArrecadacaoMes: Record<number, number> = {};
    let totalDespesasAnual = 0;
    let totalArrecadacaoAnual = 0;

    const resumoDespesasPorCategoriaMes: Record<number, Record<string, number>> = {};
    const resumoArrecadacaoPorTipoMes: Record<number, Record<string, number>> = {};

    const despesasFiltradasAno: any[] = [];
    const arrecadacoesFiltradasAno: any[] = [];

    // ────────────────────────────────────────────
    // 2. FILTRAGEM E AGREGAÇÃO POR ANO E MÊS
    // ────────────────────────────────────────────

    // A. Filtragem e Agregação de Despesas
    despesasBase.forEach((item) => {
      const { ano: anoDesp, mes: mesDesp } = extrairAnoEMes(item, 'despesa');
      
      if (anoDesp === anoFiltro && mesDesp > 0) {
        const valor = parseFloat(item.valor) || 0;
        const categoria = item.despesa_categoria || item.categoria_pagamento || "Sem Categoria";
        const banco = item.banco || item.Banco || item.nome_banco || "-";
        const formaPagamento = item.forma_pagamento || item.formaPagamento || item.Forma_Pagamento || "-";

        despesasFiltradasAno.push({ ...item, bancoCalculado: banco, formaPagamentoCalculada: formaPagamento, mesDesp });

        totalDespesasAnual += valor; 
        resumoDespesasPorCategoria[categoria] = (resumoDespesasPorCategoria[categoria] || 0) + valor;
        resumoDespesasPorBanco[banco] = (resumoDespesasPorBanco[banco] || 0) + valor;
        resumoDespesasPorFormaPagamento[formaPagamento] = (resumoDespesasPorFormaPagamento[formaPagamento] || 0) + valor;

        totalDespesasMes[mesDesp] = (totalDespesasMes[mesDesp] || 0) + valor;

        if (!resumoDespesasPorCategoriaMes[mesDesp]) {
            resumoDespesasPorCategoriaMes[mesDesp] = {};
        }
        resumoDespesasPorCategoriaMes[mesDesp][categoria] = 
            (resumoDespesasPorCategoriaMes[mesDesp][categoria] || 0) + valor;
      }
    });

    // B. Filtragem e Agregação de Arrecadações
    arrecadacoesBase.forEach((item) => {
      const { ano: anoArr, mes: mesArr } = extrairAnoEMes(item, 'receita');
      
      if (anoArr === anoFiltro && mesArr > 0) {
        const valor = parseFloat(item.valor) || 0;
        const tipo = item.tipo_arrecadacao || "Sem Tipo";
        const banco = item.banco || item.Banco || item.nome_banco || "-";
        const formaPagamento = item.forma_pagamento || item.formaPagamento || item.Forma_Pagamento || "-";
        const pagador = item.pagador || "-";

        arrecadacoesFiltradasAno.push({ 
          ...item, 
          bancoCalculado: banco, 
          formaPagamentoCalculada: formaPagamento, 
          pagadorCalculado: pagador,
          mesArr 
        });

        totalArrecadacaoAnual += valor; 
        resumoArrecadacaoPorTipo[tipo] = (resumoArrecadacaoPorTipo[tipo] || 0) + valor;
        totalArrecadacaoMes[mesArr] = (totalArrecadacaoMes[mesArr] || 0) + valor;

        if (!resumoArrecadacaoPorTipoMes[mesArr]) {
            resumoArrecadacaoPorTipoMes[mesArr] = {};
        }
        resumoArrecadacaoPorTipoMes[mesArr][tipo] = 
            (resumoArrecadacaoPorTipoMes[mesArr][tipo] || 0) + valor;
      }
    });

    // ORDENAÇÃO CRONOLÓGICA DAS LISTAS FILTRADAS
    despesasFiltradasAno.sort((a, b) => extrairTimestamp(a, 'despesa') - extrairTimestamp(b, 'despesa'));
    arrecadacoesFiltradasAno.sort((a, b) => extrairTimestamp(a, 'receita') - extrairTimestamp(b, 'receita'));

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

    // --- 3.1 ABA 1: Balanço Anual ---
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
      const nomeMesCurto = mesesNomes[mes - 1];
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

    // --- 3.2 ABA 2: Despesas por Categoria, Banco e Forma de Pagamento ---
    const sheet2 = workbook.addWorksheet("Despesas Resumo (Anual)");
    sheet2.columns = [
      { header: "Categoria / Banco / Forma Pagto", key: "item", width: 35 },
      { header: "Total (R$)", key: "total", width: 20 }
    ];

    sheet2.addRow({ item: "--- POR CATEGORIA ---", total: "" }).font = { bold: true };
    Object.keys(resumoDespesasPorCategoria).forEach(cat => {
      sheet2.addRow({ item: cat, total: resumoDespesasPorCategoria[cat] });
    });

    sheet2.addRow({ item: "--- POR BANCO ---", total: "" }).font = { bold: true };
    Object.keys(resumoDespesasPorBanco).forEach(banco => {
      sheet2.addRow({ item: banco, total: resumoDespesasPorBanco[banco] });
    });

    sheet2.addRow({ item: "--- POR FORMA DE PAGAMENTO ---", total: "" }).font = { bold: true };
    Object.keys(resumoDespesasPorFormaPagamento).forEach(forma => {
      sheet2.addRow({ item: forma, total: resumoDespesasPorFormaPagamento[forma] });
    });

    sheet2.addRow({ item: "TOTAL DESPESAS ANUAL", total: totalDespesasAnual }).font = { bold: true };
    sheet2.getColumn("total").numFmt = currencyFormat;

    // --- 3.3 ABA 3: Arrecadação por Tipo ---
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

    // --- 3.4 ABA 4: Despesas Mês-Categoria ---
    const sheet4 = workbook.addWorksheet("Despesas Mês-Categoria");
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
    
    sheet4.getRow(3).values = columns4.map(c => c.header);
    sheet4.getRow(3).eachCell(cell => Object.assign(cell, headerStyle));

    todasCategorias.forEach(categoria => {
        const row: Record<string, string | number> = { 
          categoria: categoria, 
          totalAnual: resumoDespesasPorCategoria[categoria] 
        };
        for (let m = 1; m <= 12; m++) {
            row[`m${m}`] = resumoDespesasPorCategoriaMes[m]?.[categoria] || 0;
        }
        sheet4.addRow(row);
    });

    const totalRowData4: Record<string, string | number> = { categoria: "TOTAL POR MÊS" };
    for (let m = 1; m <= 12; m++) {
        totalRowData4[`m${m}`] = totalDespesasMes[m] || 0;
    }
    totalRowData4.totalAnual = totalDespesasAnual;
    const totalRow4 = sheet4.addRow(totalRowData4);
    totalRow4.font = { bold: true };

    for (let i = 2; i <= 14; i++) {
        sheet4.getColumn(i).numFmt = currencyFormat;
    }

    // --- 3.5 ABA 5: Arrecadação Mês-Tipo ---
    const sheet5 = workbook.addWorksheet("Arrecadação Mês-Tipo");
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
    
    sheet5.getRow(3).values = columns5.map(c => c.header);
    sheet5.getRow(3).eachCell(cell => Object.assign(cell, headerStyle));

    todosTipos.forEach(tipo => {
        const row: Record<string, string | number> = { 
          tipo: tipo, 
          totalAnual: resumoArrecadacaoPorTipo[tipo] 
        };
        for (let m = 1; m <= 12; m++) {
            row[`m${m}`] = resumoArrecadacaoPorTipoMes[m]?.[tipo] || 0;
        }
        sheet5.addRow(row);
    });

    const totalRowData5: Record<string, string | number> = { tipo: "TOTAL POR MÊS" };
    for (let m = 1; m <= 12; m++) {
        totalRowData5[`m${m}`] = totalArrecadacaoMes[m] || 0;
    }
    totalRowData5.totalAnual = totalArrecadacaoAnual;
    const totalRow5 = sheet5.addRow(totalRowData5);
    totalRow5.font = { bold: true };

    for (let i = 2; i <= 14; i++) {
        sheet5.getColumn(i).numFmt = currencyFormat;
    }

    // --- 3.6 ABA 6: Detalhamento Anual de Despesas ---
    const sheet6 = workbook.addWorksheet("Detalhamento Despesas");
    const headers6 = [
      { header: "Data", key: "data", width: 15 },
      { header: "Fornecedor", key: "fornecedor", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Categoria", key: "categoria", width: 25 },
      { header: "Banco", key: "banco", width: 20 },
      { header: "Forma de Pagamento", key: "forma_pagamento", width: 22 },
      { header: "Valor (R$)", key: "valor", width: 15 },
    ];
    sheet6.columns = headers6;

    sheet6.mergeCells("A1:G1");
    sheet6.getCell("A1").value = `Detalhamento Completo de Despesas - ${anoFiltro}`;
    sheet6.getCell("A1").font = { bold: true, size: 14 };
    sheet6.getCell("A1").alignment = { horizontal: "center" };

    sheet6.getRow(2).values = headers6.map(h => h.header);
    sheet6.getRow(2).eachCell(cell => Object.assign(cell, headerStyle));

    despesasFiltradasAno.forEach((item) => {
      sheet6.addRow({
        data: formatarData(item, 'despesa'),
        fornecedor: item.fornecedor || "-",
        descricao: item.descricao || "-",
        categoria: item.despesa_categoria || item.categoria_pagamento || "-",
        banco: item.bancoCalculado,
        forma_pagamento: item.formaPagamentoCalculada,
        valor: parseFloat(item.valor) || 0,
      });
    });

    const totalRow6 = sheet6.addRow({
      descricao: "TOTAL ANUAL",
      valor: totalDespesasAnual,
    });
    totalRow6.font = { bold: true };
    sheet6.getColumn("valor").numFmt = currencyFormat;

    // --- 3.7 ABA 7: Detalhamento Anual de Arrecadações ---
    const sheet7 = workbook.addWorksheet("Detalhamento Arrecadação");
    const headers7 = [
      { header: "Data", key: "data", width: 15 },
      { header: "Pagador", key: "pagador", width: 25 },
      { header: "Tipo de Arrecadação", key: "tipo", width: 25 },
      { header: "Descrição", key: "descricao", width: 35 },
      { header: "Banco", key: "banco", width: 20 },
      { header: "Valor (R$)", key: "valor", width: 15 },
    ];
    sheet7.columns = headers7;

    sheet7.mergeCells("A1:F1");
    sheet7.getCell("A1").value = `Detalhamento Completo de Arrecadação - ${anoFiltro}`;
    sheet7.getCell("A1").font = { bold: true, size: 14 };
    sheet7.getCell("A1").alignment = { horizontal: "center" };

    sheet7.getRow(2).values = headers7.map(h => h.header);
    sheet7.getRow(2).eachCell(cell => Object.assign(cell, headerStyle));

    arrecadacoesFiltradasAno.forEach((item) => {
      sheet7.addRow({
        data: formatarData(item, 'receita'),
        pagador: item.pagadorCalculado,
        tipo: item.tipo_arrecadacao || "-",
        descricao: item.descricao || "-",
        banco: item.bancoCalculado,
        valor: parseFloat(item.valor) || 0,
      });
    });

    const totalRow7 = sheet7.addRow({
      descricao: "TOTAL ANUAL",
      valor: totalArrecadacaoAnual,
    });
    totalRow7.font = { bold: true };
    sheet7.getColumn("valor").numFmt = currencyFormat;

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

router.post("/despesa-diaria-confronto", async (req: Request, res: Response) => {
  try {
    // 1. ACEITA 'dados' OU 'despesas' NO CORPO DA REQUISIÇÃO
    const bodyDados = req.body.dados || req.body.despesas;
    const bodyReceitas = req.body.receitas || req.body.arrecadacao;

    const listaDespesasBruta = Array.isArray(bodyDados) ? bodyDados : [];
    const listaReceitasBruta = Array.isArray(bodyReceitas) ? bodyReceitas : [];

    if (listaDespesasBruta.length === 0 && listaReceitasBruta.length === 0) {
      return res.status(400).json({
        error: "É necessário fornecer ao menos despesas ou receitas.",
      });
    }

    // Helper para buscar a string de data disponível no objeto
    const obterDataBruta = (item: any, tipo: 'despesa' | 'receita') => {
      if (tipo === 'despesa') {
        return item.date_despesa || item.data || item.created_at || "";
      } else {
        return item.data_recebido || item.date_receita || item.date_arrecadacao || item.data || item.created_at || "";
      }
    };

    // Helper robusto para converter qualquer formato de data em Timestamp UTC para ordenação
    const extrairTimestamp = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return 0;

      // 1. Tenta formato ISO (Ex: "2026-09-16" ou "2026-09-16T14:30:00")
      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        const y = parseInt(matchISO[1], 10);
        const m = parseInt(matchISO[2], 10);
        const d = parseInt(matchISO[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      // 2. Tenta formato BR (Ex: "16/09/2026")
      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        const d = parseInt(matchBR[1], 10);
        const m = parseInt(matchBR[2], 10);
        const y = parseInt(matchBR[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      return 0;
    };

    // Helper para formatar a data na exibição do Excel (DD/MM/YYYY)
    const formatarData = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return "-";

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        return `${matchISO[3]}/${matchISO[2]}/${matchISO[1]}`;
      }

      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        return `${matchBR[1]}/${matchBR[2]}/${matchBR[3]}`;
      }

      return "-";
    };

    // ────────────────────────────────────────────
    // 🔍 ORDENAR AS LISTAS POR DATA
    // ────────────────────────────────────────────
    const listaDespesas = listaDespesasBruta.sort((a, b) => {
      const timeA = extrairTimestamp(a, 'despesa');
      const timeB = extrairTimestamp(b, 'despesa');
      return timeA - timeB;
    });

    const listaReceitas = listaReceitasBruta.sort((a, b) => {
      const timeA = extrairTimestamp(a, 'receita');
      const timeB = extrairTimestamp(b, 'receita');
      return timeA - timeB;
    });

    // ────────────────────────────────────────────
    // 🎯 PREPARAÇÃO DA DATA PARA TÍTULOS E ARQUIVO
    // ────────────────────────────────────────────
    const primeiroItem = listaDespesas[0] || listaReceitas[0];
    const tipoPrimeiro = listaDespesas[0] ? 'despesa' : 'receita';
    
    let dataFormatadaTitulo = "Geral";
    let nomeArquivo = "data_indefinida";

    if (primeiroItem) {
      const dataFormatada = formatarData(primeiroItem, tipoPrimeiro);
      if (dataFormatada !== "-") {
        dataFormatadaTitulo = dataFormatada;
        const [d, m, y] = dataFormatada.split("/");
        nomeArquivo = `${d}_${m}_${y}`;
      }
    }

    // ────────────────────────────────────────────
    // 📊 CRIAR WORKBOOK E CONFIGURAR COLUNAS
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Fluxo de Caixa");

    sheet.columns = [
      { key: "colA", width: 15 },
      { key: "colB", width: 25 },
      { key: "colC", width: 35 },
      { key: "colD", width: 25 },
      { key: "colE", width: 20 },
      { key: "colF", width: 22 },
      { key: "colG", width: 15 },
    ];

    const aplicarEstiloCabecalho = (row: ExcelJS.Row) => {
      row.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFDDDDDD" },
        };
      });
    };

    // ────────────────────────────────────────────
    // 1. SEÇÃO DE DESPESAS
    // ────────────────────────────────────────────
    sheet.mergeCells("A1:G1");
    const titDespesas = sheet.getCell("A1");
    titDespesas.value = `Despesas do Dia ${dataFormatadaTitulo}`;
    titDespesas.font = { bold: true, size: 16 };
    titDespesas.alignment = { horizontal: "center" };

    const headerDespesas = sheet.addRow([
      "Data", "Fornecedor", "Descrição", "Categoria", "Banco", "Forma de Pagamento", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerDespesas);

    let totalDespesas = 0;

    listaDespesas.forEach((item) => {
      const valorItem = Number(item.valor) || 0;
      totalDespesas += valorItem;

      // Fallbacks dos nomes das colunas conforme o seu JSON da tabela despesas
      const categoria = item.categoria_pagamento || item.despesa_categoria || item.categoria || "-";
      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || "-";

      sheet.addRow([
        formatarData(item, 'despesa'),
        item.fornecedor || "-",
        item.descricao || "-",
        categoria,
        bancoValor !== "" ? bancoValor : "-",
        formaPagamentoValor !== "" ? formaPagamentoValor : "-",
        valorItem
      ]);
    });

    const rowTotalDespesas = sheet.addRow(["", "", "TOTAL", "", "", "", totalDespesas]);
    rowTotalDespesas.font = { bold: true };

    sheet.addRow([]);

    // ────────────────────────────────────────────
    // 2. SEÇÃO DE RECEITAS / ARRECADAÇÃO
    // ────────────────────────────────────────────
    const linhaInicioReceitas = sheet.lastRow!.number + 1;
    sheet.mergeCells(`A${linhaInicioReceitas}:G${linhaInicioReceitas}`);
    const titReceitas = sheet.getCell(`A${linhaInicioReceitas}`);
    titReceitas.value = `Receitas ${dataFormatadaTitulo}`;
    titReceitas.font = { bold: true, size: 16 };
    titReceitas.alignment = { horizontal: "center" };

    const headerReceitas = sheet.addRow([
      "Data", "Pagador", "Tipo de arrecadação", "Descrição", "Banco", "", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerReceitas);

    let totalReceitas = 0;

    listaReceitas.forEach((item) => {
      const valorItem = Number(item.valor) || 0;
      totalReceitas += valorItem;

      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const pagadorValor = item.pagador || "-";

      sheet.addRow([
        formatarData(item, 'receita'),
        pagadorValor,
        item.tipo_arrecadacao || "-",
        item.descricao || "-",
        bancoValor !== "" ? bancoValor : "-",
        "",
        valorItem
      ]);
    });

    const rowTotalReceitas = sheet.addRow(["", "", "TOTAL", "", "", "", totalReceitas]);
    rowTotalReceitas.font = { bold: true };

    sheet.addRow([]);

    // ────────────────────────────────────────────
    // 3. SEÇÃO DE RESULTADO (DESPESAS - RECEITAS)
    // ────────────────────────────────────────────
    const resultadoFinal = totalReceitas - totalDespesas;
    const rowResultado = sheet.addRow(["", "", "Resultado", "", "", "", resultadoFinal]);
    rowResultado.font = { bold: true };

    // Formatação de Moeda na Coluna G
    sheet.getColumn("G").numFmt = "R$ #,##0.00";

    // ────────────────────────────────────────────
    // 📤 RETORNO DO ARQUIVO
    // ────────────────────────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=relatorio_financeiro_${nomeArquivo}.xlsx`
    );

    return res.send(buffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório financeiro." });
  }
});
router.post("/despesa-mensal-confronto", async (req: Request, res: Response) => {
  try {
    const { despesas, receitas, data_inicio } = req.body;

    if (!data_inicio) {
      return res.status(400).json({
        error: "A propriedade data_inicio é obrigatória para filtrar o mês.",
      });
    }

    // ────────────────────────────────────────────
    // 🎯 DETERMINAR MÊS E ANO DE REFERÊNCIA
    // ────────────────────────────────────────────
    const dataFiltroObj = new Date(data_inicio);
    const mesAlvo = dataFiltroObj.getUTCMonth(); // 0 a 11
    const anoAlvo = dataFiltroObj.getUTCFullYear();

    const nomesMeses = [
      "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
      "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ];
    const nomeMesExtenso = `${nomesMeses[mesAlvo]} / ${anoAlvo}`;
    const nomeArquivo = `${mesAlvo + 1}_${anoAlvo}`;

    // Helper para buscar a string de data disponível no objeto
    const obterDataBruta = (item: any, tipo: 'despesa' | 'receita') => {
      if (tipo === 'despesa') {
        return item.date_despesa || item.data || item.created_at || "";
      } else {
        return item.data_recebido || item.date_receita || item.date_arrecadacao || item.data || item.created_at || "";
      }
    };

    // Helper robusto para converter qualquer formato de data em Timestamp UTC
    const extrairTimestamp = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return 0;

      // 1. Tenta formato ISO (Ex: "2026-09-16" ou "2026-09-16T14:30:00")
      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        const y = parseInt(matchISO[1], 10);
        const m = parseInt(matchISO[2], 10);
        const d = parseInt(matchISO[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      // 2. Tenta formato BR (Ex: "16/09/2026")
      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        const d = parseInt(matchBR[1], 10);
        const m = parseInt(matchBR[2], 10);
        const y = parseInt(matchBR[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      return 0;
    };

    // Helper para formatar a data na exibição do Excel (DD/MM/YYYY)
    const formatarData = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return "-";

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        return `${matchISO[3]}/${matchISO[2]}/${matchISO[1]}`;
      }

      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        return `${matchBR[1]}/${matchBR[2]}/${matchBR[3]}`;
      }

      return "-";
    };

    // Helper para verificar se a data pertence ao Mês/Ano alvo
    const pertenceAoMes = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return false;

      let ano = 0;
      let mes = 0;

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        ano = parseInt(matchISO[1], 10);
        mes = parseInt(matchISO[2], 10);
      } else {
        const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (matchBR) {
          mes = parseInt(matchBR[2], 10);
          ano = parseInt(matchBR[3], 10);
        }
      }

      return ano === anoAlvo && (mes - 1) === mesAlvo;
    };

    // ────────────────────────────────────────────
    // 🔍 FILTRAR E ORDENAR OS DADOS DO MÊS POR DATA
    // ────────────────────────────────────────────
    const listaDespesasBruta = Array.isArray(despesas) ? despesas : [];
    const listaReceitasBruta = Array.isArray(receitas) ? receitas : [];

    // 1. Filtra despesas do mês e ordena do menor para o maior dia
    const listaDespesas = listaDespesasBruta
      .filter((d) => pertenceAoMes(d, 'despesa'))
      .sort((a, b) => {
        const timeA = extrairTimestamp(a, 'despesa');
        const timeB = extrairTimestamp(b, 'despesa');
        return timeA - timeB; // Ordem crescente
      });

    // 2. Filtra receitas do mês e ordena do menor para o maior dia
    const listaReceitas = listaReceitasBruta
      .filter((r) => pertenceAoMes(r, 'receita'))
      .sort((a, b) => {
        const timeA = extrairTimestamp(a, 'receita');
        const timeB = extrairTimestamp(b, 'receita');
        return timeA - timeB; // Ordem crescente
      });

    if (listaDespesas.length === 0 && listaReceitas.length === 0) {
      return res.status(400).json({
        error: `Nenhuma despesa ou receita encontrada para o mês de ${nomeMesExtenso}.`,
      });
    }

    // ────────────────────────────────────────────
    // 📊 CRIAR EXCEL E DEFINIR COLUNAS
    // ────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Relatório Mensal");

    sheet.columns = [
      { key: "colA", width: 15 },
      { key: "colB", width: 25 },
      { key: "colC", width: 35 },
      { key: "colD", width: 25 },
      { key: "colE", width: 20 },
      { key: "colF", width: 22 },
      { key: "colG", width: 15 },
    ];

    const aplicarEstiloCabecalho = (row: ExcelJS.Row) => {
      row.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFDDDDDD" },
        };
      });
    };

    // ────────────────────────────────────────────
    // 1. SEÇÃO DE DESPESAS DO MÊS
    // ────────────────────────────────────────────
    sheet.mergeCells("A1:G1");
    const titDespesas = sheet.getCell("A1");
    titDespesas.value = `Despesas do Mês: ${nomeMesExtenso}`;
    titDespesas.font = { bold: true, size: 16 };
    titDespesas.alignment = { horizontal: "center" };

    const headerDespesas = sheet.addRow([
      "Data", "Fornecedor", "Descrição", "Categoria", "Banco", "Forma de Pagamento", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerDespesas);

    let totalDespesas = 0;

    listaDespesas.forEach((item) => {
      const valorItem = Number(item.valor) || 0;
      totalDespesas += valorItem;

      const categoria = item.categoria_pagamento || item.despesa_categoria || item.categoria || "-";
      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || "-";

      sheet.addRow([
        formatarData(item, 'despesa'),
        item.fornecedor || "-",
        item.descricao || "-",
        categoria,
        bancoValor !== "" ? bancoValor : "-",
        formaPagamentoValor !== "" ? formaPagamentoValor : "-",
        valorItem
      ]);
    });

    const rowTotalDespesas = sheet.addRow(["", "", "TOTAL DESPESAS", "", "", "", totalDespesas]);
    rowTotalDespesas.font = { bold: true };

    sheet.addRow([]); // Espaçamento

    // ────────────────────────────────────────────
    // 2. SEÇÃO DE RECEITAS DO MÊS
    // ────────────────────────────────────────────
    const linhaInicioReceitas = sheet.lastRow!.number + 1;
    sheet.mergeCells(`A${linhaInicioReceitas}:G${linhaInicioReceitas}`);
    const titReceitas = sheet.getCell(`A${linhaInicioReceitas}`);
    titReceitas.value = `Receitas do Mês: ${nomeMesExtenso}`;
    titReceitas.font = { bold: true, size: 16 };
    titReceitas.alignment = { horizontal: "center" };

    const headerReceitas = sheet.addRow([
      "Data", "Pagador", "Tipo de Arrecadação", "Descrição", "Banco", "", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerReceitas);

    let totalReceitas = 0;

    listaReceitas.forEach((item) => {
      const valorItem = Number(item.valor) || 0;
      totalReceitas += valorItem;

      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const pagadorValor = item.pagador || "-";

      sheet.addRow([
        formatarData(item, 'receita'),
        pagadorValor,
        item.tipo_arrecadacao || "-",
        item.descricao || "-",
        bancoValor !== "" ? bancoValor : "-",
        "",
        valorItem
      ]);
    });

    const rowTotalReceitas = sheet.addRow(["", "", "TOTAL RECEITAS", "", "", "", totalReceitas]);
    rowTotalReceitas.font = { bold: true };

    sheet.addRow([]); // Espaçamento

    // ────────────────────────────────────────────
    // 3. RESULTADO DO MÊS (DESPESAS - RECEITAS)
    // ────────────────────────────────────────────
    const resultadoFinal = totalReceitas - totalDespesas;
    const rowResultado = sheet.addRow(["", "", "Resultado do Mês", "", "", "", resultadoFinal]);
    rowResultado.font = { bold: true };

    // Formatação de moeda na coluna de valores (G)
    sheet.getColumn("G").numFmt = "R$ #,##0.00";

    // ────────────────────────────────────────────
    // 📤 ENVIO DO EXCEL
    // ────────────────────────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=relatorio_mensal_${nomeArquivo}.xlsx`
    );

    return res.send(buffer);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório mensal de despesas e receitas." });
  }
});
router.post("/despesa-mensal-confronto-classificado", async (req: Request, res: Response) => {
  try {
    const { dados, despesas, receitas, data_inicio } = req.body;

    if (!data_inicio) {
      return res.status(400).json({ error: "data_inicio é obrigatória." });
    }

    const listaDespesasBruta = Array.isArray(despesas) ? despesas : Array.isArray(dados) ? dados : [];
    const listaReceitasBruta = Array.isArray(receitas) ? receitas : [];

    if (listaDespesasBruta.length === 0 && listaReceitasBruta.length === 0) {
      return res.status(400).json({
        error: "É necessário fornecer despesas ou receitas.",
      });
    }

    // Extrair ano e mês de data_inicio (formato esperado: YYYY-MM-DD)
    const dataISO = data_inicio.substring(0, 10);
    const partesFiltro = dataISO.split("-").map(Number);
    const anoFiltro = partesFiltro[0];
    const mesFiltro = partesFiltro[1]; // 1–12

    const dataParaNomeMes = new Date(Date.UTC(anoFiltro, mesFiltro - 1, 1));
    const nomeMes = dataParaNomeMes.toLocaleString("pt-BR", { month: "long" });

    // Helper para buscar a string de data disponível no objeto
    const obterDataBruta = (item: any, tipo: 'despesa' | 'receita') => {
      if (tipo === 'despesa') {
        return item.date_despesa || item.data || item.created_at || "";
      } else {
        return item.data_recebido || item.date_receita || item.date_arrecadacao || item.data || item.created_at || "";
      }
    };

    // Helper robusto para converter qualquer formato de data em Timestamp UTC para ordenação
    const extrairTimestamp = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return 0;

      // 1. Tenta formato ISO (Ex: "2026-09-16" ou "2026-09-16T14:30:00")
      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        const y = parseInt(matchISO[1], 10);
        const m = parseInt(matchISO[2], 10);
        const d = parseInt(matchISO[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      // 2. Tenta formato BR (Ex: "16/09/2026")
      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        const d = parseInt(matchBR[1], 10);
        const m = parseInt(matchBR[2], 10);
        const y = parseInt(matchBR[3], 10);
        const ts = new Date(Date.UTC(y, m - 1, d)).getTime();
        return isNaN(ts) ? 0 : ts;
      }

      return 0;
    };

    // Helper para formatar a data na exibição do Excel (DD/MM/YYYY)
    const formatarData = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return "-";

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        return `${matchISO[3]}/${matchISO[2]}/${matchISO[1]}`;
      }

      const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchBR) {
        return `${matchBR[1]}/${matchBR[2]}/${matchBR[3]}`;
      }

      return "-";
    };

    // Helper para verificar se a data pertence ao Mês/Ano selecionado
    const pertenceAoMes = (item: any, tipo: 'despesa' | 'receita') => {
      const dataStr = String(obterDataBruta(item, tipo) || "").trim();
      if (!dataStr) return false;

      let ano = 0;
      let mes = 0;

      const matchISO = dataStr.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (matchISO) {
        ano = parseInt(matchISO[1], 10);
        mes = parseInt(matchISO[2], 10);
      } else {
        const matchBR = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (matchBR) {
          mes = parseInt(matchBR[2], 10);
          ano = parseInt(matchBR[3], 10);
        }
      }

      return ano === anoFiltro && mes === mesFiltro;
    };

    // ────────────────────────────────────────────
    // 🔍 FILTRAR E ORDENAR OS DADOS DO MÊS POR DATA
    // ────────────────────────────────────────────
    const despesasFiltradas = listaDespesasBruta
      .filter((d) => pertenceAoMes(d, 'despesa'))
      .sort((a, b) => {
        const timeA = extrairTimestamp(a, 'despesa');
        const timeB = extrairTimestamp(b, 'despesa');
        return timeA - timeB; // Ordem crescente
      });

    const receitasFiltradas = listaReceitasBruta
      .filter((r) => pertenceAoMes(r, 'receita'))
      .sort((a, b) => {
        const timeA = extrairTimestamp(a, 'receita');
        const timeB = extrairTimestamp(b, 'receita');
        return timeA - timeB; // Ordem crescente
      });

    const workbook = new ExcelJS.Workbook();
    
    // ─────────────────────────────────────────────────────────────────
    // 📊 ABA 1: DETALHAMENTO DE DESPESAS E RECEITAS
    // ─────────────────────────────────────────────────────────────────
    const sheet1 = workbook.addWorksheet("Detalhamento");
    sheet1.columns = [
      { key: "colA", width: 15 },
      { key: "colB", width: 25 },
      { key: "colC", width: 35 },
      { key: "colD", width: 20 },
      { key: "colE", width: 20 },
      { key: "colF", width: 22 },
      { key: "colG", width: 15 },
    ];

    const aplicarEstiloCabecalho = (row: ExcelJS.Row) => {
      row.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDDDDD" } };
      });
    };

    // 1. TÍTULO DESPESAS
    sheet1.mergeCells("A1:G1");
    sheet1.getCell("A1").value = `Detalhamento de Despesas - ${nomeMes}/${anoFiltro}`;
    sheet1.getCell("A1").font = { bold: true, size: 14 };
    sheet1.getCell("A1").alignment = { horizontal: "center" };

    const headerDespesas = sheet1.addRow([
      "Data", "Descrição", "Categoria", "Subcategoria", "Banco", "Forma de Pagamento", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerDespesas);

    let totalDespesas = 0;
    const resumoCategorias: Record<string, number> = {};

    despesasFiltradas.forEach((item) => {
      const valor = Number(item.valor) || 0;
      const cat = item.despesa_categoria || item.categoria_pagamento || "Sem Categoria";
      
      totalDespesas += valor;
      resumoCategorias[cat] = (resumoCategorias[cat] || 0) + valor;

      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const formaPagamentoValor = item.forma_pagamento || item.formaPagamento || "-";

      sheet1.addRow([
        formatarData(item, 'despesa'),
        item.descricao || "-",
        cat,
        item.despesa_sub_categoria || "-",
        bancoValor !== "" ? bancoValor : "-",
        formaPagamentoValor !== "" ? formaPagamentoValor : "-",
        valor
      ]);
    });

    const rowTotalDespesas = sheet1.addRow(["", "TOTAL DESPESAS", "", "", "", "", totalDespesas]);
    rowTotalDespesas.font = { bold: true };

    sheet1.addRow([]); // Espaçamento

    // 2. TÍTULO RECEITAS
    const linIniRec = sheet1.lastRow!.number + 1;
    sheet1.mergeCells(`A${linIniRec}:G${linIniRec}`);
    sheet1.getCell(`A${linIniRec}`).value = `Detalhamento de Receitas - ${nomeMes}/${anoFiltro}`;
    sheet1.getCell(`A${linIniRec}`).font = { bold: true, size: 14 };
    sheet1.getCell(`A${linIniRec}`).alignment = { horizontal: "center" };

    const headerReceitas = sheet1.addRow([
      "Data", "Pagador", "Tipo de Arrecadação", "Descrição", "Banco", "", "Valor (R$)"
    ]);
    aplicarEstiloCabecalho(headerReceitas);

    let totalReceitas = 0;
    const resumoArrecadacao: Record<string, number> = {};

    receitasFiltradas.forEach((item) => {
      const valor = Number(item.valor) || 0;
      const tipoArrec = item.tipo_arrecadacao || "Outros / Sem Tipo";

      totalReceitas += valor;
      resumoArrecadacao[tipoArrec] = (resumoArrecadacao[tipoArrec] || 0) + valor;

      const bancoValor = item.banco || item.Banco || item.nome_banco || "-";
      const pagadorValor = item.pagador || "-";

      sheet1.addRow([
        formatarData(item, 'receita'),
        pagadorValor,
        tipoArrec,
        item.descricao || "-",
        bancoValor !== "" ? bancoValor : "-",
        "",
        valor
      ]);
    });

    const rowTotalReceitas = sheet1.addRow(["", "TOTAL RECEITAS", "", "", "", "", totalReceitas]);
    rowTotalReceitas.font = { bold: true };

    sheet1.addRow([]);

    // 3. RESULTADO DO MÊS
    const rowResultado = sheet1.addRow(["", "RESULTADO (RECEITAS - DESPESAS)", "", "", "", "", totalReceitas - totalDespesas]);
    rowResultado.font = { bold: true };

    sheet1.getColumn("G").numFmt = "R$ #,##0.00";

    // ─────────────────────────────────────────────────────────────────
    // 📈 ABA 2: RESUMO POR CATEGORIA E TIPO DE ARRECADAÇÃO
    // ─────────────────────────────────────────────────────────────────
    const sheet2 = workbook.addWorksheet("Resumo Mensal");
    sheet2.columns = [
      { key: "colA", width: 35 },
      { key: "colB", width: 25 }
    ];

    // Resumo Despesas
    sheet2.addRow(["RESUMO DE DESPESAS POR CATEGORIA", ""]).font = { bold: true, size: 12 };
    const hResDesp = sheet2.addRow(["Categoria", "Total Acumulado (R$)"]);
    aplicarEstiloCabecalho(hResDesp);

    Object.keys(resumoCategorias).forEach((cat) => {
      sheet2.addRow([cat, resumoCategorias[cat]]);
    });

    const rowTotResDesp = sheet2.addRow(["TOTAL DESPESAS", totalDespesas]);
    rowTotResDesp.font = { bold: true };

    sheet2.addRow([]); // Espaçamento

    // Resumo Receitas
    sheet2.addRow(["RESUMO DE RECEITAS POR TIPO DE ARRECADAÇÃO", ""]).font = { bold: true, size: 12 };
    const hResRec = sheet2.addRow(["Tipo de Arrecadação", "Total Acumulado (R$)"]);
    aplicarEstiloCabecalho(hResRec);

    Object.keys(resumoArrecadacao).forEach((tipo) => {
      sheet2.addRow([tipo, resumoArrecadacao[tipo]]);
    });

    const rowTotResRec = sheet2.addRow(["TOTAL RECEITAS", totalReceitas]);
    rowTotResRec.font = { bold: true };

    sheet2.addRow([]);

    // Resultado Final na Aba 2
    const rowResFinalAba2 = sheet2.addRow(["RESULTADO DO MÊS", totalReceitas - totalDespesas]);
    rowResFinalAba2.font = { bold: true };

    sheet2.getColumn("B").numFmt = "R$ #,##0.00";

    // Exportação
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=Relatorio_Classificado_${nomeMes}.xlsx`);
    return res.send(buffer);

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório classificado completo." });
  }
});
export default router;

