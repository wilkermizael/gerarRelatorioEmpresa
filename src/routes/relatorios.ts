import { Router, Request, Response } from "express";
import ExcelJS from "exceljs";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import fs from "fs";
import path from "path";
import crypto from "crypto";
import { createClient } from "@supabase/supabase-js"; // 👈 Import do Supabase adicionado

const router = Router();

// Configuração do Supabase Client
const supabaseUrl = process.env.SUPABASE_URL || "SUA_URL_DO_SUPABASE";
const supabaseKey = process.env.SUPABASE_SERVICE_KEY || "";

const supabase = createClient(supabaseUrl, supabaseKey);

// Função aux. para formatar cabeçalhos
const formatHeader = (key: string) => {
  if (key === "status") return "Status (Ativo / Inativo)";
  return key
    .replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, (l) => l.toUpperCase());
};

// Função aux. para remoção segura de arquivos temporários
const cleanupFile = (filePath: string) => {
  if (fs.existsSync(filePath)) {
    fs.unlink(filePath, (err) => {
      if (err)
        console.error(
          `⚠️ Erro ao deletar arquivo temporário ${filePath}:`,
          err,
        );
    });
  }
};

router.get("/mensagem", (req: Request, res: Response) => {
  return res.status(200).json({
    sucesso: true,
    mensagem: "✅ Backend de relatórios funcionando corretamente!",
  });
});

// =========================================================================
// 1. FILTRAR EMPRESAS XLSX
// =========================================================================
router.post("/empresas-filtradas", async (req: Request, res: Response) => {
  console.log("\n=======================================================");
  console.log("🚀 [INÍCIO] Requisição para /relatorios/empresas-filtradas");
  console.log("=======================================================");

  try {
    // 🎯 Desestrutura as variáveis direto da raiz do req.body
    const { escolhaColunaTabela, tipoConvencao, acordo } = req.body;

    console.log("📥 [BODY RECEBIDO]:", JSON.stringify(req.body, null, 2));

    // 1. Inicia a Query base do Supabase
    let query = supabase.from("empresa").select("*");

    // -------------------------------------------------------------
    // 🎯 APLICAÇÃO DOS FILTROS DIRETO DAS VARIÁVEIS
    // -------------------------------------------------------------

    // A) Filtro de Tipo de Convenção (Aceita Array na raiz)
    if (Array.isArray(tipoConvencao) && tipoConvencao.length > 0) {
      console.log("🔍 [FILTRO APLICADO] tipo_convencao IN:", tipoConvencao);
      query = query.in("tipo_convencao", tipoConvencao);
    }

    // B) Filtro de Acordos (Aceita Array na raiz: ex: ["sindical", "negocial"])
    if (Array.isArray(acordo) && acordo.length > 0) {
      console.log("🔍 [FILTRO APLICADO] acordo:", acordo);

      // Mapeia os itens do array para montar a cláusula OR dinâmica do Supabase
      const orConditions = acordo
        .map((tipo: string) => {
          const key = String(tipo).toLowerCase().trim();
          if (key === "sindical" || key === "acordo_sindical") return "acordo_sindical.eq.true";
          if (key === "negocial" || key === "acordo_negocial") return "acordo_negocial.eq.true";
          if (key === "mensalidade" || key === "acordo_mensalidade") return "acordo_mensalidade.eq.true";
          return null;
        })
        .filter(Boolean);

      if (orConditions.length > 0) {
        query = query.or(orConditions.join(","));
      }
    }

    // Ordenação alfabética por Razão Social
    query = query.order("razao_social", { ascending: true });

    // Executa a busca no Supabase
    const { data: listaEmpresas, error } = await query;

    if (error) {
      console.error("🔥 [ERRO SUPABASE]:", error);
      return res.status(500).json({ error: "Erro ao consultar banco de dados." });
    }

    console.log(`📊 [SUPABASE] Total de empresas filtradas encontradas: ${listaEmpresas?.length ?? 0}`);

    if (!listaEmpresas || listaEmpresas.length === 0) {
      return res.status(404).json({ error: "Nenhuma empresa encontrada com os filtros selecionados." });
    }

    // -------------------------------------------------------------
    // 2. MONTAGEM DAS COLUNAS DA TABELA
    // -------------------------------------------------------------
    const mapaCampos: Record<string, string | string[]> = {
      id: "id",
      createdAt: "created_at",
      dataFundacao: "data_fundacao",
      cnpj: "cnpj",
      razaoSocial: "razao_social",
      razao_social: "razao_social",
      nomeFantasia: "nome_fantasia",
      situacao: "situacao",
      status: "status",
      tipo: "tipo",
      classificacao: "classificacao",
      cnae: "cnae",
      categoria: "categoria",
      numero_funcionario: "numero_funcionario",
      tipo_convencao: "tipo_convencao",

      email: "email01",
      email01: "email01",
      email02: "email02",

      acordo_sindical: "acordo_sindical",
      acordo_negocial: "acordo_negocial",
      acordo_mensalidade: "acordo_mensalidade",

      telefone01: "telefone01",
      telefone02: "telefone02",
      nome_responsavel01: "nome_responsavel01",
      nome_responsavel02: "nome_responsavel02",
      rua: "rua",
      numero: "numero",
      bairro: "bairro",
      cidade: "cidade",
      estado: "estado",
      cep: "cep"
    };

    let colunasBanco: string[] = [];

    if (
      escolhaColunaTabela &&
      typeof escolhaColunaTabela === "object" &&
      Object.keys(escolhaColunaTabela).length > 0
    ) {
      const colunasAtivas = Object.keys(escolhaColunaTabela).filter((key) => {
        const val = escolhaColunaTabela[key];
        return val === true || String(val).toLowerCase() === "true";
      });

      colunasAtivas.forEach((key) => {
        if (key === "acordo") {
          // 🎯 Se a chave for "acordo", exibe apenas as colunas dos acordos selecionados no filtro.
          // Se nenhum foi selecionado no filtro, exibe os 3 por padrão.
          if (Array.isArray(acordo) && acordo.length > 0) {
            acordo.forEach((tipo: string) => {
              const k = String(tipo).toLowerCase().trim();
              if ((k === "sindical" || k === "acordo_sindical") && !colunasBanco.includes("acordo_sindical")) {
                colunasBanco.push("acordo_sindical");
              }
              if ((k === "negocial" || k === "acordo_negocial") && !colunasBanco.includes("acordo_negocial")) {
                colunasBanco.push("acordo_negocial");
              }
              if ((k === "mensalidade" || k === "acordo_mensalidade") && !colunasBanco.includes("acordo_mensalidade")) {
                colunasBanco.push("acordo_mensalidade");
              }
            });
          } else {
            ["acordo_sindical", "acordo_negocial", "acordo_mensalidade"].forEach((col) => {
              if (!colunasBanco.includes(col)) colunasBanco.push(col);
            });
          }
        } else {
          const mapeado = mapaCampos[key] || key;
          if (Array.isArray(mapeado)) {
            mapeado.forEach((col) => {
              if (!colunasBanco.includes(col)) colunasBanco.push(col);
            });
          } else {
            if (!colunasBanco.includes(mapeado)) colunasBanco.push(mapeado);
          }
        }
      });
    }

    // -------------------------------------------------------------
    // 3. GERAÇÃO DO ARQUIVO EXCEL
    // -------------------------------------------------------------
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Empresas Filtradas");

    // Título Principal
    sheet.addRow(["SENALBA MG - Relatório de Empresas (Filtrado)"]);
    sheet.mergeCells(1, 1, 1, Math.max(colunasBanco.length, 1));
    const titleRow = sheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    titleRow.alignment = { horizontal: "center", vertical: "middle" };
    titleRow.height = 30;

    sheet.addRow([]); // Espaçamento

    const formataCabecalho = (texto: string) =>
      texto.replace(/_/g, " ").replace(/\b\w/g, (l) => l.toUpperCase());

    // Cabeçalho da Tabela
    const headerRow = sheet.addRow(colunasBanco.map((c) => formataCabecalho(c)));
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Formatação de valores
    const formatarValor = (col: string, valor: any): string => {
      if (valor === null || valor === undefined) return "";
      const camposBooleanos = ["status", "acordo_sindical", "acordo_negocial", "acordo_mensalidade"];
      if (camposBooleanos.includes(col)) {
        return valor === true || String(valor).toLowerCase() === "true" ? "Sim" : "Não";
      }
      return String(valor);
    };

    // Preenche as linhas
    listaEmpresas.forEach((empresa: any) => {
      const linha = colunasBanco.map((col) => formatarValor(col, empresa[col]));
      sheet.addRow(linha);
    });

    // Ajusta a largura das colunas
    colunasBanco.forEach((col, i) => {
      const cabecalhoLen = formataCabecalho(col).length;
      const maxDataLen = Math.max(
        0,
        ...listaEmpresas.map((item: any) => formatarValor(col, item[col]).length)
      );
      const width = Math.max(cabecalhoLen, maxDataLen);
      sheet.getColumn(i + 1).width = Math.min(Math.max(width * 0.9 + 2, 12), 40);
    });

    // Resposta HTTP
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=empresas_filtradas_${Date.now()}.xlsx`
    );

    await workbook.xlsx.write(res);
    console.log("🎉 [SUCESSO] Relatório filtrado gerado e enviado!");
    return res.end();

  } catch (error) {
    console.error("🔥 [ERRO CRÍTICO INTERNO]:", error);
    return res.status(500).json({ error: "Erro interno ao gerar relatório." });
  }
});
// =========================================================================
// 1.1. EMPRESAS XLSX
// =========================================================================

router.post("/empresas", async (req: Request, res: Response) => {
  console.log("\n=======================================================");
  console.log("🚀 [INÍCIO] Requisição para /relatorios/empresas recebida!");
  console.log("=======================================================");

  try {
    const { escolhaColunaTabela } = req.body;

    console.log("📥 [BODY RECEBIDO]:", JSON.stringify(req.body, null, 2));

    // 1. Busca as empresas no Supabase
    const { data: listaEmpresas, error } = await supabase
      .from("empresa")
      .select("*")
      .order("razao_social", { ascending: true });

    if (error) {
      console.error("🔥 [ERRO SUPABASE]:", error);
      return res.status(500).json({ error: "Erro ao consultar banco de dados." });
    }

    if (!listaEmpresas || listaEmpresas.length === 0) {
      return res.status(404).json({ error: "Nenhuma empresa encontrada no banco." });
    }

    // 2. Mapeamento de chaves do FlutterFlow -> Coluna(s) reais do Supabase
    const mapaCampos: Record<string, string | string[]> = {
      id: "id",
      createdAt: "created_at",
      created_at: "created_at",
      dataFundacao: "data_fundacao",
      data_fundacao: "data_fundacao",
      cnpj: "cnpj",
      razaoSocial: "razao_social",
      razao_social: "razao_social",
      nomeFantasia: "nome_fantasia",
      nome_fantasia: "nome_fantasia",
      situacao: "situacao",
      status: "status",
      tipo: "tipo",
      classificacao: "classificacao",
      cnae: "cnae",
      categoria: "categoria",
      idCategoria: "id_categoria",
      id_categoria: "id_categoria",
      numeroFuncionario: "numero_funcionario",
      numeroFuncionarios: "numero_funcionario",
      numero_funcionario: "numero_funcionario",
      tipoConvencao: "tipo_convencao",
      tipo_convencao: "tipo_convencao",

      // 🎯 Mapeamento especial do Email (FlutterFlow "email" -> Supabase "email01")
      email: "email01",
      email01: "email01",
      email02: "email02",

      // 🎯 Mapeamento especial do Acordo (Expande em todos os acordos)
      acordo: ["acordo_sindical", "acordo_negocial", "acordo_mensalidade"],
      acordo_sindical: "acordo_sindical",
      acordo_negocial: "acordo_negocial",
      acordo_mensalidade: "acordo_mensalidade",

      telefone01: "telefone01",
      telefone02: "telefone02",
      nomeResponsavel01: "nome_responsavel01",
      nome_responsavel01: "nome_responsavel01",
      nomeResponsavel02: "nome_responsavel02",
      nome_responsavel02: "nome_responsavel02",
      nomeContabilidade: "nome_contabilidade",
      nome_contabilidade: "nome_contabilidade",
      emailContabilidade: "email_contabilidade",
      email_contabilidade: "email_contabilidade",
      telefoneContabilidade: "telefone_contabilidade",
      telefone_contabilidade: "telefone_contabilidade",
      contatoContabilidade: "contato_contabilidade",
      contato_contabilidade: "contato_contabilidade",
      rua: "rua",
      numero: "numero",
      complemento: "complemento",
      bairro: "bairro",
      cidade: "cidade",
      estado: "estado",
      cep: "cep",
      observacoes: "observacoes"
    };

    // 3. Seleciona APENAS as colunas com valor "true" ou true
    let colunasBanco: string[] = [];

    if (
      escolhaColunaTabela &&
      typeof escolhaColunaTabela === "object" &&
      Object.keys(escolhaColunaTabela).length > 0
    ) {
      // Aceita tanto booleano true quanto string "true"
      const colunasAtivas = Object.keys(escolhaColunaTabela).filter((key) => {
        const val = escolhaColunaTabela[key];
        return val === true || String(val).toLowerCase() === "true";
      });

      console.log("🔑 [COLUNAS ATIVAS IDENTIFICADAS]:", colunasAtivas);

      colunasAtivas.forEach((key) => {
        const mapeado = mapaCampos[key] || key;

        if (Array.isArray(mapeado)) {
          // Se for um array (como o 'acordo'), adiciona cada uma das colunas
          mapeado.forEach((col) => {
            if (!colunasBanco.includes(col)) colunasBanco.push(col);
          });
        } else {
          if (!colunasBanco.includes(mapeado)) colunasBanco.push(mapeado);
        }
      });
    }

    console.log("📋 [COLUNAS FINAIS DO EXCEL]:", colunasBanco);

    if (colunasBanco.length === 0) {
      console.warn("⚠️ Nenhuma coluna válida foi selecionada no filtro.");
      return res.status(400).json({ error: "Nenhuma coluna foi selecionada para o relatório." });
    }

    // 4. Criação do arquivo Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Empresas");

    // Título Principal
    sheet.addRow(["SENALBA MG - Relatório Geral de Empresas"]);
    sheet.mergeCells(1, 1, 1, Math.max(colunasBanco.length, 1));
    const titleRow = sheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    titleRow.alignment = { horizontal: "center", vertical: "middle" };
    titleRow.height = 30;

    sheet.addRow([]); // Espaçamento

    const formataCabecalho = (texto: string) =>
      texto.replace(/_/g, " ").replace(/\b\w/g, (l) => l.toUpperCase());

    // Cabeçalho da Tabela
    const headerRow = sheet.addRow(colunasBanco.map((c) => formataCabecalho(c)));
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Formatação de valores (Booleans, Nulos e Strings)
    const formatarValor = (col: string, valor: any): string => {
      if (valor === null || valor === undefined) return "";

      const camposBooleanos = [
        "status",
        "acordo_sindical",
        "acordo_negocial",
        "acordo_mensalidade"
      ];

      if (camposBooleanos.includes(col)) {
        return valor === true || String(valor).toLowerCase() === "true"
          ? "Sim"
          : "Não";
      }

      return String(valor);
    };

    // 5. Adiciona os dados das empresas
    listaEmpresas.forEach((empresa: any) => {
      const linha = colunasBanco.map((col) => formatarValor(col, empresa[col]));
      sheet.addRow(linha);
    });

    console.log(`✅ [EXCEL] ${listaEmpresas.length} empresas inseridas com sucesso nas colunas selecionadas.`);

    // 6. Largura automática das colunas
    colunasBanco.forEach((col, i) => {
      const cabecalhoLen = formataCabecalho(col).length;
      const maxDataLen = Math.max(
        0,
        ...listaEmpresas.map((item: any) => formatarValor(col, item[col]).length)
      );
      const width = Math.max(cabecalhoLen, maxDataLen);
      sheet.getColumn(i + 1).width = Math.min(Math.max(width * 0.9 + 2, 12), 40);
    });

    // 7. Envio do Buffer
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=empresas_${Date.now()}.xlsx`
    );

    await workbook.xlsx.write(res);
    console.log("🎉 [SUCESSO] Relatório gerado e enviado!");
    return res.end();

  } catch (error) {
    console.error("🔥 [ERRO CRÍTICO INTERNO]:", error);
    return res.status(500).json({ error: "Erro interno ao gerar relatório." });
  }
});

// =========================================================================
// 2. SINDICALIZADOS POR FILTRO (XLSX via Stream)
// =========================================================================
router.post("/sindicalizados/filtro", async (req: Request, res: Response) => {
  try {
    const { escolhaColunaTabela, filtros } = req.body;

    if (!escolhaColunaTabela || typeof escolhaColunaTabela !== "object") {
      return res.status(400).json({ error: "Configuração de colunas inválida." });
    }

    // 1. Inicia a consulta no Supabase
    let query = supabase.from("sindicalizado").select("*");

    const possuiFiltroEmpresa = filtros?.id_empresa && String(filtros.id_empresa).trim() !== "";

    // 2. Filtro de Empresa (OPCIONAL)
    if (possuiFiltroEmpresa) {
      query = query.eq("id_empresa", String(filtros.id_empresa).trim());
    }

    // 3. Filtro de Status (OPCIONAL)
    if (filtros?.status !== undefined && filtros?.status !== null && filtros?.status !== "") {
      const statusStr = String(filtros.status).trim().toLowerCase();
      if (statusStr === "ativo" || statusStr === "true") {
        query = query.eq("status", true);
      } else if (statusStr === "inativo") {
        query = query.eq("status", false);
      }
    }

    // 4. Filtro de Tipo de Desconto (OPCIONAL)
    if (filtros?.tipo_desconto && String(filtros.tipo_desconto).trim() !== "") {
      query = query.ilike("tipo_desconto", `%${String(filtros.tipo_desconto).trim()}%`);
    }

    const { data: listaFiltrada, error } = await query;

    if (error) {
      console.error("🔥 Erro ao consultar Supabase:", error);
      return res.status(500).json({ error: "Erro ao consultar banco de dados." });
    }

    console.log(`📊 Encontrados ${listaFiltrada?.length || 0} sindicalizados.`);

    if (!listaFiltrada || listaFiltrada.length === 0) {
      return res.status(404).json({ error: "Nenhum sindicalizado encontrado para os filtros informados." });
    }

    // 5. Dicionário padronizado: Mapeia TODAS as variações possíveis para o nome exato da coluna no Supabase
    const mapaCampos: Record<string, string> = {
      nome: "nome",
      cpf: "cpf",
      sexo: "sexo",
      nascimento: "nascimento",
      cargo: "cargo",
      estadoCivil: "estado_civil",
      estado_civil: "estado_civil",
      email: "email",
      telefone: "telefone",
      nomeEmpresa: "nome_empresa",
      nome_empresa: "nome_empresa",
      nMatricula: "n_matricula",
      n_matricula: "n_matricula",
      dataAdmissao: "data_admissao",
      data_admissao: "data_admissao",
      dataFiliacao: "data_filiacao",
      data_filiacao: "data_filiacao",
      tipoDesconto: "tipo_desconto",
      tipo_desconto: "tipo_desconto",
      rg: "Rg",
      Rg: "Rg",
      cidade: "cidade",
      estado: "estado",
      status: "status",
      observacoes: "observacoes",
      salario: "salario",
      pisPasep: "pis_pasep",
      pis_pasep: "pis_pasep",
      rua: "rua",
      numero: "numero",
      bairro: "bairro",
      cep: "cep",
      unidade: "unidade"
    };

    // Extrai todas as chaves marcadas como true pelo FlutterFlow
    let colunasSolicitadas = Object.keys(escolhaColunaTabela).filter(
      (key) => escolhaColunaTabela[key] === true
    );

    // Se NÃO filtrou por empresa específica, remove a coluna de empresa da exportação
    if (!possuiFiltroEmpresa) {
      colunasSolicitadas = colunasSolicitadas.filter(
        (col) => col !== "nomeEmpresa" && col !== "nome_empresa"
      );
    }

    if (colunasSolicitadas.length === 0) {
      return res.status(400).json({ error: "Nenhuma coluna selecionada para exibição." });
    }

    // Converte as chaves do FlutterFlow para as colunas exatas do banco Supabase
    // e remove duplicatas (caso venha nomeEmpresa e nome_empresa juntas)
    const colunasBanco = Array.from(
      new Set(colunasSolicitadas.map((key) => mapaCampos[key] || key))
    );

    // 6. Montagem do Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sindicalizados");

    const tituloRelatorio = possuiFiltroEmpresa
      ? "SENALBA MG - Relatório de Sindicalizados por Empresa"
      : "SENALBA MG - Relatório Geral de Associados";

    sheet.addRow([tituloRelatorio]);
    sheet.mergeCells(1, 1, 1, colunasBanco.length);
    const titleRow = sheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    titleRow.alignment = { horizontal: "center", vertical: "middle" };
    titleRow.height = 30;

    sheet.addRow([]); // Linha em branco

    // Cabeçalho das Colunas
    const headerRow = sheet.addRow(colunasBanco.map((c) => formatHeader(c)));
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Inserção das Linhas de Dados
    listaFiltrada.forEach((item: any) => {
      const linha = colunasBanco.map((colBanco) => {
        let valor = item[colBanco] ?? "";

        if (colBanco === "status") {
          valor = valor === true || String(valor).toLowerCase() === "true" ? "Ativo" : "Inativo";
        }
        return String(valor);
      });
      sheet.addRow(linha);
    });

    // Largura Dinâmica das Colunas
    colunasBanco.forEach((col, i) => {
      const maxLength = Math.max(
        formatHeader(col).length,
        ...listaFiltrada.map((item: any) => String(item[col] ?? "").length)
      );
      sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 35);
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=relatorio_associados_${Date.now()}.xlsx`);

    await workbook.xlsx.write(res);
    return res.end();

  } catch (error) {
    console.error("🔥 Erro interno:", error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});

// =========================================================================
// 3. SINDICALIZADOS POR ID (Busca no Supabase para evitar erro 413)
// =========================================================================
router.post("/sindicalizados", async (req: Request, res: Response) => {
  try {
    const { escolhaColunaTabela } = req.body;

    if (!escolhaColunaTabela || typeof escolhaColunaTabela !== "object") {
      return res
        .status(400)
        .json({ error: "Configuração de colunas inválida." });
    }

    // 1. Busca TODOS os sindicalizados direto no Supabase
    const { data: sindicalizados, error } = await supabase
      .from("sindicalizado") // Confirme o nome exato da tabela no seu banco
      .select("*");

    if (error || !sindicalizados) {
      console.error("Erro ao buscar dados no Supabase:", error);
      return res
        .status(500)
        .json({ error: "Erro ao consultar banco de dados." });
    }

    // 2. Extrai as colunas marcadas como true pelo usuário
    const colunasSelecionadas = Object.keys(escolhaColunaTabela).filter(
      (key) => escolhaColunaTabela[key] === true,
    );

    if (colunasSelecionadas.length === 0) {
      return res
        .status(400)
        .json({ error: "Nenhuma coluna selecionada para exportação." });
    }

    // 3. Monta o Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sindicalizados");

    // Título
    sheet.addRow(["SENALBA MG - Relatório de Sindicalizados"]);
    sheet.mergeCells(1, 1, 1, colunasSelecionadas.length);
    const titleRow = sheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
    titleRow.alignment = { horizontal: "center", vertical: "middle" };
    titleRow.height = 30;

    sheet.addRow([]); // Linha em branco

    // Cabeçalho das Colunas
    const headerRow = sheet.addRow(
      colunasSelecionadas.map((c) => formatHeader(c)),
    );
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

    // Adiciona todas as linhas da base
    sindicalizados.forEach((item: any) => {
      const linha = colunasSelecionadas.map((col) => {
        let valor = item[col] ?? "";
        if (col === "status") {
          valor =
            valor === true || String(valor).toLowerCase() === "true"
              ? "Ativo"
              : "Inativo";
        }
        return String(valor);
      });
      sheet.addRow(linha);
    });

    // Ajusta as larguras das colunas
    colunasSelecionadas.forEach((col, i) => {
      const maxLength = Math.max(
        formatHeader(col).length,
        ...sindicalizados.map((item: any) => String(item[col] ?? "").length),
      );
      sheet.getColumn(i + 1).width = Math.min(
        Math.max(maxLength * 0.9, 12),
        35,
      );
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=relatorio_sindicalizados_${Date.now()}.xlsx`,
    );

    await workbook.xlsx.write(res);
    return res.end();
  } catch (err) {
    console.error("Erro na geração do relatório:", err);
    return res.status(500).json({ error: "Erro interno do servidor" });
  }
});

export default router;
