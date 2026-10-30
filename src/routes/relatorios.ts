import { Router, Request, Response } from "express"; 
import ExcelJS from "exceljs"; 
import fs from "fs"; 
import { createClient } from "@supabase/supabase-js"; 

const router = Router(); 

const supabaseUrl = process.env.SUPABASE_URL || "SUA_URL_DO_SUPABASE"; 
const supabaseKey = process.env.SUPABASE_SERVICE_KEY || ""; 
const supabase = createClient(supabaseUrl, supabaseKey); 

// Função auxiliar para buscar TODOS os registros contornando o limite de 1000 do Supabase 
async function buscarTodosRegistros(tabela: string, aplicarFiltros?: (query: any) => void) { 
  let todosOsDados: any[] = []; 
  let de = 0; 
  const limitePorPagina = 1000; 
  let temMais = true; 

  while (temMais) { 
    let query = supabase.from(tabela).select("*"); 

    if (aplicarFiltros) { 
      aplicarFiltros(query); 
    } 

    const { data, error } = await query 
      .order("id", { ascending: true }) // Ordenação consistente para paginação 
      .range(de, de + limitePorPagina - 1); 

    if (error) { 
      console.error(`🔥 [ERRO BANCO - ${tabela}]:`, error); 
      throw error; 
    } 

    if (data && data.length > 0) { 
      todosOsDados = todosOsDados.concat(data); 
      de += limitePorPagina; 
      // Se retornou menos que 1000, chegou ao fim da base 
      if (data.length < limitePorPagina) { 
        temMais = false; 
      } 
    } else { 
      temMais = false; 
    } 
  } 

  return todosOsDados; 
} 

const formatHeader = (key: string) => { 
  if (key === "status") return "Status (Ativo / Inativo)"; 
  if (key === "tipo_convencao") return "Tipo de Convenção"; 
  if (key === "numero_funcionario") return "Nº de Funcionários"; 
  if (key === "razao_social") return "Razão Social"; 
  return key 
    .replace(/_/g, " ") 
    .toLowerCase() 
    .replace(/\b\w/g, (l) => l.toUpperCase()); 
}; 

// ========================================================================= 
// 1. FILTRAR EMPRESAS XLSX (Com Busca Paginada + Exibição dos Filtros no Cabeçalho)
// ========================================================================= 
router.post("/empresas-filtradas", async (req: Request, res: Response) => { 
  console.log("\n🚀 [INÍCIO] /relatorios/empresas-filtradas"); 

  try { 
    const { escolhaColunaTabela, tipoConvencao, acordo } = req.body; 

    let todosOsDados: any[] = []; 
    let de = 0; 
    const limite = 1000; 
    let temMais = true; 

    // Busca com paginação aplicando os filtros 
    while (temMais) { 
      let query = supabase.from("empresa").select("*"); 

      if (Array.isArray(tipoConvencao) && tipoConvencao.length > 0) { 
        query = query.in("tipo_convencao", tipoConvencao); 
      } 

      if (Array.isArray(acordo) && acordo.length > 0) { 
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

      const { data, error } = await query 
        .order("razao_social", { ascending: true }) 
        .range(de, de + limite - 1); 

      if (error) throw error; 

      if (data && data.length > 0) { 
        todosOsDados = todosOsDados.concat(data); 
        de += limite; 
        if (data.length < limite) temMais = false; 
      } else { 
        temMais = false; 
      } 
    } 

    console.log(`📊 Total recuperado sem limite: ${todosOsDados.length} empresas.`); 

    if (todosOsDados.length === 0) { 
      return res.status(404).json({ error: "Nenhuma empresa encontrada com os filtros." }); 
    } 

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
      numeroFuncionario: "numero_funcionario", 
      tipo_convencao: "tipo_convencao", 
      tipoConvencao: "tipo_convencao", 
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

    if (escolhaColunaTabela && typeof escolhaColunaTabela === "object") { 
      const colunasAtivas = Object.keys(escolhaColunaTabela).filter((key) => { 
        const val = escolhaColunaTabela[key]; 
        return val === true || String(val).toLowerCase() === "true"; 
      }); 

      colunasAtivas.forEach((key) => { 
        if (key === "acordo") { 
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

    // Garante exibição da coluna 'tipo_convencao' no Excel quando o filtro é aplicado
    if (Array.isArray(tipoConvencao) && tipoConvencao.length > 0) {
      if (!colunasBanco.includes("tipo_convencao")) {
        colunasBanco.push("tipo_convencao");
      }
    }

    const workbook = new ExcelJS.Workbook(); 
    const sheet = workbook.addWorksheet("Empresas Filtradas"); 

    const numColunas = Math.max(colunasBanco.length, 1);

    // 1. Título do Relatório
    sheet.addRow(["SENALBA MG - Relatório de Empresas (Filtrado)"]); 
    sheet.mergeCells(1, 1, 1, numColunas); 
    const titleRow = sheet.getRow(1); 
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } }; 
    titleRow.alignment = { horizontal: "center", vertical: "middle" }; 
    titleRow.height = 30; 

    // 2. Informações dos Filtros Aplicados
    const textoConvencao = Array.isArray(tipoConvencao) && tipoConvencao.length > 0 
      ? tipoConvencao.join(", ") 
      : "Todos";

    const textoAcordo = Array.isArray(acordo) && acordo.length > 0 
      ? acordo.join(", ") 
      : "Todos";

    const rowFiltro1 = sheet.addRow([`Convenção: ${textoConvencao}`]);
    sheet.mergeCells(2, 1, 2, numColunas);
    rowFiltro1.font = { italic: true, size: 11, color: { argb: "FF555555" } };
    rowFiltro1.alignment = { horizontal: "left", vertical: "middle" };

    const rowFiltro2 = sheet.addRow([`Acordo: ${textoAcordo}`]);
    sheet.mergeCells(3, 1, 3, numColunas);
    rowFiltro2.font = { italic: true, size: 11, color: { argb: "FF555555" } };
    rowFiltro2.alignment = { horizontal: "left", vertical: "middle" };

    // Linha em branco de separação
    sheet.addRow([]); 

    // 3. Cabeçalho das Colunas da Tabela
    const formataCabecalho = (texto: string) => formatHeader(texto);

    const headerRow = sheet.addRow(colunasBanco.map((c) => formataCabecalho(c))); 
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } }; 
    headerRow.height = 22; 
    headerRow.eachCell((cell) => { 
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } }; 
      cell.alignment = { horizontal: "center", vertical: "middle" }; 
    }); 

    const formatarValor = (col: string, valor: any): string => { 
      if (valor === null || valor === undefined) return ""; 
      const camposBooleanos = ["status", "acordo_sindical", "acordo_negocial", "acordo_mensalidade"]; 
      if (camposBooleanos.includes(col)) { 
        return valor === true || String(valor).toLowerCase() === "true" ? "Sim" : "Não"; 
      } 
      return String(valor); 
    }; 

    // 4. Preenchimento dos Dados das Empresas
    todosOsDados.forEach((empresa: any) => { 
      const linha = colunasBanco.map((col) => formatarValor(col, empresa[col])); 
      sheet.addRow(linha); 
    }); 

    // Ajuste Automático da Largura das Colunas
    colunasBanco.forEach((col, i) => { 
      const cabecalhoLen = formataCabecalho(col).length; 
      const maxDataLen = Math.max( 
        0, 
        ...todosOsDados.map((item: any) => formatarValor(col, item[col]).length) 
      ); 
      const width = Math.max(cabecalhoLen, maxDataLen); 
      sheet.getColumn(i + 1).width = Math.min(Math.max(width * 0.9 + 2, 12), 40); 
    }); 

    const buffer = await workbook.xlsx.writeBuffer(); 
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); 
    res.setHeader("Content-Disposition", `attachment; filename=empresas_filtradas_${Date.now()}.xlsx`); 
    return res.status(200).send(buffer); 

  } catch (error) { 
    console.error("🔥 Erro crítico:", error); 
    return res.status(500).json({ error: "Erro interno ao gerar relatório." }); 
  } 
}); 

// ========================================================================= 
// 3. SINDICALIZADOS POR ID / GERAL (Com Posição Paginada para > 1000 Linhas) 
// ========================================================================= 
router.post("/sindicalizados", async (req: Request, res: Response) => { 
  try { 
    const { escolhaColunaTabela } = req.body; 

    if (!escolhaColunaTabela || typeof escolhaColunaTabela !== "object") { 
      return res.status(400).json({ error: "Configuração de colunas inválida." }); 
    } 

    const sindicalizados = await buscarTodosRegistros("sindicalizado"); 

    if (!sindicalizados || sindicalizados.length === 0) { 
      return res.status(404).json({ error: "Nenum sindicalizado encontrado no banco." }); 
    } 

    console.log(`📊 Total de sindicalizados carregados: ${sindicalizados.length}`); 

    const colunasSelecionadas = Object.keys(escolhaColunaTabela).filter( 
      (key) => escolhaColunaTabela[key] === true 
    ); 

    if (colunasSelecionadas.length === 0) { 
      return res.status(400).json({ error: "Nenhuma coluna selecionada para exportação." }); 
    } 

    const workbook = new ExcelJS.Workbook(); 
    const sheet = workbook.addWorksheet("Sindicalizados"); 

    sheet.addRow(["SENALBA MG - Relatório de Sindicalizados"]); 
    sheet.mergeCells(1, 1, 1, colunasSelecionadas.length); 
    const titleRow = sheet.getRow(1); 
    titleRow.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } }; 
    titleRow.alignment = { horizontal: "center", vertical: "middle" }; 
    titleRow.height = 30; 

    sheet.addRow([]); 

    const headerRow = sheet.addRow(colunasSelecionadas.map((c) => formatHeader(c))); 
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } }; 
    headerRow.height = 22; 
    headerRow.eachCell((cell) => { 
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "4472C4" } }; 
      cell.alignment = { horizontal: "center", vertical: "middle" }; 
    }); 

    sindicalizados.forEach((item: any) => { 
      const linha = colunasSelecionadas.map((col) => { 
        let valor = item[col] ?? ""; 
        if (col === "status") { 
          valor = valor === true || String(valor).toLowerCase() === "true" ? "Ativo" : "Inativo"; 
        } 
        return String(valor); 
      }); 
      sheet.addRow(linha); 
    }); 

    colunasSelecionadas.forEach((col, i) => { 
      const maxLength = Math.max( 
        formatHeader(col).length, 
        ...sindicalizados.map((item: any) => String(item[col] ?? "").length) 
      ); 
      sheet.getColumn(i + 1).width = Math.min(Math.max(maxLength * 0.9, 12), 35); 
    }); 

    const buffer = await workbook.xlsx.writeBuffer(); 
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); 
    res.setHeader("Content-Disposition", `attachment; filename=relatorio_sindicalizados_${Date.now()}.xlsx`); 
    return res.status(200).send(buffer); 

  } catch (err) { 
    console.error("Erro na geração do relatório:", err); 
    return res.status(500).json({ error: "Erro interno do servidor" }); 
  } 
}); 

module.exports = router;