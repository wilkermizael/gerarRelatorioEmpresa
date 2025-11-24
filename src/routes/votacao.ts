import express from "express";
import ExcelJS from "exceljs";

const router = express.Router();

router.post("/votacao", async (req, res) => {
  try {
    const dados = req.body.dados;

    if (!Array.isArray(dados)) {
      return res.status(400).json({ error: "dados deve ser uma lista." });
    }

    const workbook = new ExcelJS.Workbook();

    //
    // ────────────────────────────────────────────
    //  ABA 1 — RELATÓRIO COMPLETO
    // ────────────────────────────────────────────
    //

    const sheet = workbook.addWorksheet("Relatório Votação");

    sheet.columns = [
      { header: "Nome", key: "nome", width: 30 },
      { header: "CPF", key: "cpf", width: 20 },
      { header: "Telefone", key: "telefone", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "Empresa", key: "nome_empresa", width: 30 },
      { header: "Taxa Negocial (Escolha)", key: "taxa_negocial", width: 20 },
      { header: "Opositor", key: "opositor", width: 12 },
      { header: "Outra Empresa", key: "outra", width: 30 },
    ];

    dados.forEach((item) => {
      sheet.addRow({
        nome: item.nome || "",
        cpf: item.cpf || "",
        telefone: item.telefone || "",
        email: item.email || "",
        nome_empresa: item.nome_empresa || "",
        taxa_negocial: item.taxa_negocial || "Não informado",
        opositor: item.opositor ? "Sim" : "Não",
        outra: item.outro_nome_empresa || "",
      });
    });

    // Estilo cabeçalho
    sheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDDDDD" },
      };
    });

    //
    // ────────────────────────────────────────────
    //  ABA 2 — RESUMO (CONTAGEM + EMPATE)
    // ────────────────────────────────────────────
    //

    const contagem: Record<string, number> = {};
    const totalVotos = dados.length;

    // Contar votos
    dados.forEach((item) => {
      const escolha = item.taxa_negocial || "Não informado";
      contagem[escolha] = (contagem[escolha] || 0) + 1;
    });

    // Encontrar a(s) opção(ões) mais votada(s)
    const entries = Object.entries(contagem);

    const maiorQtd = Math.max(...entries.map(([_, qtd]) => qtd));

    const empatadas = entries.filter(([_, qtd]) => qtd === maiorQtd);

    // Se só houver 1 vencedora, pega ela; se houver empate, retorna null
    const vencedora = empatadas.length === 1 ? empatadas[0][0] : null;

    const resumoSheet = workbook.addWorksheet("Resumo");

    resumoSheet.columns = [
      { header: "Opção", key: "opcao", width: 20 },
      { header: "Quantidade de Votos", key: "qtd", width: 25 },
      { header: "Percentual", key: "percentual", width: 20 },
      { header: "Vencedora", key: "vencedora", width: 15 },
    ];

    // Preencher resumo
    Object.entries(contagem).forEach(([opcao, qtd]) => {
      const percentual = ((qtd / totalVotos) * 100).toFixed(2) + "%";

      resumoSheet.addRow({
        opcao,
        qtd,
        percentual,
        vencedora:
          vencedora === opcao
            ? "SIM"
            : vencedora === null && qtd === maiorQtd
            ? "EMPATE"
            : "",
      });
    });

    // Estilizar cabeçalho
    resumoSheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEEEEEE" },
      };
    });

    //
    // ────────────────────────────────────────────
    //  ENVIAR ARQUIVO
    // ────────────────────────────────────────────
    //

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=relatorio_votacao.xlsx"
    );

    return res.send(buffer);

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Erro ao gerar relatório." });
  }
});

export default router;
