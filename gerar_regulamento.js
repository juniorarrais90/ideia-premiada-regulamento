const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, Footer, PageNumber
} = require("docx");

const FONT = "Arial";
const NAVY = "1F3864";

const pRuns = (runs, opts = {}) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 120, line: 276 },
    ...opts,
    children: runs.map(r => new TextRun({ font: FONT, size: 24, ...r })),
  });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: NAVY })],
  });

const bullet = (text, bold0 = null) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 80, line: 276 },
    children: bold0
      ? [new TextRun({ text: bold0, font: FONT, size: 24, bold: true }),
         new TextRun({ text, font: FONT, size: 24 })]
      : [new TextRun({ text, font: FONT, size: 24 })],
  });

const cb = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: cb, bottom: cb, left: cb, right: cb };

const cell = (text, w, opts = {}) =>
  new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { line: 276 },
      children: [new TextRun({ text, font: FONT, size: opts.size || 22, bold: !!opts.bold, color: opts.color })],
    })],
  });

const CW = 9026; // A4 content width with 1" margins

const children = [];

// Titulo
children.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 },
    children: [new TextRun({ text: "ARRAIS ADVOGADOS", font: FONT, size: 28, bold: true, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [new TextRun({ text: "REGULAMENTO INTERNO", font: FONT, size: 24, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: "PROGRAMA IDEIA PREMIADA", font: FONT, size: 32, bold: true, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
    children: [new TextRun({ text: "Programa de incentivo à sugestão de soluções, aperfeiçoamentos e inovações", font: FONT, size: 22, italics: true })] })
);

// Cap I
children.push(
  h1("CAPÍTULO I — DO OBJETIVO"),
  pRuns([{ text: "Art. 1º. ", bold: true },
    { text: "O Programa Ideia Premiada tem por objetivo estimular os empregados do escritório Arrais Advogados a apresentar sugestões que contribuam para a solução de problemas recorrentes, o aperfeiçoamento de rotinas e processos internos e a criação de novidades que promovam o crescimento do escritório e a melhoria do atendimento aos clientes." }]),
  pRuns([{ text: "Art. 2º. ", bold: true },
    { text: "O Programa constitui ato de gestão interna, de caráter facultativo e discricionário, instituído por liberalidade do empregador, podendo ser alterado, suspenso ou encerrado a qualquer tempo, a exclusivo critério da administração do escritório, sem que disso decorra direito adquirido aos empregados." }])
);

// Cap II
children.push(
  h1("CAPÍTULO II — DOS PARTICIPANTES"),
  pRuns([{ text: "Art. 3º. ", bold: true },
    { text: "Podem participar do Programa todos os empregados contratados sob o regime da CLT, independentemente do cargo ou função, desde que o contrato de trabalho esteja vigente na data da submissão da ideia e na data da eventual premiação." }]),
  pRuns([{ text: "Parágrafo único. ", bold: true },
    { text: "Não participam do Programa os sócios e os advogados associados, cuja contribuição para a melhoria do escritório integra a própria natureza da relação societária ou de associação." }])
);

// Cap III
children.push(
  h1("CAPÍTULO III — DAS CATEGORIAS DE IDEIAS"),
  pRuns([{ text: "Art. 4º. ", bold: true },
    { text: "As ideias submetidas ao Programa serão enquadradas em uma das seguintes categorias:" }]),
  bullet(" — proposta que resolva ou reduza significativamente um problema recorrente do dia a dia do escritório (retrabalho, falhas de comunicação, atrasos, erros em cadastros ou protocolos);", "Categoria A – Solução de problema"),
  bullet(" — proposta que torne uma rotina existente mais rápida, mais barata, mais segura ou com melhor experiência para o cliente;", "Categoria B – Aperfeiçoamento de processo"),
  bullet(" — proposta nova que amplie a captação de clientes, crie um serviço, produto ou fluxo inexistente, ou fortaleça a presença digital e institucional do escritório.", "Categoria C – Inovação e crescimento"),
  pRuns([{ text: "Parágrafo único. ", bold: true },
    { text: "Não serão admitidas sugestões que: (i) configurem mera reclamação sem proposta de solução; (ii) tratem de tema já implementado ou em implementação pelo escritório; (iii) violem a legislação, o Código de Ética e Disciplina da OAB ou o Provimento nº 205/2021 do Conselho Federal da OAB; ou (iv) dependam exclusivamente de aumento de despesa sem demonstração de retorno." }])
);

// Cap IV
children.push(
  h1("CAPÍTULO IV — DA SUBMISSÃO DAS IDEIAS"),
  pRuns([{ text: "Art. 5º. ", bold: true },
    { text: "As ideias deverão ser apresentadas por meio do Formulário de Submissão (Anexo I), encaminhado ao gestor responsável até o dia 25 de cada mês, contendo, no mínimo:" }]),
  bullet("descrição clara do problema ou da oportunidade identificada;"),
  bullet("descrição da solução proposta e de como seria implementada;"),
  bullet("benefício esperado (economia de tempo, redução de custo, ganho de clientes, redução de erros);"),
  bullet("recursos necessários para a implementação, se houver."),
  pRuns([{ text: "§ 1º. ", bold: true },
    { text: "A ideia poderá ser apresentada individualmente ou em grupo de até 3 (três) empregados, hipótese em que eventual prêmio será dividido em partes iguais entre os autores." }]),
  pRuns([{ text: "§ 2º. ", bold: true },
    { text: "Cada empregado poderá submeter quantas ideias desejar dentro do mesmo ciclo mensal." }])
);

// Cap V
children.push(
  h1("CAPÍTULO V — DA AVALIAÇÃO"),
  pRuns([{ text: "Art. 6º. ", bold: true },
    { text: "As ideias serão avaliadas mensalmente por Comitê de Avaliação composto pelo sócio-gestor e por 1 (um) advogado associado por ele designado, na primeira semana do mês subsequente ao da submissão." }]),
  pRuns([{ text: "Art. 7º. ", bold: true },
    { text: "A avaliação observará os seguintes critérios, pontuados de 1 a 5:" }]),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [2300, 5226, 1500],
    rows: [
      new TableRow({ children: [
        cell("Critério", 2300, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
        cell("O que mede", 5226, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
        cell("Peso", 1500, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
      ]}),
      new TableRow({ children: [
        cell("Aplicabilidade", 2300, { bold: true }),
        cell("Viabilidade prática de implementação com a estrutura atual do escritório", 5226),
        cell("3", 1500, { center: true }),
      ]}),
      new TableRow({ children: [
        cell("Impacto", 2300, { bold: true }),
        cell("Dimensão do benefício gerado (tempo, custo, clientes, qualidade)", 5226),
        cell("3", 1500, { center: true }),
      ]}),
      new TableRow({ children: [
        cell("Custo-benefício", 2300, { bold: true }),
        cell("Relação entre o investimento necessário e o retorno esperado", 5226),
        cell("2", 1500, { center: true }),
      ]}),
      new TableRow({ children: [
        cell("Originalidade", 2300, { bold: true }),
        cell("Grau de novidade da proposta em relação às práticas atuais", 5226),
        cell("2", 1500, { center: true }),
      ]}),
    ],
  }),
  new Paragraph({ spacing: { after: 120 }, children: [] }),
  pRuns([{ text: "§ 1º. ", bold: true },
    { text: "Serão consideradas aprovadas as ideias que atingirem pontuação ponderada igual ou superior a 70% (setenta por cento) do total possível e cuja implementação seja autorizada pela gestão." }]),
  pRuns([{ text: "§ 2º. ", bold: true },
    { text: "A aprovação da ideia não obriga o escritório a implementá-la de imediato, podendo a execução ser programada conforme a conveniência operacional e financeira." }]),
  pRuns([{ text: "§ 3º. ", bold: true },
    { text: "O resultado da avaliação será comunicado aos autores e registrado em ata simplificada, com a justificativa da decisão." }])
);

// Cap VI
children.push(
  h1("CAPÍTULO VI — DA PREMIAÇÃO"),
  pRuns([{ text: "Art. 8º. ", bold: true },
    { text: "O autor de ideia aprovada e implementada fará jus a prêmio em dinheiro, concedido por liberalidade do escritório em razão de desempenho superior ao ordinariamente esperado no exercício de suas atividades, observados os seguintes valores de referência:" }]),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [2400, 4326, 2300],
    rows: [
      new TableRow({ children: [
        cell("Nível de impacto", 2400, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
        cell("Exemplos", 4326, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
        cell("Valor de referência", 2300, { bold: true, fill: "1F3864", color: "FFFFFF", center: true }),
      ]}),
      new TableRow({ children: [
        cell("Leve", 2400, { bold: true }),
        cell("Melhoria pontual de rotina, organização de fluxo, padronização de documento", 4326),
        cell("R$ 150,00", 2300, { center: true }),
      ]}),
      new TableRow({ children: [
        cell("Médio", 2400, { bold: true }),
        cell("Economia mensurável de tempo ou custo, redução comprovada de erros ou retrabalho", 4326),
        cell("R$ 300,00", 2300, { center: true }),
      ]}),
      new TableRow({ children: [
        cell("Alto", 2400, { bold: true }),
        cell("Aumento de captação de clientes, novo serviço ou fluxo, economia relevante e permanente", 4326),
        cell("R$ 500,00 a R$ 1.000,00", 2300, { center: true }),
      ]}),
    ],
  }),
  new Paragraph({ spacing: { after: 120 }, children: [] }),
  pRuns([{ text: "§ 1º. ", bold: true },
    { text: "O enquadramento do nível de impacto e a fixação do valor do prêmio competem exclusivamente ao Comitê de Avaliação, caso a caso, à luz do resultado efetivamente verificado ou razoavelmente projetado." }]),
  pRuns([{ text: "§ 2º. ", bold: true },
    { text: "O prêmio será pago em folha, em rubrica própria de prêmio, no mês subsequente à decisão do Comitê, ou, no caso de ideias cujo resultado dependa de verificação, após a comprovação do benefício." }]),
  pRuns([{ text: "§ 3º. ", bold: true },
    { text: "A existência do Programa não assegura premiação em todos os meses: somente haverá pagamento quando houver ideia aprovada nos termos deste Regulamento." }]),
  pRuns([{ text: "§ 4º. ", bold: true },
    { text: "Havendo ideias idênticas ou substancialmente semelhantes, prevalecerá a que tiver sido submetida primeiro." }])
);

// Cap VII
children.push(
  h1("CAPÍTULO VII — DA NATUREZA JURÍDICA DO PRÊMIO"),
  pRuns([{ text: "Art. 9º. ", bold: true },
    { text: "Os prêmios concedidos no âmbito deste Programa constituem liberalidade do empregador, paga em razão de desempenho superior ao ordinariamente esperado, nos exatos termos do art. 457, §§ 2º e 4º, da CLT, com a redação dada pela Lei nº 13.467/2017, e, portanto:" }]),
  bullet("não integram a remuneração do empregado;"),
  bullet("não se incorporam ao contrato de trabalho;"),
  bullet("não constituem base de incidência de encargos trabalhistas;"),
  bullet("não integram o salário de contribuição, nos termos do art. 28, § 9º, alínea “z”, da Lei nº 8.212/1991, incluída pela Lei nº 13.467/2017."),
  pRuns([{ text: "§ 1º. ", bold: true },
    { text: "A apresentação de ideias não constitui atribuição contratual dos empregados nem condição para a manutenção do emprego, tratando-se de conduta voluntária que excede o desempenho ordinariamente esperado das funções de cada cargo." }]),
  pRuns([{ text: "§ 2º. ", bold: true },
    { text: "O prêmio não substitui, no todo ou em parte, qualquer parcela salarial ajustada, não possui caráter de contraprestação habitual e não gera expectativa de pagamento futuro, ainda que o mesmo empregado venha a ser premiado mais de uma vez." }]),
  pRuns([{ text: "§ 3º. ", bold: true },
    { text: "Cada concessão será documentada com a identificação da ideia, a justificativa do Comitê e a demonstração do desempenho superior, para fins de comprovação perante as autoridades fiscais e trabalhistas." }])
);

// Cap VIII
children.push(
  h1("CAPÍTULO VIII — DA IMPLEMENTAÇÃO E DA TITULARIDADE DAS IDEIAS"),
  pRuns([{ text: "Art. 10. ", bold: true },
    { text: "As ideias submetidas ao Programa, premiadas ou não, poderão ser livremente utilizadas, adaptadas e implementadas pelo escritório, que será o titular exclusivo dos processos, materiais, métodos e rotinas delas resultantes." }]),
  pRuns([{ text: "Parágrafo único. ", bold: true },
    { text: "A submissão da ideia implica concordância do autor com os termos deste Regulamento e autorização para uso institucional da proposta, sem prejuízo do reconhecimento interno da autoria." }])
);

// Cap IX
children.push(
  h1("CAPÍTULO IX — DAS DISPOSIÇÕES FINAIS"),
  pRuns([{ text: "Art. 11. ", bold: true },
    { text: "Este Programa vigorará por 12 (doze) meses a contar de sua divulgação, podendo ser renovado, alterado, suspenso ou encerrado a qualquer tempo pela gestão, mediante comunicação à equipe, sem direito a indenização ou compensação de qualquer natureza." }]),
  pRuns([{ text: "Art. 12. ", bold: true },
    { text: "Os casos omissos serão resolvidos pelo sócio-gestor, cujas decisões são definitivas no âmbito do Programa." }]),
  pRuns([{ text: "Art. 13. ", bold: true },
    { text: "Este Regulamento entra em vigor na data de sua divulgação à equipe." }]),
  new Paragraph({ spacing: { before: 360, after: 60 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Ubá/MG, ____ de ______________ de 2026.", font: FONT, size: 24 })] }),
  new Paragraph({ spacing: { before: 480, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "_____________________________________", font: FONT, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
    children: [new TextRun({ text: "Arrais Advogados", font: FONT, size: 24, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Sócio-gestor", font: FONT, size: 22 })] })
);

// Anexo I
const formRow = (label) =>
  new TableRow({ children: [
    cell(label, 3000, { bold: true, fill: "EDF1F8" }),
    new TableCell({
      borders,
      width: { size: 6026, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [] }), new Paragraph({ children: [] })],
    }),
  ]});

children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "ANEXO I — FORMULÁRIO DE SUBMISSÃO DE IDEIA", font: FONT, size: 26, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 160 },
    children: [new TextRun({ text: "Preencha e envie ao gestor responsável até o dia 25 de cada mês.", font: FONT, size: 22, italics: true })] }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [3000, 6026],
    rows: [
      formRow("Nome do(s) autor(es)"),
      formRow("Data da submissão"),
      formRow("Categoria (A, B ou C)"),
      formRow("Problema ou oportunidade identificada"),
      formRow("Solução proposta (como funcionaria na prática)"),
      formRow("Benefício esperado (tempo, custo, clientes, qualidade)"),
      formRow("Recursos necessários (se houver)"),
    ],
  }),
  new Paragraph({ spacing: { before: 200, after: 80 },
    children: [new TextRun({ text: "Uso exclusivo do Comitê de Avaliação:", font: FONT, size: 22, bold: true })] }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [2256, 2256, 2257, 2257],
    rows: [
      new TableRow({ children: [
        cell("Aplicabilidade (1-5)", 2256, { bold: true, fill: "EDF1F8", center: true }),
        cell("Impacto (1-5)", 2256, { bold: true, fill: "EDF1F8", center: true }),
        cell("Custo-benefício (1-5)", 2257, { bold: true, fill: "EDF1F8", center: true }),
        cell("Originalidade (1-5)", 2257, { bold: true, fill: "EDF1F8", center: true }),
      ]}),
      new TableRow({ children: [
        cell("", 2256), cell("", 2256), cell("", 2257), cell("", 2257),
      ]}),
      new TableRow({ children: [
        cell("Resultado", 2256, { bold: true, fill: "EDF1F8" }),
        cell("(   ) Aprovada    (   ) Reprovada", 2256),
        cell("Nível de impacto", 2257, { bold: true, fill: "EDF1F8" }),
        cell("", 2257),
      ]}),
    ],
  })
);

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: NAVY },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Arrais Advogados — Programa Ideia Premiada — Página ", font: FONT, size: 18, color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "666666" }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/vigilant-youthful-davinci/mnt/outputs/Regulamento_Programa_Ideia_Premiada_Arrais_Advogados.docx", buffer);
  console.log("OK");
});
