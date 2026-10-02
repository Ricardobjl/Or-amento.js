const ABAS = {
  CLIENTES: 'CLIENTES',
  SERVICOS: 'SERVICOS',
  ORCAMENTOS: 'ORCAMENTOS',
  ITENS: 'ITENS_ORCAMENTO'
};

function obterPlanilha() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function configurarBanco() {
  const ss = obterPlanilha();

  criarAba(
    ss,
    ABAS.CLIENTES,
    ['ID', 'NOME', 'TELEFONE', 'EMAIL', 'ENDERECO', 'DATA_CADASTRO']
  );

  criarAba(
    ss,
    ABAS.SERVICOS,
    ['ID', 'NOME', 'DESCRICAO', 'PRECO', 'ATIVO']
  );

  criarAba(
    ss,
    ABAS.ORCAMENTOS,
    [
      'ID',
      'NUMERO',
      'CLIENTE_ID',
      'DATA',
      'VALIDADE',
      'SUBTOTAL',
      'DESCONTO',
      'TOTAL',
      'STATUS'
    ]
  );

  criarAba(
    ss,
    ABAS.ITENS,
    [
      'ID',
      'ORCAMENTO_ID',
      'DESCRICAO',
      'QUANTIDADE',
      'VALOR_UNITARIO',
      'TOTAL'
    ]
  );

  inserirServicosIniciais();

  return 'Banco configurado com sucesso!';
}

function criarAba(ss, nome, colunas) {
  let aba = ss.getSheetByName(nome);

  if (!aba) {
    aba = ss.insertSheet(nome);
  }

  if (aba.getLastRow() === 0) {
    aba.appendRow(colunas);
  }
}

function inserirServicosIniciais() {
  const ss = obterPlanilha();
  const aba = ss.getSheetByName(ABAS.SERVICOS);

  if (aba.getLastRow() > 1) {
    return;
  }

  const servicos = [
    ['S001', 'Formatação', 'Formatação e instalação do sistema', 35, true],
    ['S002', 'Limpeza', 'Limpeza interna do computador', 30, true],
    ['S003', 'Montagem de PC', 'Montagem completa do computador', 50, true],
    ['S004', 'Manutenção', 'Diagnóstico e manutenção', 40, true],
    ['S005', 'Instalação de software', 'Instalação e configuração', 20, true]
  ];

  aba.getRange(
    aba.getLastRow() + 1,
    1,
    servicos.length,
    servicos[0].length
  ).setValues(servicos);
}