// ========================================
// LISTAR SERVIÇOS
// ========================================

function listarServicos() {

  const ss = obterPlanilha();

  const aba =
    ss.getSheetByName(ABAS.SERVICOS);

  if (!aba) {
    throw new Error(
      'Aba SERVICOS não encontrada.'
    );
  }

  const dados =
    aba.getDataRange().getValues();

  if (dados.length <= 1) {
    return [];
  }

  return dados
    .slice(1)
    .filter(linha => linha[4] === true)
    .map(linha => ({

      id: String(linha[0]),

      nome: String(linha[1]),

      descricao: String(linha[2]),

      preco: Number(linha[3])

    }));

}


// ========================================
// GERAR ID
// ========================================

function gerarId(prefixo) {

  return prefixo +
    '-' +
    Utilities.getUuid();

}


// ========================================
// GERAR NÚMERO DO ORÇAMENTO
// ========================================

function gerarNumeroOrcamento() {

  const ano =
    new Date().getFullYear();

  const aba =
    obterPlanilha()
      .getSheetByName(ABAS.ORCAMENTOS);

  const quantidade =
    Math.max(0, aba.getLastRow() - 1) + 1;

  return 'ORC-' +
    ano +
    '-' +
    String(quantidade).padStart(4, '0');

}


// ========================================
// SALVAR ORÇAMENTO
// ========================================

function salvarOrcamento(dados) {

  if (!dados) {
    throw new Error(
      'Dados do orçamento não enviados.'
    );
  }

  if (!dados.cliente ||
      !dados.cliente.nome) {

    throw new Error(
      'Cliente inválido.'
    );

  }

  if (!dados.itens ||
      dados.itens.length === 0) {

    throw new Error(
      'O orçamento precisa ter itens.'
    );

  }


  const ss =
    obterPlanilha();


  // ----------------------------------------
  // CLIENTE
  // ----------------------------------------

  const clienteId =
    gerarId('CLI');


  const abaClientes =
    ss.getSheetByName(
      ABAS.CLIENTES
    );


  abaClientes.appendRow([

    clienteId,

    dados.cliente.nome,

    dados.cliente.telefone,

    dados.cliente.email,

    dados.cliente.endereco,

    new Date()

  ]);


  // ----------------------------------------
  // CÁLCULOS
  // ----------------------------------------

  let subtotal = 0;

  dados.itens.forEach(function(item) {

    subtotal +=
      Number(item.quantidade) *
      Number(item.valor);

  });


  const desconto =
    Number(dados.desconto) || 0;


  const total =
    Math.max(
      0,
      subtotal - desconto
    );


  // ----------------------------------------
  // ORÇAMENTO
  // ----------------------------------------

  const orcamentoId =
    gerarId('ORC');


  const numero =
    gerarNumeroOrcamento();


  const agora =
    new Date();


  const validade =
    new Date(agora);


  validade.setDate(
    validade.getDate() + 7
  );


  const abaOrcamentos =
    ss.getSheetByName(
      ABAS.ORCAMENTOS
    );


  abaOrcamentos.appendRow([

    orcamentoId,

    numero,

    clienteId,

    agora,

    validade,

    subtotal,

    desconto,

    total,

    'PENDENTE'

  ]);


  // ----------------------------------------
  // ITENS
  // ----------------------------------------

  const abaItens =
    ss.getSheetByName(
      ABAS.ITENS
    );


  dados.itens.forEach(function(item) {

    const itemId =
      gerarId('ITEM');


    const quantidade =
      Number(item.quantidade);


    const valor =
      Number(item.valor);


    const totalItem =
      quantidade * valor;


    abaItens.appendRow([

      itemId,

      orcamentoId,

      item.descricao,

      quantidade,

      valor,

      totalItem

    ]);

  });


  return {

    sucesso: true,

    id: orcamentoId,

    numero: numero,

    subtotal: subtotal,

    desconto: desconto,

    total: total

  };

}