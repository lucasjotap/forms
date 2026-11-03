/*
 * Gera, no navegador, o modelo de controle (.xlsx) e a versão CSV simples.
 * Nenhuma resposta é enviada a servidor: este arquivo só monta o arquivo localmente.
 * Uso: PlanilhaModelo.buildWorkbook(ExcelJS, opts) e PlanilhaModelo.buildCsv(opts),
 * em que opts = { business, priority, activity, tool, recommendations: [{ id, title, action, check }] }
 * (recommendations também aceita só títulos). tool: 'paper' | 'phone' | 'sheet' (outro valor = neutro).
 * activity: texto do select do formulário ("Workshop…", "Projeto 1…", "Projeto 2…") ou
 * 'workshop' | 'p1' | 'p2'. Sem atividade reconhecida = modelo universal de 8 abas.
 */
(function (root) {
  'use strict';

  var EX = 'EXEMPLO — APAGUE ANTES DE USAR';
  var S = {
    leia: 'LEIA PRIMEIRO', painel: 'PAINEL', fluxo: 'FLUXO DE CAIXA', contas: 'CONTAS E COMPROMISSOS',
    clientes: 'CLIENTES E PEDIDOS', preco: 'PRECIFICAÇÃO', rotina: 'ROTINA SEMANAL', plano: 'PLANO DE MELHORIA',
    orcamento: 'ORÇAMENTO DO MÊS'
  };
  var SHEETS = [S.leia, S.painel, S.fluxo, S.contas, S.clientes, S.preco, S.rotina, S.plano];
  var COLOR = {
    dark: 'FF1F4E3D', header: 'FF2C6650', headerCalc: 'FF5F6F68', input: 'FFEEF5FC', calc: 'FFF1F3F1',
    good: 'FFE3F2E6', warn: 'FFFFF3CD', bad: 'FFFBE3E3', info: 'FFE8F0F9', white: 'FFFFFFFF', helpFill: 'FFF6F8F7',
    text: 'FF1E2B26', muted: 'FF5F6F68', exFont: 'FF8A4B00', exFill: 'FFFDF3E6', line: 'FFD3DDD7',
    goodFont: 'FF1E6B3A', warnFont: 'FF7A5600', badFont: 'FF9B2222', infoFont: 'FF1F4E79',
    secondary: 'FF6B7A74', secondaryTab: 'FFBFC6C2'
  };
  var UNIVERSAL_TABS = {};
  UNIVERSAL_TABS[S.leia] = 'FF1F5F4A'; UNIVERSAL_TABS[S.painel] = 'FF2E7D5B'; UNIVERSAL_TABS[S.fluxo] = 'FF2F6FA8';
  UNIVERSAL_TABS[S.contas] = 'FFB7791F'; UNIVERSAL_TABS[S.clientes] = 'FF7A4FA0'; UNIVERSAL_TABS[S.preco] = 'FFB0473B';
  UNIVERSAL_TABS[S.rotina] = 'FF3A8C8C'; UNIVERSAL_TABS[S.plano] = 'FF5D7A2E';
  // Sequência de verdes para as abas do roteiro da atividade; abas complementares em cinza.
  var STEP_TABS = ['FF1F4E3D', 'FF2C6650', 'FF3B7D60', 'FF4E9172', 'FF67A486', 'FF84B89D'];
  var BRL = '"R$" #,##0.00;[Red]-"R$" #,##0.00';
  var BRL_DASH = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"';
  var DATE = 'dd/mm/yyyy';
  var MONTH = '[$-416]mmmm "de" yyyy';
  var HOURS = '0.0# "h"';
  var FONT = 'Calibri';

  var FOCUS = {
    cash: { label: 'Entender e organizar o dinheiro', tab: S.fluxo, action: 'Registrar todas as movimentações por 7 dias',
      step: 'Durante 7 dias, registre no FLUXO DE CAIXA cada entrada, saída e retirada pessoal.', measure: 'Dias com registro na semana', goal: '7 de 7 dias registrados' },
    price: { label: 'Saber melhor quanto cobrar', tab: S.preco, action: 'Revisar o custo de um produto/serviço',
      step: 'Escolha um produto ou serviço e preencha uma linha da aba PRECIFICAÇÃO.', measure: 'Itens com custo revisado', goal: '1 item revisado' },
    customers: { label: 'Organizar clientes e pedidos', tab: S.clientes, action: 'Reunir pedidos em um único lugar',
      step: 'Registre em CLIENTES E PEDIDOS os pedidos e combinados que estão em aberto hoje.', measure: 'Pedidos anotados na aba', goal: 'Todos os pedidos da semana anotados' },
    routine: { label: 'Organizar o tempo e as tarefas', tab: S.rotina, action: 'Escolher até três prioridades por dia',
      step: 'Na ROTINA SEMANAL, escreva até três prioridades para amanhã.', measure: 'Dias com prioridades definidas', goal: '5 dias na semana' },
    unsure: { label: 'Ainda a definir — observe a rotina por uma semana', tab: S.fluxo, action: 'Observar a rotina por uma semana e escolher um foco',
      step: 'Registre por 7 dias o dinheiro que entra e sai e, no fim, escolha um foco no PLANO DE MELHORIA.', measure: 'Dias com registro na semana', goal: 'Escolher 1 foco ao fim da semana' }
  };

  // Variantes por atividade (08_diagnostico_e_devolutiva/LEIA_PRIMEIRO.md e caderno da oficina).
  var ACTIVITIES = {
    universal: {
      key: 'universal', label: '', subtitle: 'Um modelo para entender, organizar e melhorar sua rotina',
      order: SHEETS.slice(), primary: null, focusTabs: null, defaultTab: null, suffix: ''
    },
    workshop: {
      key: 'workshop', label: 'Workshop de Educação Financeira e Planejamento', suffix: '_oficina',
      subtitle: 'Workshop de Educação Financeira e Planejamento · acompanhe a oficina aba por aba',
      order: [S.leia, S.fluxo, S.contas, S.orcamento, S.plano, S.painel, S.clientes, S.preco, S.rotina],
      primary: [S.leia, S.fluxo, S.contas, S.orcamento, S.plano, S.painel], focusTabs: [], defaultTab: S.fluxo,
      steps: {}
    },
    p1: {
      key: 'p1', label: 'Projeto 1 · Gestão de pequenos negócios', suffix: '_projeto1',
      subtitle: 'Projeto 1 · Gestão de pequenos negócios — registro financeiro, caixa, custos e preço',
      intro: 'Neste projeto, o foco é registrar o dinheiro, acompanhar o caixa, conhecer os custos e revisar critérios de preço. As abas coloridas seguem essa ordem; as cinzas são complementares.',
      order: [S.leia, S.painel, S.fluxo, S.contas, S.preco, S.plano, S.clientes, S.rotina],
      primary: [S.leia, S.painel, S.fluxo, S.contas, S.preco, S.plano], focusTabs: [S.fluxo, S.contas, S.preco], defaultTab: S.fluxo
    },
    p2: {
      key: 'p2', label: 'Projeto 2 · Organização e desenvolvimento', suffix: '_projeto2',
      subtitle: 'Projeto 2 · Organização e desenvolvimento — rotina, prioridades, atendimento e metas',
      intro: 'Neste projeto, o foco é a rotina, as prioridades, o atendimento e a comunicação com clientes e as metas. As abas coloridas seguem essa ordem; as cinzas são complementares.',
      order: [S.leia, S.painel, S.rotina, S.clientes, S.plano, S.fluxo, S.contas, S.preco],
      primary: [S.leia, S.painel, S.rotina, S.clientes, S.plano], focusTabs: [S.rotina, S.clientes, S.plano], defaultTab: S.rotina
    }
  };
  ACTIVITIES.workshop.steps[S.fluxo] = 'Oficina · passo 1 de 4';
  ACTIVITIES.workshop.steps[S.contas] = 'Oficina · passo 2 de 4';
  ACTIVITIES.workshop.steps[S.orcamento] = 'Oficina · passo 3 de 4';
  ACTIVITIES.workshop.steps[S.plano] = 'Oficina · passo 4 de 4';

  function activityFor(value) {
    var v = String(value || '').trim();
    if (ACTIVITIES[v] && v !== 'universal') return ACTIVITIES[v];
    if (/workshop|oficina/i.test(v)) return ACTIVITIES.workshop;
    if (/projeto\s*1\b/i.test(v)) return ACTIVITIES.p1;
    if (/projeto\s*2\b/i.test(v)) return ACTIVITIES.p2;
    return ACTIVITIES.universal;
  }
  function startTabFor(activity, focus) {
    if (!activity.defaultTab) return focus.tab;
    return activity.focusTabs.indexOf(focus.tab) >= 0 ? focus.tab : activity.defaultTab;
  }

  // Meta simples para cada sugestão do diagnóstico (id da pergunta em avaliacao.html).
  var SUGGESTION_GOALS = {
    record: '5 de 7 dias com registro', separate: 'Todas as retiradas da semana marcadas', price: '1 item revisado',
    bills: 'Próximos 7 dias listados toda semana', customers: 'Lista revisada em 5 dias da semana', routine: 'Até 3 prioridades em 5 dias'
  };
  var FOCUS_SUGGESTION = { cash: 'record', price: 'price', customers: 'customers', routine: 'routine' };
  var TOOLS = { paper: 'paper', phone: 'phone', sheet: 'sheet' };
  function clean(value, max) { return String(value || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }
  function normalizeSuggestion(item) {
    if (typeof item === 'string' || !item) return { id: '', title: clean(item, 160), action: '', check: '' };
    return { id: clean(item.id, 20), title: clean(item.title, 160), action: clean(item.action, 400), check: clean(item.check, 400) };
  }

  function personalize(opts) {
    opts = opts || {};
    var name = String(opts.business || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, 100);
    var focus = FOCUS[opts.priority] || FOCUS.unsure;
    var activity = activityFor(opts.activity);
    var recs = (opts.recommendations || []).slice(0, 3).map(normalizeSuggestion).filter(function (s) { return s.title; });
    return {
      name: name || 'Meu negócio', hasName: !!name, focus: focus, focusKey: FOCUS[opts.priority] ? opts.priority : 'unsure',
      recommendations: recs, activity: activity, tool: TOOLS[opts.tool] || 'neutral',
      startTab: startTabFor(activity, focus), today: opts.today ? new Date(opts.today) : new Date()
    };
  }
  function isWorkshop(p) { return p.activity.key === 'workshop'; }
  function isSecondary(p, name) { return !!p.activity.primary && p.activity.primary.indexOf(name) < 0; }
  function tabColor(p, name) {
    var a = p.activity;
    if (!a.primary) return UNIVERSAL_TABS[name];
    var k = a.primary.indexOf(name);
    return k < 0 ? COLOR.secondaryTab : STEP_TABS[Math.min(k, STEP_TABS.length - 1)];
  }

  function colLetter(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function q(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }
  function ref(sheet, range) { return "'" + sheet + "'!" + range; }
  function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }
  function cfFill(argb) { return { type: 'pattern', pattern: 'solid', bgColor: { argb: argb } }; }
  function font(o) { var f = { name: FONT, size: 11, color: { argb: COLOR.text } }; for (var k in o) f[k] = o[k]; return f; }
  var THIN = { style: 'thin', color: { argb: COLOR.line } };
  var BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  function utcDate(base, offsetDays) { return new Date(Date.UTC(base.getFullYear(), base.getMonth(), base.getDate() + (offsetDays || 0))); }
  // Altura aproximada para textos com quebra (o Excel não ajusta células mescladas sozinho).
  function textHeight(text, widthChars, size, minHeight) {
    var lines = 0, cpl = Math.max(8, widthChars * (11 / (size || 11)) * 1.05);
    String(text || '').split('\n').forEach(function (part) { lines += Math.max(1, Math.ceil(part.length / cpl)); });
    return Math.max(minHeight || 15, Math.ceil(lines * (size || 11) * 1.32 + 6));
  }

  function listValidation(items, strict) {
    var f = q(items.join(','));
    if (f.length > 255) throw new Error('Lista de validação longa demais: ' + f);
    return {
      type: 'list', allowBlank: true, formulae: [f], showErrorMessage: true,
      errorStyle: strict ? 'stop' : 'information',
      errorTitle: strict ? 'Escolha uma opção' : 'Valor fora da lista sugerida',
      error: strict ? 'Escolha uma das opções da lista.' : 'Este valor não está na lista sugerida. Pode manter se fizer sentido para o seu negócio.'
    };
  }
  var DATE_VALIDATION = { type: 'date', operator: 'greaterThan', allowBlank: true, formulae: [new Date(Date.UTC(2000, 0, 1))], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Data inválida', error: 'Digite uma data, por exemplo 05/03/2026.' };
  var MONEY_VALIDATION = { type: 'decimal', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [0], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valor inválido', error: 'Digite um número maior ou igual a zero, sem o "R$".' };
  var HOURS_VALIDATION = { type: 'decimal', operator: 'between', allowBlank: true, formulae: [0, 1000], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Horas inválidas', error: 'Digite horas em número. Ex.: 1,5 para uma hora e meia.' };

  var LISTS = {
    tipoFluxo: ['Entrada', 'Saída', 'Retirada pessoal', 'Saldo inicial'],
    categoria: ['Venda de produto', 'Prestação de serviço', 'Encomenda', 'Material/mercadoria', 'Embalagem', 'Transporte', 'Aluguel/espaço', 'Contas (água/luz/internet)', 'Taxas e tarifas', 'Impostos', 'Manutenção/equipamento', 'Divulgação', 'Retirada pessoal', 'Outros'],
    pagamento: ['Dinheiro', 'Pix', 'Cartão de débito', 'Cartão de crédito', 'Boleto', 'Transferência', 'A prazo/fiado', 'Outro'],
    simNao: ['Sim', 'Não'],
    tipoConta: ['A pagar', 'A receber'],
    categoriaConta: ['Fornecedor', 'Aluguel/espaço', 'Contas (água/luz/internet)', 'Empréstimo/parcelamento', 'Impostos e taxas', 'Venda a prazo', 'Serviço prestado', 'Encomenda', 'Outros'],
    statusConta: ['Em aberto', 'Parcial', 'Pago', 'Recebido', 'Cancelado'],
    statusPedido: ['Novo', 'A fazer', 'Em andamento', 'Aguardando cliente', 'Concluído', 'Cancelado'],
    contato: ['WhatsApp', 'Telefone', 'Presencial', 'E-mail', 'Rede social', 'Outro'],
    statusTarefa: ['A fazer', 'Em andamento', 'Concluída', 'Adiada', 'Cancelada'],
    statusPlano: ['A fazer', 'Em andamento', 'Concluída', 'Pausada', 'Cancelada'],
    tipoNegocio: ['Vendo produtos', 'Presto serviços', 'Faço produtos sob encomenda', 'Faço os dois', 'Ainda estou entendendo'],
    incomeCats: ['Venda de produto', 'Prestação de serviço', 'Encomenda'],
    expenseCats: ['Material/mercadoria', 'Embalagem', 'Transporte', 'Aluguel/espaço', 'Contas (água/luz/internet)', 'Taxas e tarifas', 'Impostos', 'Manutenção/equipamento', 'Divulgação', 'Outros']
  };

  // Linhas de dados (cabeçalho na linha 4, dados a partir da 5).
  var FIRST = 5;
  var ROWS = { fluxo: 300, contas: 150, clientes: 150, preco: 60, rotina: 150, plano: 40 };
  var PRINT_ROWS = 25;
  function lastRow(key) { return FIRST + ROWS[key] - 1; }

  function sheetOptions(p, name, frozenRows) {
    var view = { showGridLines: false, zoomScale: 100 };
    if (frozenRows) { view.state = 'frozen'; view.xSplit = 0; view.ySplit = frozenRows; view.topLeftCell = 'A' + (frozenRows + 1); }
    return { properties: { tabColor: { argb: tabColor(p, name) } }, views: [view] };
  }
  function pageSetup(ws, area, landscape, titleRows, fitOnePage) {
    ws.pageSetup.paperSize = 9;
    ws.pageSetup.orientation = landscape ? 'landscape' : 'portrait';
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = fitOnePage ? 1 : 0;
    ws.pageSetup.horizontalCentered = true;
    ws.pageSetup.margins = { left: 0.35, right: 0.35, top: 0.45, bottom: 0.5, header: 0.25, footer: 0.25 };
    ws.pageSetup.printArea = area;
    if (titleRows) ws.pageSetup.printTitlesRow = titleRows;
    ws.headerFooter.oddFooter = '&L&8&A&R&8Página &P de &N';
  }
  function exampleRule(col) {
    return { type: 'expression', formulae: ['LEFT($' + col + FIRST + ',7)="EXEMPLO"'], style: { font: { italic: true, color: { argb: COLOR.exFont } }, fill: cfFill(COLOR.exFill) } };
  }
  function exprRule(formula, bg, fg, bold) {
    var style = { fill: cfFill(bg) };
    if (fg) style.font = { color: { argb: fg }, bold: !!bold };
    return { type: 'expression', formulae: [formula], style: style };
  }
  function withPriority(rules, start) { return rules.map(function (r, k) { r.priority = (start || 1) + k; return r; }); }

  function titleBand(ws, range, text, p, name, size) {
    ws.mergeCells(range);
    var cell = ws.getCell(range.split(':')[0]);
    cell.value = text;
    cell.font = font({ size: size || 16, bold: true, color: { argb: COLOR.white } });
    cell.fill = fill(isSecondary(p, name) ? COLOR.secondary : COLOR.dark);
    cell.alignment = { vertical: 'middle', indent: 1 };
    return cell;
  }
  function stepTitle(p, name, text) {
    var step = p.activity.steps && p.activity.steps[name];
    return step ? step + ' — ' + text : text;
  }
  function secondaryNote(p, name) {
    if (!isSecondary(p, name)) return '';
    return isWorkshop(p) ? 'Aba complementar: não faz parte do roteiro de hoje. Use depois da oficina, se fizer sentido. ' : 'Aba complementar para esta atividade. ';
  }

  // Nunca reduz a altura de uma linha já usada pela tabela principal.
  function raise(ws, r, h) { var row = ws.getRow(r); if (!row.height || row.height < h) row.height = h; }

  // Quadro lateral (fora da área de impressão) com rótulo à esquerda e valor/fórmula à direita.
  function sideBox(ws, labelCol, valueCol, startRow, heading, items, p, name) {
    ws.mergeCells(labelCol + startRow + ':' + valueCol + startRow);
    var h = ws.getCell(labelCol + startRow);
    h.value = heading;
    h.font = font({ bold: true, size: 10.5, color: { argb: COLOR.white } });
    h.fill = fill(isSecondary(p, name) ? COLOR.secondary : COLOR.header);
    h.alignment = { vertical: 'middle', indent: 1 };
    raise(ws, startRow, 22);
    var row = startRow + 1;
    items.forEach(function (it) {
      var lc = ws.getCell(labelCol + row), vc = ws.getCell(valueCol + row);
      if (it.note) {
        ws.mergeCells(labelCol + row + ':' + valueCol + row);
        lc.value = it.formula ? { formula: it.note } : it.note;
        lc.font = font({ size: 9.5, italic: !it.formula, bold: !!it.formula, color: { argb: it.formula ? COLOR.text : COLOR.muted } });
        lc.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
        if (it.formula) lc.fill = fill(COLOR.info);
        raise(ws, row, it.height || 20);
        if (it.cf) ws.addConditionalFormatting({ ref: labelCol + row, rules: withPriority(it.cf) });
        row++;
        return;
      }
      lc.value = it.label;
      lc.font = font({ size: 10.5, bold: !!it.strong });
      lc.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
      lc.fill = fill(it.input ? COLOR.info : COLOR.calc);
      lc.border = BOX;
      if (it.formula) vc.value = { formula: it.formula };
      vc.numFmt = it.fmt || BRL;
      vc.font = font({ size: 10.5, bold: !!it.strong });
      vc.fill = fill(it.input ? COLOR.input : COLOR.calc);
      vc.border = BOX;
      vc.alignment = { horizontal: it.fmt === DATE || it.list ? 'center' : 'right', vertical: 'middle' };
      if (it.validation) ws.dataValidations.add(valueCol + row, it.validation);
      if (it.strong) { lc.border = vc.border = { top: { style: 'thin', color: { argb: COLOR.muted } }, left: THIN, right: THIN, bottom: { style: 'double', color: { argb: COLOR.muted } } }; }
      raise(ws, row, it.height || 20);
      row++;
    });
    return row;
  }

  function buildDataSheet(wb, p, s) {
    var ws = wb.addWorksheet(s.name, sheetOptions(p, s.name, 4));
    var cols = s.columns, n = cols.length, L = colLetter(n), end = lastRow(s.key);
    // Quem prefere papel recebe linhas impressas mais altas, para escrever à mão.
    function rowHeight(r) { return p.tool === 'paper' && r < FIRST + PRINT_ROWS ? 24 : 18; }
    ws.columns = cols.map(function (c) { return { width: c.w }; });
    var totalWidth = cols.reduce(function (a, c) { return a + c.w; }, 0);

    titleBand(ws, 'A1:' + L + '1', stepTitle(p, s.name, s.title), p, s.name);
    ws.getRow(1).height = 30;

    ws.mergeCells('A2:' + L + '2');
    var h = ws.getCell('A2');
    var helpText = secondaryNote(p, s.name) + s.help;
    h.value = helpText;
    h.font = font({ size: 10, color: { argb: COLOR.muted } });
    h.fill = fill(COLOR.helpFill);
    h.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
    ws.getRow(2).height = textHeight(helpText, totalWidth - 2, 10, 30);

    cols.forEach(function (c, i) {
      var calc = c.kind === 'calc', opt = c.kind === 'opt';
      var tag = ws.getCell(3, i + 1);
      tag.value = calc ? '= calculado' : opt ? 'opcional' : '✎ preencher';
      tag.font = font({ size: 8.5, italic: true, color: { argb: calc || opt ? COLOR.muted : COLOR.infoFont } });
      tag.alignment = { horizontal: 'center', vertical: 'bottom' };
      var hd = ws.getCell(4, i + 1);
      hd.value = c.h;
      hd.font = font({ bold: true, size: 10.5, color: { argb: COLOR.white } });
      hd.fill = fill(calc ? COLOR.headerCalc : isSecondary(p, s.name) ? COLOR.secondary : COLOR.header);
      hd.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
      hd.border = { top: THIN, bottom: THIN, left: { style: 'thin', color: { argb: COLOR.white } }, right: { style: 'thin', color: { argb: COLOR.white } } };
    });
    ws.getRow(3).height = 15;
    ws.getRow(4).height = 32;

    for (var r = FIRST; r <= end; r++) {
      ws.getRow(r).height = rowHeight(r);
      for (var i = 0; i < n; i++) {
        var c = cols[i], cell = ws.getCell(r, i + 1);
        cell.fill = fill(c.kind === 'calc' ? COLOR.calc : COLOR.input);
        cell.border = BOX;
        cell.font = font({ size: 10.5 });
        if (c.fmt) cell.numFmt = c.fmt;
        cell.alignment = c.fmt === BRL || c.align === 'right' ? { horizontal: 'right', vertical: 'middle' }
          : c.fmt === DATE || c.align === 'center' ? { horizontal: 'center', vertical: 'middle' } : { wrapText: !!c.wrap, vertical: 'middle' };
        if (c.f) cell.value = { formula: c.f(r) };
      }
    }
    (s.examples || []).forEach(function (row, k) {
      var height = rowHeight(FIRST + k);
      row.forEach(function (v, i) {
        if (v === undefined || v === null || cols[i].f) return;
        ws.getCell(FIRST + k, i + 1).value = v;
        if (cols[i].wrap && typeof v === 'string') height = Math.max(height, textHeight(v, cols[i].w - 1, 10.5, 18));
      });
      ws.getRow(FIRST + k).height = height;
    });

    cols.forEach(function (c, i) {
      if (!c.v) return;
      var letter = colLetter(i + 1);
      ws.dataValidations.add(letter + FIRST + ':' + letter + end, typeof c.v === 'function' ? c.v(letter) : c.v);
    });

    ws.autoFilter = 'A4:' + L + end;
    var rules = (s.rules ? s.rules(end) : []).concat([exampleRule(s.exampleCol)]);
    ws.addConditionalFormatting({ ref: 'A' + FIRST + ':' + L + end, rules: withPriority(rules) });
    if (s.extraCF) s.extraCF(end).forEach(function (cf) { ws.addConditionalFormatting(cf); });
    pageSetup(ws, 'A1:' + L + (FIRST + PRINT_ROWS - 1), true, '4:4', true);
    if (s.extra) s.extra(ws, end);
    return ws;
  }

  var EXAMPLE_NOTE = ' Linhas marcadas "' + EX + '" são fictícias: servem só para aprender e não entram em nenhum total.';

  function fluxoSheet(wb, p) {
    var crit = '"<>Não",L$' + FIRST + ':L';
    function saldo(r) {
      var range = function (col) { return col + '$' + FIRST + ':' + col + r; };
      return 'IF(AND(H' + r + '="",I' + r + '=""),"",IF(L' + r + '=' + q(EX) + ',"(exemplo)",IF(K' + r + '="Não","(não realizado)",' +
        'SUMIFS(' + range('H') + ',' + range('K') + ',' + crit + r + ',' + q('<>' + EX) + ')-' +
        'SUMIFS(' + range('I') + ',' + range('K') + ',' + crit + r + ',' + q('<>' + EX) + '))))';
    }
    var t = p.today, ws = isWorkshop(p);
    var end = lastRow('fluxo');
    function R(col) { return '$' + col + '$' + FIRST + ':$' + col + '$' + end; }
    var NEX = q('<>' + EX), real = R('K') + ',"<>Não",' + R('L') + ',' + NEX;
    var nF = '(COUNTIFS(' + R('H') + ',"<>",' + R('L') + ',' + NEX + ')+COUNTIFS(' + R('I') + ',"<>",' + R('L') + ',' + NEX + '))';
    function guard(f) { return 'IF(' + nF + '=0,"—",' + f + ')'; }
    var examples = ws ? [
      [utcDate(t, -6), 'Saldo inicial', 'Caso A (fictício): dinheiro disponível no início', null, null, null, 'Dinheiro', 300, null, null, 'Sim', EX],
      [utcDate(t, -5), 'Entrada', 'Caso A: recebimento por afiação (serviço)', 'Prestação de serviço', 'Cliente (fictício)', 'Afiação', 'Pix', 240, null, null, 'Sim', EX],
      [utcDate(t, -5), 'Saída', 'Caso A: insumos de afiação', 'Material/mercadoria', 'Fornecedor (fictício)', null, 'Pix', null, 80, null, 'Sim', EX],
      [utcDate(t, -4), 'Entrada', 'Caso A: pagamento de cliente antigo', 'Prestação de serviço', 'Cliente (fictício)', 'Afiação', 'Dinheiro', 150, null, null, 'Sim', EX],
      [utcDate(t, -3), 'Saída', 'Caso A: energia do período', 'Contas (água/luz/internet)', null, null, 'Boleto', null, 90, null, 'Sim', EX],
      [utcDate(t, -2), 'Retirada pessoal', 'Caso A: retirada para uso pessoal (não é despesa do serviço)', 'Retirada pessoal', null, null, 'Dinheiro', null, 100, null, 'Sim', EX]
    ] : [
      [utcDate(t, -4), 'Entrada', 'Venda de 2 camisetas (produto)', 'Venda de produto', 'Cliente A', 'Camiseta', 'Pix', 120, null, null, 'Sim', EX],
      [utcDate(t, -3), 'Entrada', 'Afiação de 3 facas (serviço)', 'Prestação de serviço', 'Cliente B', 'Afiação', 'Dinheiro', 45, null, null, 'Sim', EX],
      [utcDate(t, -3), 'Saída', 'Compra de material para os próximos pedidos', 'Material/mercadoria', 'Fornecedor X', 'Material', 'Pix', null, 60, null, 'Sim', EX],
      [utcDate(t, -2), 'Retirada pessoal', 'Dinheiro do negócio usado em casa', 'Retirada pessoal', null, null, 'Pix', null, 50, null, 'Sim', EX],
      [utcDate(t, -1), 'Entrada', 'Encomenda entregue, cliente vai pagar depois (combinado)', 'Encomenda', 'Cliente C', 'Bolo sob encomenda', 'A prazo/fiado', 80, null, null, 'Não', EX]
    ];
    var help = ws
      ? 'Caderno p. 7 · Meu controle simples. Caixa é movimento de dinheiro: registre só quando o dinheiro efetivamente entra ou sai. Comece com uma linha "Saldo inicial" (dinheiro disponível no começo). ' +
        'Use Entrada OU Saída, nunca as duas na mesma linha. Compra pessoal paga com dinheiro do negócio = "Retirada pessoal": reduz o caixa, mas não é despesa do negócio. ' +
        'Venda a receber não entra aqui: anote em CONTAS E COMPROMISSOS. Pode usar números fictícios. As linhas de exemplo são o Caso A do caderno, sem as respostas.' + EXAMPLE_NOTE
      : 'Registre só o que já aconteceu. Use a coluna Entrada OU a coluna Saída (nunca as duas na mesma linha). Comece, se quiser, com uma linha "Saldo inicial". Dinheiro do negócio usado em casa = tipo "Retirada pessoal": ele sai do caixa, mas não é despesa do negócio. ' +
        'Se o valor foi só combinado e ainda não foi pago/recebido, marque "Não" em Pago/recebido? — ele não entra no saldo (prefira anotá-lo em CONTAS E COMPROMISSOS).' + EXAMPLE_NOTE;
    return buildDataSheet(wb, p, {
      name: S.fluxo, key: 'fluxo', exampleCol: 'L',
      title: ws ? 'MEU CONTROLE SIMPLES — fluxo de caixa' : 'FLUXO DE CAIXA — uma linha para cada dinheiro que entra ou sai',
      help: help,
      columns: [
        { h: 'Data', w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Tipo', w: 15, v: listValidation(LISTS.tipoFluxo, true) },
        { h: 'Descrição', w: 34, wrap: true },
        { h: 'Categoria', w: 21, v: listValidation(LISTS.categoria, false) },
        { h: 'Cliente/fornecedor', w: 17, wrap: true, kind: 'opt' },
        { h: 'Produto/serviço', w: 15, wrap: true, kind: 'opt' },
        { h: 'Forma de pagamento', w: 13, v: listValidation(LISTS.pagamento, false), kind: 'opt' },
        { h: 'Entrada (R$)', w: 13.5, fmt: BRL, v: { type: 'custom', allowBlank: true, formulae: ['AND(ISNUMBER(H' + FIRST + '),H' + FIRST + '>=0,ISBLANK(I' + FIRST + '))'], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Entrada ou saída', error: 'Digite um valor positivo. Se esta linha já tem Saída, use outra linha para a entrada.' } },
        { h: 'Saída (R$)', w: 13.5, fmt: BRL, v: { type: 'custom', allowBlank: true, formulae: ['AND(ISNUMBER(I' + FIRST + '),I' + FIRST + '>=0,ISBLANK(H' + FIRST + '))'], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Entrada ou saída', error: 'Digite um valor positivo. Se esta linha já tem Entrada, use outra linha para a saída.' } },
        { h: 'Saldo acumulado (R$)', w: 15, fmt: BRL, kind: 'calc', f: saldo },
        { h: 'Pago/ recebido?', w: 10, align: 'center', v: listValidation(LISTS.simNao, true) },
        { h: 'Observação', w: 30, wrap: true }
      ],
      examples: examples,
      rules: function () {
        return [
          exprRule('AND($H' + FIRST + '<>"",$I' + FIRST + '<>"")', COLOR.bad, COLOR.badFont, true),
          exprRule('OR(AND(OR($B' + FIRST + '="Entrada",$B' + FIRST + '="Saldo inicial"),$I' + FIRST + '<>""),AND(OR($B' + FIRST + '="Saída",$B' + FIRST + '="Retirada pessoal"),$H' + FIRST + '<>""))', COLOR.bad, COLOR.badFont, true),
          exprRule('$K' + FIRST + '="Não"', COLOR.warn, COLOR.warnFont)
        ];
      },
      extraCF: function (e) {
        return [{ ref: 'B' + FIRST + ':B' + e, rules: [
          { type: 'cellIs', operator: 'equal', formulae: ['"Entrada"'], priority: 10, style: { font: { color: { argb: COLOR.goodFont }, bold: true } } },
          { type: 'cellIs', operator: 'equal', formulae: ['"Saída"'], priority: 11, style: { font: { color: { argb: COLOR.badFont }, bold: true } } },
          { type: 'cellIs', operator: 'equal', formulae: ['"Retirada pessoal"'], priority: 12, style: { font: { color: { argb: COLOR.infoFont }, bold: true } } },
          { type: 'cellIs', operator: 'equal', formulae: ['"Saldo inicial"'], priority: 13, style: { font: { color: { argb: COLOR.muted }, bold: true } } }
        ] }];
      },
      extra: function (sheet) {
        sheet.getColumn(13).width = 3;
        sheet.getColumn(14).width = 56;
        sheet.getColumn(15).width = 16;
        var row = sideBox(sheet, 'N', 'O', 4, 'CONFERÊNCIA (sem exemplos)', [
          { label: 'Saldo inicial', formula: guard('SUMIFS(' + R('H') + ',' + R('B') + ',"Saldo inicial",' + real + ')') },
          { label: '+ Total de entradas', formula: guard('SUMIFS(' + R('H') + ',' + R('B') + ',"<>Saldo inicial",' + real + ')') },
          { label: '− Total de saídas (inclui retiradas)', formula: guard('SUMIFS(' + R('I') + ',' + real + ')') },
          { label: '= Saldo final', strong: true, formula: guard('SUMIFS(' + R('H') + ',' + real + ')-SUMIFS(' + R('I') + ',' + real + ')') },
          { note: 'Saldo inicial + total de entradas − total de saídas = saldo final. Deve ser igual ao último saldo da coluna J. Saldo não é lucro.' },
          { label: 'Retiradas pessoais (já dentro das saídas)', formula: guard('SUMIFS(' + R('I') + ',' + R('B') + ',"Retirada pessoal",' + real + ')') },
          { label: 'Marcado "Não" (fora do saldo)', formula: guard('SUMIFS(' + R('H') + ',' + R('K') + ',"Não",' + R('L') + ',' + NEX + ')+SUMIFS(' + R('I') + ',' + R('K') + ',"Não",' + R('L') + ',' + NEX + ')') }
        ], p, S.fluxo);
        sideBox(sheet, 'N', 'O', row + 1, 'ANTES DE FECHAR O DIA', [
          { label: 'Conferi se registrei o recebimento apenas uma vez?', input: true, fmt: '@', list: true, validation: listValidation(LISTS.simNao, true) },
          { label: 'Separei as retiradas pessoais?', input: true, fmt: '@', list: true, validation: listValidation(LISTS.simNao, true) },
          { label: 'Vi contas que vencem antes dos recebimentos?', input: true, fmt: '@', list: true, validation: listValidation(LISTS.simNao, true) }
        ], p, S.fluxo);
      }
    });
  }

  function contasSheet(wb, p) {
    var t = p.today, ws = isWorkshop(p), end = lastRow('contas');
    function R(col) { return '$' + col + '$' + FIRST + ':$' + col + '$' + end; }
    function F(col) { return ref(S.fluxo, '$' + col + '$' + FIRST + ':$' + col + '$' + lastRow('fluxo')); }
    var NEX = q('<>' + EX);
    var nF = '(COUNTIFS(' + F('H') + ',"<>",' + F('L') + ',' + NEX + ')+COUNTIFS(' + F('I') + ',"<>",' + F('L') + ',' + NEX + '))';
    var nC = 'COUNTIFS(' + R('A') + ',"<>",' + R('K') + ',' + NEX + ')';
    var realF = F('K') + ',"<>Não",' + F('L') + ',' + NEX;
    function open(tipo) {
      var base = R('A') + ',' + q(tipo) + ',' + R('K') + ',' + NEX + ',' + R('E') + ',"<="&$N$5';
      return 'IF(' + nC + '=0,"—",SUMIFS(' + R('F') + ',' + base + ',' + R('H') + ',"Em aberto")+SUMIFS(' + R('F') + ',' + base + ',' + R('H') + ',"Parcial")-SUMIFS(' + R('G') + ',' + base + ',' + R('H') + ',"Parcial"))';
    }
    var none = 'AND(' + nF + '=0,' + nC + '=0)';
    var examples = ws ? [
      ['A receber', 'Caso A (fictício): serviço de afiação concluído, pago só na próxima semana', 'Cliente (fictício)', 'Serviço prestado', utcDate(t, 7), 180, 0, 'Em aberto', null, 'Não contar como recebido antes da hora', EX],
      ['A receber', 'Caso B (fictício): venda de roupas de R$ 200 — R$ 50 pagos, falta o resto (produto)', 'Cliente (fictício)', 'Venda a prazo', utcDate(t, 5), 150, 0, 'Em aberto', null, 'Confirmar a data com o cliente', EX],
      ['A pagar', 'Caso B (fictício): compromissos da semana seguinte', 'Fornecedor (fictício)', 'Fornecedor', utcDate(t, 6), 450, 0, 'Em aberto', null, 'Ver se o saldo previsto cobre', EX]
    ] : [
      ['A pagar', 'Aluguel do espaço', 'Proprietário', 'Aluguel/espaço', utcDate(t, 5), 300, 0, 'Em aberto', null, 'Separar o valor até sexta', EX],
      ['A receber', 'Serviço de manutenção entregue; cliente pagou metade', 'Cliente D', 'Serviço prestado', utcDate(t, 3), 150, 75, 'Parcial', null, 'Lembrar o cliente por mensagem', EX],
      ['A pagar', 'Conta de internet', 'Operadora', 'Contas (água/luz/internet)', utcDate(t, -2), 90, 0, 'Em aberto', null, 'Pagar hoje', EX]
    ];
    return buildDataSheet(wb, p, {
      name: S.contas, key: 'contas', exampleCol: 'K',
      title: ws ? 'COMPROMISSOS FUTUROS — o que vou receber e pagar' : 'CONTAS E COMPROMISSOS — o que pagar e o que receber',
      help: (ws ? 'Caderno p. 7 · Compromissos futuros: não misturar com o caixa realizado. Planejar é antecipar datas. ' : 'Aqui ficam PREVISÕES: contas a pagar e valores a receber, com vencimento. ') +
        'Isso não é dinheiro em caixa. Quando pagar ou receber, atualize o Status aqui e registre o movimento no FLUXO DE CAIXA. O quadro PREVISÃO, à direita, mostra o saldo previsto e o que acontece se nada for recebido. ' +
        'Cores: ✖ vermelho = vencido · ⚠ amarelo = vence nos próximos 7 dias · ✔ verde = pago/recebido.' + EXAMPLE_NOTE,
      columns: [
        { h: 'Tipo', w: 11, v: listValidation(LISTS.tipoConta, true) },
        { h: 'Descrição', w: 34, wrap: true },
        { h: 'Pessoa/empresa relacionada', w: 18, wrap: true, kind: 'opt' },
        { h: 'Categoria', w: 23, v: listValidation(LISTS.categoriaConta, false) },
        { h: 'Data de vencimento', w: 12, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Valor previsto', w: 13.5, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Valor pago/ recebido', w: 13.5, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Status', w: 12, v: listValidation(LISTS.statusConta, true) },
        { h: 'Dias até o vencimento', w: 19, kind: 'calc', align: 'right', fmt: '"faltam "0" dias";"⚠ vencida há "0" dias";"⚠ vence hoje"',
          f: function (r) { return 'IF(NOT(ISNUMBER(E' + r + ')),"",IF(OR(H' + r + '="Pago",H' + r + '="Recebido",H' + r + '="Cancelado"),"",E' + r + '-TODAY()))'; } },
        { h: 'Próxima ação', w: 24, wrap: true },
        { h: 'Observação', w: 22, wrap: true }
      ],
      examples: examples,
      rules: function () {
        var openSt = '$H' + FIRST + '<>"Pago",$H' + FIRST + '<>"Recebido",$H' + FIRST + '<>"Cancelado"';
        return [
          exprRule('OR($H' + FIRST + '="Pago",$H' + FIRST + '="Recebido")', COLOR.good, COLOR.goodFont),
          exprRule('AND(ISNUMBER($E' + FIRST + '),$E' + FIRST + '<TODAY(),' + openSt + ')', COLOR.bad, COLOR.badFont, true),
          exprRule('AND(ISNUMBER($E' + FIRST + '),$E' + FIRST + '>=TODAY(),$E' + FIRST + '<=TODAY()+7,' + openSt + ')', COLOR.warn, COLOR.warnFont)
        ];
      },
      extra: function (sheet) {
        sheet.getColumn(12).width = 3;
        sheet.getColumn(13).width = 56;
        sheet.getColumn(14).width = 16;
        sideBox(sheet, 'M', 'N', 4, 'PREVISÃO (não é dinheiro em caixa)', [
          { label: '✎ Olhar compromissos até', input: true, fmt: DATE, formula: 'TODAY()+7', validation: DATE_VALIDATION },
          { label: 'Saldo atual (FLUXO DE CAIXA)', formula: 'IF(' + nF + '=0,"—",SUMIFS(' + F('H') + ',' + realF + ')-SUMIFS(' + F('I') + ',' + realF + '))' },
          { label: '+ A receber até a data (em aberto)', formula: open('A receber') },
          { label: '− A pagar até a data (em aberto)', formula: open('A pagar') },
          { label: '= Saldo previsto', strong: true, formula: 'IF(' + none + ',"—",N(N6)+N(N7)-N(N8))' },
          { label: 'E se nada for recebido?', formula: 'IF(' + none + ',"—",N(N6)-N(N8))' },
          { note: 'IF(' + none + ',"ℹ Registre o caixa e os compromissos para ver a previsão.",IF(N(N9)<0,          "⚠ Mesmo recebendo tudo, o saldo previsto fica negativo. Planeje antes do vencimento.",IF(N(N10)<0,"⚠ Se os recebimentos atrasarem, o saldo fica negativo. Combine alternativas.","✔ Pelo previsto, o saldo cobre os compromissos até a data.")))',
                      formula: true, cf: [exprRule('LEFT($M11,1)="⚠"', COLOR.warn, COLOR.warnFont, true), exprRule('LEFT($M11,1)="✔"', COLOR.good, COLOR.goodFont, true)] },
          { note: 'Previsão não é promessa: não conte pendentes como recebidos. Sem data, não entra.' }
        ], p, S.contas);
      }
    });
  }

  function clientesSheet(wb, p) {
    var t = p.today;
    return buildDataSheet(wb, p, {
      name: S.clientes, key: 'clientes', exampleCol: 'K',
      title: 'CLIENTES E PEDIDOS — pedidos, encomendas e serviços combinados',
      help: 'Privacidade: registre só o necessário. Use apelido ou referência curta do cliente (ex.: "Ana — vizinha"). Não anote CPF, senhas, número de cartão, dados bancários ou informações pessoais desnecessárias. ' +
        'Valor combinado é previsão; o dinheiro só entra no caixa quando for recebido (registre no FLUXO DE CAIXA). Cores: ✖ prazo vencido · ⚠ prazo em até 3 dias · ✔ concluído.' + EXAMPLE_NOTE,
      columns: [
        { h: 'Data de entrada', w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Cliente ou referência curta', w: 20, wrap: true },
        { h: 'Produto/serviço', w: 19, wrap: true },
        { h: 'Descrição do combinado', w: 32, wrap: true },
        { h: 'Prazo prometido', w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Valor combinado', w: 13.5, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Valor recebido', w: 13.5, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Status', w: 17, v: listValidation(LISTS.statusPedido, true) },
        { h: 'Próximo contato', w: 11.5, fmt: DATE, v: DATE_VALIDATION, kind: 'opt' },
        { h: 'Forma de contato', w: 13, v: listValidation(LISTS.contato, false), kind: 'opt' },
        { h: 'Observação', w: 26, wrap: true }
      ],
      examples: [
        [utcDate(t, -3), 'Cliente A (loja)', 'Camisetas personalizadas', '10 camisetas tamanho M, estampa da frente', utcDate(t, 4), 350, 100, 'Em andamento', utcDate(t, 2), 'WhatsApp', EX],
        [utcDate(t, -2), 'Cliente B (oficina)', 'Afiação de tesouras', '4 tesouras; cliente busca no balcão', utcDate(t, 1), 60, 0, 'A fazer', utcDate(t, 1), 'Presencial', EX],
        [utcDate(t, -6), 'Cliente E', 'Reparo de cadeira (serviço)', 'Trocar assento; aguardando escolha do tecido', utcDate(t, -1), 120, 0, 'Aguardando cliente', utcDate(t, 0), 'Telefone', EX]
      ],
      rules: function () {
        var open = '$H' + FIRST + '<>"Concluído",$H' + FIRST + '<>"Cancelado"';
        return [
          exprRule('$H' + FIRST + '="Concluído"', COLOR.good, COLOR.goodFont),
          exprRule('$H' + FIRST + '="Cancelado"', COLOR.calc, COLOR.muted),
          exprRule('AND(ISNUMBER($E' + FIRST + '),$E' + FIRST + '<TODAY(),' + open + ')', COLOR.bad, COLOR.badFont, true),
          exprRule('AND(ISNUMBER($E' + FIRST + '),$E' + FIRST + '<=TODAY()+3,' + open + ')', COLOR.warn, COLOR.warnFont)
        ];
      }
    });
  }

  function precoSheet(wb, p) {
    function blank(r) { return 'AND(A' + r + '="",COUNT(B' + r + ':G' + r + ')=0)'; }
    return buildDataSheet(wb, p, {
      name: S.preco, key: 'preco', exampleCol: 'M',
      title: 'PRECIFICAÇÃO — entenda o custo de um produto ou serviço',
      help: 'Ferramenta de aprendizagem, não calculadora fiscal. Custo direto = material + embalagem + transporte + outros · Custo do trabalho = horas × valor desejado por hora · Custo total = custo direto + custo do trabalho · Diferença = preço cobrado − custo total. ' +
        'ATENÇÃO: a diferença NÃO é lucro líquido. Impostos, taxas, aluguel, contas fixas, perdas e outros custos podem não estar incluídos. Use a última coluna para anotar o que ainda falta descobrir.' + EXAMPLE_NOTE,
      columns: [
        { h: 'Produto ou serviço', w: 22, wrap: true },
        { h: 'Material/ mercadoria', w: 12.5, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Embalagem', w: 12, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Transporte', w: 12, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Outros custos diretos', w: 12, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Horas de trabalho', w: 10, fmt: HOURS, align: 'right', v: HOURS_VALIDATION },
        { h: 'Valor desejado por hora', w: 12, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Custo direto estimado', w: 13, fmt: BRL, kind: 'calc', f: function (r) { return 'IF(' + blank(r) + ',"",SUM(B' + r + ':E' + r + '))'; } },
        { h: 'Custo do trabalho', w: 13, fmt: BRL, kind: 'calc', f: function (r) { return 'IF(' + blank(r) + ',"",N(F' + r + ')*N(G' + r + '))'; } },
        { h: 'Custo total estimado', w: 13, fmt: BRL, kind: 'calc', f: function (r) { return 'IF(' + blank(r) + ',"",SUM(H' + r + ':I' + r + '))'; } },
        { h: 'Preço atualmente cobrado', w: 13, fmt: BRL, v: MONEY_VALIDATION },
        { h: 'Diferença para o custo', w: 13, fmt: BRL, kind: 'calc', f: function (r) { return 'IF(OR(' + blank(r) + ',NOT(ISNUMBER(K' + r + '))),"",K' + r + '-J' + r + ')'; } },
        { h: 'O que ainda falta descobrir', w: 32, wrap: true }
      ],
      examples: [
        ['Bolo de pote (produto)', 6, 1.5, 0.5, 0, 0.5, 20, null, null, null, 15, null, EX + ': falta incluir gás e energia'],
        ['Afiação de faca (serviço)', 1, 0, 2, 0.5, 0.25, 30, null, null, null, 15, null, EX + ': falta incluir desgaste da pedra/máquina']
      ],
      rules: function () {
        return [
          exprRule('AND(ISNUMBER($L' + FIRST + '),$L' + FIRST + '<0)', COLOR.bad, COLOR.badFont, true),
          exprRule('AND(ISNUMBER($L' + FIRST + '),$L' + FIRST + '>=0)', COLOR.good, COLOR.goodFont)
        ];
      }
    });
  }

  var REVIEW_CELL = 'O12';
  function rotinaSheet(wb, p) {
    var t = p.today;
    return buildDataSheet(wb, p, {
      name: S.rotina, key: 'rotina', exampleCol: 'A',
      title: 'ROTINA SEMANAL — prioridades, tarefas e revisão da semana',
      help: 'Escolha até 3 prioridades e liste as tarefas da semana. No fim da semana, faça o ritual de revisão (quadro à direita) e anote a data da revisão: o PAINEL mostra essa data. ' +
        'Cores: ✖ prazo vencido e tarefa não concluída · ✔ concluída.' + EXAMPLE_NOTE,
      columns: [
        { h: 'Semana', w: 16, wrap: true },
        { h: 'Data', w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Prioridade 1', w: 20, wrap: true },
        { h: 'Prioridade 2', w: 20, wrap: true },
        { h: 'Prioridade 3', w: 20, wrap: true },
        { h: 'Tarefa', w: 26, wrap: true },
        { h: 'Prazo', w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: 'Status', w: 13, v: listValidation(LISTS.statusTarefa, true) },
        { h: 'Tempo estimado', w: 10, fmt: HOURS, align: 'right', v: HOURS_VALIDATION, kind: 'opt' },
        { h: 'Tempo real', w: 10, fmt: HOURS, align: 'right', v: HOURS_VALIDATION, kind: 'opt' },
        { h: 'O que ficou pendente', w: 22, wrap: true },
        { h: 'Próximo passo', w: 22, wrap: true }
      ],
      examples: [
        [EX, utcDate(t, -1), 'Registrar entradas e saídas do dia', 'Entregar pedido do Cliente A', 'Revisar o preço de 1 serviço', 'Anotar o caixa no fim do dia (5 min)', utcDate(t, 0), 'Em andamento', 0.5, 0.25, 'Faltou anotar uma venda em dinheiro', 'Deixar o caderno/celular perto do caixa'],
        [EX, utcDate(t, -1), null, null, null, 'Comprar material para a encomenda', utcDate(t, -2), 'Concluída', 1, 1.5, null, 'Comprar em maior quantidade na próxima vez']
      ],
      rules: function () {
        return [
          exprRule('$H' + FIRST + '="Concluída"', COLOR.good, COLOR.goodFont),
          exprRule('AND($F' + FIRST + '<>"",ISNUMBER($G' + FIRST + '),$G' + FIRST + '<TODAY(),$H' + FIRST + '<>"Concluída",$H' + FIRST + '<>"Cancelada")', COLOR.bad, COLOR.badFont, true)
        ];
      },
      extra: function (ws) {
        ws.getColumn(13).width = 3;
        ws.getColumn(14).width = 40;
        ws.getColumn(15).width = 40;
        ws.mergeCells('N4:O4');
        var head = ws.getCell('N4');
        head.value = 'RITUAL DE REVISÃO SEMANAL (15 minutos)';
        head.font = font({ bold: true, size: 10.5, color: { argb: COLOR.white } });
        head.fill = fill(isSecondary(p, S.rotina) ? COLOR.secondary : COLOR.header);
        head.alignment = { vertical: 'middle', indent: 1 };
        ['O que entrou?', 'O que saiu?', 'O que vence na próxima semana?', 'Qual pedido/serviço está parado?', 'Qual tarefa gerou mais resultado?', 'Qual ação de melhoria será mantida?'].forEach(function (question, k) {
          var row = 5 + k;
          var qc = ws.getCell('N' + row);
          qc.value = (k + 1) + '. ' + question;
          qc.font = font({ size: 10.5 });
          qc.fill = fill(COLOR.calc);
          qc.border = BOX;
          qc.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
          var ac = ws.getCell('O' + row);
          ac.fill = fill(COLOR.input);
          ac.border = BOX;
          ac.alignment = { wrapText: true, vertical: 'middle' };
          raise(ws, row, 22);
        });
        ws.mergeCells('N11:O11');
        ws.getCell('N11').value = 'Dica: use o PAINEL, o FLUXO DE CAIXA, as CONTAS e os PEDIDOS para responder.';
        ws.getCell('N11').font = font({ size: 9.5, italic: true, color: { argb: COLOR.muted } });
        ws.getCell('N11').alignment = { wrapText: true, vertical: 'middle', indent: 1 };
        raise(ws, 11, 20);
        var label = ws.getCell('N12');
        label.value = '✎ Data da última revisão semanal';
        label.font = font({ bold: true, size: 10.5, color: { argb: COLOR.infoFont } });
        label.fill = fill(COLOR.info);
        label.border = BOX;
        label.alignment = { vertical: 'middle', indent: 1 };
        var date = ws.getCell(REVIEW_CELL);
        date.numFmt = DATE;
        date.fill = fill(COLOR.input);
        date.border = BOX;
        date.alignment = { horizontal: 'left', vertical: 'middle' };
        raise(ws, 12, 22);
        ws.dataValidations.add(REVIEW_CELL, DATE_VALIDATION);
        ws.mergeCells('N13:O13');
        ws.getCell('N13').value = 'O PAINEL mostra esta data e avisa quando passar de 7 dias.';
        ws.getCell('N13').font = font({ size: 9.5, italic: true, color: { argb: COLOR.muted } });
        ws.getCell('N13').alignment = { indent: 1 };
      }
    });
  }

  // Linhas do próprio plano: foco escolhido + sugestões do diagnóstico (não são EXEMPLO; a pessoa edita ou apaga).
  function planRows(p) {
    var f = p.focus, rows = [], match = FOCUS_SUGGESTION[p.focusKey];
    var merged = p.recommendations.some(function (s) { return s.id && s.id === match; });
    if (!merged) rows.push(['Foco sugerido no diagnóstico: ' + f.label, '(anote um exemplo real da sua rotina)', f.action, f.step, null, null, null, f.measure, f.goal, 'A fazer', null, null]);
    p.recommendations.forEach(function (s, k) {
      var isFocus = merged && s.id === match;
      rows.push([(isFocus ? 'Foco + sugestão ' : 'Sugestão ') + (k + 1) + ' do diagnóstico (edite à vontade): ' + s.title,
        '(anote um exemplo real da sua rotina)', s.action || null, isFocus ? f.step : null, null, null, null,
        s.check ? 'Como conferir: ' + s.check : null, s.title.indexOf('Entenda melhor') === 0 ? 'Responder a pergunta de novo depois de observar' : (SUGGESTION_GOALS[s.id] || null),
        'A fazer', null, null]);
    });
    return rows;
  }

  function planoSheet(wb, p) {
    var ws = isWorkshop(p);
    var own = planRows(p);
    var examples = ws ? own.concat([
      ['Esqueço de registrar vendas ou serviços', 'Não sei quanto entrou na semana', 'Registrar as entradas e saídas ao fim de cada dia', 'Separar 5 minutos no fim do dia para anotar no FLUXO DE CAIXA', 'Eu', null, null, 'Quantidade de dias com registro conferido', '5 de 7 dias', 'A fazer', EX, 'Manter por mais uma semana?'],
      ['Tenho pagamentos antes dos recebimentos', 'Conta venceu antes de o cliente pagar', 'Olhar os compromissos da semana toda segunda-feira', 'Anotar hoje, em CONTAS E COMPROMISSOS, o que vence nos próximos 7 dias', 'Eu', null, null, 'Semanas em que vi o aperto antes do vencimento', '4 semanas seguidas', 'A fazer', EX, 'Negociar prazo antes do vencimento?']
    ]) : own.concat([
      ['Não sei quanto entrou e saiu na semana', 'Fim do mês sem saber para onde foi o dinheiro', 'Registrar todas as movimentações por 7 dias', 'Separar 5 minutos no fim do dia para anotar no FLUXO DE CAIXA', 'Eu', null, null, 'Dias com registro', '7 de 7 dias', 'A fazer', EX, 'Manter por mais uma semana?'],
      ['Pedidos esquecidos ou espalhados', 'Encomenda atrasou porque o combinado estava no meio das mensagens', 'Reunir pedidos em um único lugar', 'Anotar hoje todos os pedidos em aberto em CLIENTES E PEDIDOS', 'Eu', null, null, 'Pedidos entregues no prazo', 'Nenhum pedido esquecido na semana', 'A fazer', EX, 'Revisar a lista toda manhã?'],
      ['Não sei se o preço cobre os custos', 'Vendo bastante, mas sobra pouco dinheiro', 'Revisar o custo de um produto/serviço', 'Preencher uma linha da PRECIFICAÇÃO com o item mais vendido', 'Eu', null, null, 'Itens com custo revisado', '1 item revisado', 'A fazer', EX, 'Ajustar o preço ou manter?']
    ]);
    var H = ws
      ? ['Problema que quero reduzir', 'Evidência ou exemplo', 'Ação concreta que vou experimentar', 'Primeiro passo', 'Responsável', 'Quando começo', 'Quando reviso', 'Como vou perceber se funcionou', 'Meta simples', 'Status', 'Resultado observado', 'Próxima decisão']
      : ['Problema observado', 'Evidência ou exemplo', 'Ação escolhida', 'Primeiro passo', 'Responsável', 'Data de início', 'Prazo', 'Como medir', 'Meta simples', 'Status', 'Resultado observado', 'Próxima decisão'];
    return buildDataSheet(wb, p, {
      name: S.plano, key: 'plano', exampleCol: 'K',
      title: ws ? 'UMA AÇÃO DE MELHORIA — meu plano de ação' : 'PLANO DE MELHORIA — uma ação pequena de cada vez',
      help: (ws ? 'Caderno p. 8 · Meu plano de ação: qual problema quero reduzir, qual ação concreta vou experimentar, quando começo e quando reviso, e como vou perceber se funcionou. A ação precisa de prazo, primeiro passo e indicador. ' :
        'Conecte o que você observou a uma ação simples e mensurável. ') +
        'Mantenha apenas UMA ação com status "Em andamento": ela aparece no PAINEL. As primeiras linhas trazem o foco e as sugestões do seu diagnóstico (ação e como conferir) — são sugestões suas, não exemplos: ajuste, escolha uma ou apague as que não servirem. ' +
        'Resultado observado é o que de fato aconteceu; deixe em branco até conferir.' + EXAMPLE_NOTE,
      columns: [
        { h: H[0], w: 24, wrap: true },
        { h: H[1], w: 22, wrap: true, kind: 'opt' },
        { h: H[2], w: 26, wrap: true },
        { h: H[3], w: 28, wrap: true },
        { h: H[4], w: 12, wrap: true, kind: 'opt' },
        { h: H[5], w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: H[6], w: 11.5, fmt: DATE, v: DATE_VALIDATION },
        { h: H[7], w: 22, wrap: true },
        { h: H[8], w: 16, wrap: true },
        { h: H[9], w: 13, v: listValidation(LISTS.statusPlano, true) },
        { h: H[10], w: 22, wrap: true },
        { h: H[11], w: 20, wrap: true }
      ],
      examples: examples,
      rules: function () {
        return [
          exprRule('$J' + FIRST + '="Concluída"', COLOR.good, COLOR.goodFont),
          exprRule('$J' + FIRST + '="Em andamento"', COLOR.info, COLOR.infoFont, true),
          exprRule('AND(ISNUMBER($G' + FIRST + '),$G' + FIRST + '<TODAY(),$J' + FIRST + '<>"Concluída",$J' + FIRST + '<>"Cancelada")', COLOR.bad, COLOR.badFont, true)
        ];
      }
    });
  }

  function orcamentoSheet(wb, p) {
    var name = S.orcamento;
    var ws = wb.addWorksheet(name, sheetOptions(p, name, 4));
    ws.columns = [{ width: 32 }, { width: 15 }, { width: 16 }, { width: 15 }, { width: 27 }, { width: 30 }];
    function F(col) { return ref(S.fluxo, '$' + col + '$' + FIRST + ':$' + col + '$' + lastRow('fluxo')); }
    var NEX = q('<>' + EX), M = '$B$3';
    var base = F('A') + ',">="&DATE(YEAR(' + M + '),MONTH(' + M + '),1),' + F('A') + ',"<"&DATE(YEAR(' + M + '),MONTH(' + M + ')+1,1),' + F('K') + ',"<>Não",' + F('L') + ',' + NEX;
    var noMonth = 'NOT(ISNUMBER(' + M + '))';

    titleBand(ws, 'A1:F1', stepTitle(p, name, 'ORÇAMENTO DO MÊS — planejado x realizado'), p, name);
    ws.getRow(1).height = 30;
    ws.mergeCells('A2:F2');
    var help = 'Orçamento é o plano do mês. Em Planejado, escreva quanto espera receber e gastar em cada categoria. O Realizado vem sozinho do FLUXO DE CAIXA: soma o que foi pago/recebido no mês, pela Categoria de cada linha ' +
      '(exemplos e valores marcados "Não" não entram). Planejado em branco = não planejado, não é zero. Ex.: planejei R$ 300 de material e o realizado foi R$ 350 → ⚠ acima do planejado: o que ajustar no próximo mês? ' +
      'Previsão não é promessa, e resultado de caixa não é lucro.';
    ws.getCell('A2').value = help;
    ws.getCell('A2').font = font({ size: 10, color: { argb: COLOR.muted } });
    ws.getCell('A2').fill = fill(COLOR.helpFill);
    ws.getCell('A2').alignment = { wrapText: true, vertical: 'middle', indent: 1 };
    ws.getRow(2).height = textHeight(help, 130, 10, 30);

    var lab = ws.getCell('A3');
    lab.value = '✎ Mês do orçamento';
    lab.font = font({ bold: true, size: 10.5, color: { argb: COLOR.infoFont } });
    lab.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
    ws.mergeCells('B3:C3');
    var month = ws.getCell('B3');
    month.value = new Date(Date.UTC(p.today.getFullYear(), p.today.getMonth(), 1));
    month.numFmt = MONTH;
    month.font = font({ bold: true, size: 10.5 });
    month.fill = fill(COLOR.input);
    month.border = BOX;
    month.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.dataValidations.add('B3', DATE_VALIDATION);
    ws.mergeCells('D3:F3');
    ws.getCell('D3').value = 'Digite qualquer data do mês desejado (ex.: 01/12/2026).';
    ws.getCell('D3').font = font({ size: 9.5, italic: true, color: { argb: COLOR.muted } });
    ws.getCell('D3').alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(3).height = 24;

    ['Categoria', 'Planejado (R$)', 'Realizado no FLUXO (R$)', 'Diferença (realizado − planejado)', 'Situação', 'Observação'].forEach(function (text, i) {
      var c = ws.getCell(4, i + 1);
      c.value = text;
      c.font = font({ bold: true, size: 10.5, color: { argb: COLOR.white } });
      c.fill = fill(i >= 2 && i <= 4 ? COLOR.headerCalc : COLOR.header);
      c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
      c.border = { top: THIN, bottom: THIN, left: { style: 'thin', color: { argb: COLOR.white } }, right: { style: 'thin', color: { argb: COLOR.white } } };
    });
    ws.getRow(4).height = 32;

    var row = 5;
    function section(text) {
      ws.mergeCells('A' + row + ':F' + row);
      var c = ws.getCell('A' + row);
      c.value = text;
      c.font = font({ bold: true, size: 10.5, color: { argb: COLOR.dark } });
      c.fill = fill(COLOR.good);
      c.alignment = { vertical: 'middle', indent: 1 };
      c.border = { bottom: { style: 'thin', color: { argb: COLOR.header } } };
      ws.getRow(row).height = 22;
      row++;
    }
    function situation(r, kind) {
      var good = kind === 'in' ? '"✔ Atingiu o planejado"' : '"✔ Dentro do planejado"';
      var bad = kind === 'in' ? '"⚠ Abaixo do planejado (até agora)"' : '"⚠ Acima do planejado"';
      var test = kind === 'in' ? 'N(C' + r + ')>=B' + r : 'N(C' + r + ')<=B' + r;
      return 'IF($A' + r + '="","",IF(NOT(ISNUMBER(B' + r + ')),IF(N(C' + r + ')>0,"ℹ Sem valor planejado",""),IF(' + test + ',' + good + ',' + bad + ')))';
    }
    function line(opts) {
      var r = row, cells = [];
      for (var i = 1; i <= 6; i++) cells.push(ws.getCell(r, i));
      var a = cells[0];
      a.value = opts.label || null;
      a.fill = fill(opts.calcLabel ? COLOR.calc : COLOR.input);
      a.font = font({ size: 10.5, bold: !!opts.total, italic: !!opts.calcLabel && !opts.total });
      a.alignment = { vertical: 'middle', indent: 1, wrapText: true };
      var b = cells[1];
      if (opts.planned) b.value = { formula: opts.planned(r) };
      b.fill = fill(opts.planned ? COLOR.calc : COLOR.input);
      b.numFmt = BRL;
      if (!opts.planned) ws.dataValidations.add('B' + r, MONEY_VALIDATION);
      cells[2].value = { formula: opts.real(r) };
      cells[2].numFmt = BRL_DASH;
      cells[3].value = { formula: 'IF(OR($A' + r + '="",NOT(ISNUMBER(B' + r + '))),"",N(C' + r + ')-B' + r + ')' };
      cells[3].numFmt = '+"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00';
      cells[4].value = { formula: situation(r, opts.kind) };
      [cells[2], cells[3], cells[4]].forEach(function (c) { c.fill = fill(COLOR.calc); });
      cells[5].fill = fill(COLOR.input);
      cells[5].alignment = { wrapText: true, vertical: 'middle' };
      cells.forEach(function (c, i) {
        c.border = opts.total ? { top: { style: 'thin', color: { argb: COLOR.muted } }, bottom: { style: 'double', color: { argb: COLOR.muted } }, left: THIN, right: THIN } : BOX;
        if (i >= 1 && i <= 3) { c.font = font({ size: 10.5, bold: !!opts.total }); c.alignment = { horizontal: 'right', vertical: 'middle' }; }
        if (i === 4) { c.font = font({ size: 10 }); c.alignment = { vertical: 'middle', indent: 1 }; }
        if (i === 5) c.font = font({ size: 10 });
      });
      ws.getRow(r).height = 20;
      row++;
      return r;
    }
    function realBy(col, typeCrit) {
      return function (r) { return 'IF(OR($A' + r + '="",' + noMonth + '),"",SUMIFS(' + F(col) + ',' + F('D') + ',$A' + r + ',' + F('B') + ',' + typeCrit + ',' + base + '))'; };
    }
    function block(title, cats, extraRows, col, typeCrit, kind, otherLabel, totalLabel) {
      section(title);
      var first = row;
      cats.forEach(function (cat) { line({ label: cat, real: realBy(col, typeCrit), kind: kind }); });
      for (var k = 0; k < extraRows; k++) line({ real: realBy(col, typeCrit), kind: kind });
      var lastCat = row - 1;
      line({ label: otherLabel, calcLabel: true, kind: kind, real: function () {
        return 'IF(' + noMonth + ',"",MAX(0,SUMIFS(' + F(col) + ',' + F('B') + ',' + typeCrit + ',' + base + ')-SUM(C' + first + ':C' + lastCat + ')))';
      } });
      var total = line({ label: totalLabel, calcLabel: true, total: true, kind: kind,
        planned: function (r) { return 'IF(COUNT(B' + first + ':B' + (r - 1) + ')=0,"",SUM(B' + first + ':B' + (r - 1) + '))'; },
        real: function (r) { return 'IF(' + noMonth + ',"",SUM(C' + first + ':C' + (r - 1) + '))'; } });
      ws.getRow(row).height = 8;
      row++;
      return total;
    }
    var tIn = block('RECEITAS — dinheiro que espero receber no mês', LISTS.incomeCats, 2, 'H', '"<>Saldo inicial"', 'in', 'Outras entradas / sem categoria', 'Total de receitas');
    var tOut = block('DESPESAS DO NEGÓCIO — o que espero pagar no mês', LISTS.expenseCats, 2, 'I', '"<>Retirada pessoal"', 'out', 'Outras saídas / sem categoria', 'Total de despesas');
    section('RETIRADAS PESSOAIS — dinheiro do negócio usado em casa (não é despesa do negócio)');
    var tRet = line({ label: 'Retirada pessoal', calcLabel: true, kind: 'out', real: function () {
      return 'IF(' + noMonth + ',"",SUMIFS(' + F('I') + ',' + F('B') + ',"Retirada pessoal",' + base + '))';
    } });
    ws.getRow(row).height = 8;
    row++;
    section('RESULTADO DE CAIXA DO MÊS — receitas − despesas − retiradas (não é lucro)');
    var res = line({ label: 'Resultado de caixa', calcLabel: true, total: true, kind: 'in',
      planned: function () { return 'IF(AND(B' + tIn + '="",B' + tOut + '="",NOT(ISNUMBER(B' + tRet + '))),"",N(B' + tIn + ')-N(B' + tOut + ')-N(B' + tRet + '))'; },
      real: function () { return 'IF(' + noMonth + ',"",N(C' + tIn + ')-N(C' + tOut + ')-N(C' + tRet + '))'; } });
    ws.getCell('E' + res).value = { formula: 'IF(NOT(ISNUMBER(B' + res + ')),"ℹ Preencha o planejado para comparar",IF(N(C' + res + ')>=B' + res + ',"✔ Igual ou melhor que o planejado","⚠ Abaixo do planejado (até agora)"))' };
    ws.mergeCells('A' + row + ':F' + row);
    var note = ws.getCell('A' + row);
    note.value = 'O resultado de caixa mostra dinheiro que entrou e saiu no mês. Lucro exige considerar também estoque, valores a receber, obrigações e outros fatores. Esta aba não é contabilidade oficial.';
    note.font = font({ size: 9.5, italic: true, color: { argb: COLOR.muted } });
    note.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 30;

    ws.addConditionalFormatting({ ref: 'E6:E' + res, rules: withPriority([
      exprRule('LEFT($E6,1)="⚠"', COLOR.warn, COLOR.warnFont, true),
      exprRule('LEFT($E6,1)="✔"', COLOR.good, COLOR.goodFont, false)
    ]) });
    pageSetup(ws, 'A1:F' + row, false, null, true);
    return ws;
  }

  function painelSheet(wb, p) {
    var ws = wb.addWorksheet(S.painel, sheetOptions(p, S.painel, 0));
    ws.columns = [{ width: 2 }, { width: 20 }, { width: 18 }, { width: 20 }, { width: 18 }, { width: 20 }, { width: 18 }, { width: 3 }, { width: 40 }, { width: 16 }];
    function rng(sheet, key, col) { return ref(sheet, '$' + col + '$' + FIRST + ':$' + col + '$' + lastRow(key)); }
    function fl(c) { return rng(S.fluxo, 'fluxo', c); }
    function ct(c) { return rng(S.contas, 'contas', c); }
    function cl(c) { return rng(S.clientes, 'clientes', c); }
    function rt(c) { return rng(S.rotina, 'rotina', c); }
    function pl(c) { return rng(S.plano, 'plano', c); }
    var NEX = q('<>' + EX), EXQ = q(EX);
    var START = '$C$3', END = '$E$3';
    var per = fl('A') + ',">="&' + START + ',' + fl('A') + ',"<"&(' + END + '+1)';
    var real = fl('K') + ',"<>Não",' + fl('L') + ',' + NEX;
    var notInitial = fl('B') + ',"<>Saldo inicial"';
    function ctOpen(tipo, extra) {
      var base = ct('A') + ',' + q(tipo) + ',' + ct('K') + ',' + NEX + (extra ? ',' + extra : '');
      return 'SUMIFS(' + ct('F') + ',' + base + ',' + ct('H') + ',"Em aberto")+SUMIFS(' + ct('F') + ',' + base + ',' + ct('H') + ',"Parcial")-SUMIFS(' + ct('G') + ',' + base + ',' + ct('H') + ',"Parcial")';
    }

    // Cálculos auxiliares (colunas I/J, agrupadas e ocultas; clique no "+" acima da coluna para ver).
    var H = {};
    var helpers = [
      ['nF', 'Movimentos registrados (sem exemplos)', 'COUNTIFS(' + fl('H') + ',"<>",' + fl('L') + ',' + NEX + ')+COUNTIFS(' + fl('I') + ',"<>",' + fl('L') + ',' + NEX + ')', '0'],
      ['saldo', 'Saldo realizado', 'SUMIFS(' + fl('H') + ',' + real + ')-SUMIFS(' + fl('I') + ',' + real + ')', BRL],
      ['nP', 'Movimentos no período', 'COUNTIFS(' + per + ',' + fl('L') + ',' + NEX + ')', '0'],
      ['entP', 'Entradas no período (sem saldo inicial)', 'SUMIFS(' + fl('H') + ',' + notInitial + ',' + per + ',' + real + ')', BRL],
      ['saiP', 'Saídas no período (sem retiradas)', 'SUMIFS(' + fl('I') + ',' + fl('B') + ',"<>Retirada pessoal",' + per + ',' + real + ')', BRL],
      ['retP', 'Retiradas pessoais no período', 'SUMIFS(' + fl('I') + ',' + fl('B') + ',"Retirada pessoal",' + per + ',' + real + ')', BRL],
      ['resP', 'Resultado de caixa do período', null, BRL],
      ['pag7', 'A pagar: vencidas + próximos 7 dias', ctOpen('A pagar', ct('E') + ',"<="&(TODAY()+7)'), BRL],
      ['pagAll', 'A pagar: total em aberto', ctOpen('A pagar'), BRL],
      ['venc', 'Contas a pagar vencidas (qtd.)', 'COUNTIFS(' + ct('A') + ',"A pagar",' + ct('H') + ',"Em aberto",' + ct('E') + ',"<"&TODAY(),' + ct('K') + ',' + NEX + ')+COUNTIFS(' + ct('A') + ',"A pagar",' + ct('H') + ',"Parcial",' + ct('E') + ',"<"&TODAY(),' + ct('K') + ',' + NEX + ')', '0'],
      ['rec', 'A receber: total em aberto', ctOpen('A receber'), BRL],
      ['nCP', 'Contas a pagar registradas', 'COUNTIFS(' + ct('A') + ',"A pagar",' + ct('K') + ',' + NEX + ')', '0'],
      ['nCR', 'Valores a receber registrados', 'COUNTIFS(' + ct('A') + ',"A receber",' + ct('K') + ',' + NEX + ')', '0'],
      ['pedA', 'Pedidos/serviços em aberto', 'COUNTIFS(' + cl('B') + ',"<>",' + cl('H') + ',"<>Concluído",' + cl('H') + ',"<>Cancelado",' + cl('K') + ',' + NEX + ')', '0'],
      ['pedAtr', 'Pedidos com prazo vencido', 'COUNTIFS(' + cl('B') + ',"<>",' + cl('E') + ',"<"&TODAY(),' + cl('H') + ',"<>Concluído",' + cl('H') + ',"<>Cancelado",' + cl('K') + ',' + NEX + ')', '0'],
      ['nPed', 'Pedidos registrados', 'COUNTIFS(' + cl('B') + ',"<>",' + cl('K') + ',' + NEX + ')', '0'],
      ['tarA', 'Tarefas atrasadas', 'COUNTIFS(' + rt('F') + ',"<>",' + rt('G') + ',"<"&TODAY(),' + rt('H') + ',"<>Concluída",' + rt('H') + ',"<>Cancelada",' + rt('A') + ',' + NEX + ')', '0'],
      ['nTar', 'Tarefas registradas', 'COUNTIFS(' + rt('F') + ',"<>",' + rt('A') + ',' + NEX + ')', '0'],
      ['acaoPos', 'Linha da ação Em andamento', 'IFERROR(MATCH("Em andamento",' + pl('J') + ',0),0)', '0'],
      ['nEx', 'Linhas de exemplo restantes', 'COUNTIF(' + fl('L') + ',' + EXQ + ')+COUNTIF(' + ct('K') + ',' + EXQ + ')+COUNTIF(' + cl('K') + ',' + EXQ + ')+COUNTIF(' + rt('A') + ',' + EXQ + ')+COUNTIF(' + pl('K') + ',' + EXQ + ')+COUNTIF(' + rng(S.preco, 'preco', 'M') + ',"EXEMPLO*")', '0']
    ];
    ws.getCell('I2').value = 'Cálculos auxiliares (não editar)';
    ws.getCell('I2').font = font({ bold: true, size: 9.5, color: { argb: COLOR.muted } });
    helpers.forEach(function (hp, k) { H[hp[0]] = '$J$' + (3 + k); });
    H.rev = ref(S.rotina, '$O$12');
    helpers.forEach(function (hp, k) {
      var row = 3 + k;
      if (hp[0] === 'resP') hp[2] = H.entP + '-' + H.saiP + '-' + H.retP;
      var lc = ws.getCell('I' + row), vc = ws.getCell('J' + row);
      lc.value = hp[1];
      lc.font = font({ size: 9.5, color: { argb: COLOR.muted } });
      lc.fill = fill(COLOR.calc);
      vc.value = { formula: hp[2] };
      vc.numFmt = hp[3];
      vc.font = font({ size: 9.5, color: { argb: COLOR.muted } });
      vc.fill = fill(COLOR.calc);
    });
    [9, 10].forEach(function (c) { var col = ws.getColumn(c); col.outlineLevel = 1; col.hidden = true; });

    titleBand(ws, 'B1:G1', 'PAINEL — ' + p.name, p, S.painel, 18);
    ws.getRow(1).height = 34;
    ws.mergeCells('B2:G2');
    ws.getCell('B2').value = 'Resumo automático das outras abas. Não digite nos cartões. Linhas de exemplo não entram nos números.';
    ws.getCell('B2').font = font({ size: 10, italic: true, color: { argb: COLOR.muted } });
    ws.getCell('B2').alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(2).height = 20;

    ws.getCell('B3').value = '✎ Período: de';
    ws.getCell('B3').font = font({ bold: true, size: 10.5, color: { argb: COLOR.infoFont } });
    ws.getCell('B3').alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell('C3').value = { formula: 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)' };
    ws.getCell('D3').value = 'até';
    ws.getCell('D3').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell('D3').font = font({ bold: true, size: 10.5, color: { argb: COLOR.infoFont } });
    ws.getCell('E3').value = { formula: 'TODAY()' };
    ['C3', 'E3'].forEach(function (a) {
      var c = ws.getCell(a);
      c.numFmt = DATE; c.fill = fill(COLOR.input); c.border = BOX; c.font = font({ bold: true, size: 10.5 });
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.dataValidations.add(a, DATE_VALIDATION);
    });
    ws.mergeCells('F3:G3');
    ws.getCell('F3').value = 'Padrão: mês atual. Pode digitar outras datas.';
    ws.getCell('F3').font = font({ size: 9.5, italic: true, color: { argb: COLOR.muted } });
    ws.getCell('F3').alignment = { wrapText: true, vertical: 'middle', indent: 1 };
    ws.getRow(3).height = 22;

    ws.mergeCells('B4:G4');
    var alert = ws.getCell('B4');
    alert.value = { formula: 'IF(' + H.nEx + '>0,"⚠ Ainda há "&' + H.nEx + '&" linha(s) de EXEMPLO nas abas. Elas não entram no painel; apague-as antes de usar.","✔ Nenhuma linha de exemplo encontrada.")' };
    alert.font = font({ bold: true, size: 10.5 });
    alert.fill = fill(COLOR.info);
    alert.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(4).height = 22;
    ws.addConditionalFormatting({ ref: 'B4:G4', rules: withPriority([
      exprRule(H.nEx + '>0', COLOR.warn, COLOR.warnFont, true),
      exprRule(H.nEx + '=0', COLOR.good, COLOR.goodFont, true)
    ]) });

    var guideLines = isWorkshop(p)
      ? [['➜ Hoje, na oficina: ', 'comece pela aba ' + S.fluxo + ' e siga os passos 1 a 4.'], ['★ Foco sugerido pelo diagnóstico: ', p.focus.label], ['➜ Depois da oficina: ', p.focus.step]]
      : [['★ Foco sugerido pelo diagnóstico: ', p.focus.label], ['➜ Primeiro passo: ', p.focus.step], ['➜ Comece pela aba: ', p.startTab]];
    guideLines.forEach(function (line, k) {
      var row = 5 + k;
      ws.mergeCells('B' + row + ':G' + row);
      var c = ws.getCell('B' + row);
      c.value = { richText: [{ text: line[0], font: font({ bold: true, size: 10.5, color: { argb: COLOR.dark } }) }, { text: line[1], font: font({ size: 10.5 }) }] };
      c.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
      ws.getRow(row).height = textHeight(line[0] + line[1], 110, 10.5, 20);
    });
    ws.getRow(8).height = 8;
    var noData = '"ℹ Ainda não há dados suficientes"';
    var cards = [
      { label: 'Saldo acumulado atual', fmt: BRL, value: 'IF(' + H.nF + '=0,"—",' + H.saldo + ')',
        msg: 'IF(' + H.nF + '=0,' + noData + ',IF(' + H.saldo + '<0,"✖ Saldo negativo: precisa de ação","✔ Acompanhado (só pago/recebido)"))',
        cf: [['AND(' + H.nF + '>0,' + H.saldo + '<0)', 'bad'], [H.nF + '>0', 'good']] },
      { label: 'Entradas no período', fmt: BRL, value: 'IF(' + H.nP + '=0,"—",' + H.entP + ')',
        msg: 'IF(' + H.nP + '=0,' + noData + ',"ℹ Dinheiro que realmente entrou (sem saldo inicial)")' },
      { label: 'Saídas no período (sem retiradas)', fmt: BRL, value: 'IF(' + H.nP + '=0,"—",' + H.saiP + ')',
        msg: 'IF(' + H.nP + '=0,' + noData + ',"ℹ Gastos do negócio pagos")' },
      { label: 'Retiradas pessoais no período', fmt: BRL, value: 'IF(' + H.nP + '=0,"—",' + H.retP + ')',
        msg: 'IF(' + H.nP + '=0,' + noData + ',"ℹ Usado em casa: não é despesa do negócio")' },
      { label: 'Resultado de caixa do período', fmt: BRL, value: 'IF(' + H.nP + '=0,"—",' + H.resP + ')',
        msg: 'IF(' + H.nP + '=0,' + noData + ',IF(' + H.resP + '<0,"⚠ Saiu mais do que entrou (não é lucro)","✔ Entrou mais do que saiu (não é lucro)"))',
        cf: [['AND(' + H.nP + '>0,' + H.resP + '<0)', 'warn'], [H.nP + '>0', 'good']] },
      { label: 'Contas a pagar: vencidas e próximos 7 dias', fmt: BRL, value: 'IF(' + H.nCP + '=0,"—",' + H.pag7 + ')',
        msg: 'IF(' + H.nCP + '=0,"ℹ Nenhuma conta a pagar registrada",IF(' + H.venc + '>0,"✖ "&' + H.venc + '&" vencida(s): precisa de ação",IF(' + H.pag7 + '>0,"⚠ Há contas nos próximos 7 dias","✔ Nada a pagar nos próximos 7 dias")))',
        cf: [['AND(' + H.nCP + '>0,' + H.venc + '>0)', 'bad'], ['AND(' + H.nCP + '>0,' + H.pag7 + '>0)', 'warn'], [H.nCP + '>0', 'good']] },
      { label: 'Valores a receber (em aberto)', fmt: BRL, value: 'IF(' + H.nCR + '=0,"—",' + H.rec + ')',
        msg: 'IF(' + H.nCR + '=0,"ℹ Nenhum valor a receber registrado","ℹ Previsto: ainda não é dinheiro em caixa")' },
      { label: 'Pedidos/serviços em aberto', fmt: '0', value: 'IF(' + H.nPed + '=0,"—",' + H.pedA + ')',
        msg: 'IF(' + H.nPed + '=0,"ℹ Ainda não há pedidos registrados",IF(' + H.pedAtr + '>0,"✖ "&' + H.pedAtr + '&" com prazo vencido","✔ Nenhum pedido atrasado"))',
        cf: [['AND(' + H.nPed + '>0,' + H.pedAtr + '>0)', 'bad'], [H.nPed + '>0', 'good']] },
      { label: 'Tarefas atrasadas', fmt: '0', value: 'IF(' + H.nTar + '=0,"—",' + H.tarA + ')',
        msg: 'IF(' + H.nTar + '=0,"ℹ Ainda não há tarefas registradas",IF(' + H.tarA + '>0,"⚠ Reveja prazos ou reagende","✔ Nenhuma tarefa atrasada"))',
        cf: [['AND(' + H.nTar + '>0,' + H.tarA + '>0)', 'warn'], [H.nTar + '>0', 'good']] },
      { label: 'Ação de melhoria atual', text: true,
        value: 'IF(' + H.acaoPos + '>0,IF(INDEX(' + pl('C') + ',' + H.acaoPos + ')="","(ação sem descrição)",INDEX(' + pl('C') + ',' + H.acaoPos + ')),' + q('Sugestão: ' + p.focus.action) + ')',
        msg: 'IF(' + H.acaoPos + '>0,"✔ Em andamento no PLANO DE MELHORIA","ℹ Marque uma ação como Em andamento")',
        cf: [[H.acaoPos + '>0', 'good']] },
      { label: 'Data da última revisão', fmt: DATE, value: 'IF(ISNUMBER(' + H.rev + '),' + H.rev + ',"—")',
        msg: 'IF(NOT(ISNUMBER(' + H.rev + ')),"ℹ Anote a data na ROTINA SEMANAL",IF(TODAY()-' + H.rev + '>7,"⚠ Mais de 7 dias sem revisão","✔ Revisão em dia"))',
        cf: [['AND(ISNUMBER(' + H.rev + '),TODAY()-' + H.rev + '>7)', 'warn'], ['ISNUMBER(' + H.rev + ')', 'good']] },
      { label: 'Movimentos registrados', fmt: '0', value: 'IF(' + H.nF + '=0,"—",' + H.nF + ')',
        msg: 'IF(' + H.nF + '<7,"ℹ Meta inicial: 7 dias de registros","✔ Registro em andamento")' }
    ];
    var slots = [['B', 'C'], ['D', 'E'], ['F', 'G']];
    var cfPriority = 10;
    var WHITE_EDGE = { style: 'medium', color: { argb: COLOR.white } };
    cards.forEach(function (card, k) {
      var cols = slots[k % 3], top = 9 + Math.floor(k / 3) * 4;
      var range = function (row) { return cols[0] + row + ':' + cols[1] + row; };
      [top, top + 1, top + 2].forEach(function (row) { ws.mergeCells(range(row)); });
      var lc = ws.getCell(cols[0] + top), vc = ws.getCell(cols[0] + (top + 1)), mc = ws.getCell(cols[0] + (top + 2));
      lc.value = card.label;
      lc.font = font({ bold: true, size: 10, color: { argb: COLOR.dark } });
      lc.alignment = { wrapText: true, vertical: 'bottom', indent: 1 };
      vc.value = { formula: card.value };
      vc.font = font({ bold: true, size: card.text ? 11 : 17 });
      vc.alignment = { wrapText: true, vertical: 'middle', horizontal: card.text ? 'left' : 'center', indent: card.text ? 1 : 0 };
      if (card.fmt) vc.numFmt = card.fmt;
      mc.value = { formula: card.msg };
      mc.font = font({ size: 9, color: { argb: COLOR.muted } });
      mc.alignment = { wrapText: true, vertical: 'top', indent: 1 };
      [lc, vc, mc].forEach(function (c) { c.fill = fill(COLOR.info); });
      lc.border = { top: WHITE_EDGE, left: WHITE_EDGE, right: WHITE_EDGE };
      vc.border = { left: WHITE_EDGE, right: WHITE_EDGE };
      mc.border = { bottom: WHITE_EDGE, left: WHITE_EDGE, right: WHITE_EDGE };
      if (card.cf) {
        var palette = { good: [COLOR.good, COLOR.goodFont], warn: [COLOR.warn, COLOR.warnFont], bad: [COLOR.bad, COLOR.badFont] };
        ws.addConditionalFormatting({ ref: cols[0] + top + ':' + cols[1] + (top + 2), rules: card.cf.map(function (rule) {
          var r = exprRule(rule[0], palette[rule[1]][0], palette[rule[1]][1], false);
          r.priority = cfPriority++;
          return r;
        }) });
      }
    });
    [9, 13, 17, 21].forEach(function (top) {
      ws.getRow(top).height = 30;
      ws.getRow(top + 1).height = top === 21 ? 44 : 32;
      ws.getRow(top + 2).height = 28;
      ws.getRow(top + 3).height = 8;
    });

    // "Gráficos" em barras de texto (REPT), compatíveis com Excel e Google Planilhas.
    var row = 25;
    function bar(valueRef, maxRange) {
      return 'IF(MAX(' + maxRange + ')<=0,"",REPT("█",ROUND(10*N(' + valueRef + ')/MAX(' + maxRange + '),0)))';
    }
    function section(titleText, headers, rows, guard, valueCols) {
      ws.mergeCells('B' + row + ':G' + row);
      var tc = ws.getCell('B' + row);
      tc.value = titleText;
      tc.font = font({ bold: true, size: 11, color: { argb: COLOR.white } });
      tc.fill = fill(isSecondary(p, S.painel) ? COLOR.secondary : COLOR.header);
      tc.alignment = { indent: 1, vertical: 'middle' };
      ws.getRow(row).height = 22;
      var first = row + 3, last = row + 2 + rows.length;
      var dataRange = valueCols.map(function (c) { return '$' + c + '$' + first + ':$' + c + '$' + last; }).join(',');
      ws.mergeCells('B' + (row + 1) + ':G' + (row + 1));
      var mc = ws.getCell('B' + (row + 1));
      mc.value = { formula: 'IF(OR(' + guard + ',SUM(' + dataRange + ')=0),"ℹ Ainda não há dados suficientes para este gráfico.","Barras proporcionais ao maior valor desta tabela (sem exemplos).")' };
      mc.font = font({ size: 9, italic: true, color: { argb: COLOR.muted } });
      mc.alignment = { vertical: 'middle', indent: 1 };
      ws.getRow(row + 1).height = 18;
      headers.forEach(function (hd, i) {
        var c = ws.getCell(row + 2, i + 2);
        c.value = hd;
        c.font = font({ bold: true, size: 10 });
        c.fill = fill(COLOR.calc);
        c.border = BOX;
        c.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
      });
      ws.getRow(row + 2).height = 18;
      rows.forEach(function (cells, k) {
        ws.getRow(first + k).height = 17;
        cells.forEach(function (cellSpec, i) {
          if (!cellSpec) return;
          var c = ws.getCell(first + k, i + 2);
          c.value = cellSpec.f ? { formula: cellSpec.f(first + k, first, last) } : cellSpec.v;
          if (cellSpec.fmt) c.numFmt = cellSpec.fmt;
          c.font = font(cellSpec.color ? { color: { argb: cellSpec.color }, size: 10 } : { size: 10 });
          c.border = BOX;
          c.alignment = cellSpec.fmt === BRL || cellSpec.fmt === '0' ? { horizontal: 'right', vertical: 'middle' } : { wrapText: true, vertical: 'middle', indent: 1 };
        });
      });
      row = last + 2;
    }

    var months = [];
    for (var m = 5; m >= 0; m--) {
      months.push((function (back) {
        var start = 'DATE(YEAR(TODAY()),MONTH(TODAY())-' + back + ',1)';
        var next = 'DATE(YEAR(TODAY()),MONTH(TODAY())-' + back + '+1,1)';
        var crit = fl('A') + ',">="&' + start + ',' + fl('A') + ',"<"&' + next + ',' + notInitial + ',' + real;
        return [
          { f: function () { return start; }, fmt: '[$-416]mmm/yyyy' },
          { f: function () { return 'IF(' + H.nF + '=0,"—",SUMIFS(' + fl('H') + ',' + crit + '))'; }, fmt: BRL },
          { f: function () { return 'IF(' + H.nF + '=0,"—",SUMIFS(' + fl('I') + ',' + crit + '))'; }, fmt: BRL },
          { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$D$' + b); }, color: COLOR.goodFont },
          { f: function (r, a, b) { return bar('D' + r, '$C$' + a + ':$D$' + b); }, color: COLOR.badFont }
        ];
      })(m));
    }
    section('Entradas x saídas por mês (últimos 6 meses)', ['Mês', 'Entradas', 'Saídas + retiradas', 'Entradas', 'Saídas'], months, H.nF + '=0', ['C', 'D']);

    var catRows = LISTS.expenseCats.map(function (cat) {
      return [
        { v: cat },
        { f: function () { return 'IF(' + H.nP + '=0,"—",SUMIFS(' + fl('I') + ',' + fl('B') + ',"Saída",' + fl('D') + ',' + q(cat) + ',' + per + ',' + real + '))'; }, fmt: BRL },
        { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$C$' + (b)); }, color: COLOR.badFont }
      ];
    });
    catRows.push([
      { v: 'Outras categorias / sem categoria' },
      { f: function (r, a) { return 'IF(' + H.nP + '=0,"—",MAX(0,' + H.saiP + '-SUM($C$' + a + ':$C$' + (r - 1) + ')))'; }, fmt: BRL },
      { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$C$' + b); }, color: COLOR.badFont }
    ]);
    section('Despesas do período por categoria (sem retiradas pessoais)', ['Categoria', 'Valor', 'Proporção'], catRows, H.nP + '=0', ['C']);

    section('Contas em aberto: a pagar x a receber (previsão, não é caixa)', ['Tipo', 'Valor em aberto', 'Proporção'], [
      [{ v: 'A pagar' }, { f: function () { return 'IF(' + H.nCP + '=0,"—",' + H.pagAll + ')'; }, fmt: BRL }, { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$C$' + b); }, color: COLOR.badFont }],
      [{ v: 'A receber' }, { f: function () { return 'IF(' + H.nCR + '=0,"—",' + H.rec + ')'; }, fmt: BRL }, { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$C$' + b); }, color: COLOR.goodFont }]
    ], 'AND(' + H.nCP + '=0,' + H.nCR + '=0)', ['C']);

    section('Pedidos/serviços por status', ['Status', 'Quantidade', 'Proporção'], LISTS.statusPedido.map(function (status) {
      return [
        { v: status },
        { f: function () { return 'IF(' + H.nPed + '=0,"—",COUNTIFS(' + cl('B') + ',"<>",' + cl('H') + ',' + q(status) + ',' + cl('K') + ',' + NEX + '))'; }, fmt: '0' },
        { f: function (r, a, b) { return bar('C' + r, '$C$' + a + ':$C$' + b); }, color: COLOR.infoFont }
      ];
    }), H.nPed + '=0', ['C']);

    pageSetup(ws, 'B1:G' + (row - 1), false, null, false);
    ws.getRow(24).addPageBreak();
    return ws;
  }

  var STEP_TEXT = {};
  STEP_TEXT[S.painel] = 'PAINEL: resumo que se atualiza sozinho — você não digita nele.';
  STEP_TEXT[S.fluxo] = 'FLUXO DE CAIXA: anote cada dinheiro que entra ou sai, uma linha por movimento. Dinheiro do negócio usado em casa é "Retirada pessoal".';
  STEP_TEXT[S.contas] = 'CONTAS E COMPROMISSOS: anote o que precisa pagar e o que vai receber, com a data. Previsão não é dinheiro em caixa.';
  STEP_TEXT[S.clientes] = 'CLIENTES E PEDIDOS: reúna pedidos, encomendas e serviços combinados em um só lugar.';
  STEP_TEXT[S.preco] = 'PRECIFICAÇÃO: revise o custo de UM produto ou serviço por vez.';
  STEP_TEXT[S.plano] = 'PLANO DE MELHORIA: escolha UMA ação pequena e marque o status "Em andamento".';
  STEP_TEXT[S.rotina] = 'ROTINA SEMANAL: uma vez por semana, faça a revisão de 15 minutos e anote a data.';
  function link(sheet, text) { return { text: text || sheet, hyperlink: "#'" + sheet + "'!A1", tooltip: 'Ir para a aba ' + sheet }; }
  var PAPER_SHEETS = [S.fluxo, S.contas, S.orcamento, S.clientes, S.preco, S.rotina, S.plano];
  function paperTabs(p) {
    return p.activity.order.filter(function (name) { return PAPER_SHEETS.indexOf(name) >= 0 && !isSecondary(p, name); });
  }

  function leiaSheet(wb, p) {
    var ws = wb.addWorksheet(S.leia, sheetOptions(p, S.leia, 0));
    var A = p.activity, workshop = isWorkshop(p);
    ws.columns = [{ width: 2 }, { width: 7 }, { width: 64 }, { width: 28 }, { width: 2 }];
    var WIDTH = 92;
    var row = 1;
    function rich(parts, size, color) {
      if (typeof parts === 'string') return parts;
      return { richText: parts.map(function (part) {
        return typeof part === 'string' ? { text: part, font: font({ size: size, color: { argb: color || COLOR.text } }) }
          : { text: part[0], font: font({ size: size, bold: true, color: { argb: COLOR.dark } }) };
      }) };
    }
    function plain(parts) { return typeof parts === 'string' ? parts : parts.map(function (x) { return typeof x === 'string' ? x : x[0]; }).join(''); }
    function line(icon, text, opts) {
      opts = opts || {};
      var size = opts.size || 10.5;
      var b = ws.getCell('B' + row), c = ws.getCell('C' + row);
      if (!opts.where) ws.mergeCells('C' + row + ':D' + row);
      b.value = icon || null;
      c.value = text === null ? null : rich(text, size, opts.color);
      b.alignment = { horizontal: 'center', vertical: 'middle' };
      c.alignment = { wrapText: true, vertical: 'middle', indent: opts.fill ? 1 : 0 };
      b.font = font({ bold: true, color: { argb: opts.iconColor || COLOR.header }, size: size });
      if (typeof text === 'string' || text === null) c.font = font({ bold: !!opts.bold, size: size, italic: !!opts.italic, color: { argb: opts.color || COLOR.text } });
      if (opts.fill) { b.fill = fill(opts.fill); c.fill = fill(opts.fill); }
      if (opts.where) {
        var d = ws.getCell('D' + row);
        d.value = opts.where.sheet ? link(opts.where.sheet, '➜ ' + opts.where.sheet) : opts.where.text;
        d.font = font({ size: 10, bold: !!opts.where.sheet, underline: !!opts.where.sheet, color: { argb: opts.where.sheet ? COLOR.infoFont : COLOR.muted } });
        d.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
        if (opts.fill) d.fill = fill(opts.fill);
        [b, c, d].forEach(function (cell) { cell.border = { bottom: THIN }; });
      }
      ws.getRow(row).height = opts.height || textHeight(plain(text || ''), opts.where ? 62 : WIDTH - 2, size, 20);
      row++;
      return c;
    }
    function heading(text) {
      ws.getRow(row).height = 10;
      row++;
      ws.mergeCells('B' + row + ':D' + row);
      var h = ws.getCell('B' + row);
      h.value = text;
      h.font = font({ bold: true, size: 11.5, color: { argb: COLOR.white } });
      h.fill = fill(COLOR.header);
      h.alignment = { indent: 1, vertical: 'middle' };
      ws.getRow(row).height = 22;
      row++;
    }

    ws.mergeCells('B1:D1');
    ws.getCell('B1').value = 'Meu negócio — controle simples';
    ws.getCell('B1').font = font({ size: 22, bold: true, color: { argb: COLOR.white } });
    ws.getCell('B1').fill = fill(COLOR.dark);
    ws.getCell('B1').alignment = { indent: 1, vertical: 'bottom' };
    ws.getRow(1).height = 40;
    ws.mergeCells('B2:D2');
    ws.getCell('B2').value = A.subtitle;
    ws.getCell('B2').font = font({ size: 11.5, italic: true, color: { argb: 'FFDCEBE3' } });
    ws.getCell('B2').fill = fill(COLOR.dark);
    ws.getCell('B2').alignment = { indent: 1, vertical: 'top' };
    ws.getRow(2).height = 26;
    ws.getRow(3).height = 10;
    row = 4;
    line('ℹ', 'Os dados desta planilha são seus. Use exemplos fictícios apenas para aprender.' + (workshop ? ' Na oficina, você pode trabalhar só com números fictícios.' : ''), { bold: true, fill: COLOR.info, iconColor: COLOR.infoFont });
    if (A.intro) line('➜', A.intro, { color: COLOR.text });

    heading('SEU PONTO DE PARTIDA');
    line('★', [['Negócio: '], p.name], { size: 11.5 });
    if (workshop) {
      line('➜', [['Hoje, na oficina: '], 'comece pela aba ', [S.fluxo], ' (Meu controle simples) e siga o roteiro abaixo.']);
      line('★', [['Foco sugerido pelo seu diagnóstico: '], p.focus.label + '. Ele já aparece na 1ª linha da aba ' + S.plano + '.']);
      line('➜', [['Depois da oficina: '], p.focus.step]);
    } else {
      line('★', [['Foco sugerido: '], p.focus.label]);
      line('➜', [['Primeiro passo: '], p.focus.step]);
      line('➜', [['Comece pela aba: '], p.startTab + ' (depois veja o resultado no PAINEL).'], { where: { sheet: p.startTab } });
    }
    p.recommendations.forEach(function (s, k) {
      var parts = [['Sugestão ' + (k + 1) + ' do diagnóstico: '], s.title + '.'];
      if (s.action) parts.push([' Experimente: '], s.action);
      if (s.check) parts.push([' Como conferir: '], s.check);
      line('•', parts, { size: 10 });
    });
    line(null, 'Estas sugestões vêm das suas respostas de hoje e foram montadas no seu aparelho. Não são nota nem diagnóstico definitivo do negócio.' +
      (p.recommendations.length ? ' Elas já aparecem nas primeiras linhas da aba ' + S.plano + ', para você ajustar ou apagar.' : ''), { italic: true, color: COLOR.muted, size: 9.5 });

    if (p.tool === 'paper') {
      heading('PREFERE PAPEL? PODE IMPRIMIR E PREENCHER À MÃO');
      line('🖨', [['Imprima as abas que vai usar. '], 'Cada aba de registro (' + paperTabs(p).join(', ') + ') sai em uma folha A4 deitada, com cabeçalho e linhas em branco para escrever à mão.']);
      line('•', 'Para folhas limpas, apague antes as linhas marcadas "' + EX + '". Em Arquivo › Imprimir, escolha "Imprimir Planilhas Ativas".');
      line('•', 'Os totais e o PAINEL só calculam o que for digitado. Se quiser, passe as anotações do papel para o arquivo uma vez por semana.');
    } else if (p.tool === 'phone') {
      heading('VAI USAR NO CELULAR?');
      line('📱', [['Abra no app gratuito Google Planilhas ou Microsoft Excel. '], 'Pelo navegador do celular a planilha pode aparecer sem cores ou sem as listas de opções.']);
      line('•', 'Gire o celular na horizontal. Toque numa célula azul e use a setinha ou a lista para escolher uma opção.');
      line('•', 'Comece por uma aba só (' + p.startTab + '). Se o arquivo não abrir, volte à página do diagnóstico e baixe a versão simples (CSV).');
    }

    if (workshop) {
      heading('COMO ESTA PLANILHA ACOMPANHA A OFICINA DE HOJE');
      var hr = row;
      [['B', 'Caderno'], ['C', 'Etapa da oficina'], ['D', 'Onde fazer']].forEach(function (h) {
        var c = ws.getCell(h[0] + hr);
        c.value = h[1];
        c.font = font({ bold: true, size: 9.5, color: { argb: COLOR.muted } });
        c.alignment = { horizontal: h[0] === 'B' ? 'center' : 'left', vertical: 'middle', indent: h[0] === 'B' ? 0 : 1 };
        c.border = { bottom: { style: 'thin', color: { argb: COLOR.muted } } };
      });
      ws.getRow(hr).height = 18;
      row++;
      var paper = COLOR.white, sheetFill = COLOR.good;
      [
        ['01–02', [['Antes de calcular · Quatro ideias. '], 'Conversa, diagnóstico e conceitos (pp. 2–3).'], { text: 'No caderno' }, paper],
        ['03–05', [['Casos A, B e C. '], 'Resolva no papel (pp. 4–6). O Caso A aparece como EXEMPLO no FLUXO DE CAIXA, sem as respostas.'], { text: 'No caderno' }, paper],
        ['06', [['Meu controle simples. '], 'Saldo inicial + entradas − saídas = saldo final (p. 7).'], { sheet: S.fluxo }, sheetFill],
        ['06', [['Compromissos futuros. '], 'A receber e a pagar, sem misturar com o caixa. Veja o saldo previsto (p. 7).'], { sheet: S.contas }, sheetFill],
        ['+', [['Orçamento do mês. '], 'Planejado x realizado por categoria. Complementa o caderno, como previsto no plano da oficina.'], { sheet: S.orcamento }, sheetFill],
        ['07', [['Uma ação de melhoria. '], 'Problema, ação concreta, quando começo e reviso, como vou perceber se funcionou (p. 8).'], { sheet: S.plano }, sheetFill],
        ['07', [['Verificação final e avaliação breve. '], 'Só no caderno (p. 8). A planilha não dá nota.'], { text: 'No caderno' }, paper],
        ['depois', [['Acompanhar. '], 'O PAINEL resume tudo. Abas em cinza são complementares, para usar se fizer sentido.'], { sheet: S.painel }, paper]
      ].forEach(function (step) {
        line(step[0], step[1], { where: step[2], fill: step[3], size: 10.5, iconColor: COLOR.dark });
      });

      heading('QUATRO IDEIAS PARA USAR HOJE');
      line('1', [['Caixa é movimento de dinheiro. '], 'Registre quando o dinheiro efetivamente entra ou sai. Venda a receber fica numa lista separada até o recebimento.']);
      line('2', [['Saldo não é lucro. '], 'Saldo final = saldo inicial + entradas − saídas. Lucro exige considerar receitas, custos e despesas do período.']);
      line('3', [['Separe as finalidades. '], 'Compra pessoal com dinheiro do negócio é retirada pessoal: reduz o caixa, mas não é despesa do negócio.']);
      line('4', [['Planejar é antecipar datas. '], 'Se o saldo previsto ficar negativo, identifique o momento do aperto antes do vencimento. Previsão não é recebimento garantido.']);
    } else {
      heading('COMO USAR EM ' + (A.primary ? A.order.length - 1 : 7) + ' PASSOS');
      var steps = A.primary ? A.order.slice(1) : [S.painel, S.fluxo, S.contas, S.clientes, S.preco, S.plano, S.rotina];
      steps.forEach(function (sheet, k) {
        var text = STEP_TEXT[sheet];
        if (!A.primary && sheet === S.painel) text = 'Leia esta aba e abra o PAINEL. Ele se atualiza sozinho — você não digita nele.';
        line(String(k + 1), text + (isSecondary(p, sheet) ? ' (complementar)' : ''), { where: { sheet: sheet }, color: isSecondary(p, sheet) ? COLOR.muted : COLOR.text });
      });
    }

    heading('LEGENDA DE CORES (sempre com símbolo ou texto)');
    line('✎', 'Azul claro: células para você preencher. "opcional" no topo da coluna = pode deixar em branco.', { fill: COLOR.input, iconColor: COLOR.infoFont });
    line('=', 'Cinza: células calculadas por fórmula. Não digite nelas.', { fill: COLOR.calc, iconColor: COLOR.muted });
    line('✔', 'Verde: situação acompanhada, pago, recebido ou concluído.', { fill: COLOR.good, iconColor: COLOR.goodFont });
    line('⚠', 'Amarelo: atenção — vence em breve, ainda não realizado ou merece revisão.', { fill: COLOR.warn, iconColor: COLOR.warnFont });
    line('✖', 'Vermelho suave: precisa de ação — vencido, atrasado ou preenchimento a corrigir.', { fill: COLOR.bad, iconColor: COLOR.badFont });
    line('EX', 'Texto marrom em itálico: linha de EXEMPLO fictícia. Não entra em nenhum total. Apague antes de usar.', { fill: COLOR.exFill, iconColor: COLOR.exFont, italic: true, color: COLOR.exFont });

    heading('BOM SABER');
    var tips = workshop ? [
      ['•', 'Para apagar um exemplo: selecione a linha inteira, clique com o botão direito e escolha Excluir. O PAINEL avisa enquanto houver exemplos.'],
      ['•', 'Em branco é "não informado", não é zero. Preencha só o que você sabe.'],
      ['🔒', 'Privacidade: use apelidos ou referências curtas para clientes. Não registre CPF, senhas, número de cartão ou dados bancários.'],
      ['!', 'A planilha é uma ferramenta de organização: não dá nota ao negócio, não comprova aprendizado e não promete lucro. A verificação final e a avaliação da oficina ficam no caderno. Não é contabilidade oficial nem orientação tributária.']
    ] : [
      ['•', 'Não é preciso preencher tudo no primeiro dia. Comece com 7 dias de registros no FLUXO DE CAIXA.'],
      ['•', 'Apague as linhas marcadas "' + EX + '": selecione a linha inteira, clique com o botão direito e escolha Excluir. O PAINEL avisa enquanto houver exemplos e não os soma.'],
      ['•', 'Faça uma cópia do arquivo de vez em quando como backup (ex.: meu-negocio_2026-11.xlsx).'],
      ['•', 'Revisão semanal: uma vez por semana, olhe o PAINEL, responda as perguntas do ritual na aba ROTINA SEMANAL e anote a data da revisão.'],
      ['🔒', 'Privacidade: use apelidos ou referências curtas para clientes. Não registre CPF, senhas, número de cartão, dados bancários completos ou informações pessoais desnecessárias.'],
      ['•', 'As listas de opções são sugestões. Em Categoria, Forma de pagamento e Forma de contato você pode digitar outro valor que faça sentido para o seu negócio.'],
      ['!', 'Esta planilha é uma ferramenta de aprendizagem e organização. Não é contabilidade oficial e não substitui contador nem orientação jurídica ou tributária. Os números mostram caixa, não lucro.']
    ];
    tips.forEach(function (item) { line(item[0], item[1], { color: item[0] === '!' ? COLOR.muted : COLOR.text }); });

    if (!workshop) {
      heading('QUAL É O SEU TIPO DE NEGÓCIO? (opcional — só orienta a leitura)');
      var choice = line('✎', null, { fill: COLOR.input, iconColor: COLOR.infoFont });
      var choiceAddr = 'C' + (row - 1);
      choice.border = BOX;
      ws.dataValidations.add(choiceAddr, listValidation(LISTS.tipoNegocio, true));
      var guide = line('➜', null, { height: 34 });
      guide.value = { formula: 'IF(' + choiceAddr + '="","Escolha uma opção na célula azul acima (clique e use a setinha).",' +
        'IF(' + choiceAddr + '="Vendo produtos","Comece por FLUXO DE CAIXA e PRECIFICAÇÃO (material/mercadoria e embalagem).",' +
        'IF(' + choiceAddr + '="Presto serviços","Comece por FLUXO DE CAIXA e CLIENTES E PEDIDOS; na PRECIFICAÇÃO, valorize as horas de trabalho.",' +
        'IF(' + choiceAddr + '="Faço produtos sob encomenda","Use CLIENTES E PEDIDOS para prazos e CONTAS E COMPROMISSOS para sinais e valores a receber.",' +
        'IF(' + choiceAddr + '="Faço os dois","Use a coluna Produto/serviço para separar o que é venda e o que é serviço.",' +
        '"Comece só pelo FLUXO DE CAIXA por 7 dias e depois escolha um foco no PLANO DE MELHORIA.")))))' };
    }

    pageSetup(ws, 'B1:D' + (row - 1), false, null, true);
    return ws;
  }

  var BUILDERS = {};
  BUILDERS[S.leia] = leiaSheet; BUILDERS[S.painel] = painelSheet; BUILDERS[S.fluxo] = fluxoSheet; BUILDERS[S.contas] = contasSheet;
  BUILDERS[S.clientes] = clientesSheet; BUILDERS[S.preco] = precoSheet; BUILDERS[S.rotina] = rotinaSheet; BUILDERS[S.plano] = planoSheet;
  BUILDERS[S.orcamento] = orcamentoSheet;

  // O ExcelJS não ajusta alturas de linha: estimamos a largura do texto (Calibri) e só aumentamos alturas.
  var NARROW = 'iljtfrI.,;:!|\'()[]{} ', WIDE = 'mwMW—@%█', UPPER = /[A-ZÀ-Ý]/;
  function textUnits(text, size, bold) {
    var u = 0;
    for (var k = 0; k < text.length; k++) {
      var ch = text.charAt(k);
      u += NARROW.indexOf(ch) >= 0 ? 0.52 : WIDE.indexOf(ch) >= 0 ? 1.45 : UPPER.test(ch) ? 1.18 : /[0-9]/.test(ch) ? 0.98 : 0.9;
    }
    return u * (size / 11) * (bold ? 1.08 : 1);
  }
  function wrappedLines(text, width, size, bold) {
    var lines = 0;
    String(text).split('\n').forEach(function (para) {
      var used = 0, n = 1, space = textUnits(' ', size, bold);
      para.split(' ').forEach(function (word) {
        var w = textUnits(word, size, bold);
        if (used > 0 && used + space + w > width) { n++; used = 0; }
        if (w > width) { n += Math.floor(w / width); used = w % width; } else used += (used > 0 ? space : 0) + w;
      });
      lines += n;
    });
    return lines;
  }
  function cellText(cell) {
    var v = cell.value;
    if (v === null || v === undefined) return null;
    if (v.richText) return { text: v.richText.map(function (r) { return r.text; }).join(''), size: Math.max.apply(null, v.richText.map(function (r) { return (r.font && r.font.size) || 11; })), bold: v.richText.some(function (r) { return r.font && r.font.bold; }) };
    if (v.formula) {
      var strings = (v.formula.match(/"(?:[^"]|"")*"/g) || []).map(function (s) { return s.slice(1, -1).replace(/""/g, '"'); });
      if (!strings.length) return null;
      return { text: strings.reduce(function (a, b) { return b.length > a.length ? b : a; }), formula: true };
    }
    if (v.text) return { text: String(v.text) };
    if (typeof v === 'string') return { text: v };
    return null;
  }
  function fitRows(ws) {
    var merges = {};
    (ws.model.merges || []).forEach(function (range) {
      var parts = range.split(':'), a = ws.getCell(parts[0]), b = ws.getCell(parts[1]);
      merges[parts[0]] = { c1: a.col, c2: b.col, r1: a.row, r2: b.row };
    });
    function colWidth(c) { var col = ws.getColumn(c); return col.hidden ? 0 : (col.width || 8.43); }
    ws.eachRow({ includeEmpty: false }, function (row, r) {
      var need = 0;
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        if (cell.isMerged && cell.master !== cell) return;
        var m = merges[cell.address] || { c1: c, c2: c, r1: r, r2: r };
        if (m.r2 !== m.r1) return;
        var info = cellText(cell);
        if (!info || !info.text) return;
        var f = cell.font || {}, size = info.size || f.size || 11, bold = info.bold || !!f.bold;
        var al = cell.alignment || {};
        var width = 0;
        for (var k = m.c1; k <= m.c2; k++) width += colWidth(k);
        width = (width - 0.8 - (al.indent || 0) * 1.35) * 0.88;
        var lines = 1;
        if (al.wrapText) lines = wrappedLines(info.text, width, size, bold);
        else if (m.c2 > m.c1 && textUnits(info.text, size, bold) > width) {
          cell.alignment = Object.assign({}, al, { wrapText: true });
          lines = wrappedLines(info.text, width, size, bold);
        }
        need = Math.max(need, Math.ceil(lines * size * 1.37 + (lines > 1 ? 5 : 3)));
      });
      if (need > (row.height || 15)) row.height = Math.min(409, need);
    });
  }

  function buildWorkbook(ExcelJS, opts) {
    var p = personalize(opts);
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Oficina de Educação Financeira e Planejamento — modelo gerado no navegador';
    wb.title = 'Meu negócio — controle simples' + (p.activity.label ? ' · ' + p.activity.label : '');
    wb.created = p.today;
    wb.modified = p.today;
    wb.calcProperties.fullCalcOnLoad = true;
    p.activity.order.forEach(function (name) { fitRows(BUILDERS[name](wb, p)); });
    wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: 'visible' }];
    return wb;
  }

  function csvCell(value) {
    var s = value === null || value === undefined ? '' : String(value);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return '"' + s.replace(/"/g, '""') + '"';
  }
  function money(n) { return n === null || n === undefined ? '' : n.toFixed(2).replace('.', ','); }
  function brDate(d) { return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + d.getUTCFullYear(); }

  var CSV_HEADER = ['Data', 'Tipo', 'Descrição', 'Categoria', 'Cliente/fornecedor', 'Produto/serviço', 'Forma de pagamento', 'Entrada (R$)', 'Saída (R$)', 'Pago/recebido?', 'Observação'];
  function buildCsv(opts) {
    var p = personalize(opts), t = p.today;
    var intro = 'Modelo simples (CSV) para ' + p.name + (p.activity.label ? ' · ' + p.activity.label : '') + '. Foco sugerido: ' + p.focus.label + '. Primeiro passo: ' + p.focus.step;
    var rows = [
      CSV_HEADER,
      ['', 'Orientação', intro, '', '', '', '', '', '', '', 'Apague esta linha. Use uma linha por movimento; Entrada OU Saída, nunca as duas.'],
      [brDate(utcDate(t, -4)), 'Entrada', 'Venda de 2 camisetas (produto)', 'Venda de produto', 'Cliente A', 'Camiseta', 'Pix', money(120), '', 'Sim', EX],
      [brDate(utcDate(t, -3)), 'Entrada', 'Afiação de 3 facas (serviço)', 'Prestação de serviço', 'Cliente B', 'Afiação', 'Dinheiro', money(45), '', 'Sim', EX],
      [brDate(utcDate(t, -3)), 'Saída', 'Compra de material, embalagem; "linha extra"', 'Material/mercadoria', 'Fornecedor X', 'Material', 'Pix', '', money(60), 'Sim', EX],
      [brDate(utcDate(t, -2)), 'Retirada pessoal', 'Dinheiro do negócio usado em casa (não é despesa do negócio)', 'Retirada pessoal', '', '', 'Pix', '', money(50), 'Sim', EX],
      [brDate(utcDate(t, -1)), 'Entrada', 'Encomenda entregue; cliente paga depois (ainda não entrou no caixa)', 'Encomenda', 'Cliente C', 'Bolo sob encomenda', 'A prazo/fiado', money(80), '', 'Não', EX]
    ];
    return '\ufeff' + rows.map(function (r) { return r.map(csvCell).join(';'); }).join('\r\n') + '\r\n';
  }

  function slug(value) {
    return String(value || 'meu-negocio').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'meu-negocio';
  }

  // Texto curto para o formulário explicar o arquivo antes do download.
  function describe(opts) {
    var p = personalize(opts), a = p.activity;
    var text;
    if (a.key === 'workshop') {
      text = 'Uma planilha que acompanha a oficina, aba por aba: Meu controle simples (FLUXO DE CAIXA) → Compromissos futuros → Orçamento do mês → Uma ação de melhoria. Comece pela aba LEIA PRIMEIRO e depois vá para ' + p.startTab + '.';
    } else if (a.primary) {
      text = 'Uma planilha com ' + a.order.length + ' abas, ordenadas para o ' + a.label.split(' · ')[0] + ' (abas em cinza são complementares). Comece pela aba LEIA PRIMEIRO e depois vá para ' + p.startTab + '.';
    } else {
      text = 'Uma planilha com 8 abas que serve para quem vende produtos, presta serviços ou faz os dois. Comece pela aba LEIA PRIMEIRO e depois vá para ' + p.startTab + '.';
    }
    if (p.recommendations.length) text += ' Suas sugestões (ação e como conferir) já vêm na aba ' + S.plano + ', para você ajustar.';
    if (p.tool === 'paper') text += ' Prefere papel? As abas de registro podem ser impressas em A4, com linhas em branco para escrever à mão.';
    if (p.tool === 'phone') text += ' No celular, abra com o app Google Planilhas ou Excel; se não abrir, use a versão simples (CSV) abaixo.';
    return { activity: a.key, tool: p.tool, sheets: a.order.slice(), startTab: p.startTab, text: text };
  }

  root.PlanilhaModelo = {
    SHEETS: SHEETS.slice(),
    EXAMPLE_MARKER: EX,
    CSV_HEADER: CSV_HEADER.slice(),
    sheetsFor: function (activity) { return activityFor(activity).order.slice(); },
    activityKey: function (activity) { return activityFor(activity).key; },
    focusFor: function (priority) { return FOCUS[priority] || FOCUS.unsure; },
    describe: describe,
    buildWorkbook: buildWorkbook,
    buildCsv: buildCsv,
    fileName: function (business, ext, activity) { return slug(business) + '_modelo_controle_negocio' + activityFor(activity).suffix + '.' + ext; }
  };
})(typeof window !== 'undefined' ? window : globalThis);
