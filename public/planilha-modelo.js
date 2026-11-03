/*
 * Gera, no navegador, a planilha simples de controle (.xlsx) e a versão CSV.
 * Nenhuma resposta é enviada a servidor: este arquivo só monta o arquivo localmente.
 * Uso: PlanilhaModelo.buildWorkbook(ExcelJS, opts) e PlanilhaModelo.buildCsv(opts),
 * em que opts = { business } (outros campos do formulário são ignorados).
 */
(function (root) {
  'use strict';

  var SHEETS = ['MEU CAIXA', 'CONTAS A PAGAR E RECEBER', 'PEDIDOS', 'PREÇO'];
  var HEAD = 5, FIRST = 6;
  var C = {
    title: 'FF1F4E3D', header: 'FF2C6650', white: 'FFFFFFFF', text: 'FF1E2B26', muted: 'FF5F6F68',
    help: 'FFF4F7F5', total: 'FFE3F2E6', calc: 'FFF1F3F1', example: 'FF8A8F8C', exampleFill: 'FFFAFAF7',
    line: 'FFD3DDD7', good: 'FFE3F2E6', goodFont: 'FF1E6B3A', warn: 'FFFFF3CD', warnFont: 'FF7A5600',
    bad: 'FFFBE3E3', badFont: 'FF9B2222'
  };
  var TABS = ['FF2C6650', 'FFB7791F', 'FF2F6FA8', 'FFB0473B'];
  var BRL = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"';
  var DATE = 'dd/mm/yyyy';
  var THIN = { style: 'thin', color: { argb: C.line } };
  var BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }
  function font(o) { var f = { name: 'Calibri', size: 11, color: { argb: C.text } }; for (var k in o) f[k] = o[k]; return f; }
  function col(n) { return String.fromCharCode(64 + n); }
  function day(base, offset) { return new Date(Date.UTC(base.getFullYear(), base.getMonth(), base.getDate() + (offset || 0))); }
  function cleanName(business) { return String(business || '').replace(/\s+/g, ' ').trim().slice(0, 60) || 'Meu negócio'; }

  var LIST_SIM_NAO = { type: 'list', allowBlank: true, formulae: ['"Sim,Não"'], showErrorMessage: true, errorTitle: 'Escolha na lista', error: 'Use Sim ou Não.' };
  var V_DATE = { type: 'date', operator: 'greaterThan', allowBlank: true, formulae: [new Date(Date.UTC(2000, 0, 1))], showErrorMessage: true, errorTitle: 'Data', error: 'Digite uma data, por exemplo 05/11/2026.' };
  var V_MONEY = { type: 'decimal', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [0], showErrorMessage: true, errorTitle: 'Valor', error: 'Digite um valor em reais, sem sinal de menos.' };
  var V_HOURS = { type: 'decimal', operator: 'between', allowBlank: true, formulae: [0, 1000], showErrorMessage: true, errorTitle: 'Horas', error: 'Digite as horas, por exemplo 1,5.' };

  // Monta uma aba: título, ajuda curta, linha de totais, uma linha de exemplo e a tabela.
  function sheet(wb, idx, spec) {
    var n = spec.columns.length, L = col(n), last = FIRST + spec.rows - 1;
    var ws = wb.addWorksheet(SHEETS[idx], {
      properties: { tabColor: { argb: TABS[idx] } },
      views: [{ state: 'frozen', ySplit: HEAD, showGridLines: false }],
      pageSetup: {
        paperSize: 9, orientation: spec.landscape ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        horizontalCentered: true, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 }
      },
      headerFooter: { oddFooter: '&L&A&RPágina &P de &N' }
    });
    ws.columns = spec.columns.map(function (c) { return { width: c.w }; });
    var totalWidth = spec.columns.reduce(function (a, c) { return a + c.w; }, 0);

    ws.mergeCells('A1:' + L + '1');
    var t = ws.getCell('A1');
    t.value = spec.title;
    t.font = font({ bold: true, size: 15, color: { argb: C.white } });
    t.fill = fill(C.title);
    t.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(1).height = 32;

    ws.mergeCells('A2:' + L + '2');
    var h = ws.getCell('A2');
    h.value = spec.help;
    h.font = font({ size: 11, color: { argb: C.muted } });
    h.fill = fill(C.help);
    h.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
    ws.getRow(2).height = Math.max(22, Math.ceil(spec.help.length / (totalWidth * 1.05)) * 15 + 8);

    ws.getRow(3).height = 24;
    spec.totals.forEach(function (tot) {
      if (tot.label.from !== tot.label.to) ws.mergeCells(tot.label.from + '3:' + tot.label.to + '3');
      var lc = ws.getCell(tot.label.from + '3');
      lc.value = tot.label.text;
      lc.font = font({ bold: !tot.note, italic: !!tot.note, size: tot.note ? 10.5 : 11, color: { argb: tot.note ? C.muted : C.title } });
      lc.alignment = { vertical: 'middle', horizontal: tot.note ? 'left' : 'right', indent: 1, wrapText: !!tot.note };
      lc.fill = fill(C.total);
      if (tot.cell) {
        var vc = ws.getCell(tot.cell + '3');
        vc.value = { formula: tot.formula };
        vc.numFmt = tot.fmt || BRL;
        vc.font = font({ bold: true, size: 12, color: { argb: C.title } });
        vc.fill = fill(C.total);
        vc.alignment = { vertical: 'middle', horizontal: 'right' };
      }
    });

    ws.getRow(4).height = 20;
    spec.example.forEach(function (v, i) {
      var c = ws.getCell(4, i + 1);
      if (v !== null && v !== undefined) c.value = v;
      c.font = font({ italic: true, size: 10.5, color: { argb: C.example } });
      c.fill = fill(C.exampleFill);
      c.border = { bottom: THIN };
      c.numFmt = spec.columns[i].fmt || 'General';
      c.alignment = { vertical: 'middle', horizontal: spec.columns[i].align || (spec.columns[i].fmt ? 'right' : 'left'), indent: spec.columns[i].fmt ? 0 : 1 };
    });

    ws.getRow(HEAD).height = 34;
    spec.columns.forEach(function (c, i) {
      var cell = ws.getCell(HEAD, i + 1);
      cell.value = c.h;
      cell.font = font({ bold: true, color: { argb: C.white } });
      cell.fill = fill(c.calc ? C.muted : C.header);
      cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
      cell.border = BOX;
    });

    for (var r = FIRST; r <= last; r++) {
      ws.getRow(r).height = 20;
      spec.columns.forEach(function (c, i) {
        var cell = ws.getCell(r, i + 1);
        cell.border = BOX;
        cell.font = font({ bold: !!c.calc && !!c.bold });
        cell.alignment = { vertical: 'middle', horizontal: c.align || (c.fmt && c.fmt !== DATE ? 'right' : (c.fmt === DATE ? 'center' : 'left')), wrapText: !!c.wrap, indent: c.fmt ? 0 : 1 };
        if (c.fmt) cell.numFmt = c.fmt;
        if (c.calc) {
          cell.fill = fill(C.calc);
          cell.value = { formula: c.calc(r) };
        } else if (c.v) {
          cell.dataValidation = c.v;
        }
      });
    }
    if (spec.rules) spec.rules(ws, last);
    ws.pageSetup.printArea = 'A1:' + L + (FIRST + 24);
    ws.pageSetup.printTitlesRow = HEAD + ':' + HEAD;
    return ws;
  }

  function rule(ref, formula, bg, fg, bold) {
    return { ref: ref, rules: [{ type: 'expression', priority: 1, formulae: [formula], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: bg } }, font: { color: { argb: fg }, bold: !!bold } } }] };
  }

  function caixa(wb, name, today) {
    return sheet(wb, 0, {
      title: 'MEU CAIXA — ' + name,
      help: 'Anote cada dinheiro que entra e que sai, no dia em que acontece. Na 1ª linha, coloque em "Entrou" o dinheiro que você já tem hoje. Dinheiro do negócio usado em casa também é saída.',
      rows: 200,
      columns: [
        { h: 'Data', w: 13, fmt: DATE, v: V_DATE },
        { h: 'O que foi', w: 42, wrap: true },
        { h: 'Entrou (R$)', w: 15, fmt: BRL, v: V_MONEY },
        { h: 'Saiu (R$)', w: 15, fmt: BRL, v: V_MONEY },
        { h: 'Saldo (R$)', w: 16, fmt: BRL, bold: true, calc: function (r) { return 'IF(AND(C' + r + '="",D' + r + '=""),"",SUM(C$' + FIRST + ':C' + r + ')-SUM(D$' + FIRST + ':D' + r + '))'; } }
      ],
      totals: [
        { label: { from: 'A', to: 'B', text: 'Totais até agora' }, cell: 'C', formula: 'SUM(C' + FIRST + ':C205)' },
        { label: { from: 'D', to: 'D', text: '' }, cell: 'D', formula: 'SUM(D' + FIRST + ':D205)' },
        { label: { from: 'E', to: 'E', text: '' }, cell: 'E', formula: 'C3-D3' }
      ],
      example: [day(today, -1), 'Exemplo: venda de 2 bolos', 40, null, null]
    });
  }

  function contas(wb, name, today) {
    return sheet(wb, 1, {
      title: 'CONTAS A PAGAR E RECEBER — ' + name,
      help: 'Anote o que você vai pagar ou receber e quando vence. Quando pagar ou receber, marque "Sim" e anote também no MEU CAIXA.',
      rows: 100, landscape: true,
      columns: [
        { h: 'Vencimento', w: 13, fmt: DATE, v: V_DATE },
        { h: 'O que é', w: 34, wrap: true },
        { h: 'Pagar ou receber?', w: 15, align: 'center', v: { type: 'list', allowBlank: true, formulae: ['"A pagar,A receber"'], showErrorMessage: true, errorTitle: 'Escolha na lista', error: 'Use A pagar ou A receber.' } },
        { h: 'Valor (R$)', w: 15, fmt: BRL, v: V_MONEY },
        { h: 'Já foi pago ou recebido?', w: 15, align: 'center', v: LIST_SIM_NAO },
        { h: 'Situação', w: 20, align: 'center', bold: true, calc: function (r) {
          var a = 'A' + r;
          return 'IF(OR(' + a + '="",D' + r + '=""),"",IF(E' + r + '="Sim","Concluída",IF(' + a + '<TODAY(),"Atrasada",IF(' + a + '=TODAY(),"Vence hoje",IF(' + a + '-TODAY()=1,"Vence amanhã",IF(' + a + '-TODAY()<=7,"Vence em "&(' + a + '-TODAY())&" dias","Em dia"))))))';
        } }
      ],
      totals: [
        { label: { from: 'A', to: 'C', text: 'Falta pagar' }, cell: 'D', formula: 'SUMIFS(D' + FIRST + ':D105,C' + FIRST + ':C105,"A pagar",E' + FIRST + ':E105,"<>Sim")' },
        { label: { from: 'E', to: 'E', text: 'Falta receber' }, cell: 'F', formula: 'SUMIFS(D' + FIRST + ':D105,C' + FIRST + ':C105,"A receber",E' + FIRST + ':E105,"<>Sim")' }
      ],
      example: [day(today, 5), 'Exemplo: aluguel do ponto', 'A pagar', 300, 'Não', null],
      rules: function (ws, last) {
        var ref = 'F' + FIRST + ':F' + last;
        ws.addConditionalFormatting(rule(ref, '$F' + FIRST + '="Atrasada"', C.bad, C.badFont, true));
        ws.addConditionalFormatting(rule(ref, 'LEFT($F' + FIRST + ',5)="Vence"', C.warn, C.warnFont, true));
        ws.addConditionalFormatting(rule(ref, '$F' + FIRST + '="Concluída"', C.good, C.goodFont));
      }
    });
  }

  function pedidos(wb, name, today) {
    return sheet(wb, 2, {
      title: 'PEDIDOS — ' + name,
      help: 'Um pedido ou serviço combinado por linha. Use só o apelido do cliente: não anote CPF, telefone ou dados pessoais.',
      rows: 100, landscape: true,
      columns: [
        { h: 'Data do pedido', w: 13, fmt: DATE, v: V_DATE },
        { h: 'Cliente (apelido)', w: 18, wrap: true },
        { h: 'O que pediu', w: 32, wrap: true },
        { h: 'Valor (R$)', w: 14, fmt: BRL, v: V_MONEY },
        { h: 'Entregar até', w: 13, fmt: DATE, v: V_DATE },
        { h: 'Pago?', w: 10, align: 'center', v: LIST_SIM_NAO },
        { h: 'Entregue?', w: 11, align: 'center', v: LIST_SIM_NAO },
        { h: 'Situação', w: 17, align: 'center', bold: true, calc: function (r) {
          return 'IF(C' + r + '="","",IF(AND(F' + r + '="Sim",G' + r + '="Sim"),"Concluído",IF(AND(G' + r + '<>"Sim",ISNUMBER(E' + r + '),E' + r + '<TODAY()),"Atrasado",IF(G' + r + '="Sim","Falta receber",IF(F' + r + '="Sim","Falta entregar","Em aberto")))))';
        } }
      ],
      totals: [
        { label: { from: 'A', to: 'C', text: 'Falta receber dos pedidos' }, cell: 'D', formula: 'SUMIFS(D' + FIRST + ':D105,C' + FIRST + ':C105,"<>",F' + FIRST + ':F105,"<>Sim")' },
        { label: { from: 'E', to: 'G', text: 'Pedidos ainda não entregues' }, cell: 'H', fmt: '0;-0;"–"', formula: 'COUNTIFS(C' + FIRST + ':C105,"<>",G' + FIRST + ':G105,"<>Sim")' }
      ],
      example: [today, 'Ana (vizinha)', 'Exemplo: bolo de aniversário', 80, day(today, 3), 'Não', 'Não', null],
      rules: function (ws, last) {
        var ref = 'H' + FIRST + ':H' + last;
        ws.addConditionalFormatting(rule(ref, '$H' + FIRST + '="Atrasado"', C.bad, C.badFont, true));
        ws.addConditionalFormatting(rule(ref, 'LEFT($H' + FIRST + ',5)="Falta"', C.warn, C.warnFont));
        ws.addConditionalFormatting(rule(ref, '$H' + FIRST + '="Concluído"', C.good, C.goodFont));
      }
    });
  }

  function preco(wb, name) {
    return sheet(wb, 3, {
      title: 'PREÇO — ' + name,
      help: 'Uma linha por produto ou serviço. Custo total = material + outros custos + horas × valor da sua hora. Sobra = preço que você cobra − custo total.',
      rows: 30, landscape: true,
      columns: [
        { h: 'Produto ou serviço', w: 26, wrap: true },
        { h: 'Material (R$)', w: 13, fmt: BRL, v: V_MONEY },
        { h: 'Outros custos (R$)', w: 13, fmt: BRL, v: V_MONEY },
        { h: 'Horas de trabalho', w: 11, fmt: '0.0#', v: V_HOURS },
        { h: 'Valor da sua hora (R$)', w: 13, fmt: BRL, v: V_MONEY },
        { h: 'Custo total (R$)', w: 14, fmt: BRL, calc: function (r) { return 'IF(A' + r + '="","",N(B' + r + ')+N(C' + r + ')+N(D' + r + ')*N(E' + r + '))'; } },
        { h: 'Preço que você cobra (R$)', w: 14, fmt: BRL, v: V_MONEY },
        { h: 'Sobra (R$)', w: 13, fmt: BRL, bold: true, calc: function (r) { return 'IF(OR(A' + r + '="",G' + r + '=""),"",G' + r + '-F' + r + ')'; } }
      ],
      totals: [
        { note: true, label: { from: 'A', to: 'H', text: 'A sobra não é lucro: aluguel, contas, taxas e impostos ainda não entraram na conta. Se ficar vermelha, o preço não cobre nem o custo.' } }
      ],
      example: ['Exemplo: bolo de pote', 4.5, 1, 0.5, 20, 15.5, 18, 2.5],
      rules: function (ws, last) {
        ws.addConditionalFormatting(rule('H' + FIRST + ':H' + last, 'AND(ISNUMBER($H' + FIRST + '),$H' + FIRST + '<0)', C.bad, C.badFont, true));
      }
    });
  }

  function buildWorkbook(ExcelJS, opts) {
    var name = cleanName(opts && opts.business), today = new Date();
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Oficina de Educação Financeira e Planejamento — planilha gerada no navegador';
    wb.title = 'Planilha de controle — ' + name;
    wb.created = today;
    wb.calcProperties = { fullCalcOnLoad: true };
    caixa(wb, name, today);
    contas(wb, name, today);
    pedidos(wb, name, today);
    preco(wb, name);
    wb.views = [{ activeTab: 0, firstSheet: 0 }];
    return wb;
  }

  function csvCell(v) { return '"' + String(v === null || v === undefined ? '' : v).replace(/"/g, '""') + '"'; }
  function brDate(d) { return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + d.getUTCFullYear(); }
  var CSV_HEADER = ['Data', 'O que foi', 'Entrou (R$)', 'Saiu (R$)'];

  function buildCsv() {
    var t = new Date();
    var rows = [
      CSV_HEADER,
      [brDate(day(t, -1)), 'Exemplo: venda de 2 bolos (apague esta linha)', '40,00', ''],
      [brDate(day(t, -1)), 'Exemplo: compra de material (apague esta linha)', '', '15,00']
    ];
    return '\ufeff' + rows.map(function (r) { return r.map(csvCell).join(';'); }).join('\r\n') + '\r\n';
  }

  function slug(value) {
    return String(value || 'meu-negocio').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'meu-negocio';
  }

  function describe() {
    return {
      sheets: SHEETS.slice(),
      text: 'Uma planilha simples com 4 abas: MEU CAIXA, CONTAS A PAGAR E RECEBER, PEDIDOS e PREÇO. Comece pelo MEU CAIXA.'
    };
  }

  root.PlanilhaModelo = {
    SHEETS: SHEETS.slice(),
    CSV_HEADER: CSV_HEADER.slice(),
    describe: describe,
    buildWorkbook: buildWorkbook,
    buildCsv: buildCsv,
    fileName: function (business, ext) { return slug(business) + '_planilha_controle.' + ext; }
  };
})(typeof window !== 'undefined' ? window : globalThis);
