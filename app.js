/* =====================================================================
   PAINEL OPERACIONAL · RJ15 · Inventário
   Arquivo de ESTRUTURA (lógica): leitura dos Excel, confronto de IDs,
   cálculo de status, prioridade por onda, dashboard e PDF.
   Interface: index.html
   ===================================================================== */
(function (root) {
  'use strict';

  /* ---------- utilidades ---------- */
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9/ ]+/g, ' ').replace(/\s+/g, ' ').trim();
  const txt = v => String(v == null ? '' : v).trim();
  const KEY = v => txt(v).toUpperCase().replace(/\s+/g, ' ');
  // remove sujeira de leitura no início do código (ex.: ">H1_PM1" vira "H1_PM1")
  const cleanCode = v => txt(v).replace(/^[^A-Za-z0-9]+/, '').trim();
  const parseBR = s => { const m = txt(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/); return m ? new Date(+m[3], m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : 0; };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------- reconhecimento automático de colunas ---------- */
  const SYN = {
    id: ['id pacote', 'id do pacote', 'id', 'pacote', 'package id', 'package', 'parcel id', 'parcel', 'shipment id', 'tracking', 'tracking number', 'tracking id', 'waybill', 'barcode', 'codigo', 'etiqueta'],
    gaiola: ['gaiola', 'id gaiola', 'id da gaiola', 'cage', 'cage id', 'cage number', 'cage name', 'cage code', 'container'],
    rota: ['rota', 'route', 'route id', 'route name', 'route code', 'id rota', 'rota atual', 'rota separacao'],
    desc: ['descricao', 'description', 'item description', 'bairro', 'neighborhood', 'district', 'cidade', 'city', 'endereco', 'address', 'destino', 'destination', 'produto', 'product'],
    ciclo: ['ciclo', 'cycle', 'turno', 'periodo', 'period', 'am/pm', 'shift'],
    status: ['status gaiola', 'cage status', 'status', 'situacao', 'estado', 'state'],
    opBipou: ['operador bipou', 'operador que bipou', 'bipado por', 'scanned by', 'scanner', 'operador', 'operator', 'operator name', 'audited by', 'usuario', 'auditor', 'nome operador', 'user'],
    opFinal: ['operador finalizou', 'operador que finalizou', 'finalizado por', 'finalizou', 'usuario finalizou', 'operador final', 'finished by', 'finalized by', 'closed by', 'completed by', 'operator finished', 'finalizer'],
    tipo: ['tipo', 'type', 'divergencia', 'discrepancy', 'ocorrencia', 'issue', 'resultado', 'result'],
    onda: ['onda', 'wave', 'wave number']
  };
  Object.assign(SYN, {
    parada: ['parada', 'stop', 'stop number'],
    ponto: ['ponto', 'point'],
    local: ['localizacao do pacote', 'localizacao', 'location', 'package location'],
    detalhe: ['detalhe do pacote', 'detalhe', 'detail', 'package detail'],
    placa: ['placa', 'plate', 'license plate', 'vehicle'],
    data: ['data auditoria', 'data da auditoria', 'data', 'date', 'audit date', 'data hora', 'timestamp'],
    idRota: ['id da rota', 'id rota', 'route id'],
    rotaOrig: ['rota original', 'original route'],
    doca: ['doca', 'dock'],
    transp: ['transportadora', 'carrier'],
    tipoVeic: ['tipo de veiculo', 'vehicle type'],
    tipoRota: ['tipo de rota', 'route type'],
    tipoArea: ['tipo de area', 'area type'],
    cluster: ['cluster']
  });
  SYN.rota.push('n rota', 'no rota', 'numero rota', 'numero da rota', 'id da rota', 'route number', 'rota otimizada', 'optimized route');
  SYN.opBipou.unshift('rep auditoria', 'rep', 'representante');
  SYN.status = SYN.status.filter(s => s !== 'estado' && s !== 'state');
  SYN.tipo.unshift('estado', 'state');
  const ORDER = {
    base: ['gaiola', 'desc', 'rota', 'id'],
    sep: ['status', 'ciclo', 'local', 'detalhe', 'parada', 'ponto', 'gaiola', 'desc', 'rota', 'id'],
    aud: ['opFinal', 'opBipou', 'tipo', 'status', 'ciclo', 'data', 'gaiola', 'desc', 'rota', 'id'],
    adu: ['opBipou', 'tipo', 'placa', 'data', 'rota', 'id'],
    rotas: ['onda', 'idRota', 'rotaOrig', 'rota', 'doca', 'transp', 'placa', 'tipoVeic', 'tipoRota', 'tipoArea', 'cluster']
  };
  const REQ = { sep: ['id', 'rota'], aud: ['id', 'gaiola'], adu: ['id', 'rota'], rotas: ['rota', 'onda'] };
  const LABELS = {
    id: 'ID do pacote', gaiola: 'Gaiola', rota: 'Rota', desc: 'Descrição', ciclo: 'Ciclo (AM/PM)',
    status: 'Status da gaiola', opBipou: 'Operador que bipou', opFinal: 'Operador que finalizou',
    tipo: 'Estado (correto / a mais / faltante)', onda: 'Onda', parada: 'Parada', ponto: 'Ponto',
    local: 'Localização do pacote', detalhe: 'Detalhe do pacote', placa: 'Placa', data: 'Data da auditoria',
    idRota: 'ID da rota', rotaOrig: 'Rota original', doca: 'Doca', transp: 'Transportadora',
    tipoVeic: 'Tipo de veículo', tipoRota: 'Tipo de rota', tipoArea: 'Tipo de área', cluster: 'Cluster'
  };
  const KINDS = {
    sep: { nome: 'Planilha de separação', ajuda: 'IDs e suas gaiolas/rotas (referência para o confronto)' },
    aud: { nome: 'Planilha de auditoria', ajuda: 'Gaiola, estado do ID (correto / a mais / faltante) e rep auditoria' },
    adu: { nome: 'Planilha de aduana', ajuda: 'Rota, placa, estado do ID e rep auditoria (limpa os números após "|" da rota)' },
    rotas: { nome: 'Lista de rotas / ondas', ajuda: 'Onda, rota otimizada, doca, transportadora, placa e cluster (prioridade)' }
  };

  function autoMap(kind, headers) {
    const hn = headers.map(norm), map = {}, used = new Set(), order = ORDER[kind];
    for (const f of order) {
      const i = hn.findIndex((h, ix) => !used.has(ix) && SYN[f].includes(h));
      if (i >= 0) { map[f] = headers[i]; used.add(i); }
    }
    for (const f of order) {
      if (map[f]) continue;
      const syns = SYN[f].filter(s => s.length >= 4).sort((a, b) => b.length - a.length);
      let hit = -1;
      for (const s of syns) { hit = hn.findIndex((h, ix) => !used.has(ix) && h.includes(s)); if (hit >= 0) break; }
      if (hit >= 0) { map[f] = headers[hit]; used.add(hit); }
    }
    return map;
  }

  const ALL_SYN = [].concat(...Object.values(SYN));
  function parseSheetRows(aoa) {
    let best = 0, bestScore = -1;
    for (let i = 0; i < Math.min(aoa.length, 15); i++) {
      const sc = aoa[i].filter(c => ALL_SYN.includes(norm(c))).length;
      if (sc > bestScore) { bestScore = sc; best = i; }
    }
    const seen = {};
    const headers = (aoa[best] || []).map((h, i) => {
      let n = txt(h) || ('Coluna ' + (i + 1));
      if (seen[n]) { seen[n]++; n = n + ' (' + seen[n] + ')'; } else seen[n] = 1;
      return n;
    });
    const rows = [];
    for (let r = best + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row || row.every(c => txt(c) === '')) continue;
      const o = {};
      headers.forEach((h, i) => { o[h] = row[i] == null ? '' : row[i]; });
      rows.push(o);
    }
    return { headers, rows };
  }

  /* ---------- classificadores ---------- */
  function cyc(...vals) {
    for (const v of vals) {
      const n = norm(v);
      if (!n) continue;
      if (/\bpm\d*\b|tarde|noite|afternoon|evening|night/.test(n)) return 'PM';
      if (/\bam\d*\b|manha|madrugada|morning/.test(n)) return 'AM';
    }
    return '';
  }
  const isCancel = s => /cancel/.test(norm(s));
  const isAberta = s => /nao fech|n fech|abert|pendente|open|nao finaliz|andamento|pending|not clos|unclos|not final|not complet|incomplete|in progress/.test(norm(s));
  const isFechada = s => !isAberta(s) && /fechad|closed|finaliz|final|conclu|complet|done/.test(norm(s));
  const parseOnda = v => { const m = txt(v).match(/\d+/); return m ? parseInt(m[0], 10) : null; };

  const STATUS = {
    fechada: { label: 'Fechada 100%', cor: '#16a34a' },
    ajustada: { label: 'Ajustada', cor: '#0d9488' },
    a_mais: { label: 'Com a mais', cor: '#d97706' },
    faltante: { label: 'Com faltante', cor: '#dc2626' },
    cancelado: { label: 'Cancelada', cor: '#6b7280' },
    pendente: { label: 'Não auditada', cor: '#2563eb' },
    aberta: { label: 'Não fechada', cor: '#7c3aed' },
    em_auditoria: { label: 'Em auditoria', cor: '#4f46e5' }
  };
  const PENDENTES = ['pendente', 'aberta', 'em_auditoria'];

  /* =====================================================================
     NÚCLEO: confronto dos IDs
     ===================================================================== */
  function compute(raw, map) {
    raw = raw || {}; map = map || {};
    const col = (kind, row, f) => {
      const h = map[kind] && map[kind][f]; let v = h ? txt(row[h]) : '';
      if (f === 'rota' || f === 'gaiola') { if (kind === 'adu') v = v.split('|')[0]; v = cleanCode(v); } // aduana: remove "| 477222964"
      return v;
    };
    const pick = i => ({ parada: i.parada || '', ponto: i.ponto || '', local: i.local || '', detalhe: i.detalhe || '' });
    const has = (kind, f) => !!(map[kind] && map[kind][f]);

    // Base de dados = verdade do sistema
    const base = new Map();
    ((raw.base && raw.base.rows) || []).forEach(r => {
      const id = KEY(col('base', r, 'id')); if (!id) return;
      base.set(id, { gaiola: col('base', r, 'gaiola'), rota: col('base', r, 'rota'), desc: col('base', r, 'desc') });
    });

    const G = new Map();
    const get = name => {
      if (!G.has(name)) G.set(name, {
        gaiola: name, rotas: {}, descs: {}, ciclo: '', expected: new Set(), bipped: new Set(),
        exMissing: new Set(), exExtra: new Set(), opB: new Set(), opF: new Set(), opById: new Map(),
        sepRows: 0, audRows: 0, data: '', sepStatus: [], audFinal: false, cancel: false, aberta: false, fechadaSep: false
      });
      return G.get(name);
    };
    const inc = (o, k) => { if (k) o[k] = (o[k] || 0) + 1; };

    // Separação
    const sepInfo = new Map(); let sepHasIds = false;
    ((raw.sep && raw.sep.rows) || []).forEach(r => {
      const name = txt(col('sep', r, 'gaiola')) || txt(col('sep', r, 'rota')); if (!name) return;
      const g = get(name), id = KEY(col('sep', r, 'id')), rota = col('sep', r, 'rota'), desc = col('sep', r, 'desc'), st = col('sep', r, 'status');
      g.sepRows++; inc(g.rotas, rota); inc(g.descs, desc);
      g.ciclo = g.ciclo || cyc(col('sep', r, 'ciclo'), name);
      if (st) { if (isCancel(st)) g.cancel = true; else if (isAberta(st)) g.aberta = true; else if (isFechada(st)) g.fechadaSep = true; }
      if (id) { sepHasIds = true; g.expected.add(id); sepInfo.set(id, { gaiola: name, rota, desc, parada: col('sep', r, 'parada'), ponto: col('sep', r, 'ponto'), local: col('sep', r, 'local'), detalhe: col('sep', r, 'detalhe') }); }
    });

    // Auditoria
    let audLoaded = !!(raw.aud && raw.aud.rows && raw.aud.rows.length);
    // se a auditoria já traz o Estado de cada ID (Correto / A mais / Faltante), ele é a fonte da verdade
    const audHasTipo = has('aud', 'tipo') && ((raw.aud && raw.aud.rows) || []).some(r => col('aud', r, 'tipo'));
    const audHasFinalInfo = has('aud', 'opFinal') || has('aud', 'status');
    ((raw.aud && raw.aud.rows) || []).forEach(r => {
      const id = KEY(col('aud', r, 'id'));
      let name = txt(col('aud', r, 'gaiola'));
      if (!name && id && sepInfo.has(id)) name = sepInfo.get(id).gaiola;
      if (!name) return;
      const g = get(name), tipo = norm(col('aud', r, 'tipo')), ob = col('aud', r, 'opBipou'), of = col('aud', r, 'opFinal'), st = col('aud', r, 'status');
      g.audRows++; inc(g.rotas, col('aud', r, 'rota')); inc(g.descs, col('aud', r, 'desc'));
      g.ciclo = g.ciclo || cyc(col('aud', r, 'ciclo'), name);
      const dt = col('aud', r, 'data'); if (dt && (!g.data || parseBR(dt) >= parseBR(g.data))) g.data = dt;
      const isMiss = /falt|miss|short|lost/.test(tipo);
      if (ob) { (isMiss ? g.opF : g.opB).add(ob); if (id) g.opById.set(id, ob); }
      if (of) { g.opF.add(of); g.audFinal = true; }
      if (st) { if (isCancel(st)) g.cancel = true; else if (isFechada(st)) g.audFinal = true; }
      if (id) {
        if (/falt|miss|short|lost/.test(tipo)) g.exMissing.add(id);
        else {
          g.bipped.add(id);
          if (/mais|sobr|excess|extra|surplus|overage|over/.test(tipo)) g.exExtra.add(id);
          if (ob) g.opById.set(id, ob);
        }
      }
    });

    const top = o => Object.entries(o).sort((a, b) => b[1] - a[1])[0];
    const gaiolas = [];
    G.forEach(g => {
      const audited = g.audRows > 0;
      const finalizada = audited && (audHasFinalInfo ? g.audFinal : true);
      const inSep = g.sepRows > 0;
      const items = []; let extraN = 0, missingN = 0, regN = 0;

      if (audited) {
        const extraSet = new Set(g.exExtra);
        if (!audHasTipo && sepHasIds && inSep) g.bipped.forEach(id => { if (!g.expected.has(id)) extraSet.add(id); });
        const missSet = new Set(g.exMissing);
        if (!audHasTipo && finalizada && sepHasIds) g.expected.forEach(id => { if (!g.bipped.has(id)) missSet.add(id); });

        extraSet.forEach(id => {
          const info = base.get(id) || sepInfo.get(id) || {};
          const correta = info.gaiola || '';
          let tipo;
          if (!correta) tipo = 'nao_localizado'; else if (KEY(correta) === KEY(g.gaiola)) tipo = 'regularizado'; else tipo = 'a_mais';
          if (tipo === 'regularizado') regN++; else extraN++;
          items.push({ id, tipo, gaiolaAtual: g.gaiola, gaiolaCorreta: correta, rotaCorreta: info.rota || '', desc: info.desc || '', ...pick(info), operador: g.opById.get(id) || [...g.opB].join(', ') });
        });
        missSet.forEach(id => {
          const info = base.get(id) || sepInfo.get(id) || {};
          const onde = info.gaiola || '';
          let tipo;
          if (onde && KEY(onde) !== KEY(g.gaiola)) tipo = 'regularizado_falt'; else tipo = 'faltante';
          if (tipo === 'faltante') missingN++; else regN++;
          items.push({ id, tipo, gaiolaAtual: g.gaiola, gaiolaCorreta: onde, rotaCorreta: info.rota || '', desc: info.desc || '', ...pick(info), operador: g.opById.get(id) || [...g.opF].join(', ') });
        });
      }

      let status;
      if (g.cancel) status = 'cancelado';
      else if (!audited) status = g.aberta ? 'aberta' : 'pendente';
      else if (!finalizada) status = 'em_auditoria';
      else if (missingN > 0) status = 'faltante';
      else if (extraN > 0) status = 'a_mais';
      else if (regN > 0) status = 'ajustada';
      else status = 'fechada';

      const rotaT = top(g.rotas), descT = top(g.descs);
      const rotasCorretas = [...new Set(items.filter(i => i.tipo === 'a_mais' && i.rotaCorreta).map(i => i.rotaCorreta))];
      gaiolas.push({
        gaiola: g.gaiola, ciclo: g.ciclo || cyc(g.gaiola) || 'N/D', rota: rotaT ? rotaT[0] : g.gaiola, desc: descT ? descT[0] : '',
        status, audited, finalizada, inSep, data: g.data, esperados: g.expected.size, bipados: g.bipped.size,
        extraN, missingN, regN, opBipou: [...g.opB], opFinal: [...g.opF], rotasCorretas, items
      });
    });
    gaiolas.sort((a, b) => a.gaiola.localeCompare(b.gaiola, 'pt-BR', { numeric: true }));

    // Rotas / ondas / prioridade
    const ondaMap = new Map(), infoMap = new Map();
    ((raw.rotas && raw.rotas.rows) || []).forEach(r => {
      const rt = KEY(col('rotas', r, 'rota')); if (!rt) return;
      ondaMap.set(rt, parseOnda(col('rotas', r, 'onda')));
      infoMap.set(rt, { nome: col('rotas', r, 'rota'), idRota: col('rotas', r, 'idRota'), rotaOrig: col('rotas', r, 'rotaOrig'), doca: col('rotas', r, 'doca'), transp: col('rotas', r, 'transp'), placa: col('rotas', r, 'placa'), tipoVeic: col('rotas', r, 'tipoVeic'), tipoRota: col('rotas', r, 'tipoRota'), tipoArea: col('rotas', r, 'tipoArea'), cluster: col('rotas', r, 'cluster') });
    });
    const R = new Map();
    const rget = (k, nome) => { if (!R.has(k)) R.set(k, { rota: nome, onda: ondaMap.has(k) ? ondaMap.get(k) : null, gaiolas: [], info: infoMap.get(k) || {} }); return R.get(k); };
    ondaMap.forEach((o, k) => rget(k, (infoMap.get(k) && infoMap.get(k).nome) || k));
    gaiolas.forEach(g => { if (g.rota) rget(KEY(g.rota), g.rota).gaiolas.push(g); });
    const rotas = [];
    R.forEach(r => {
      const vivas = r.gaiolas.filter(g => g.status !== 'cancelado');
      const ok = vivas.filter(g => g.status === 'fechada' || g.status === 'ajustada').length;
      r.total = vivas.length; r.ok = ok; r.pend = vivas.length - ok;
      r.faltante = vivas.filter(g => g.status === 'faltante').length;
      r.a_mais = vivas.filter(g => g.status === 'a_mais').length;
      r.nao_auditadas = vivas.filter(g => PENDENTES.includes(g.status)).length;
      r.ring = r.gaiolas.length === 0 ? 'gray' : (r.pend > 0 ? 'red' : 'green');
      rotas.push(r);
    });
    const rank = { red: 0, green: 1, gray: 2 };
    rotas.sort((a, b) => rank[a.ring] - rank[b.ring]
      || (a.onda == null ? 99 : a.onda) - (b.onda == null ? 99 : b.onda)
      || b.pend - a.pend || String(a.rota).localeCompare(String(b.rota), 'pt-BR', { numeric: true }));
    let p = 0; rotas.forEach(r => { r.prioridade = r.ring === 'red' ? ++p : null; });

    // KPIs
    const cnt = {}; Object.keys(STATUS).forEach(k => cnt[k] = 0);
    gaiolas.forEach(g => cnt[g.status]++);
    const vivas = gaiolas.length - cnt.cancelado;
    const okN = cnt.fechada + cnt.ajustada;
    const kpi = { total: gaiolas.length, vivas, ok: okN, pct: vivas ? okN / vivas * 100 : 0, cnt, meta: 100 };
    const ciclos = {};
    ['AM', 'PM', 'N/D'].forEach(c => { ciclos[c] = {}; Object.keys(STATUS).forEach(k => ciclos[c][k] = 0); ciclos[c].total = 0; });
    gaiolas.forEach(g => { const c = ciclos[g.ciclo] ? g.ciclo : 'N/D'; ciclos[c][g.status]++; ciclos[c].total++; });
    Object.values(ciclos).forEach(c => {
      const v = c.total - c.cancelado; c.pct = v ? (c.fechada + c.ajustada) / v * 100 : 0;
    });
    const ondas = {};
    rotas.forEach(r => { const k = r.onda == null ? 'Sem onda' : 'Onda ' + r.onda; (ondas[k] = ondas[k] || { red: 0, green: 0, gray: 0 })[r.ring]++; });

    // Aduana: auditoria por rota/placa (o "| número" da rota já é removido em col())
    const A = new Map();
    ((raw.adu && raw.adu.rows) || []).forEach(r => {
      const rota = col('adu', r, 'rota'); if (!rota) return;
      const k = KEY(rota);
      if (!A.has(k)) A.set(k, { rota, placas: new Set(), ops: new Set(), total: 0, correto: 0, a_mais: 0, faltante: 0, outros: 0, data: '', items: [] });
      const a = A.get(k), t = norm(col('adu', r, 'tipo')), id = col('adu', r, 'id'), pl = col('adu', r, 'placa'), op = col('adu', r, 'opBipou'), dt = col('adu', r, 'data');
      a.total++; if (pl) a.placas.add(pl); if (op) a.ops.add(op);
      if (dt && (!a.data || parseBR(dt) >= parseBR(a.data))) a.data = dt;
      let tipo = 'outros';
      if (/falt|miss|short|lost/.test(t)) tipo = 'faltante';
      else if (/mais|sobr|excess|extra|surplus|over/.test(t)) tipo = 'a_mais';
      else if (/^(corret|ok$|correct)/.test(t)) tipo = 'correto';
      a[tipo]++; a.items.push({ id, tipo, estado: col('adu', r, 'tipo'), placa: pl, operador: op, rota });
    });
    const aduana = [...A.values()].map(a => Object.assign(a, { placas: [...a.placas], ops: [...a.ops], diverg: a.a_mais + a.faltante }))
      .sort((x, y) => y.diverg - x.diverg || String(x.rota).localeCompare(String(y.rota), 'pt-BR', { numeric: true }));
    rotas.forEach(r => { const a = A.get(KEY(r.rota)); r.aduDiv = a ? a.diverg : 0; });
    gaiolas.forEach(g => { const r = R.get(KEY(g.rota)); g.onda = r ? r.onda : null; g.doca = (r && r.info && r.info.doca) || ''; });

    return { gaiolas, rotas, kpi, ciclos, ondas, aduana, sepHasIds, audLoaded, baseSize: base.size };
  }

  /* =====================================================================
     DADOS DE EXEMPLO (para testar o painel sem planilhas)
     ===================================================================== */
  function buildDemo() {
    let seed = 11; const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    const bairros = ['Tijuca', 'Méier', 'Copacabana', 'Madureira', 'Bangu', 'Campo Grande', 'Barra', 'Penha', 'Ilha', 'Centro', 'Realengo', 'Jacarepaguá'];
    const ops = ['Carlos Silva', 'Ana Souza', 'João Lima', 'Marcos Pinto', 'Fernanda Reis'];
    const rotas = []; for (let i = 1; i <= 12; i++) rotas.push({ Rota: 'R-' + (100 + i), Onda: 'Onda ' + (((i - 1) % 5) + 1) });
    const base = [], sep = [], aud = [], plano = []; let pid = 7000000000;
    for (let i = 1; i <= 48; i++) {
      const ciclo = i % 2 ? 'AM' : 'PM', rota = rotas[i % 12].Rota, nome = 'G-' + ciclo + '-' + String(i).padStart(3, '0');
      const n = 6 + Math.floor(rnd() * 7), ids = [];
      for (let k = 0; k < n; k++) { const id = String(pid++); ids.push(id); base.push({ 'ID Pacote': id, 'Gaiola': nome, 'Rota': rota, 'Descrição': bairros[(i + k) % 12] }); }
      const r = rnd(); let cen = r < 0.12 ? 'pend' : r < 0.2 ? 'canc' : r < 0.5 ? 'ok' : r < 0.65 ? 'falt' : r < 0.78 ? 'mais' : r < 0.9 ? 'aj' : 'emaud';
      plano.push({ nome, ciclo, rota, ids, cen });
    }
    plano.forEach((p, i) => {
      const sepStatus = p.cen === 'canc' ? 'Cancelada' : (p.cen === 'pend' && i % 2 ? 'Não fechada' : 'Fechada');
      p.ids.forEach(id => sep.push({ 'ID Pacote': id, 'Gaiola': p.nome, 'Rota': p.rota, 'Ciclo': p.ciclo, 'Status': sepStatus }));
    });
    plano.forEach((p, i) => {
      if (p.cen === 'pend' || p.cen === 'canc') return;
      const op = ops[i % 5], opF = p.cen === 'emaud' ? '' : ops[(i + 2) % 5];
      let bip = p.ids.slice();
      if (p.cen === 'falt') bip = bip.slice(0, bip.length - 1 - (i % 2));
      if (p.cen === 'aj') { const alvo = bip.pop(); const b = base.find(x => x['ID Pacote'] === alvo); b['Gaiola'] = plano[(i + 1) % 48].nome; b['Rota'] = plano[(i + 1) % 48].rota; }
      if (p.cen === 'mais') { const outro = plano[(i + 5) % 48].ids[0]; bip.push(outro); }
      if (p.cen === 'ok' && i % 7 === 0) { /* gaiola OK */ }
      bip.forEach(id => aud.push({ 'ID Pacote': id, 'Gaiola': p.nome, 'Rota': p.rota, 'Operador Bipou': op, 'Operador Finalizou': opF, 'Tipo': '' }));
    });
    const mk = (kind, rows, nome) => { const headers = Object.keys(rows[0]); return { headers, rows, fileName: nome }; };
    const raw = { sep: mk('sep', sep, 'Separação (exemplo)'), aud: mk('aud', aud, 'Auditoria (exemplo)'), rotas: mk('rotas', rotas, 'Lista de rotas (exemplo)') };
    const map = {}; Object.keys(raw).forEach(k => map[k] = autoMap(k, raw[k].headers));
    return { raw, map };
  }

  const Core = { compute, autoMap, buildDemo, STATUS, PENDENTES, KINDS, LABELS, ORDER, parseSheetRows, norm, esc };
  root.PainelCore = Core;
  if (typeof module !== 'undefined' && module.exports) module.exports = Core;
  if (typeof document === 'undefined') return;

  /* =====================================================================
     INTERFACE (somente no navegador)
     ===================================================================== */
  const S = { raw: {}, map: {}, res: null, tab: 'dashboard', ciclo: 'todos', filtro: 'todas', busca: '', divTipo: 'todos', aduFiltro: 'todas', updated: null };
  const $ = s => document.querySelector(s);
  const $$ = s => [...document.querySelectorAll(s)];
  let charts = {};
  const fmt = n => Number(n).toLocaleString('pt-BR');
  const pct = n => n.toFixed(1).replace('.', ',') + '%';
  const badge = st => `<span class="badge" style="--c:${STATUS[st].cor}">${STATUS[st].label}</span>`;
  const MAXROWS = 400;

  /* ---- persistência (IndexedDB) ---- */
  const DB = {
    open() { return new Promise((res, rej) => { const r = indexedDB.open('painel-operacional-rj15', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
    async get(k) { const d = await this.open(); return new Promise((res, rej) => { const q = d.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
    async set(k, v) { const d = await this.open(); return new Promise((res, rej) => { const t = d.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); }
  };
  const salvar = async () => { try { await DB.set('estado', { raw: S.raw, map: S.map, updated: S.updated }); } catch (e) { /* sem persistência */ } };

  /* ---- leitura de arquivo ---- */
  function fixCell(h, v) {
    if (typeof v === 'number') {
      if (/data|date|hora/.test(norm(h)) && v > 30000 && v < 80000) {
        const d = new Date(Math.round((v - 25569) * 86400) * 1000), p = n => String(n).padStart(2, '0');
        return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear() + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes());
      }
      return String(v); // IDs grandes sem notação científica
    }
    return v;
  }
  function lerArquivo(kind, file) {
    return new Promise((resolve, reject) => {
      if (!root.XLSX) { reject(new Error('Biblioteca de Excel não carregou (verifique a internet).')); return; }
      const fr = new FileReader();
      fr.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array' });
          let best = null;
          wb.SheetNames.forEach(n => {
            const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: true });
            if (!best || aoa.length > best.aoa.length) best = { aoa, n };
          });
          if (!best || !best.aoa.length) throw new Error('Planilha vazia');
          const { headers, rows } = parseSheetRows(best.aoa);
          rows.forEach(r => headers.forEach(h => { r[h] = fixCell(h, r[h]); }));
          resolve({ headers, rows, fileName: file.name, aba: best.n });
        } catch (err) { reject(err); }
      };
      fr.onerror = () => reject(fr.error);
      fr.readAsArrayBuffer(file);
    });
  }
  async function carregar(kind, file) {
    if (!file) return;
    toast('Lendo ' + file.name + '…');
    try {
      const d = await lerArquivo(kind, file);
      S.raw[kind] = d; S.map[kind] = autoMap(kind, d.headers);
      const faltam = (REQ[kind] || []).filter(f => !S.map[kind][f]);
      recalc();
      toast(KINDS[kind].nome + ': ' + fmt(d.rows.length) + ' linhas' + (faltam.length ? ' · confira o mapeamento de colunas' : ''), faltam.length ? 'warn' : 'ok');
    } catch (e) { toast('Erro ao ler: ' + e.message, 'err'); }
  }

  function recalc() {
    S.res = compute(S.raw, S.map); S.updated = new Date(); render(); renderDrawer(); salvar();
  }

  /* ---- toast ---- */
  let tt;
  function toast(msg, tipo) {
    const t = $('#toast'); t.textContent = msg; t.className = 'toast show ' + (tipo || '');
    clearTimeout(tt); tt = setTimeout(() => t.classList.remove('show'), 3800);
  }

  /* ---- drawer (menu hambúrguer) ---- */
  function renderDrawer() {
    $('#uploads').innerHTML = Object.keys(KINDS).map(k => {
      const d = S.raw[k], ok = !!d;
      return `<div class="up ${ok ? 'ok' : ''}">
        <div class="up-h"><b>${KINDS[k].nome}</b><span class="dot"></span></div>
        <div class="up-s">${KINDS[k].ajuda}</div>
        <div class="up-f">${ok ? `${esc(d.fileName)} · ${fmt(d.rows.length)} linhas` : 'Nenhum arquivo carregado'}</div>
        <div class="up-b">
          <label class="btn sm">Carregar Excel<input type="file" accept=".xlsx,.xls,.xlsm,.csv" data-up="${k}" hidden></label>
          ${ok ? `<button class="btn sm ghost" data-map="${k}">Mapear colunas</button>` : ''}
        </div></div>`;
    }).join('');
    $$('[data-up]').forEach(i => i.onchange = e => { carregar(i.dataset.up, e.target.files[0]); e.target.value = ''; });
    $$('[data-map]').forEach(b => b.onclick = () => abrirMapa(b.dataset.map));
  }
  const abrirDrawer = v => { $('#drawer').classList.toggle('open', v); $('#scrim').classList.toggle('show', v); };

  /* ---- mapeamento manual de colunas ---- */
  function abrirMapa(kind) {
    const d = S.raw[kind];
    const campos = ORDER[kind];
    $('#modal-body').innerHTML = `<h3>Mapear colunas · ${KINDS[kind].nome}</h3>
      <p class="mut">O painel reconhece os cabeçalhos automaticamente. Ajuste aqui se alguma coluna estiver errada.</p>
      <div class="form">${campos.map(f => `<label>${LABELS[f]}<select data-f="${f}"><option value="">— não usar —</option>${d.headers.map(h => `<option ${S.map[kind][f] === h ? 'selected' : ''}>${esc(h)}</option>`).join('')}</select></label>`).join('')}</div>
      <div class="row end"><button class="btn ghost" id="m-x">Cancelar</button><button class="btn" id="m-ok">Aplicar</button></div>`;
    $('#modal').classList.add('show');
    $('#m-x').onclick = fecharModal;
    $('#m-ok').onclick = () => { const m = {}; $$('#modal-body select').forEach(s => { if (s.value) m[s.dataset.f] = s.value; }); S.map[kind] = m; fecharModal(); recalc(); };
  }
  const fecharModal = () => $('#modal').classList.remove('show');

  /* ---- filtros ---- */
  function filtrar() {
    const q = norm(S.busca);
    const base = S.res.gaiolas.filter(g => {
      if (S.ciclo !== 'todos' && g.ciclo !== S.ciclo) return false;
      if (q && !(norm(g.gaiola).includes(q) || norm(g.rota).includes(q) || g.items.some(i => norm(i.id).includes(q)))) return false;
      return true;
    });
    const pass = (g, f) => f === 'todas' ? true : f === 'fechada' ? g.status === 'fechada' : f === 'a_mais' ? g.extraN > 0 : f === 'faltante' ? g.missingN > 0 :
      f === 'cancelado' ? g.status === 'cancelado' : f === 'ajustada' ? g.status === 'ajustada' : f === 'pendente' ? PENDENTES.includes(g.status) : true;
    return { base, pass, lista: base.filter(g => pass(g, S.filtro)) };
  }

  /* ---- views ---- */
  const vazio = () => `<div class="empty card"><div class="empty-ico">▦</div><h2>Nenhum dado carregado</h2>
    <p>Abra o menu <b>☰</b> e carregue as planilhas (Separação, Auditoria e Lista de rotas) para ver o resultado do confronto.</p>
    <div class="row center"><button class="btn" data-act="menu">Abrir menu</button><button class="btn ghost" data-act="demo">Ver com dados de exemplo</button></div></div>`;

  function kpiCard(t, v, cor, sub) { return `<div class="kpi card" style="--c:${cor}"><div class="kpi-t">${t}</div><div class="kpi-v">${v}</div>${sub ? `<div class="kpi-s">${sub}</div>` : ''}</div>`; }

  function vDashboard() {
    const { kpi, ciclos } = S.res, c = kpi.cnt, atingiu = kpi.pct >= kpi.meta - 1e-9;
    const cor = atingiu ? '#16a34a' : kpi.pct >= 80 ? '#d97706' : '#dc2626';
    const linhas = ['AM', 'PM', 'N/D'].filter(k => ciclos[k].total).map(k => `<tr><td><b>${k}</b></td><td>${ciclos[k].total}</td><td>${ciclos[k].fechada}</td><td>${ciclos[k].ajustada}</td><td>${ciclos[k].a_mais}</td><td>${ciclos[k].faltante}</td><td>${ciclos[k].pendente + ciclos[k].aberta + ciclos[k].em_auditoria}</td><td>${ciclos[k].cancelado}</td><td><b>${pct(ciclos[k].pct)}</b></td></tr>`).join('');
    const prio = S.res.rotas.filter(r => r.ring === 'red').slice(0, 5);
    return `<div class="grid g-hero">
      <div class="card meta" style="--c:${cor}"><div class="meta-h"><span>Gaiolas fechadas · meta ${kpi.meta}%</span><span class="pill" style="--c:${cor}">${atingiu ? 'Meta atingida' : 'Abaixo da meta'}</span></div>
        <div class="meta-v">${pct(kpi.pct)}</div><div class="bar"><i style="width:${Math.min(100, kpi.pct)}%"></i><b style="left:100%"></b></div>
        <div class="meta-s">${fmt(kpi.ok)} de ${fmt(kpi.vivas)} gaiolas fechadas sem erro (100% + ajustadas)${c.cancelado ? ' · ' + c.cancelado + ' canceladas fora da conta' : ''}</div></div>
      <div class="grid g-kpi">
        ${kpiCard('Total de gaiolas', fmt(kpi.total), '#14213d')}
        ${kpiCard('Fechadas 100%', fmt(c.fechada), STATUS.fechada.cor)}
        ${kpiCard('Ajustadas', fmt(c.ajustada), STATUS.ajustada.cor)}
        ${kpiCard('Com a mais', fmt(c.a_mais), STATUS.a_mais.cor)}
        ${kpiCard('Com faltante', fmt(c.faltante), STATUS.faltante.cor)}
        ${kpiCard('Falta auditar', fmt(c.pendente + c.aberta + c.em_auditoria), STATUS.pendente.cor)}
      </div></div>
      <div class="grid g-2">
        <div class="card"><h3>Situação das gaiolas</h3><div class="chart"><canvas id="ch-status"></canvas></div></div>
        <div class="card"><h3>Resultado por ciclo</h3><div class="chart"><canvas id="ch-ciclo"></canvas></div></div>
      </div>
      <div class="grid g-2">
        <div class="card"><h3>Resultado por ciclo (AM / PM)</h3><div class="tw"><table><thead><tr><th>Ciclo</th><th>Gaiolas</th><th>Fechadas</th><th>Ajust.</th><th>A mais</th><th>Faltante</th><th>Pend.</th><th>Canc.</th><th>% fechadas</th></tr></thead><tbody>${linhas || '<tr><td colspan="9" class="mut">Sem dados</td></tr>'}</tbody></table></div></div>
        <div class="card"><div class="row between"><h3>Prioridades agora</h3><button class="btn sm ghost" data-tab="prioridade">Ver todas</button></div>
          ${prio.length ? prio.map(r => `<div class="mini red"><b>#${r.prioridade} · ${esc(r.rota)}</b><span>${r.onda ? 'Onda ' + r.onda : 'Sem onda'}</span><span>${r.pend} pendente(s)</span></div>`).join('') : '<p class="mut">Nenhuma rota pendente. Carregue a lista de rotas para ver as prioridades por onda.</p>'}
        </div></div>`;
  }

  function vGaiolas() {
    const { base, pass, lista } = filtrar();
    const chips = [['todas', 'Todas'], ['fechada', 'Fechadas 100%'], ['a_mais', 'Com a mais'], ['faltante', 'Com faltante'], ['ajustada', 'Ajustadas'], ['pendente', 'Falta auditar'], ['cancelado', 'Canceladas']];
    const rows = lista.slice(0, MAXROWS).map(g => `<tr class="click" data-g="${esc(g.gaiola)}"><td><b>${esc(g.gaiola)}</b></td><td>${g.ciclo}</td><td>${esc(g.rota)}</td><td>${g.onda ? 'Onda ' + g.onda : '–'}</td><td>${esc(g.doca) || '–'}</td><td>${g.bipados}/${g.esperados}</td><td>${g.extraN || '–'}</td><td>${g.missingN || '–'}</td><td>${badge(g.status)}</td><td>${esc(g.opBipou.join(', ')) || '–'}</td><td>${esc(g.opFinal.join(', ')) || '–'}</td></tr>`).join('');
    return `<div class="card"><div class="filters">
      <div class="seg">${['todos', 'AM', 'PM'].map(c => `<button class="${S.ciclo === c ? 'on' : ''}" data-ciclo="${c}">${c === 'todos' ? 'AM + PM' : c}</button>`).join('')}</div>
      <input id="busca" class="search" placeholder="Buscar gaiola, rota ou ID do pacote…" value="${esc(S.busca)}"></div>
      <div class="chips">${chips.map(([k, l]) => `<button class="chip ${S.filtro === k ? 'on' : ''}" data-filtro="${k}">${l} <b>${base.filter(g => pass(g, k)).length}</b></button>`).join('')}</div>
      <div class="tw"><table><thead><tr><th>Gaiola</th><th>Ciclo</th><th>Rota</th><th>Onda</th><th>Doca</th><th>Bipados/Esperados</th><th>A mais</th><th>Faltante</th><th>Situação</th><th>Operador (bipou)</th><th>Operador (finalizou)</th></tr></thead><tbody>${rows || '<tr><td colspan="11" class="mut">Nenhuma gaiola neste filtro.</td></tr>'}</tbody></table></div>
      ${lista.length > MAXROWS ? `<p class="mut">Mostrando ${MAXROWS} de ${fmt(lista.length)}. Use a busca para refinar.</p>` : ''}</div>`;
  }

  function vAuditoria() {
    const gs = S.res.gaiolas;
    const falta = gs.filter(g => PENDENTES.includes(g.status));
    const divs = [];
    gs.forEach(g => g.items.forEach(i => { if (i.tipo !== 'regularizado' && i.tipo !== 'regularizado_falt') divs.push(Object.assign({ ciclo: g.ciclo, rota: g.rota }, i)); }));
    const f = S.divTipo, lista = divs.filter(d => f === 'todos' || (f === 'a_mais' && (d.tipo === 'a_mais' || d.tipo === 'nao_localizado')) || (f === 'faltante' && d.tipo === 'faltante'));
    const tipoTxt = t => t === 'a_mais' ? '<span class="badge" style="--c:#d97706">A mais</span>' : t === 'nao_localizado' ? '<span class="badge" style="--c:#d97706">A mais · ID não localizado</span>' : '<span class="badge" style="--c:#dc2626">Faltante</span>';
    return `<div class="card"><div class="row between"><h3>Gaiolas que faltam auditar <span class="count">${falta.length}</span></h3></div>
      <div class="tw"><table><thead><tr><th>Gaiola</th><th>Ciclo</th><th>Rota</th><th>Descrição</th><th>Pacotes</th><th>Situação</th><th>Operador (bipou)</th></tr></thead><tbody>
      ${falta.slice(0, MAXROWS).map(g => `<tr class="click" data-g="${esc(g.gaiola)}"><td><b>${esc(g.gaiola)}</b></td><td>${g.ciclo}</td><td>${esc(g.rota)}</td><td>${esc(g.desc) || '–'}</td><td>${g.esperados}</td><td>${badge(g.status)}</td><td>${esc(g.opBipou.join(', ')) || '–'}</td></tr>`).join('') || '<tr><td colspan="7" class="mut">Todas as gaiolas já foram auditadas.</td></tr>'}</tbody></table></div></div>
      <div class="card"><div class="row between"><h3>Divergências por ID <span class="count">${lista.length}</span></h3>
        <div class="seg">${[['todos', 'Todos'], ['a_mais', 'A mais'], ['faltante', 'Faltante']].map(([k, l]) => `<button class="${f === k ? 'on' : ''}" data-div="${k}">${l}</button>`).join('')}</div></div>
      <p class="mut">Cada ID foi confrontado com a Separação: a gaiola/rota corretas vêm de onde o ID foi separado.</p>
      <div class="tw"><table><thead><tr><th>ID</th><th>Tipo</th><th>Gaiola atual</th><th>Rota</th><th>Parada</th><th>Ponto</th><th>Localização do pacote</th><th>Detalhe</th><th>Gaiola correta</th><th>Rota correta p/ atrelar</th><th>Operador</th></tr></thead><tbody>
      ${lista.slice(0, MAXROWS).map(d => `<tr><td><b>${esc(d.id)}</b></td><td>${tipoTxt(d.tipo)}</td><td>${esc(d.gaiolaAtual)}</td><td>${esc(d.rota)}</td><td>${esc(d.parada) || '–'}</td><td>${esc(d.ponto) || '–'}</td><td>${esc(d.local) || '–'}</td><td>${esc(d.detalhe) || '–'}</td><td>${esc(d.gaiolaCorreta) || 'não consta na separação'}</td><td>${esc(d.rotaCorreta) || '–'}</td><td>${esc(d.operador) || '–'}</td></tr>`).join('') || '<tr><td colspan="11" class="mut">Nenhuma divergência.</td></tr>'}</tbody></table></div>
      ${lista.length > MAXROWS ? `<p class="mut">Mostrando ${MAXROWS} de ${fmt(lista.length)}. Gere o PDF ou use a aba Gaiolas para detalhar.</p>` : ''}</div>`;
  }

  function vConclusao() {
    const gs = S.res.gaiolas, ok = gs.filter(g => g.status === 'fechada'), aj = gs.filter(g => g.status === 'ajustada'), { kpi } = S.res;
    const tb = (arr, aj2) => arr.slice(0, MAXROWS).map(g => `<tr class="click" data-g="${esc(g.gaiola)}"><td><b>${esc(g.gaiola)}</b></td><td>${g.ciclo}</td><td>${esc(g.rota)}</td><td>${g.bipados}/${g.esperados}</td>${aj2 ? `<td>${g.regN} ID(s) regularizado(s) pelo sistema</td>` : ''}<td>${esc(g.opFinal.join(', ')) || '–'}</td></tr>`).join('');
    return `<div class="grid g-kpi3">${kpiCard('Conclusão da auditoria', pct(kpi.pct), kpi.pct >= 100 ? '#16a34a' : '#d97706', 'meta 100%')}${kpiCard('Fechadas 100%', fmt(ok.length), STATUS.fechada.cor)}${kpiCard('Ajustadas e fechadas sem erro', fmt(aj.length), STATUS.ajustada.cor)}</div>
      <div class="card"><h3>Gaiolas fechadas 100% <span class="count">${ok.length}</span></h3><div class="tw"><table><thead><tr><th>Gaiola</th><th>Ciclo</th><th>Rota</th><th>Bipados/Esperados</th><th>Operador (finalizou)</th></tr></thead><tbody>${tb(ok, false) || '<tr><td colspan="5" class="mut">Nenhuma ainda.</td></tr>'}</tbody></table></div></div>
      <div class="card"><h3>Gaiolas ajustadas (fechadas sem erro) <span class="count">${aj.length}</span></h3><p class="mut">Tiveram divergência na bipagem, mas o confronto com a Separação confirmou que o ID está correto.</p><div class="tw"><table><thead><tr><th>Gaiola</th><th>Ciclo</th><th>Rota</th><th>Bipados/Esperados</th><th>Ajuste</th><th>Operador (finalizou)</th></tr></thead><tbody>${tb(aj, true) || '<tr><td colspan="6" class="mut">Nenhuma ainda.</td></tr>'}</tbody></table></div></div>`;
  }

  function vPrioridade() {
    const rs = S.res.rotas, ondas = S.res.ondas;
    const info = S.raw.rotas ? '' : '<div class="notice">Carregue a <b>Lista de rotas</b> (rota + onda) para classificar as rotas por onda.</div>';
    const keys = Object.keys(ondas).sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));
    return `<div class="card"><div class="row between wrap"><h3>Prioridade por onda</h3><div class="row">
        <label class="btn sm">Carregar lista de rotas<input type="file" accept=".xlsx,.xls,.xlsm,.csv" data-up="rotas" hidden></label>
        <label class="btn sm ghost">Carregar separação<input type="file" accept=".xlsx,.xls,.xlsm,.csv" data-up="sep" hidden></label></div></div>${info}
      <div class="chips">${keys.map(k => `<span class="chip static">${k} <b class="r">${ondas[k].red}</b><b class="g">${ondas[k].green}</b></span>`).join('')}</div>
      <p class="mut"><span class="dotc r"></span> com pendência (prioridade, aparece primeiro) &nbsp; <span class="dotc g"></span> tudo certo &nbsp; <span class="dotc n"></span> sem gaiola na separação</p></div>
      <div class="grid g-routes">${rs.map(r => `<div class="route ${r.ring}"><div class="route-h"><b>${esc(r.rota)}</b><span class="onda">${r.onda ? 'Onda ' + r.onda : 'Sem onda'}</span></div>
        ${r.ring === 'red' ? `<div class="prio">Prioridade #${r.prioridade}</div>` : r.ring === 'green' ? '<div class="okm">Tudo certo</div>' : '<div class="nm">Sem gaiolas na separação</div>'}
        <div class="route-s">${r.ok}/${r.total} gaiolas fechadas</div>
        <div class="route-s">${r.info && r.info.doca ? 'Doca ' + esc(r.info.doca) : ''}${r.info && r.info.transp ? ' · ' + esc(r.info.transp) : ''}</div>
        <div class="route-s">${r.info ? esc(r.info.placa || '') : ''}${r.info && r.info.cluster ? ' · ' + esc(r.info.cluster) : ''}</div>
        <div class="tags">${r.faltante ? `<span class="tag" style="--c:#dc2626">${r.faltante} faltante</span>` : ''}${r.a_mais ? `<span class="tag" style="--c:#d97706">${r.a_mais} a mais</span>` : ''}${r.nao_auditadas ? `<span class="tag" style="--c:#2563eb">${r.nao_auditadas} s/ auditar</span>` : ''}${r.aduDiv ? `<span class="tag" style="--c:#0d9488">${r.aduDiv} na aduana</span>` : ''}</div></div>`).join('')}</div>`;
  }

  function vAduana() {
    const A = S.res.aduana, soma = k => A.reduce((s, a) => s + a[k], 0);
    const lista = A.filter(a => S.aduFiltro === 'todas' || a.diverg > 0);
    const divs = []; A.forEach(a => a.items.forEach(i => { if (i.tipo === 'a_mais' || i.tipo === 'faltante') divs.push(i); }));
    const tTxt = t => t === 'a_mais' ? '<span class="badge" style="--c:#d97706">A mais</span>' : '<span class="badge" style="--c:#dc2626">Faltante</span>';
    return `<div class="grid g-kpi3">${kpiCard('Pacotes na aduana', fmt(soma('total')), '#14213d')}${kpiCard('Corretos', fmt(soma('correto')), STATUS.fechada.cor)}${kpiCard('A mais / Faltante', fmt(soma('a_mais')) + ' / ' + fmt(soma('faltante')), STATUS.a_mais.cor)}</div>
      <div class="card"><div class="row between"><h3>Aduana por rota <span class="count">${lista.length}</span></h3>
        <div class="seg">${[['todas', 'Todas'], ['div', 'Com divergência']].map(([k, l]) => `<button class="${S.aduFiltro === k ? 'on' : ''}" data-adu="${k}">${l}</button>`).join('')}</div></div>
      <div class="tw"><table><thead><tr><th>Rota</th><th>Placa</th><th>Pacotes</th><th>Corretos</th><th>A mais</th><th>Faltante</th><th>Rep auditoria</th><th>Última auditoria</th></tr></thead><tbody>
      ${lista.slice(0, MAXROWS).map(a => `<tr><td><b>${esc(a.rota)}</b></td><td>${esc(a.placas.join(', ')) || '–'}</td><td>${a.total}</td><td>${a.correto}</td><td>${a.a_mais || '–'}</td><td>${a.faltante || '–'}</td><td>${esc(a.ops.join(', ')) || '–'}</td><td>${esc(a.data) || '–'}</td></tr>`).join('') || '<tr><td colspan="8" class="mut">Nenhuma rota neste filtro.</td></tr>'}</tbody></table></div></div>
      <div class="card"><h3>Divergências por ID na aduana <span class="count">${divs.length}</span></h3>
      <div class="tw"><table><thead><tr><th>ID</th><th>Estado</th><th>Rota</th><th>Placa</th><th>Rep auditoria</th></tr></thead><tbody>
      ${divs.slice(0, MAXROWS).map(i => `<tr><td><b>${esc(i.id)}</b></td><td>${tTxt(i.tipo)}</td><td>${esc(i.rota)}</td><td>${esc(i.placa) || '–'}</td><td>${esc(i.operador) || '–'}</td></tr>`).join('') || '<tr><td colspan="5" class="mut">Nenhuma divergência na aduana.</td></tr>'}</tbody></table></div></div>`;
  }

  function vRelatorio() {
    return `<div class="card center-card"><h3>Relatório em PDF</h3><p class="mut">O PDF reúne: meta e % de gaiolas fechadas, resultado por ciclo (AM/PM), rotas prioritárias por onda, gaiolas com faltante e a mais (com gaiola/rota corretas), gaiolas que faltam auditar e a conclusão da auditoria.</p>
      <button class="btn lg" data-act="pdf">Gerar PDF agora</button></div>`;
  }

  /* ---- detalhe da gaiola ---- */
  function abrirGaiola(nome) {
    const g = S.res.gaiolas.find(x => x.gaiola === nome); if (!g) return;
    const tipoTxt = t => ({ a_mais: 'A mais', nao_localizado: 'A mais (ID não localizado)', faltante: 'Faltante', regularizado: 'Regularizado (sistema confirma aqui)', regularizado_falt: 'Regularizado (sistema indica outra gaiola)' }[t]);
    $('#modal-body').innerHTML = `<div class="row between"><h3>${esc(g.gaiola)} ${badge(g.status)}</h3><button class="btn sm ghost" id="m-x">Fechar</button></div>
      <div class="facts"><div><span>Ciclo</span><b>${g.ciclo}</b></div><div><span>Rota</span><b>${esc(g.rota) || '–'}</b></div><div><span>Bipados/Esperados</span><b>${g.bipados}/${g.esperados}</b></div><div><span>Operador que bipou</span><b>${esc(g.opBipou.join(', ')) || '–'}</b></div><div><span>Operador que finalizou</span><b>${esc(g.opFinal.join(', ')) || '–'}</b></div><div><span>Rota correta p/ atrelar</span><b>${esc(g.rotasCorretas.join(', ')) || '–'}</b></div><div><span>Última auditoria</span><b>${esc(g.data) || '–'}</b></div></div>
      <div class="tw"><table><thead><tr><th>ID</th><th>Resultado</th><th>Gaiola correta</th><th>Rota correta</th><th>Parada</th><th>Ponto</th><th>Localização</th><th>Detalhe</th></tr></thead><tbody>${g.items.map(i => `<tr><td><b>${esc(i.id)}</b></td><td>${tipoTxt(i.tipo)}</td><td>${esc(i.gaiolaCorreta) || '–'}</td><td>${esc(i.rotaCorreta) || '–'}</td><td>${esc(i.parada) || '–'}</td><td>${esc(i.ponto) || '–'}</td><td>${esc(i.local) || '–'}</td><td>${esc(i.detalhe) || '–'}</td></tr>`).join('') || '<tr><td colspan="8" class="mut">Sem divergências de ID nesta gaiola.</td></tr>'}</tbody></table></div>`;
    $('#modal').classList.add('show'); $('#m-x').onclick = fecharModal;
  }

  /* ---- gráficos ---- */
  function desenharGraficos() {
    if (!root.Chart || !$('#ch-status')) return;
    Object.values(charts).forEach(c => c.destroy()); charts = {};
    const ks = Object.keys(STATUS), cnt = S.res.kpi.cnt;
    charts.status = new Chart($('#ch-status'), { type: 'doughnut', data: { labels: ks.map(k => STATUS[k].label), datasets: [{ data: ks.map(k => cnt[k]), backgroundColor: ks.map(k => STATUS[k].cor), borderWidth: 2, borderColor: '#fff' }] }, options: { maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'right', labels: { boxWidth: 12 } } } } });
    const cs = ['AM', 'PM', 'N/D'].filter(c => S.res.ciclos[c].total);
    charts.ciclo = new Chart($('#ch-ciclo'), { type: 'bar', data: { labels: cs, datasets: ks.map(k => ({ label: STATUS[k].label, data: cs.map(c => S.res.ciclos[c][k]), backgroundColor: STATUS[k].cor, borderRadius: 4 })) }, options: { maintainAspectRatio: false, scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } } } });
  }

  /* ---- render geral ---- */
  const VIEWS = { dashboard: vDashboard, gaiolas: vGaiolas, auditoria: vAuditoria, conclusao: vConclusao, prioridade: vPrioridade, aduana: vAduana, relatorio: vRelatorio };
  function render() {
    $$('.tab').forEach(t => t.classList.toggle('on', t.dataset.tab === S.tab));
    const temDados = S.res && S.res.gaiolas.length;
    const mostra = temDados || (S.tab === 'prioridade' && S.res && S.res.rotas.length) || (S.tab === 'aduana' && S.res && S.res.aduana.length);
    $('#view').innerHTML = mostra ? VIEWS[S.tab]() : vazio();
    $('#updated').textContent = S.updated ? 'Atualizado ' + S.updated.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }) : 'Sem dados';
    if (temDados) { const k = S.res.kpi; const el = $('#top-pct'); el.textContent = pct(k.pct); el.className = 'top-pct ' + (k.pct >= 100 ? 'g' : k.pct >= 80 ? 'y' : 'r'); }
    ligar(); if (S.tab === 'dashboard' && temDados) desenharGraficos();
  }
  function ligar() {
    $$('[data-tab]').forEach(b => b.onclick = () => { S.tab = b.dataset.tab; render(); window.scrollTo(0, 0); });
    $$('[data-ciclo]').forEach(b => b.onclick = () => { S.ciclo = b.dataset.ciclo; render(); });
    $$('[data-filtro]').forEach(b => b.onclick = () => { S.filtro = b.dataset.filtro; render(); });
    $$('[data-div]').forEach(b => b.onclick = () => { S.divTipo = b.dataset.div; render(); });
    $$('[data-adu]').forEach(b => b.onclick = () => { S.aduFiltro = b.dataset.adu; render(); });
    $$('[data-g]').forEach(r => r.onclick = () => abrirGaiola(r.dataset.g));
    $$('#view [data-up]').forEach(i => i.onchange = e => { carregar(i.dataset.up, e.target.files[0]); e.target.value = ''; });
    $$('[data-act]').forEach(b => b.onclick = () => { const a = b.dataset.act; if (a === 'menu') abrirDrawer(true); if (a === 'demo') { const d = buildDemo(); S.raw = d.raw; S.map = d.map; recalc(); } if (a === 'pdf') gerarPDF(); });
    const bu = $('#busca'); if (bu) bu.oninput = e => { S.busca = e.target.value; const pos = e.target.selectionStart; render(); const n = $('#busca'); n.focus(); n.setSelectionRange(pos, pos); };
  }

  /* =====================================================================
     RELATÓRIO PDF
     ===================================================================== */
  function gerarPDF() {
    if (!S.res || !S.res.gaiolas.length) { toast('Carregue os dados antes de gerar o PDF', 'warn'); return; }
    if (!root.jspdf) { toast('Biblioteca de PDF não carregou (verifique a internet).', 'err'); return; }
    const doc = new root.jspdf.jsPDF({ unit: 'mm', format: 'a4' }), W = 210;
    const { kpi, ciclos, gaiolas, rotas } = S.res, c = kpi.cnt;
    const titulo = (t, y) => { doc.setFontSize(13); doc.setTextColor(20, 33, 61); doc.setFont(undefined, 'bold'); doc.text(t, 14, y); doc.setFont(undefined, 'normal'); return y + 3; };
    const head = { fillColor: [20, 33, 61], textColor: 255 };
    doc.setFillColor(20, 33, 61); doc.rect(0, 0, W, 28, 'F');
    doc.setTextColor(255); doc.setFontSize(18); doc.setFont(undefined, 'bold'); doc.text('Painel operacional · RJ15', 14, 13);
    doc.setFontSize(10); doc.setFont(undefined, 'normal'); doc.text('Relatório de inventário e auditoria de gaiolas · ' + new Date().toLocaleString('pt-BR'), 14, 21);
    const cor = kpi.pct >= 100 ? [22, 163, 74] : kpi.pct >= 80 ? [217, 119, 6] : [220, 38, 38];
    doc.setTextColor(...cor); doc.setFontSize(34); doc.setFont(undefined, 'bold'); doc.text(pct(kpi.pct), 14, 46);
    doc.setFontSize(10); doc.setTextColor(90); doc.setFont(undefined, 'normal');
    doc.text(`Gaiolas fechadas sem erro (meta ${kpi.meta}%): ${kpi.ok} de ${kpi.vivas}  ·  canceladas fora da conta: ${c.cancelado}`, 14, 53);
    let y = 60;
    doc.autoTable({ startY: y, head: [['Total', 'Fechadas 100%', 'Ajustadas', 'Com a mais', 'Com faltante', 'Falta auditar', 'Canceladas']], body: [[kpi.total, c.fechada, c.ajustada, c.a_mais, c.faltante, c.pendente + c.aberta + c.em_auditoria, c.cancelado]], headStyles: head, styles: { halign: 'center' }, margin: { left: 14, right: 14 } });
    y = doc.lastAutoTable.finalY + 10; y = titulo('Resultado por ciclo', y);
    doc.autoTable({ startY: y, head: [['Ciclo', 'Gaiolas', 'Fechadas', 'Ajustadas', 'A mais', 'Faltante', 'Falta auditar', 'Canceladas', '% fechadas']], body: ['AM', 'PM', 'N/D'].filter(k => ciclos[k].total).map(k => [k, ciclos[k].total, ciclos[k].fechada, ciclos[k].ajustada, ciclos[k].a_mais, ciclos[k].faltante, ciclos[k].pendente + ciclos[k].aberta + ciclos[k].em_auditoria, ciclos[k].cancelado, pct(ciclos[k].pct)]), headStyles: head, margin: { left: 14, right: 14 } });
    y = doc.lastAutoTable.finalY + 10;
    if (charts.status) { try { doc.addImage(charts.status.toBase64Image(), 'PNG', 14, y, 90, 55); } catch (e) { } }
    if (charts.ciclo) { try { doc.addImage(charts.ciclo.toBase64Image(), 'PNG', 108, y, 88, 55); } catch (e) { } }

    doc.addPage(); y = titulo('Prioridade por rota e onda (rotas com pendência primeiro)', 18);
    doc.autoTable({ startY: y, head: [['Prior.', 'Rota', 'Onda', 'Doca', 'Transportadora', 'Gaiolas OK/Total', 'Faltante', 'A mais', 'Falta auditar', 'Situação']], body: rotas.map(r => [r.prioridade || '–', r.rota, r.onda ? 'Onda ' + r.onda : 'Sem onda', (r.info && r.info.doca) || '–', (r.info && r.info.transp) || '–', r.ok + '/' + r.total, r.faltante, r.a_mais, r.nao_auditadas, r.ring === 'red' ? 'PENDENTE' : r.ring === 'green' ? 'Tudo certo' : 'Sem gaiolas']), headStyles: head, margin: { left: 14, right: 14 }, styles: { fontSize: 8.5 },
      didParseCell: d => { if (d.section === 'body' && d.column.index === 9) { const t = d.cell.raw; d.cell.styles.textColor = t === 'PENDENTE' ? [220, 38, 38] : t === 'Tudo certo' ? [22, 163, 74] : [107, 114, 128]; d.cell.styles.fontStyle = 'bold'; } } });

    const divs = []; gaiolas.forEach(g => g.items.forEach(i => { if (i.tipo === 'a_mais' || i.tipo === 'nao_localizado' || i.tipo === 'faltante') divs.push([i.id, i.tipo === 'faltante' ? 'Faltante' : 'A mais', i.gaiolaAtual, g.rota, i.gaiolaCorreta || 'não consta', i.rotaCorreta || '–', i.operador || '–']); }));
    doc.addPage(); y = titulo('Divergências por ID (a mais e faltante)', 18);
    doc.autoTable({ startY: y, head: [['ID', 'Tipo', 'Gaiola atual', 'Rota', 'Gaiola correta', 'Rota correta', 'Operador']], body: divs.length ? divs : [['Nenhuma divergência', '', '', '', '', '', '']], headStyles: head, margin: { left: 14, right: 14 }, styles: { fontSize: 8 } });

    const falta = gaiolas.filter(g => PENDENTES.includes(g.status));
    y = doc.lastAutoTable.finalY + 10; if (y > 250) { doc.addPage(); y = 18; }
    y = titulo('Gaiolas que faltam auditar', y);
    doc.autoTable({ startY: y, head: [['Gaiola', 'Ciclo', 'Rota', 'Pacotes', 'Situação']], body: falta.length ? falta.map(g => [g.gaiola, g.ciclo, g.rota, g.esperados, STATUS[g.status].label]) : [['Todas auditadas', '', '', '', '']], headStyles: head, margin: { left: 14, right: 14 }, styles: { fontSize: 8.5 } });

    if (S.res.aduana.length) {
      doc.addPage(); y = titulo('Aduana por rota', 18);
      doc.autoTable({ startY: y, head: [['Rota', 'Placa', 'Pacotes', 'Corretos', 'A mais', 'Faltante', 'Rep auditoria']], body: S.res.aduana.map(a => [a.rota, a.placas.join(', ') || '–', a.total, a.correto, a.a_mais, a.faltante, a.ops.join(', ') || '–']), headStyles: head, margin: { left: 14, right: 14 }, styles: { fontSize: 8.5 } });
    }
    const fin = gaiolas.filter(g => g.status === 'fechada' || g.status === 'ajustada');
    doc.addPage(); y = titulo('Conclusão da auditoria: gaiolas fechadas e ajustadas', 18);
    doc.autoTable({ startY: y, head: [['Gaiola', 'Ciclo', 'Rota', 'Situação', 'Operador que finalizou']], body: fin.length ? fin.map(g => [g.gaiola, g.ciclo, g.rota, STATUS[g.status].label, g.opFinal.join(', ') || '–']) : [['Nenhuma ainda', '', '', '', '']], headStyles: head, margin: { left: 14, right: 14 }, styles: { fontSize: 8.5 } });

    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) { doc.setPage(i); doc.setFontSize(8); doc.setTextColor(140); doc.text('Painel operacional RJ15 · página ' + i + ' de ' + n, 14, 290); }
    doc.save('Painel-operacional-RJ15-' + new Date().toISOString().slice(0, 10) + '.pdf');
    toast('PDF gerado', 'ok');
  }

  /* ---- inicialização ---- */
  async function init() {
    $('#burger').onclick = () => abrirDrawer(true);
    $('#drawer-x').onclick = () => abrirDrawer(false);
    $('#scrim').onclick = () => abrirDrawer(false);
    $('#modal').onclick = e => { if (e.target.id === 'modal') fecharModal(); };
    document.addEventListener('keydown', e => { if (e.key === 'Escape') { abrirDrawer(false); fecharModal(); } });
    $('#btn-pdf').onclick = gerarPDF;
    $('#btn-limpar').onclick = () => { if (confirm('Limpar todos os dados carregados?')) { S.raw = {}; S.map = {}; S.res = null; S.updated = null; render(); renderDrawer(); salvar(); abrirDrawer(false); } };
    try { const e = await DB.get('estado'); if (e && e.raw) { S.raw = e.raw; S.map = e.map; delete S.raw.base; delete S.map.base; S.updated = e.updated ? new Date(e.updated) : null; S.res = compute(S.raw, S.map); } } catch (err) { }
    render(); renderDrawer();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})(typeof window !== 'undefined' ? window : globalThis);
