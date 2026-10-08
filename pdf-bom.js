/* ---------------- lista de material dentro de PDF ---------------- */
// Lê os textos de cada página do PDF (com a posição de cada um) e procura, em QUALQUER ponto
// de QUALQUER folha, as tabelas de lista de material: a tabela não tem lugar fixo no desenho
// (pode estar no canto, no meio, em cima do carimbo, com o cabeçalho em cima ou embaixo das
// linhas, duas lado a lado...). O que identifica a tabela é a linha de cabeçalho — Item, Qtd,
// Título/Descrição, Especificação, Material, Massa —, reconhecida pelo NOME das colunas, com o
// mesmo vocabulário que o app usa para planilhas. Achado o cabeçalho, as linhas da tabela são
// os textos logo acima ou logo abaixo dele, alinhados com as colunas, até o texto parar.
//
// O resultado é uma tabela comum (cabeçalho + linhas), que entra no app pelo mesmo caminho do
// .xlsx — as mesmas regras de agrupamento, conjunto, massa por peça etc. valem para o PDF.
//
// A leitura do PDF usa o pdf.js (Mozilla), carregado do CDN só na primeira vez que um PDF é
// aberto. A parte que acha a tabela (tabelasDaPagina) não depende dele: recebe só os textos
// com posição, e por isso é testada direto no Node (tools/test_pdf.js).
(function(raiz){
  "use strict";

  const PDFJS_VERSAO = "4.10.38";
  const PDFJS_CDNS = [
    `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSAO}/`,
    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSAO}/build/`,
  ];
  let pdfjsPromessa = null;

  // carrega o pdf.js uma vez só. O worker é criado a partir de um blob que importa o script do
  // CDN — assim funciona também com a página aberta direto do disco (file://), onde o
  // navegador não deixa criar worker de outro endereço. Se nem assim der, o pdf.js roda sem
  // worker (mais lento, mas lê igual).
  function carregarPdfjs(){
    if(pdfjsPromessa) return pdfjsPromessa;
    pdfjsPromessa = (async () => {
      let ultimoErro = null;
      for(const base of PDFJS_CDNS){
        try{
          const lib = await import(base + "pdf.min.mjs");
          lib.GlobalWorkerOptions.workerSrc = base + "pdf.worker.min.mjs";
          try{
            const codigo = `import "${base}pdf.worker.min.mjs";`;
            const url = URL.createObjectURL(new Blob([codigo], {type:"text/javascript"}));
            lib.GlobalWorkerOptions.workerPort = new Worker(url, {type:"module"});
          }catch(e){ /* sem worker: o pdf.js cai sozinho no modo sem worker */ }
          return lib;
        }catch(e){ ultimoErro = e; }
      }
      pdfjsPromessa = null;   // deixa tentar de novo depois (ex.: a internet voltou)
      throw new Error("não consegui carregar o leitor de PDF (pdf.js) — confira a conexão com a internet"
        + (ultimoErro && ultimoErro.message ? ` (${ultimoErro.message})` : ""));
    })();
    return pdfjsPromessa;
  }

  // textos de cada página: [{numero, itens:[{s, x0, x1, y, h}]}], com y medido de cima para
  // baixo e já na orientação em que a folha é vista (página girada sai desvirada). Texto
  // inclinado ou na vertical (cotas, rótulos do desenho) fica de fora — tabela é horizontal.
  async function textosDoPdf(dados, aoProgredir, lib){
    const pdfjs = lib || await carregarPdfjs();
    const doc = await pdfjs.getDocument({data: dados, verbosity: 0, isEvalSupported: false}).promise;
    const paginas = [];
    try{
      for(let n = 1; n <= doc.numPages; n++){
        if(aoProgredir) aoProgredir(n, doc.numPages);
        const pag = await doc.getPage(n);
        const vp = pag.getViewport({scale: 1});
        const tc = await pag.getTextContent();
        const itens = [];
        for(const it of tc.items){
          if(!it.str || !it.str.trim()) continue;
          const t = pdfjs.Util.transform(vp.transform, it.transform);
          const h = Math.hypot(t[2], t[3]);
          const escala = Math.hypot(t[0], t[1]);
          if(!(h > 0) || !(escala > 0)) continue;
          if(Math.abs(t[1]) > 0.02 * escala || t[0] <= 0) continue;   // fora da horizontal
          const larg = (it.width || 0) * (vp.scale || 1);
          itens.push({s: it.str, x0: t[4], x1: t[4] + larg, y: t[5], h});
        }
        paginas.push({numero: n, itens});
        pag.cleanup();
      }
    } finally {
      try{ await doc.destroy(); }catch(e){}
    }
    return paginas;
  }

  /* ---------- montagem das linhas de texto ---------- */
  const mediana = v => { if(!v.length) return 0; const a = [...v].sort((p, q) => p - q); return a[a.length >> 1]; };

  // junta os pedaços que estão na mesma altura (mesma linha de base, com folga)
  function agruparEmLinhas(itens){
    const ord = [...itens].sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const linhas = [];
    for(const it of ord){
      const ult = linhas[linhas.length - 1];
      if(ult && Math.abs(it.y - ult.y) <= 0.45 * Math.min(it.h, ult.h)) ult.itens.push(it);
      else linhas.push({y: it.y, h: it.h, itens: [it]});
    }
    for(const l of linhas){
      l.itens.sort((a, b) => a.x0 - b.x0);
      l.h = mediana(l.itens.map(i => i.h));
    }
    return linhas;
  }

  // pedaços de texto lado a lado: com espaço só quando há espaço de verdade entre eles
  // (o pdf.js às vezes corta uma palavra em dois pedaços colados)
  function juntarNaLinha(pecas){
    const ord = [...pecas].sort((a, b) => a.x0 - b.x0);
    let txt = "", fim = null;
    for(const p of ord){
      const s = p.s.trim();
      if(!s) continue;
      if(txt && (fim === null || p.x0 - fim > 0.18 * p.h) && !/\s$/.test(txt)) txt += " ";
      txt += s; fim = p.x1;
    }
    return txt;
  }
  function juntarCelula(pecas){
    const linhas = agruparEmLinhas(pecas);
    return linhas.map(l => juntarNaLinha(l.itens)).filter(Boolean).join(" ").replace(/\s+/g, " ").trim();
  }

  /* ---------- cabeçalho ---------- */
  // quebra um pedaço de texto em palavras, com a posição estimada de cada uma — o pdf.js
  // costuma entregar "Item Qty." como um pedaço só, e no cabeçalho cada palavra pode ser
  // uma coluna
  function palavras(it){
    const s = it.s;
    const n = s.length || 1;
    const larg = it.x1 - it.x0;
    const out = [];
    const re = /\S+/g; let m;
    while((m = re.exec(s))){
      out.push({s: m[0], x0: it.x0 + larg * m.index / n, x1: it.x0 + larg * (m.index + m[0].length) / n, y: it.y, h: it.h});
    }
    return out;
  }

  // classifica o título de uma coluna; se o título inteiro não for reconhecido, vale a
  // primeira palavra reconhecida ("Nº ITEM" -> item, "Total QTY" -> qtd)
  function campoDoTitulo(titulo, classificar){
    const c = classificar(titulo);
    if(c) return c;
    for(const p of String(titulo).split(/\s+/)){
      const cp = classificar(p);
      if(cp) return cp;
    }
    return null;
  }

  // palavras de uma linha -> células de cabeçalho. Duas palavras próximas ficam na mesma
  // célula ("Mass [kg]", "Item No.", "Part Number"), menos quando as duas sozinhas já são
  // nomes de coluna ("Item Qty." são duas colunas).
  function celulasDoCabecalho(linha, classificar){
    const ws = linha.itens.flatMap(palavras).sort((a, b) => a.x0 - b.x0);
    const cels = [];
    for(const w of ws){
      const cur = cels[cels.length - 1];
      const perto = cur && (w.x0 - cur.x1) < 1.2 * Math.max(w.h, cur.h);
      if(perto && !(classificar(w.s) && classificar(cur.palavras[cur.palavras.length - 1]))){
        cur.palavras.push(w.s); cur.x1 = Math.max(cur.x1, w.x1);
      } else {
        cels.push({palavras: [w.s], x0: w.x0, x1: w.x1, y: w.y, h: w.h});
      }
    }
    for(const c of cels){
      c.titulo = c.palavras.join(" ");
      c.campo = c.palavras.length <= 5 ? campoDoTitulo(c.titulo, classificar) : null;
    }
    return cels;
  }

  // uma linha pode ter mais de uma tabela (duas listas lado a lado): quando um nome de coluna
  // se repete, começa outra tabela. Colunas sem nome conhecido nas pontas são texto de fora
  // da tabela; no meio, são colunas extras da própria tabela.
  function cabecalhosDaLinha(linha, classificar){
    const cels = celulasDoCabecalho(linha, classificar);
    const grupos = [];
    let atual = [], vistos = new Set();
    for(const c of cels){
      if(c.campo && vistos.has(c.campo)){ grupos.push(atual); atual = []; vistos = new Set(); }
      atual.push(c);
      if(c.campo) vistos.add(c.campo);
    }
    grupos.push(atual);
    const out = [];
    for(let g of grupos){
      while(g.length && !g[0].campo) g = g.slice(1);
      while(g.length && !g[g.length - 1].campo) g = g.slice(0, -1);
      const campos = new Set(g.filter(c => c.campo).map(c => c.campo));
      if(campos.size < 3 || !campos.has("qtd")) continue;
      if(!["item","titulo","descricao","especificacao","dimensoes"].some(c => campos.has(c))) continue;
      // cabeçalho é texto curto, de colunas afastadas — frase de nota não é cabeçalho
      if(g.length > 14) continue;
      out.push(g);
    }
    return out;
  }

  /* ---------- colunas ---------- */
  // Bordas das colunas. Os desenhos centralizam o título na célula, mas o valor pode estar
  // centralizado ou encostado à esquerda (a especificação comprida começa logo na borda) —
  // então "a coluna cujo título está mais perto" erra. Três modelos de borda são testados e
  // fica o que deixa menos textos atravessando bordas:
  //  - título centralizado: cada borda é o espelho da anterior pelo centro do título;
  //  - título encostado à esquerda: a borda fica logo antes de cada título;
  //  - ponto médio entre os centros dos títulos (o mais simples, de reserva).
  function bordasCandidatas(cab){
    const n = cab.length;
    const cand = [];
    const valida = b => {
      for(let j = 0; j < n; j++){
        if(!(b[j + 1] > b[j])) return false;
        if(cab[j].x0 < b[j] - 0.5 || cab[j].x1 > b[j + 1] + 0.5) return false;   // o título cabe na coluna
      }
      return true;
    };
    const centros = cab.map(c => (c.x0 + c.x1) / 2);
    const h = mediana(cab.map(c => c.h)) || 6;
    // 1) centralizado: b[j+1] = 2·centro[j] − b[j]; varia só a primeira borda
    for(let b0 = cab[0].x0 - 8 * h; b0 <= cab[0].x0; b0 += 0.25){
      const b = [b0];
      for(let j = 0; j < n; j++) b.push(2 * centros[j] - b[j]);
      if(valida(b)) cand.push({modelo: "centro", b});
    }
    // 2) encostado à esquerda
    for(let p = 0.5; p <= 1.6 * h; p += 0.5){
      const b = cab.map(c => c.x0 - p);
      b.push(cab[n - 1].x1 + Math.max(p, (cab[n - 1].x1 - cab[n - 1].x0)));
      if(valida(b)) cand.push({modelo: "esquerda", b});
    }
    // 3) ponto médio entre centros
    const b = [cab[0].x0 - 1.5 * h];
    for(let j = 0; j < n - 1; j++) b.push((centros[j] + centros[j + 1]) / 2);
    b.push(cab[n - 1].x1 + 1.5 * h);
    if(valida(b)) cand.push({modelo: "meio", b});
    return cand;
  }
  // texto atravessando borda (com uma folga pequena: a especificação comprida às vezes vaza
  // um pouquinho na coluna do lado)
  function cruzamentos(b, pecas){
    let pen = 0;
    for(const p of pecas){
      for(let j = 1; j < b.length - 1; j++){
        if(b[j] > p.x0 + 1.5 && b[j] < p.x1 - 1.5) pen++;
      }
    }
    return pen;
  }
  function escolherBordas(cab, pecas){
    const cand = bordasCandidatas(cab);
    if(!cand.length) return null;
    const dentro = pecas.filter(p => { const c = (p.x0 + p.x1) / 2; return c > cand[0].b[0] - 50 && c < cand[0].b[cand[0].b.length - 1] + 50; });
    let melhor = null;
    const porModelo = {centro: [], esquerda: [], meio: []};
    for(const c of cand){
      c.pen = cruzamentos(c.b, dentro);
      porModelo[c.modelo].push(c);
    }
    for(const m of ["centro", "esquerda", "meio"]){
      const lista = porModelo[m];
      if(!lista.length) continue;
      const min = Math.min(...lista.map(c => c.pen));
      // várias posições empatadas: fica a do meio da faixa (a mais folgada)
      const emp = lista.filter(c => c.pen === min);
      const esc = emp[emp.length >> 1];
      if(!melhor || esc.pen < melhor.pen) melhor = esc;
    }
    return melhor.b;
  }
  function colunaDe(b, p){
    const c = (p.x0 + p.x1) / 2;
    if(c < b[0] || c > b[b.length - 1]) return -1;
    for(let j = 0; j < b.length - 1; j++) if(c >= b[j] && c <= b[j + 1]) return j;
    return -1;
  }

  /* ---------- linhas da tabela ---------- */
  const RE_ITEM = /^(?:[A-Za-z]{0,3}[-.]?)?\d{1,4}(?:[.\-/]\d{1,4}){0,3}[A-Za-z]?\.?$/;
  const RE_NUM = /^\d+(?:[.,]\d+)?(?:\s*(?:pç|pc|pcs|un|und|unid|ea|x))?\.?$/i;

  // a partir do cabeçalho, anda para cima ou para baixo pegando as linhas de texto da tabela
  // até o texto parar (um vão grande), outro cabeçalho, ou um texto que atravessa várias
  // colunas (título da tabela, nota, carimbo)
  function blocoNaDirecao(linhasTab, iCab, dir, ctx){
    const out = [];
    let ultY = linhasTab[iCab].y;
    const vaos = [];
    for(let i = iCab + dir; i >= 0 && i < linhasTab.length; i += dir){
      const l = linhasTab[i];
      const vao = Math.abs(l.y - ultY);
      const limite = Math.max(3.2 * ctx.h, vaos.length >= 2 ? 2.4 * mediana(vaos) : 0);
      if(vao > limite) break;
      if(l.cabecalho) break;
      const atravessa = l.itens.some(p => {
        let k = 0; for(let j = 1; j < ctx.b.length - 1; j++) if(ctx.b[j] > p.x0 + 2 && ctx.b[j] < p.x1 - 2) k++;
        return k >= 2;
      });
      if(atravessa) break;
      out.push(l);
      vaos.push(vao);
      ultY = l.y;
    }
    return out;
  }

  function ehAncora(p, ctx){
    const s = p.s.trim();
    if(ctx.colAncora === ctx.colItem) return RE_ITEM.test(s);
    return RE_NUM.test(s);
  }

  // o pdf.js às vezes entrega o texto de duas células vizinhas num pedaço só ("BLIND FLANGE
  // DIN 2527, DN 100…" = Título + Especificação). Pedaço que atravessa uma borda com boa
  // parte de cada lado é cortado no espaço entre palavras mais perto da borda. (Só vazar um
  // pouco na coluna do lado — especificação comprida — não corta.)
  function cortarNasBordas(p, b){
    const larg = p.x1 - p.x0;
    for(let j = 1; j < b.length - 1; j++){
      const esq = b[j] - p.x0, dir = p.x1 - b[j];
      if(esq < 0.25 * larg || dir < 0.25 * larg || esq < 1.5 * p.h || dir < 1.5 * p.h) continue;
      let melhor = -1, dist = Infinity;
      for(let k = 1; k < p.s.length; k++){
        if(p.s[k] !== " " || p.s[k - 1] === " ") continue;
        const x = p.x0 + larg * k / p.s.length;
        if(Math.abs(x - b[j]) < dist){ dist = Math.abs(x - b[j]); melhor = k; }
      }
      if(melhor < 0 || dist > 2.5 * p.h) continue;
      const xc = p.x0 + larg * melhor / p.s.length;
      const a = {...p, s: p.s.slice(0, melhor), x1: xc};
      const c = {...p, s: p.s.slice(melhor + 1), x0: xc + larg / p.s.length};
      return [...cortarNasBordas(a, b), ...cortarNasBordas(c, b)];
    }
    return [p];
  }

  // distribui os pedaços de texto pelas linhas da tabela. Cada linha começa no número do
  // item. Em cada coluna, linhas de texto bem juntinhas (até ~1,25× a altura da letra) são a
  // mesma célula quebrada em várias linhas; a célula vai para o item cuja altura ela cobre —
  // vale para célula centralizada na altura e para célula encostada em cima. Célula que não
  // cobre nenhum item vai para o vizinho que ainda não tem nada naquela coluna.
  function montarLinhas(bloco, ctx){
    const pecas = [];
    for(const l of bloco) for(const p0 of l.itens){
      for(const p of cortarNasBordas(p0, ctx.b)){
        const col = colunaDe(ctx.b, p);
        if(col >= 0) pecas.push({...p, col});
      }
    }
    // âncoras: número do item na coluna Item (ou, sem coluna Item, a quantidade)
    const ancoras = pecas.filter(p => p.col === ctx.colAncora && ehAncora(p, ctx)).sort((a, b) => a.y - b.y);
    if(!ancoras.length) return [];
    const linhas = ancoras.map(a => ({y: a.y, h: a.h, cel: ctx.b.slice(1).map(() => [])}));
    const passo = linhas.length > 1 ? mediana(linhas.slice(1).map((r, k) => r.y - linhas[k].y)) : 3 * ctx.h;
    const maisPerto = y => linhas.reduce((m, r) => Math.abs(r.y - y) < Math.abs(m.y - y) ? r : m, linhas[0]);
    for(let col = 0; col < ctx.b.length - 1; col++){
      const daColuna = pecas.filter(p => p.col === col).sort((a, b) => a.y - b.y);
      // blocos = células (uma ou mais linhas de texto coladas)
      const blocos = [];
      for(const p of daColuna){
        const ult = blocos[blocos.length - 1];
        const fimUlt = ult ? ult[ult.length - 1] : null;
        if(ult && p.y - fimUlt.y <= 1.25 * Math.max(p.h, fimUlt.h)) ult.push(p);
        else blocos.push([p]);
      }
      // 1º as células que cobrem a altura de algum item; depois as soltas, já sabendo quais
      // linhas têm texto próprio naquela coluna
      const soltos = [];
      for(const bl of blocos){
        const y0 = bl[0].y - 0.6 * bl[0].h, y1 = bl[bl.length - 1].y + 0.6 * bl[bl.length - 1].h;
        const cobertas = linhas.filter(r => r.y >= y0 && r.y <= y1);
        if(cobertas.length === 1) cobertas[0].cel[col].push(...bl);
        else if(cobertas.length > 1) for(const p of bl) maisPerto(p.y).cel[col].push(p);
        else soltos.push(bl);
      }
      for(const bl of soltos){
        // não cobre nenhum item: vai para o vizinho de cima ou de baixo
        const yc = (bl[0].y + bl[bl.length - 1].y) / 2;
        const A = [...linhas].reverse().find(r => r.y < yc);
        const B = linhas.find(r => r.y > yc);
        const vazia = r => r && !r.cel[col].length;
        let alvo = null;
        if(A && B){
          if(vazia(B) && !vazia(A)) alvo = B;
          else if(vazia(A) && !vazia(B)) alvo = A;
          else alvo = (yc - A.y) <= (B.y - yc) ? A : B;
        } else {
          // antes da primeira ou depois da última linha da tabela: só se estiver colado nela
          const r = A || B;
          if(r && Math.abs(r.y - yc) <= Math.max(1.7 * ctx.h, 0.6 * passo)) alvo = r;
        }
        if(alvo && Math.abs(alvo.y - yc) <= Math.max(2.5 * passo, 3 * ctx.h)) alvo.cel[col].push(...bl);
      }
    }
    return linhas.map(r => r.cel.map(juntarCelula));
  }

  // título da tabela ("LISTA DE MATERIAL", "BILL OF MATERIAL - NEW LINE ..."): a linha de
  // texto colada no cabeçalho, do lado oposto às linhas da tabela
  function tituloDaTabela(linhasTodas, cabLinha, dir, ctx){
    const larg = ctx.b[ctx.b.length - 1] - ctx.b[0];
    const cand = linhasTodas
      .filter(l => (dir > 0 ? l.y < cabLinha.y : l.y > cabLinha.y) && Math.abs(l.y - cabLinha.y) <= 3.2 * ctx.h)
      .sort((a, b) => Math.abs(a.y - cabLinha.y) - Math.abs(b.y - cabLinha.y));
    for(const l of cand){
      const pecas = l.itens.filter(p => {
        const c = (p.x0 + p.x1) / 2;
        return c > ctx.b[0] - 0.1 * larg && c < ctx.b[ctx.b.length - 1] + 0.1 * larg;
      });
      const txt = juntarNaLinha(pecas);
      if(txt && /[A-Za-zÀ-ú]{3}/.test(txt)) return txt;
    }
    return "";
  }

  // todas as tabelas de lista de material de uma página
  function tabelasDaPagina(itens, classificar){
    const linhas = agruparEmLinhas(itens);
    const cabecalhos = [];
    for(const l of linhas){
      for(const cab of cabecalhosDaLinha(l, classificar)){ cabecalhos.push({linha: l, cab}); l.cabecalho = true; }
    }
    const tabelas = [];
    for(const {linha, cab} of cabecalhos){
      const h = mediana(cab.map(c => c.h)) || linha.h;
      // primeira passada: bordas pelos títulos só, para separar o texto que é da tabela
      const prov = bordasCandidatas(cab);
      if(!prov.length) continue;
      const larg0 = prov[0].b;
      const x0 = larg0[0] - 2 * h, x1 = larg0[larg0.length - 1] + 2 * h;
      const pertoX = itens.filter(p => { const c = (p.x0 + p.x1) / 2; return c >= x0 && c <= x1; });
      const pertoY = pertoX.filter(p => Math.abs(p.y - linha.y) < 400);
      const b = escolherBordas(cab, pertoY.filter(p => p.y !== linha.y)) || larg0;
      const ctx = {b, h,
        colItem: cab.findIndex(c => c.campo === "item"),
        colQtd: cab.findIndex(c => c.campo === "qtd")};
      ctx.colAncora = ctx.colItem >= 0 ? ctx.colItem : ctx.colQtd;
      // colunas de texto: todas menos as numéricas (item, quantidade, massa, área, comprimento)
      ctx.colTexto = cab.map((c, j) => j).filter(j => !["item","qtd","massa","area","comprimento"].includes(cab[j].campo));
      // só o texto que cai dentro das colunas desta tabela
      const dentro = itens.filter(p => colunaDe(b, p) >= 0);
      const linhasTab = agruparEmLinhas(dentro);
      for(const l of linhasTab) l.cabecalho = linhas.some(o => o.cabecalho && o !== linha && Math.abs(o.y - l.y) < 0.3 * h
        && o.itens.some(p => colunaDe(b, p) >= 0));
      let iCab = linhasTab.findIndex(l => Math.abs(l.y - linha.y) <= 0.45 * h);
      if(iCab < 0) continue;
      linhasTab[iCab].cabecalho = false;
      // cabeçalho em cima (linhas para baixo) ou embaixo (linhas para cima, como na lista
      // sobre o carimbo): fica o lado que tem mais itens
      let melhor = null;
      for(const dir of [1, -1]){
        const bloco = blocoNaDirecao(linhasTab, iCab, dir, ctx);
        // linha de verdade tem o que a peça é escrito em alguma coluna de texto; só números
        // (cotas do desenho alinhadas por acaso com a tabela) não é item
        const rows = montarLinhas(bloco, ctx)
          .filter(r => r.some((v, j) => ctx.colTexto.includes(j) && /\p{L}{2}/u.test(v)));
        if(rows.length && (!melhor || rows.length > melhor.rows.length)) melhor = {dir, rows};
      }
      if(!melhor) continue;
      tabelas.push({
        y: linha.y, x: b[0],
        titulo: tituloDaTabela(linhas, linha, melhor.dir, ctx),
        colunas: cab.map(c => ({titulo: c.titulo, campo: c.campo})),
        linhas: melhor.rows,
      });
    }
    // ordem de leitura: de cima para baixo, da esquerda para a direita
    tabelas.sort((a, b) => a.y - b.y || a.x - b.x);
    return tabelas;
  }

  // junta as tabelas de todas as folhas numa tabela só (cabeçalho + linhas), com as colunas
  // reconhecidas encaixadas pelo campo e duas colunas a mais: a folha e o título da tabela.
  function tabelaDasPaginas(paginas, classificar){
    const colunas = [];   // [{titulo, campo}]
    const chave = c => c.campo || ("?" + String(c.titulo).toLowerCase());
    const achadas = [];
    for(const pg of paginas){
      for(const t of tabelasDaPagina(pg.itens, classificar)){
        for(const c of t.colunas) if(!colunas.some(o => chave(o) === chave(c))) colunas.push(c);
        achadas.push({pagina: pg.numero, ...t});
      }
    }
    const comTitulo = achadas.some(t => t.titulo);
    const cabecalho = [...colunas.map(c => c.titulo), "Folha", ...(comTitulo ? ["Tabela"] : [])];
    const linhas = [cabecalho];
    for(const t of achadas){
      const idx = t.colunas.map(c => colunas.findIndex(o => chave(o) === chave(c)));
      for(const r of t.linhas){
        const out = colunas.map(() => "");
        r.forEach((v, j) => { if(idx[j] >= 0) out[idx[j]] = v; });
        out.push(String(t.pagina));
        if(comTitulo) out.push(t.titulo || "");
        linhas.push(out);
      }
    }
    const semTexto = paginas.filter(p => !p.itens.length).map(p => p.numero);
    return {
      tabela: linhas,
      tabelas: achadas.map(t => ({pagina: t.pagina, titulo: t.titulo, linhas: t.linhas.length})),
      paginas: paginas.length,
      semTexto,
    };
  }

  async function lerListaDoPdf(dados, classificar, aoProgredir, lib){
    const paginas = await textosDoPdf(dados, aoProgredir, lib);
    return tabelaDasPaginas(paginas, classificar);
  }

  const api = {carregarPdfjs, textosDoPdf, tabelasDaPagina, tabelaDasPaginas, lerListaDoPdf};
  raiz.PdfBom = api;
  if(typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
