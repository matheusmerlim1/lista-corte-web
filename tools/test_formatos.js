/*
 * Teste dos outros formatos de lista: abre a página num Chrome/Edge sem janela e confere que
 * a ferramenta acha os campos sozinha, qualquer que seja o jeito de o documento escrevê-los.
 *
 *  - BOM exportada do SolidWorks (referencia/*.xlsx): cabeçalho em inglês, tipo na
 *    DESCRIPTION, medida na Dimensions, norma na Specification, massa por peça ("8.912 kg").
 *  - Lista com valores já calculados: Comprimento unitário, Área unitária, Massa unitária.
 *
 * Uso: node tools/test_formatos.js
 */
const { spawn } = require("child_process");
const fs = require("fs"), os = require("os"), path = require("path");

const RAIZ = path.join(__dirname, "..");
const PASTA_REF = path.join(RAIZ, "referencia");
const NAVEGADOR = ["C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe"].find(p => fs.existsSync(p));
const esperar = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const perfil = fs.mkdtempSync(path.join(os.tmpdir(), "p9t-"));
  const nav = spawn(NAVEGADOR, ["--headless=new", "--disable-gpu", "--allow-file-access-from-files",
    "--remote-debugging-port=9343", `--user-data-dir=${perfil}`, "about:blank"], { stdio: "ignore" });
  let alvo = null;
  for (let i = 0; i < 40 && !alvo; i++) {
    await esperar(250);
    try { alvo = (await (await fetch("http://127.0.0.1:9343/json/list")).json()).find(t => t.type === "page"); } catch (e) { /* subindo */ }
  }
  const ws = new WebSocket(alvo.webSocketDebuggerUrl);
  await new Promise(r => ws.addEventListener("open", r));
  let id = 0; const pend = new Map(); const logs = [];
  ws.addEventListener("message", ev => {
    const m = JSON.parse(ev.data);
    if (m.id && pend.has(m.id)) { pend.get(m.id)(m); pend.delete(m.id); }
    if (m.method === "Runtime.exceptionThrown") logs.push(String((m.params.exceptionDetails.exception || {}).description).split("\n")[0]);
  });
  const send = (metodo, params = {}) => new Promise(res => { const i = ++id; pend.set(i, res); ws.send(JSON.stringify({ id: i, method: metodo, params })); });
  const rodar = async expr => {
    const r = await send("Runtime.evaluate", { expression: expr, returnByValue: true, awaitPromise: true });
    if (r.result.exceptionDetails) throw new Error(String((r.result.exceptionDetails.exception || {}).description).split("\n")[0]);
    return r.result.result.value;
  };

  await send("Runtime.enable"); await send("Page.enable");
  await send("Page.navigate", { url: encodeURI("file:///" + RAIZ.replace(/\\/g, "/") + "/index.html") });
  await esperar(1500);

  const ok = [], falhas = [];
  const passo = async (oque, fn) => {
    try { const r = await fn(); if (r === false) falhas.push(oque); else ok.push(oque); }
    catch (e) { falhas.push(oque + " — " + String(e.message).split("\n")[0]); }
  };
  const perto = (a, b, tol) => Math.abs(a - b) <= tol;

  async function soltar(arquivo) {
    const b64 = fs.readFileSync(arquivo).toString("base64");
    await rodar(`(async () => {
      const bin = atob(${JSON.stringify(b64)});
      const buf = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      const arq = new File([buf], ${JSON.stringify(path.basename(arquivo))}, {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const dt = new DataTransfer();
      dt.items.add(arq);
      window.dispatchEvent(new DragEvent("dragenter", {dataTransfer: dt, bubbles:true}));
      window.dispatchEvent(new DragEvent("drop", {dataTransfer: dt, bubbles:true}));
      await new Promise(r => setTimeout(r, 900));
      return true;
    })()`);
    await esperar(600);
  }
  /** grupos da lista, já interpretados */
  const grupos = async () => JSON.parse(await rodar(`JSON.stringify(lastGrouped.map(g => {
    const c = classifyGroup(g);
    return {esp:g.especificacao, desc:g.descricao, mat:g.material, qtd:g.qtd, massa:g.massa,
            tipo:c.tipo, ok:c.ok, e:c.thickness, w:c.width, h:c.height, L:c.length,
            ident:c.identidade, estimada:c.estimada, area:areaM2DaLinha(g, c)};
  }))`));

  /* ---------- BOM do SolidWorks ---------- */
  const bom = fs.existsSync(PASTA_REF) ? fs.readdirSync(PASTA_REF).find(f => /BOM.*\.xlsx$/i.test(f)) : null;
  if (bom) {
    await soltar(path.join(PASTA_REF, bom));
    const g = await grupos();

    await passo("BOM do SolidWorks: nenhuma linha fica de fora", async () => {
      const linhas = await rodar(`registrosOriginais().length`);
      if (linhas !== 36) throw new Error("linhas lidas: " + linhas + " de 36");
      return true;
    });

    await passo("BOM: tipo vem da DESCRIPTION, medida da Dimensions", async () => {
      const chapas = g.filter(x => x.tipo === "chapa"), barras = g.filter(x => x.tipo === "barra");
      if (chapas.length !== 5) throw new Error("chapas: " + chapas.length + " — " + chapas.map(x => x.desc).join(" | "));
      if (barras.length !== 10) throw new Error("perfis/barras: " + barras.length);
      const naoLidas = [...chapas, ...barras].filter(x => !x.ok);
      if (naoLidas.length) throw new Error("não interpretadas: " + naoLidas.map(x => x.esp + " " + x.desc).join(" | "));
      return true;
    });

    await passo("BOM: massa por peça (e decimal com ponto) vira total da linha", async () => {
      const b = g.find(x => /203 mm/.test(x.desc));         // 8 × 0.495 kg
      if (!b || !perto(b.massa, 3.96, 0.001)) throw new Error("barra chata 203 mm: " + (b && b.massa));
      const c = g.find(x => /300 x 300/.test(x.desc));       // 1 × "8.912 kg"
      if (!c || !perto(c.massa, 8.912, 0.001)) throw new Error("chapa 300×300: " + (c && c.massa));
      return true;
    });

    await passo("BOM: espessura e tamanho da chapa completados pelo código da peça", async () => {
      const c1 = g.find(x => x.tipo === "chapa" && x.w === 300 && x.h === 300);
      if (!c1 || !perto(c1.e, 12.7, 0.01)) throw new Error("chapa 1/2\" 300×300 não saiu: " + JSON.stringify(g.filter(x=>x.tipo==="chapa")));
      const c2 = g.find(x => x.tipo === "chapa" && x.w === 120 && x.h === 230);
      if (!c2 || !perto(c2.e, 9.53, 0.01)) throw new Error("chapa 3/8\" 120×230: " + JSON.stringify(c2));
      const est = g.filter(x => x.tipo === "chapa" && x.estimada);
      if (est.length !== 2) throw new Error("chapas só com espessura (estimadas pela massa): " + est.length);
      return true;
    });

    await passo("BOM: comprimento de perfil — descrição, ou massa ÷ kg/m", async () => {
      const L = g.find(x => /1855 mm/.test(x.desc));
      if (!L || L.L !== 1855 || /,$/.test(L.ident)) throw new Error("cantoneira 1855: " + JSON.stringify(L));
      const U = g.find(x => /kg\/m/.test(x.desc));
      // 28,584 kg ÷ 9,3 kg/m = 3,074 m
      if (!U || !perto(U.L, 3074, 1) || U.estimada !== "massa") throw new Error("perfil U: " + JSON.stringify(U));
      return true;
    });

    await passo("BOM: norma entra no agrupamento e na descrição dos itens soltos", async () => {
      const nl6 = g.filter(x => /NL6/.test(x.esp) && /^M6/.test(x.desc));
      if (nl6.length !== 2) throw new Error("arruelas NL6 M6 (DIN 25201 e 25201-4 são peças diferentes): " + nl6.map(x => x.desc).join(" | "));
      const porca = g.find(x => /PORCA BAIXA/.test(x.esp) && /^M6/.test(x.desc));
      if (!porca || porca.qtd !== 40) throw new Error("porca M6 (23 + 17, códigos diferentes) devia somar 40: " + (porca && porca.qtd));
      return true;
    });

    await passo("BOM: consolidado no padrão do documento devolve todas as colunas", async () => {
      const r = JSON.parse(await rodar(`(() => { const s = buildPadraoOriginalSheet(); return JSON.stringify({h:s.headers, rows:s.rows}); })()`));
      const esperado = ["ITEM NO.","PART NUMBER","DESCRIPTION","Dimensions","Specification","Material","Treating_1","QTY.","Mass","Area"];
      if (r.h.join("|") !== esperado.join("|")) throw new Error("colunas: " + r.h.join(" | "));
      const porca = r.rows.find(l => /PORCA BAIXA/.test(l[2]) && l[3] === "M6");
      if (!porca) throw new Error("porca M6 não está na consolidada");
      if (!/90453A114.*;.*90453A114/.test(porca[1])) throw new Error("códigos das duas linhas somadas: " + porca[1]);
      if (porca[7] !== 40) throw new Error("qtd consolidada: " + porca[7]);
      if (!perto(porca[8], 0.003, 1e-6)) throw new Error("massa devia voltar por peça (0.003): " + porca[8]);
      const trat = r.rows.find(l => /ZINC FLAKE/.test(l[6]));
      if (!trat) throw new Error("Treating_1 sumiu");
      const chapa = r.rows.find(l => /^CHAPA-300x300/.test(l[1]));
      if (chapa[3] !== '1/2"') throw new Error("Dimensions devia sair como veio: " + chapa[3]);
      return true;
    });
  } else {
    falhas.push("BOM do SolidWorks: coloque o .xlsx em referencia/ para rodar este teste");
  }

  /* ---------- lista com valores já calculados ---------- */
  await passo("lista com valores unitários: lê comprimento, área e massa por peça", async () => {
    const r = JSON.parse(await rodar(`(() => {
      limparColunasDoEstudo(); lastGrouped = []; linhasDeConjunto = [];
      const t = [
        ["LISTA DE MATERIAL"],
        ["Item","Qtd.","Título","Especificação","Comprimento unitário (m)","Área unitária (m²)","Material","Massa unitária (kg)"],
        ["1","1","Chapa de reforço",'#3/4" (19 mm) x 500 x 500 mm',"-","0.089","ASTM A131 Gr AH36","13.2"],
        ["7","2","Chapa reforço",'3/4" (19 mm) x 559 x 465 mm',"-","0.117","ASTM A131 Gr AH36","17.3"],
        ["8","3","Perfil W",'W150 x 13,0',"2.45","-","ASTM A572 Gr50","31.85"],
        ["12","1","Eslinga içamento pilar 1","Eslinga 2 pernas (CMT 4,9 ton) com anelão e sapatilho : 1 perna com 2880 e 1 perna com 3100 mm.","-","-","Aço liga","12"]
      ];
      const lida = lerTabela(t, "Lista calculada.xlsx", "planilha");
      lastGrouped = ordenarLista(groupRows(lida.rows));
      const gs = lastGrouped.map(g => { const c = classifyGroup(g); return {esp:g.especificacao, qtd:g.qtd, massa:g.massa,
        tipo:c.tipo, ok:c.ok, L:c.length, area:areaM2DaLinha(g,c), comp:comprimentoMDaLinha(g,c)}; });
      const sh = buildPadraoOriginalSheet();
      return JSON.stringify({gs, h:sh.headers, rows:sh.rows});
    })()`));
    const [c1, c7, w, esl] = r.gs;
    if (!perto(c1.area, 0.089, 1e-6)) throw new Error("área da chapa 1 devia ser a do documento (0.089): " + c1.area);
    if (!perto(c7.area, 0.234, 1e-6)) throw new Error("chapa 7: 2 × 0.117 = 0.234 m², saiu " + c7.area);
    if (!perto(c7.massa, 34.6, 1e-6)) throw new Error("chapa 7: massa 2 × 17.3 = 34.6, saiu " + c7.massa);
    if (w.tipo !== "barra" || w.L !== 2450 || !perto(w.comp, 7.35, 1e-6)) throw new Error("perfil W: " + JSON.stringify(w));
    if (esl.tipo !== "outro" || esl.massa !== 12) throw new Error("eslinga: " + JSON.stringify(esl));
    if (r.h.length !== 8) throw new Error("colunas devolvidas: " + r.h.join(" | "));
    const l7 = r.rows.find(l => l[0] === "7");
    if (l7[5] !== 0.117 || l7[7] !== 17.3) throw new Error("linha 7 devia voltar por peça: " + JSON.stringify(l7));
    const l8 = r.rows.find(l => l[0] === "8");
    if (l8[4] !== 2.45 || l8[5] !== "-") throw new Error("linha 8: " + JSON.stringify(l8));
    return true;
  });

  await passo("colado do Excel com decimal brasileiro continua igual", async () => {
    const r = JSON.parse(await rodar(`(() => {
      limparColunasDoEstudo();
      const lida = parseRows("Item\\tQtd\\tEspecificação\\tDescrição\\tMaterial\\tMassa\\n1\\t2\\tChapa\\t#1/4\\" (6,35 mm) x 100 x 200 mm\\tASTM A36\\t1.264\\n2\\t1\\tTubo\\tØ48,30 SCH.40 x 1.200,5 mm\\tASTM A106\\t0,5", "colado");
      return JSON.stringify(lida.rows.map(x => ({m:x.massa, d:x.descricao})));
    })()`));
    if (r[0].m !== 1264 && r[0].m !== 1.264) throw new Error("massa: " + r[0].m);
    if (r[1].m !== 0.5) throw new Error("massa 0,5: " + r[1].m);
    return true;
  });

  /* ---------- item escrito à mão ---------- */
  await passo("texto escrito é separado em Título, Especificação, Norma e comprimento total", async () => {
    const it = JSON.parse(await rodar(`JSON.stringify(interpretarItemEscrito(
      "TUBO Ø1.1/2'' SCH. 40 ASME B 36.10\\nComprimento Total dos Tubos - 113,7 m"))`));
    if (it.tipo !== "Tubo") throw new Error("título: " + it.tipo);
    if (it.descricao !== 'Ø1.1/2" SCH. 40') throw new Error("especificação: " + it.descricao);
    if (it.norma !== "ASME B36.10") throw new Error("norma: " + it.norma);
    if (it.comprimento !== 113.7 || !it.comprimentoTotal) throw new Error("comprimento: " + it.comprimento + " total=" + it.comprimentoTotal);
    const um = JSON.parse(await rodar(`JSON.stringify(interpretarItemEscrito(
      '4 Chapa #1/4" (6,35 mm) x 100 x 200 mm AISI 316 massa unitária 1,1 kg'))`));
    if (um.qtd !== 4 || um.tipo !== "Chapa" || um.material !== "AISI 316" || um.massa !== 1.1 || !um.massaUnit)
      throw new Error("item numa linha: " + JSON.stringify(um));
    if (!/^#1\/4" \(6,35 mm\) x 100 x 200 mm$/.test(um.descricao)) throw new Error("especificação da chapa: " + um.descricao);
    const ret = JSON.parse(await rodar(`JSON.stringify(interpretarItemEscrito("240 x 120 x 6.35 mm Tubo Retangular"))`));
    if (ret.qtd != null) throw new Error("240 não é quantidade: " + JSON.stringify(ret));
    return true;
  });

  await passo("unidade escrita manda: 113.7 m é metro, 6000 mm é milímetro", async () => {
    const casos = JSON.parse(await rodar(`JSON.stringify([
      "TUBO Ø1.1/2'' SCH. 40 113.7 m",
      "TUBO Ø2'' SCH. 80 - 113,7 metros",
      "Tubo Ø1'' SCH. 40 comprimento total 11370 cm",
      "4 Tubo Ø1'' SCH. 40 comprimento 6000 mm",
      "4 Tubo Ø1'' SCH. 40 comprimento 6000",
      "Tubo Ø1'' SCH. 40 comprimento total 113,7",
      "Chapa #1/4\\" (6,35 mm) x 100 x 200 mm massa 1,2 t",
      "Parafuso M6 x 20 ISO 4017 massa unitária 5 g"
    ].map(interpretarItemEscrito).map(i => ({c:i.comprimento, t:i.comprimentoTotal, u:i.unidadeComprimento,
      m:i.massa, d:i.descricao, s:i.suposto})))`));
    const esp = [[113.7, true, "m"], [113.7, true, "m"], [113.7, true, "cm"], [6, false, "mm"], [6, false, "mm"], [113.7, true, "m"]];
    esp.forEach(([c, t, u], i) => {
      if (!perto(casos[i].c, c, 1e-9) || casos[i].t !== t || casos[i].u !== u)
        throw new Error(`caso ${i + 1}: ${JSON.stringify(casos[i])}`);
    });
    if (!casos[4].s.length || !casos[5].s.length) throw new Error("sem unidade devia avisar o que foi suposto");
    if (casos[0].s.length) throw new Error("com unidade não devia avisar: " + casos[0].s);
    if (/113/.test(casos[0].d)) throw new Error("o comprimento ficou na especificação: " + casos[0].d);
    if (casos[6].m !== 1200) throw new Error("1,2 t = 1200 kg, saiu " + casos[6].m);
    if (!perto(casos[7].m, 0.005, 1e-12)) throw new Error("5 g = 0,005 kg, saiu " + casos[7].m);
    if (!/6,35 mm/.test(casos[6].d)) throw new Error("mm da espessura não pode virar comprimento: " + casos[6].d);
    return true;
  });

  await passo("os dois campos de escrever têm exemplo com as palavras-chave e prévia", async () => {
    const r = JSON.parse(await rodar(`(() => {
      const ajuda = id => { const el = document.getElementById(id); return el && {
        exemplos: el.querySelectorAll(".whelp__ex").length, chaves: el.querySelectorAll(".whelp__chaves tbody tr").length }; };
      els.pasteArea.value = "TUBO Ø1.1/2'' SCH. 40 ASME B 36.10\\nComprimento Total dos Tubos - 113.7 m";
      els.pasteArea.dispatchEvent(new Event("input"));
      const previa = document.getElementById("pastePreview").textContent;
      document.querySelector("#pasteHelp [data-usar-exemplo='0']").click();
      const usado = els.pasteArea.value;
      els.pasteArea.value = ""; els.pasteArea.dispatchEvent(new Event("input"));
      return JSON.stringify({paste: ajuda("pasteHelp"), quick: ajuda("quickItemHelp"), previa, usado,
        vazia: document.getElementById("pastePreview").innerHTML});
    })()`));
    if (!r.paste || r.paste.exemplos < 1 || r.paste.chaves < 6) throw new Error("ajuda do campo de colar: " + JSON.stringify(r.paste));
    if (!r.quick || r.quick.exemplos < 1) throw new Error("ajuda do campo da seção 2: " + JSON.stringify(r.quick));
    if (!/Título\s*Tubo/.test(r.previa) || !/113,7 m/.test(r.previa) || !/ASME B36\.10/.test(r.previa))
      throw new Error("prévia: " + r.previa);
    if (!/^TUBO Ø1\.1\/2'' SCH\. 40 ASME B 36\.10\nComprimento Total/.test(r.usado)) throw new Error("usar exemplo: " + JSON.stringify(r.usado));
    if (r.vazia) throw new Error("prévia devia sumir com o campo vazio");
    return true;
  });

  const lm = fs.existsSync(PASTA_REF) ? fs.readdirSync(PASTA_REF).find(f => /^LM .*\.xlsx$/i.test(f)) : null;
  if (lm) {
    await passo("LM + tubo escrito somado: vira barras inteiras no corte", async () => {
      await rodar(`limparColunasDoEstudo(); lastGrouped = []; linhasDeConjunto = []; ordemGlobal = 0;`);
      await soltar(path.join(PASTA_REF, lm));
      const r = JSON.parse(await rodar(`(() => {
        const antes = lastGrouped.length, conjuntos = linhasDeConjunto.length;
        els.pasteArea.value = "TUBO Ø1.1/2'' SCH. 40 ASME B 36.10\\nComprimento Total dos Tubos - 113,7 m";
        els.addPasteBtn.click();
        const g = lastGrouped.find(x => /Ø1.1\\/2/.test(x.descricao));
        const c = g && classifyGroup(g);
        const barra = lastBarraResults.find(b => b && /Ø1.1\\/2/.test(b.identidade));
        const tr = lastGrouped.find(x => /240 x 120/.test(x.descricao));
        const sh = buildPadraoOriginalSheet();
        const linha = sh.rows.find(l => l.some(v => /Ø1.1\\/2/.test(String(v))));
        return JSON.stringify({antes, conjuntos, depois: lastGrouped.length, tipo: c && c.tipo, L: c && c.L, len: c && c.length,
          comp: g && comprimentoMDaLinha(g, c), barras: barra && barra.count, lengthM: barra && barra.lengthM,
          ret: tr && classifyGroup(tr), headers: sh.headers, linha, status: els.parseStatus.textContent});
      })()`));
      if (r.conjuntos !== 1) throw new Error("'Conjunto Soldado' devia ficar fora da soma: " + r.conjuntos);
      if (r.antes !== 9) throw new Error("itens da LM: " + r.antes);
      if (r.depois !== 10) throw new Error("depois de somar o tubo: " + r.depois);
      if (r.tipo !== "barra" || !perto(r.comp, 113.7, 1e-6)) throw new Error("tubo: " + JSON.stringify(r));
      // 9 barras de 12 m (11,997 m úteis com 3 mm de perda) + a sobra numa barra de 6 m
      if (r.barras !== 10 || !perto(r.lengthM, 113.7, 1e-6)) throw new Error("barras: " + r.barras + ", " + r.lengthM + " m");
      if (!r.ret.ok || !perto(r.ret.length, 965, 5) || r.ret.estimada !== "massa") throw new Error("tubo retangular: " + JSON.stringify(r.ret));
      if (!r.headers.includes("Norma") || !r.linha.includes("ASME B36.10")) throw new Error("norma na devolução: " + r.headers.join("|") + " / " + JSON.stringify(r.linha));
      if (!r.headers.includes("Massa [kg]")) throw new Error("colunas da LM sumiram: " + r.headers.join("|"));
      console.log(`     (${r.barras} barras · ${r.headers.join(" | ")})`);
      return true;
    });

    await passo("seção 2: item escrito numa linha entra separado nas colunas", async () => {
      const r = JSON.parse(await rodar(`(() => {
        const antes = lastGrouped.length;
        const inp = document.getElementById("quickItemInput");
        inp.value = '4 Chapa #1/4" (6,35 mm) x 100 x 200 mm AISI 316 massa unitária 1,1 kg';
        inp.dispatchEvent(new KeyboardEvent("keydown", {key:"Enter", bubbles:true}));
        const g = lastGrouped.find(x => /100 x 200/.test(x.descricao));
        return JSON.stringify({antes, depois: lastGrouped.length, g: g && {esp:g.especificacao, mat:g.material, qtd:g.qtd, massa:g.massa},
          vazio: inp.value, lido: document.getElementById("quickItemStatus").textContent});
      })()`));
      if (r.depois !== r.antes + 1) throw new Error("não entrou: " + r.antes + " → " + r.depois);
      if (!r.g || r.g.esp !== "Chapa" || r.g.mat !== "AISI 316" || r.g.qtd !== 4 || !perto(r.g.massa, 4.4, 1e-9))
        throw new Error("colunas: " + JSON.stringify(r.g));
      if (r.vazio !== "" || !/Título\s*Chapa/.test(r.lido)) throw new Error("retorno na tela: " + r.lido);
      return true;
    });
  }

  if (process.argv.includes("--shot")) {
    // abre a ajuda e escreve um item, para ver o exemplo e a prévia na tela
    await rodar(`(() => {
      document.querySelectorAll(".whelp").forEach(d => d.open = true);
      const q = document.getElementById("quickItemInput");
      if (q) { q.value = "4 TUBO Ø1.1/2'' SCH. 40 ASME B 36.10 ASTM A106 Gr.B comprimento 6000"; q.dispatchEvent(new Event("input")); }
      return true;
    })()`);
    await send("Emulation.setDeviceMetricsOverride", { width: 1100, height: 900, deviceScaleFactor: 1, mobile: false });
    for (const [id, nome] of [["pasteHelp", "p9_ajuda_colar.png"], ["quickItemHelp", "p9_ajuda_secao2.png"]]) {
      const c = JSON.parse(await rodar(`(() => { const el = document.getElementById(${JSON.stringify(id)}).closest("section");
        el.scrollIntoView(); const r = el.getBoundingClientRect();
        return JSON.stringify({x:r.left + scrollX, y:r.top + scrollY, w:r.width, h:r.height}); })()`));
      const r = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true,
        clip: { x: c.x, y: c.y, width: c.w, height: Math.min(c.h, 2400), scale: 1 } });
      fs.writeFileSync(path.join(os.tmpdir(), nome), Buffer.from(r.result.data, "base64"));
      console.log("     tela: " + path.join(os.tmpdir(), nome));
    }
  }
  ws.close(); nav.kill();
  console.log("ok:", ok.length);
  ok.forEach(x => console.log("  ✓", x));
  if (falhas.length) { console.log("FALHAS:"); falhas.forEach(x => console.log("  ✗", x)); }
  if (logs.length) { console.log("ERROS DE CONSOLE:"); logs.slice(0, 5).forEach(l => console.log("  !", l)); }
  process.exit(falhas.length || logs.length ? 1 : 0);
})();
