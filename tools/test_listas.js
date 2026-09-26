/*
 * Teste da leitura e da consolidação: abre a página num Chrome sem janela, solta as listas
 * de referência (simulando o arrastar-e-soltar) e confere item por item.
 *
 * Uso: node tools/test_listas.js
 */
const { spawn } = require("child_process");
const fs = require("fs"), os = require("os"), path = require("path");

const RAIZ = path.join(__dirname, "..");
const CHROME = ["C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"].find(p => fs.existsSync(p));
const esperar = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const perfil = fs.mkdtempSync(path.join(os.tmpdir(), "p8t-"));
  const chrome = spawn(CHROME, ["--headless=new", "--disable-gpu", "--allow-file-access-from-files",
    "--remote-debugging-port=9342", `--user-data-dir=${perfil}`, "about:blank"], { stdio: "ignore" });
  let alvo = null;
  for (let i = 0; i < 40 && !alvo; i++) {
    await esperar(250);
    try { alvo = (await (await fetch("http://127.0.0.1:9342/json/list")).json()).find(t => t.type === "page"); } catch (e) { /* subindo */ }
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

  /** solta um arquivo na página, como se fosse arrastado do Windows */
  async function soltar(nomeArquivo) {
    const b64 = fs.readFileSync(path.join(RAIZ, "referencias", nomeArquivo)).toString("base64");
    await rodar(`(async () => {
      const bin = atob(${JSON.stringify(b64)});
      const buf = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      const arq = new File([buf], ${JSON.stringify(nomeArquivo)}, {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const dt = new DataTransfer();
      dt.items.add(arq);
      window.dispatchEvent(new DragEvent("dragenter", {dataTransfer: dt, bubbles:true}));
      window.dispatchEvent(new DragEvent("drop", {dataTransfer: dt, bubbles:true}));
      await new Promise(r => setTimeout(r, 900));
      return true;
    })()`);
    await esperar(600);
  }

  await passo("nada cobre a pagina quando ela abre", async () => {
    const r = JSON.parse(await rodar(`(() => {
      // quem esta no meio da tela? tem que ser o campo de colar ou o conteudo, nunca uma camada
      const x = innerWidth/2, y = innerHeight/2;
      const alvo = document.elementFromPoint(x, y);
      const caixa = document.getElementById("pasteBox");
      const aviso = caixa && caixa.querySelector(".paste-box__hint");
      return JSON.stringify({
        alvo: alvo ? (alvo.id || alvo.className || alvo.tagName) : "(nada)",
        avisoVisivel: !!(aviso && getComputedStyle(aviso).display !== "none"),
        cobrindo: [...document.body.querySelectorAll("*")].filter(el => {
          const e = getComputedStyle(el);
          if(e.position !== "fixed" || e.display === "none" || e.visibility === "hidden") return false;
          const c = el.getBoundingClientRect();
          return c.width > innerWidth*0.9 && c.height > innerHeight*0.9;
        }).map(el => el.id || el.className || el.tagName)
      });
    })()`));
    if (r.avisoVisivel) throw new Error("o aviso de soltar arquivo aparece sem estar arrastando nada");
    if (r.cobrindo.length) throw new Error("camada cobrindo a pagina: " + r.cobrindo.join(", "));
    return true;
  });

  await passo("arrastar realca o campo de colar, e sair desfaz", async () => {
    const estado = async evento => JSON.parse(await rodar(`(() => {
      const dt = new DataTransfer();
      dt.items.add(new File([new Uint8Array([1])], "x.xlsx"));
      window.dispatchEvent(new DragEvent(${JSON.stringify(evento)}, {dataTransfer: dt, bubbles:true}));
      const caixa = document.getElementById("pasteBox");
      const aviso = caixa.querySelector(".paste-box__hint");
      return JSON.stringify({realce: caixa.classList.contains("is-drag"),
                             visivel: getComputedStyle(aviso).display !== "none"});
    })()`));
    const entrou = await estado("dragenter");
    if (!entrou.realce || !entrou.visivel) throw new Error("arrastando o arquivo, o aviso nao apareceu no campo");
    const saiu = await estado("dragleave");
    if (saiu.realce || saiu.visivel) throw new Error("o aviso continuou depois de sair");
    return true;
  });

  await passo("soltar a primeira lista carrega os itens", async () => {
    await soltar("LISTA DE MATERIAL - DEMOLIÇÃO TUBULAÇÃO AREA 01.xlsx");
    const r = JSON.parse(await rodar(`JSON.stringify({
      itens: lastGrouped.length,
      descr: lastGrouped.map(g => g.especificacao),
      juntas: lastGrouped.filter(g => /junta/i.test(g.especificacao)).length
    })`));
    if (r.itens !== 5) throw new Error("itens lidos: " + r.itens + " (" + r.descr.join(", ") + ")");
    if (r.juntas !== 2) throw new Error("juntas na lista: " + r.juntas);
    return true;
  });

  await passo("a ordem é a mesma do documento", async () => {
    const primeiro = await rodar(`lastGrouped[0].especificacao`);
    if (!/parafuso/i.test(primeiro)) throw new Error("primeiro item: " + primeiro);
    return true;
  });

  await passo("soltar a segunda lista soma sem apagar a primeira", async () => {
    await soltar("LISTA DE MATERIAL - DEMOLIÇÃO TUBULAÇÃO AREA 02.xlsx");
    const r = JSON.parse(await rodar(`JSON.stringify({itens: lastGrouped.length, fontes: currentSources().map(s=>s.label)})`));
    if (r.itens !== 12) throw new Error("itens depois de somar: " + r.itens);
    if (r.fontes.length !== 2) throw new Error("listas carregadas: " + r.fontes.join(", "));
    return true;
  });

  await passo("linha de conjunto não entra como item", async () => {
    await soltar("LISTA DE MATERIAL - LINHA DE INCÊNDIO.xlsx");
    const r = JSON.parse(await rodar(`JSON.stringify({
      conjuntos: linhasDeConjunto.map(l => l.especificacao),
      naLista: lastGrouped.filter(g => /linha de inc/i.test(g.especificacao)).length
    })`));
    if (!r.conjuntos.some(c => /linha de inc/i.test(c))) throw new Error("conjunto não registrado: " + r.conjuntos.join(", "));
    if (r.naLista !== 0) throw new Error("a linha de conjunto entrou como item");
    return true;
  });

  await passo("colunas da lista final são a união das listas", async () => {
    const cols = JSON.parse(await rodar(`JSON.stringify(colunasDaLista().map(c => c.titulo))`));
    if (cols.length < 6) throw new Error("colunas: " + cols.join(" | "));
    const sh = JSON.parse(await rodar(`JSON.stringify(buildPadraoOriginalSheet())`));
    if (sh.headers.join("|") !== cols.join("|")) throw new Error("cabeçalho da planilha diferente do registro");
    if (!sh.rows.length) throw new Error("planilha sem linhas");
    return true;
  });

  await passo("lista com colunas diferentes mescla com traço", async () => {
    // tabela A: Descrição, Qtd, Massa   |   tabela B: Especificação, Qtd, Massa
    const r = JSON.parse(await rodar(`(() => {
      limparColunasDoEstudo();
      const a = lerTabela([["Descrição","Qtd","Massa"],["Chapa #1/4\\" x 100 x 100 mm","2","1,5"]], "A.xlsx", "planilha");
      const b = lerTabela([["Especificação","Qtd","Massa"],["Junta de Vedação","3","0"]], "B.xlsx", "planilha");
      lastGrouped = ordenarLista(groupRows([...a.rows, ...b.rows]));
      const sh = buildPadraoOriginalSheet();
      return JSON.stringify({headers: sh.headers, rows: sh.rows});
    })()`));
    const cab = r.headers.join(" | ");
    if (!/Descrição/.test(cab) || !/Especificação/.test(cab)) throw new Error("faltou coluna na união: " + cab);
    const comTraco = r.rows.filter(l => l.includes("-")).length;
    if (comTraco < 2) throw new Error("esperava traço nas colunas que cada lista não tinha: " + JSON.stringify(r.rows));
    return true;
  });

  await passo("padrão do documento sai separado, uma aba por lista", async () => {
    // recarrega as três listas de referência
    await rodar(`limparColunasDoEstudo(); lastGrouped = []; linhasDeConjunto = [];`);
    await soltar("LISTA DE MATERIAL - DEMOLIÇÃO TUBULAÇÃO AREA 01.xlsx");
    await soltar("LISTA DE MATERIAL - SUPORTE 1.xlsx");
    const r = JSON.parse(await rodar(`JSON.stringify(buildPadraoOriginalSheets().map(a => ({
      nome: a.name, cols: a.headers, linhas: a.rows.length })))`));
    if (r.length !== 2) throw new Error("abas geradas: " + r.map(a=>a.nome).join(", "));
    if (!r.every(a => a.linhas > 0)) throw new Error("aba vazia: " + JSON.stringify(r));
    if (!r[0].cols.join("|").includes("Título")) throw new Error("colunas da aba 1: " + r[0].cols.join(" | "));
    return true;
  });

  await passo("consolidada sai tudo numa aba só", async () => {
    const r = JSON.parse(await rodar(`JSON.stringify(buildConsolidadoPages(
      buildConsolidadoCortavelRows(), buildConsolidadoOutroRows(), {mode:"unica", size:14})
      .map(p => ({nome:p.name, linhas:p.rows.length})))`));
    if (r.length !== 1) throw new Error("abas da consolidada: " + r.map(x=>x.nome).join(", "));
    if (!r[0].linhas) throw new Error("consolidada sem linhas");
    return true;
  });

  await passo("o Excel gerado abre e tem as abas certas", async () => {
    const b64 = await rodar(`(async () => {
      const blob = buildXlsxWorkbookBlob(buildPadraoOriginalSheets());
      const buf = new Uint8Array(await blob.arrayBuffer());
      let s = ""; for (let i = 0; i < buf.length; i++) s += String.fromCharCode(buf[i]);
      return btoa(s);
    })()`);
    require("fs").writeFileSync(require("path").join(require("os").tmpdir(), "p8_saida.xlsx"), Buffer.from(b64, "base64"));
    return true;
  });

  await passo("as sete listas de referência entram inteiras", async () => {
    await rodar(`limparColunasDoEstudo(); lastGrouped = []; linhasDeConjunto = []; ordemGlobal = 0;`);
    const arquivos = fs.readdirSync(path.join(RAIZ, "referencias")).filter(f => /\.xlsx$/i.test(f));
    for (const nome of arquivos) await soltar(nome);
    const r = JSON.parse(await rodar(`JSON.stringify({
      fontes: currentSources().length,
      itens: lastGrouped.length,
      linhas: registrosOriginais().length,
      conjuntos: linhasDeConjunto.length,
      abas: buildPadraoOriginalSheets().length
    })`));
    if (r.fontes !== arquivos.length) throw new Error(`listas carregadas: ${r.fontes} de ${arquivos.length}`);
    if (r.abas !== arquivos.length) throw new Error(`abas no padrão do documento: ${r.abas}`);
    // 53 linhas de dados nas sete planilhas, menos 7 linhas de conjunto
    if (r.linhas + r.conjuntos < 40) throw new Error(`linhas lidas: ${r.linhas} + ${r.conjuntos} conjunto(s)`);
    console.log(`     (${r.fontes} listas · ${r.linhas} linhas · ${r.itens} itens agrupados · ${r.conjuntos} conjuntos)`);
    return true;
  });

  await passo("exportações geram o arquivo nos dois modos", async () => {
    const r = JSON.parse(await rodar(`(async () => {
      const sep = buildPadraoOriginalSheets();
      const junto = [buildPadraoOriginalSheet()];
      const cons = buildConsolidadoPages(buildConsolidadoCortavelRows(), buildConsolidadoOutroRows(), {mode:"unica", size:14});
      const b64 = async blob => {
        const buf = new Uint8Array(await blob.arrayBuffer());
        let t = ""; for (let i = 0; i < buf.length; i++) t += String.fromCharCode(buf[i]);
        return btoa(t);
      };
      const b1 = buildXlsxWorkbookBlob(sep), b2 = buildXlsxWorkbookBlob(junto);
      return JSON.stringify({abasSep: sep.length, colsJunto: junto[0].headers.length,
        linhasJunto: junto[0].rows.length, abasCons: cons.length, bytes: [b1.size, b2.size],
        arquivos: {separado: await b64(b1), junto: await b64(b2)}});
    })()`));
    if (r.abasSep < 7) throw new Error("abas separadas: " + r.abasSep);
    if (r.abasCons !== 1) throw new Error("consolidada em " + r.abasCons + " abas");
    if (!r.bytes.every(b => b > 2000)) throw new Error("arquivo pequeno demais: " + r.bytes.join(", "));
    for (const [nome, dados] of Object.entries(r.arquivos))
      fs.writeFileSync(path.join(os.tmpdir(), `p8_sete_${nome}.xlsx`), Buffer.from(dados, "base64"));
    console.log(`     (separado: ${r.abasSep} abas · junto: ${r.colsJunto} colunas, ${r.linhasJunto} linhas · consolidada: ${r.abasCons} aba)`);
    return true;
  });

  /** o que a previa desenhou na tela, aba por aba */
  const previaNaTela = async (conteudo, formato, modoAbas) => JSON.parse(await rodar(`(() => {
    els.exportContentSel.value = ${JSON.stringify(conteudo)};
    els.exportFormatSel.value = ${JSON.stringify(formato)};
    if (${JSON.stringify(modoAbas)} && els.pageModeSel) els.pageModeSel.value = ${JSON.stringify(modoAbas)};
    els.exportFormatSel.dispatchEvent(new Event("change", {bubbles:true}));
    els.previewBtn.click();
    return JSON.stringify([...document.querySelectorAll("#previewBody .preview-page")].map(p => ({
      titulo: p.querySelector("h4").textContent.trim(),
      colunas: [...p.querySelectorAll("thead th")].map(t => t.textContent.trim()),
      linhas: p.querySelectorAll("tbody tr").length
    })));
  })()`));

  await passo("prévia mostra tudo numa aba só quando é isso que foi pedido", async () => {
    const abas = await previaNaTela("consolidado", "lista", "unica");
    if (abas.length !== 1) throw new Error("a prévia mostrou " + abas.length + " blocos: " +
      abas.map(a => a.titulo).join(" / "));
    if (abas[0].colunas.length !== 6) throw new Error("colunas na prévia: " + abas[0].colunas.join(" | "));
    const doPlano = JSON.parse(await rodar(`JSON.stringify(planoDeExportacao("consolidado","lista")
      .map(a => ({nome:a.name, linhas:a.rows.length})))`));
    if (doPlano.length !== 1 || doPlano[0].linhas !== abas[0].linhas)
      throw new Error("a prévia não bate com o plano de exportação");
    console.log(`     (1 aba "${doPlano[0].nome}" · ${abas[0].linhas} linhas · ${abas[0].colunas.join(" | ")})`);
    return true;
  });

  await passo("pedindo separado, a prévia separa", async () => {
    const abas = await previaNaTela("consolidado", "lista", "auto");
    if (abas.length < 2) throw new Error("a prévia continuou com " + abas.length + " bloco(s)");
    console.log("     (" + abas.map(a => a.titulo.split("·")[1].trim() + ": " + a.linhas).join(" · ") + ")");
    return true;
  });

  await passo("prévia do padrão do documento traz as colunas de cada lista", async () => {
    const abas = await previaNaTela("consolidado", "original", null);
    if (abas.length !== 7) throw new Error("abas na prévia: " + abas.length);
    const seis = abas.every(a => a.colunas.length === 6);
    if (!seis) throw new Error("colunas diferentes do documento: " + abas.map(a=>a.colunas.length).join(","));
    if (!abas.some(a => a.titulo.includes("SUPORTE 1"))) throw new Error("nomes de aba não vieram: " +
      abas.map(a=>a.titulo).join(" / "));
    return true;
  });

  await passo("prévia bate com o arquivo gerado", async () => {
    await rodar(`els.exportContentSel.value="consolidado"; els.exportFormatSel.value="lista";
                 if(els.pageModeSel) els.pageModeSel.value="unica";`);
    const r = JSON.parse(await rodar(`(async () => {
      const plano = planoDeExportacao("consolidado", "lista");
      const blob = await buildListaXlsxFromTemplate(plano.map(a => ({name:a.name, rows:a.fonte})));
      const buf = new Uint8Array(await blob.arrayBuffer());
      let t = ""; for (let i = 0; i < buf.length; i++) t += String.fromCharCode(buf[i]);
      return JSON.stringify({abas: plano.map(a => ({nome:a.name, linhas:a.rows.length})), b64: btoa(t)});
    })()`));
    fs.writeFileSync(path.join(os.tmpdir(), "p8_previa.xlsx"), Buffer.from(r.b64, "base64"));
    if (r.abas.length !== 1) throw new Error("plano com " + r.abas.length + " abas");
    console.log(`     (arquivo salvo em ${path.join(os.tmpdir(), "p8_previa.xlsx")} — 1 aba, ${r.abas[0].linhas} linhas)`);
    return true;
  });

  if (process.argv.includes("--shot")) {
    const r = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true });
    fs.writeFileSync(path.join(os.tmpdir(), "p8_tela.png"), Buffer.from(r.result.data, "base64"));
  }
  ws.close(); chrome.kill();
  console.log("ok:", ok.length);
  ok.forEach(x => console.log("  ✓", x));
  if (falhas.length) { console.log("FALHAS:"); falhas.forEach(x => console.log("  ✗", x)); }
  if (logs.length) { console.log("ERROS DE CONSOLE:"); logs.slice(0, 5).forEach(l => console.log("  !", l)); }
  process.exit(falhas.length || logs.length ? 1 : 0);
})();
