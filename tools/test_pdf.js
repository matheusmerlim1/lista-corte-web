/*
 * Teste da leitura de lista de material em PDF: gera PDFs de desenho de mentira (sem
 * nenhum documento de cliente), abre a página num Chrome/Edge sem janela e confere que a
 * ferramenta acha a tabela sozinha, onde quer que ela esteja na folha.
 *
 *  - A1.pdf: lista no meio da folha, cabeçalho em inglês EM CIMA, título centralizado e
 *    valores encostados à esquerda, especificação quebrada em duas linhas, notas ao lado e
 *    cotas do desenho alinhadas com a tabela por acaso.
 *  - B2.pdf: duas folhas; a lista fica na 2ª, no canto de baixo, em português, com o
 *    cabeçalho EMBAIXO (linhas crescendo para cima, como sobre o carimbo).
 *  - C3.pdf: desenho sem lista de material — tem que aparecer no relatório como "nenhuma".
 *
 * O pdf.js vem do CDN: precisa de internet.
 * Uso: node tools/test_pdf.js
 */
const { spawn } = require("child_process");
const fs = require("fs"), os = require("os"), path = require("path");

const RAIZ = path.join(__dirname, "..");
const NAVEGADOR = ["C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe"].find(p => fs.existsSync(p));
const esperar = ms => new Promise(r => setTimeout(r, ms));

/* ---------- PDF mínimo, com Helvetica e texto posicionado ---------- */
// textos: [{x, y, s, t}] com y de cima para baixo (como na folha) e t = tamanho da letra
function gerarPdf(folhas){
  const L = 1191, A = 842;                      // A3 deitada, em pontos
  const objs = [];
  const add = s => { objs.push(s); return objs.length; };
  const fonte = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
  const idsPag = [];
  const raizPaginas = objs.length + 1 + folhas.length * 2;   // reservado: vem depois das páginas
  for(const textos of folhas){
    const esc = s => s.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
    const cont = "BT\n" + textos.map(t => `/F1 ${t.t || 7} Tf 1 0 0 1 ${t.x.toFixed(2)} ${(A - t.y).toFixed(2)} Tm (${esc(t.s)}) Tj`).join("\n") + "\nET";
    const bytes = Buffer.from(cont, "latin1");
    const c = add(`<< /Length ${bytes.length} >>\nstream\n${bytes.toString("latin1")}\nendstream`);
    idsPag.push(add(`<< /Type /Page /Parent ${raizPaginas} 0 R /MediaBox [0 0 ${L} ${A}] /Resources << /Font << /F1 ${fonte} 0 R >> >> /Contents ${c} 0 R >>`));
  }
  add(`<< /Type /Pages /Kids [${idsPag.map(i => i + " 0 R").join(" ")}] /Count ${idsPag.length} >>`);
  const cat = add(`<< /Type /Catalog /Pages ${raizPaginas} 0 R >>`);
  let out = "%PDF-1.4\n";
  const pos = [];
  objs.forEach((o, i) => { pos.push(Buffer.byteLength(out, "latin1")); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(out, "latin1");
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + pos.map(p => String(p).padStart(10, "0") + " 00000 n \n").join("");
  out += `trailer\n<< /Size ${objs.length + 1} /Root ${cat} 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return Buffer.from(out, "latin1");
}
// largura aproximada da Helvetica (só para centralizar os títulos)
const largura = (s, t = 7) => [...s].reduce((w, c) => w + (/[A-ZÀ-Ý]/.test(c) ? 0.68 : /[a-z]/.test(c) ? 0.52 : /\d/.test(c) ? 0.556 : c === " " ? 0.278 : 0.4), 0) * t;
const centro = (cx, s, y, t = 7) => ({x: cx - largura(s, t) / 2, y, s, t});

function pdfA(){
  const t = [];
  // carimbo, notas e cotas espalhadas
  t.push({x: 900, y: 780, s: "FDOC-TESTE-001", t: 12}, {x: 820, y: 120, s: "NOTES", t: 9},
         {x: 820, y: 135, s: "1. Dimensions in mm;"}, {x: 820, y: 150, s: "2. Do not scale drawing;"});
  // tabela: colunas [borda esquerda, direita]
  const col = [[300, 330], [330, 360], [360, 470], [470, 650], [650, 730], [730, 790]];
  const cx = j => (col[j][0] + col[j][1]) / 2;
  const y0 = 300;
  t.push(centro((300 + 790) / 2, "BILL OF MATERIAL - SUPPORT A", y0 - 18, 8));
  ["Item", "Qty.", "Title", "Specification", "Material", "Mass [kg]"].forEach((h, j) => t.push(centro(cx(j), h, y0)));
  const linhas = [
    ["1", "1", "Support A", ["-"], "-", "12.4"],
    ["1.1", "2", "Plate", ["#1/4\" (6,35 mm) x 150 x 80 mm"], "ASTM A36", "1.2"],
    ["1.2", "2", "Square Tube", ["2\" x 2\" x 1/4\" x 1075 mm"], "ASTM A36", "9.2"],
    ["2", "8", "Stud bolt with nuts", ["Stud Bolt: A193 Gr. B7, Zn-Ni Coated,", "M20 x 100 mm; Nuts: A194 Gr. 2H"], "ASTM A193 Gr. B7", "0.4"],
    ["3", "4", "Hex Nut", ["ASME B18.2.2 - 5/8\"-11 UNC Grade 2"], "Carbon Steel", "0.1"],
  ];
  linhas.forEach((l, i) => {
    const y = y0 + 18 + i * 18;
    t.push(centro(cx(0), l[0], y), centro(cx(1), l[1], y), centro(cx(2), l[2], y));
    // especificação encostada à esquerda; em duas linhas fica centralizada na altura
    const n = l[3].length;
    l[3].forEach((s, k) => t.push({x: col[3][0] + 3, y: y + (k - (n - 1) / 2) * 8.5, s, t: 7}));
    t.push(centro(cx(4), l[4], y), centro(cx(5), l[5], y));
  });
  // cotas do desenho que caem por acaso embaixo da tabela, nas mesmas colunas
  t.push(centro(cx(0), "52", y0 + 18 * 5 + 60), centro(cx(1), "20", y0 + 18 * 5 + 60), centro(cx(2), "36", y0 + 18 * 5 + 60));
  return gerarPdf([t]);
}

function pdfB(){
  const folha1 = [{x: 100, y: 100, s: "VISTA GERAL", t: 10}, {x: 300, y: 400, s: "1250"}, {x: 500, y: 420, s: "A"}];
  const t = [{x: 950, y: 800, s: "FDOC-TESTE-002", t: 12}];
  // canto de baixo, à direita: cabeçalho embaixo e linhas para cima
  const col = [[760, 790], [790, 825], [825, 920], [920, 1060], [1060, 1120], [1120, 1170]];
  const cx = j => (col[j][0] + col[j][1]) / 2;
  const yCab = 740;
  ["ITEM", "QTD", "DESCRIÇÃO", "DIMENSÕES", "MATERIAL", "PESO (kg)"].forEach((h, j) => t.push(centro(cx(j), h, yCab)));
  t.push(centro((760 + 1170) / 2, "LISTA DE MATERIAL", yCab + 16, 8));
  const linhas = [   // de baixo para cima, como nos carimbos
    ["1", "2", "CHAPA", "#3/8\" (9,53 mm) x 180 x 180 mm", "ASTM A36", "4,8"],
    ["2", "6", "CANTONEIRA", "L 2\" x 2\" x 1/4\" x 750 mm", "ASTM A36", "13,5"],
    ["3", "12", "PARAFUSO SEXTAVADO", "M16 x 60 - ISO 4017", "AÇO INOX", "1,2"],
  ];
  linhas.forEach((l, i) => {
    const y = yCab - 16 - i * 16;
    l.forEach((v, j) => t.push(j === 3 ? {x: col[3][0] + 3, y, s: v, t: 7} : centro(cx(j), v, y)));
  });
  return gerarPdf([folha1, t]);
}

function pdfC(){
  return gerarPdf([[{x: 100, y: 100, s: "PLANTA BAIXA - SEM LISTA", t: 10}, {x: 400, y: 300, s: "Item 3 ver detalhe"}, {x: 400, y: 320, s: "2500"}]]);
}

(async () => {
  const perfil = fs.mkdtempSync(path.join(os.tmpdir(), "pdft-"));
  const nav = spawn(NAVEGADOR, ["--headless=new", "--disable-gpu", "--allow-file-access-from-files",
    "--remote-debugging-port=9344", `--user-data-dir=${perfil}`, "about:blank"], { stdio: "ignore" });
  let alvo = null;
  for (let i = 0; i < 40 && !alvo; i++) {
    await esperar(250);
    try { alvo = (await (await fetch("http://127.0.0.1:9344/json/list")).json()).find(t => t.type === "page"); } catch (e) { /* subindo */ }
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

  // põe os arquivos num <input type=file> e dispara o change, como a pessoa escolhendo
  async function escolher(inputId, arquivos){
    const lista = arquivos.map(([nome, buf, caminho]) => ({nome, b64: buf.toString("base64"), caminho}));
    return rodar(`(async () => {
      const dt = new DataTransfer();
      for(const a of ${JSON.stringify(lista)}){
        const bin = atob(a.b64); const u = new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++) u[i] = bin.charCodeAt(i);
        const f = new File([u], a.nome, {type: a.nome.endsWith(".pdf") ? "application/pdf" : "application/octet-stream"});
        if(a.caminho) f.caminho = a.caminho;
        dt.items.add(f);
      }
      const inp = document.getElementById(${JSON.stringify(inputId)});
      inp.files = dt.files;
      inp.dispatchEvent(new Event("change"));
      // espera a leitura terminar (o status sai de "Lendo…")
      for(let i=0;i<300;i++){
        await new Promise(r=>setTimeout(r,100));
        const s = document.getElementById("parseStatus").textContent;
        if(s && !/^Lendo/.test(s)) break;
      }
      return {
        status: document.getElementById("parseStatus").textContent,
        relatorio: [...document.querySelectorAll("#importReport li")].map(li => li.textContent),
      };
    })()`);
  }
  // linhas de origem de todos os grupos (o que foi lido de cada linha da lista)
  const origem = () => rodar(`JSON.stringify(lastGrouped.flatMap(g => g.origem.map(o => ({
    item:o.item, qtd:o.qtd, massa:o.massa, tipo:o.especificacao, desc:o.descricao, mat:o.material, fonte:o.fonte,
    folha:(o.extras||{})["extra:folha"]}))))`).then(JSON.parse);
  const conjuntos = () => rodar(`linhasDeConjunto.map(r => r.item + " " + r.especificacao)`);

  await passo("PDF com a lista no meio da folha (cabeçalho em cima)", async () => {
    const r = await escolher("xlsxInput", [["A1.pdf", pdfA()]]);
    const o = await origem();
    const it = k => o.find(x => x.item === k);
    if (o.length !== 4) throw new Error(`linhas lidas: ${o.length} (${o.map(x => x.item).join(",")}) — ${r.status}`);
    if (!it("1.1") || it("1.1").tipo !== "Plate" || !/6,35 mm\) x 150 x 80/.test(it("1.1").desc)) throw new Error("1.1: " + JSON.stringify(it("1.1")));
    if (it("1.2").qtd !== 2 || Math.abs(it("1.2").massa - 18.4) > 0.01) throw new Error("1.2 massa por peça × qtd: " + JSON.stringify(it("1.2")));
    if (!/A193 Gr. B7, Zn-Ni Coated, M20 x 100 mm; Nuts: A194 Gr.2H/.test(it("2").desc)) throw new Error("célula em duas linhas: " + it("2").desc);
    if (it("2").mat !== "ASTM A193 Gr. B7") throw new Error("material: " + it("2").mat);
    if (o.some(x => x.item === "52")) throw new Error("cota do desenho entrou como item");
    const cj = await conjuntos();
    if (!cj.some(c => /^1 Support A/.test(c))) throw new Error("linha de conjunto: " + cj.join(" | "));
    if (!/1 tabela/.test(r.relatorio.join(" "))) throw new Error("relatório: " + r.relatorio.join(" | "));
  });

  await passo("PDF com a lista no canto, em português, cabeçalho embaixo (folha 2)", async () => {
    const r = await escolher("xlsxAddInput", [["B2.pdf", pdfB()]]);
    const o = (await origem()).filter(x => x.fonte === "B2.pdf");
    if (o.length !== 3) throw new Error(`linhas lidas: ${o.length} — ${r.status} — ${r.relatorio.join(" | ")}`);
    const p = o.find(x => x.item === "3");
    if (!p || p.qtd !== 12 || p.tipo !== "PARAFUSO SEXTAVADO" || !/M16 x 60/.test(p.desc)) throw new Error("item 3: " + JSON.stringify(p));
    if (Math.abs(p.massa - 1.2) > 1e-6) throw new Error("PESO sem 'unitário' é o total da linha: " + p.massa);
    if (o.some(x => x.folha !== "2")) throw new Error("folha: " + o.map(x => x.folha).join(","));
    if (!/folha\(s\) 2 de 2/.test(r.relatorio.join(" "))) throw new Error("relatório: " + r.relatorio.join(" | "));
  });

  await passo("pasta com PDFs e um PDF sem lista: lê os que têm e avisa do que não tem", async () => {
    const r = await escolher("pastaInput", [["A1.pdf", pdfA(), "desenhos/A1.pdf"], ["C3.pdf", pdfC(), "desenhos/C3.pdf"],
      ["B2.pdf", pdfB(), "desenhos/sub/B2.pdf"], ["leiame.txt", Buffer.from("x"), "desenhos/leiame.txt"]]);
    const o = await origem();
    if (o.length !== 7) throw new Error(`linhas: ${o.length} — ${r.status}`);
    const rel = r.relatorio.join(" | ");
    if (!/C3\.pdf — nenhuma lista de material/.test(rel)) throw new Error("relatório sem o aviso do C3: " + rel);
    if (!/desenhos\/sub\/B2\.pdf/.test(rel)) throw new Error("relatório sem o caminho na pasta: " + rel);
    if (!/outro tipo/.test(r.status)) throw new Error("não avisou do .txt ignorado: " + r.status);
    // A1 e B2 chegam como listas separadas (dá pra remover uma sem a outra)
    const fontes = await rodar(`[...document.querySelectorAll("#sourcesList .source-chip b")].map(b=>b.textContent)`);
    if (fontes.join(",") !== "A1.pdf,B2.pdf") throw new Error("listas carregadas: " + fontes.join(","));
  });

  await passo("exportação no padrão do documento leva a coluna Folha", async () => {
    const cols = await rodar(`colunasDaLista().map(c => c.titulo)`);
    if (!cols.includes("Folha")) throw new Error("colunas: " + cols.join(", "));
  });

  if (logs.length) falhas.push("erros no console: " + logs.join(" | "));
  ws.close(); nav.kill();
  try { fs.rmSync(perfil, { recursive: true, force: true }); } catch (e) { /* o navegador ainda segura */ }
  console.log(ok.map(o => "✓ " + o).join("\n"));
  if (falhas.length) { console.log(falhas.map(f => "✗ " + f).join("\n")); process.exit(1); }
  console.log(`\n${ok.length} verificações ok.`);
})();
