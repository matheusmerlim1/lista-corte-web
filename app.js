/*
 * Lista & Corte — lógica da aplicação.
 *
 * Tudo num arquivo só (sem módulos/bundler) pra rodar direto no navegador sem servidor nem
 * build — basta abrir index.html. Seções, em ordem:
 *   1. Tema claro/escuro
 *   2. Utilidades numéricas + parsing da colagem/Excel (linhas → grupos)
 *   3. Interpretação da descrição (chapa: espessura/dimensões; barra: bitola/comprimento)
 *   4. Encaixe 2D (chapas) e 1D (barras/tubos) — os dois algoritmos de corte
 *   5. Modelos .xlsx embutidos em base64 + leitura/escrita de .xlsx na mão (sem libs externas)
 *   6. Geradores de planilha (Excel simples / Lista Preliminar / Lista consolidada) e paginação
 *   7. Wiring da interface: tabela editável de itens agrupados, seções de otimização de
 *      corte (só aparecem para o conteúdo "Resumo de corte"), pré-visualização e exportação
 *
 * Ver README.md para o funcionamento do ponto de vista do usuário.
 */

/* ---------------- tema claro/escuro ---------------- */
const themeBtn = document.getElementById("themeBtn");
const themeIcon = document.getElementById("themeIcon");
const themeText = document.getElementById("themeText");
function readSavedTheme(){ try{ return localStorage.getItem("listaCorteTheme"); }catch(e){ return null; } }
function saveTheme(v){ try{ localStorage.setItem("listaCorteTheme", v); }catch(e){} }
function applyTheme(theme){
  document.documentElement.setAttribute("data-theme", theme);
  if(theme==="dark"){ themeIcon.textContent="☾"; themeText.textContent="Tela escura"; }
  else { themeIcon.textContent="☀"; themeText.textContent="Tela clara"; }
}
applyTheme(readSavedTheme() || "light");
themeBtn.addEventListener("click", ()=>{
  const cur = document.documentElement.getAttribute("data-theme")==="dark" ? "dark" : "light";
  const next = cur==="dark" ? "light" : "dark";
  applyTheme(next); saveTheme(next);
});

/* ---------------- utilidades numéricas ---------------- */
function fmt(v, d=2){
  if(v==null || isNaN(v)) return "—";
  return Number(v).toLocaleString("pt-BR", {minimumFractionDigits:d, maximumFractionDigits:d});
}
function toNumBR(s){
  if(s==null) return NaN;
  s = String(s).trim();
  if(/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) s = s.replace(/\./g,"").replace(",", ".");
  else s = s.replace(",", ".");
  return parseFloat(s);
}
function toNumFlex(s){ return parseFloat(String(s).replace(",", ".")); }

// ordena uma lista de linhas (qualquer objeto com `.descricao`) alfabeticamente pela
// descrição — usado nas listas consolidada e compactada (o resumo de corte continua
// ordenado pelo plano de corte, que é mais útil ali do que ordem alfabética).
function sortByDescricao(rows){
  return [...rows].sort((a,b)=> String(a.descricao||"").localeCompare(String(b.descricao||""), "pt-BR", {sensitivity:"base"}));
}

/* ---------------- parsing da colagem (Excel) ---------------- */
// linhas "vazias" (sem especificação real) ou marcadas como esboço/placeholder são
// descartadas silenciosamente — não são um erro, são estrutura da planilha (cabeçalhos de
// seção, itens ainda não definidos etc.)
function isJunkRow(especificacao){
  const esp = String(especificacao||"").trim();
  if(!esp) return true;
  if(/^esbo[cç]o\s+remover$/i.test(esp)) return true;
  return false;
}
// colapsa espaços internos duplicados, tabs perdidos e espaços não separáveis (NBSP) —
// sem isso, duas linhas que parecem idênticas na tela (o HTML colapsa espaços ao exibir)
// podem ter texto diferente por baixo e não se agrupar (ex: "ASTM A36" com espaço duplo).
function normalizeText(s){
  return String(s||"")
    .replace(/\s+/g, " ")
    // "SCH. 80", "Gr. 50" etc. viram "SCH.80", "Gr.50" — mesma abreviação digitada com ou
    // sem espaço depois do ponto não pode virar grupo diferente na hora de agrupar peças
    .replace(/\.\s+(?=\d)/g, ".")
    // "SCH 80" (sem ponto nenhum) é a mesma abreviação que "SCH.80" — sem isso, a falta do
    // ponto (não só o espaço depois dele) também vira grupo diferente na hora de agrupar.
    .replace(/\b(SCH|Gr)\.?\s*(?=\d)/gi, "$1.")
    .trim();
}
// material às vezes vem com um "Aço" solto no final ("ASTM A36 Aço") em parte das linhas e
// sem em outras ("ASTM A36") — mesmo material, só um qualificador redundante ("aço" = "steel")
// digitado a mais; sem isso as duas viram grupos de corte separados só por essa diferença.
// Só remove quando "Aço" está sozinho no final — "Aço carbono"/"Aço inox" (material dos
// parafusos/arruelas) não é afetado, porque "Aço" ali não é a última palavra.
function normalizeMaterial(s){
  return normalizeText(s).replace(/\s+Aço$/i, "").trim();
}
function buildRow(item, qtd, especificacao, descricao, material, massa){
  return {item, qtd, especificacao:normalizeText(especificacao), descricao:normalizeText(descricao), material:normalizeMaterial(material), massa};
}

function parseRows(text){
  const lines = text.split(/\r\n|\r|\n/).filter(l=>l.trim().length>0);
  const parsed = [];
  const problems = [];
  let startIdx = 0;
  if(lines.length){
    const firstCols = lines[0].split("\t");
    if(firstCols.length>=6){
      const qtdTest = toNumBR(firstCols[1]);
      const massaTest = toNumBR(firstCols[5]);
      if(isNaN(qtdTest) || isNaN(massaTest)) startIdx = 1;
    }
  }
  for(let i=startIdx;i<lines.length;i++){
    const cols = lines[i].split("\t").map(c=>c.trim());
    if(cols.length<6){ problems.push(`Linha ${i+1}: esperadas 6 colunas (Item, Qtd, Especificação, Descrição, Material, Massa), encontradas ${cols.length} — ignorada.`); continue; }
    const [item, qtdStr, especificacao, descricao, material, massaStr] = cols;
    if(isJunkRow(especificacao)) continue;
    const qtd = toNumBR(qtdStr);
    const massa = toNumBR(massaStr);
    if(isNaN(qtd) || isNaN(massa)){ problems.push(`Linha ${i+1}: Qtd ou Massa não numérico — ignorada.`); continue; }
    parsed.push(buildRow(item, qtd, especificacao, descricao, material, massa));
  }
  return {rows:parsed, problems};
}

function groupRows(rows){
  const map = new Map();
  for(const r of rows){
    const key = [r.especificacao, r.descricao, r.material].join("||");
    if(!map.has(key)) map.set(key, {especificacao:r.especificacao, descricao:r.descricao, material:r.material, qtd:0, massa:0});
    const g = map.get(key);
    g.qtd += r.qtd;
    g.massa += r.massa;
  }
  return [...map.values()];
}

/* ---------------- interpretação da descrição (espessura / dimensões / diâmetro) ---------------- */
function parseDescricao(descricao){
  let s = " " + descricao + " ";
  let thickness = null;
  const mThick = s.match(/\(\s*([\d]+(?:[.,]\d+)?)\s*mm\s*\)/i);
  if(mThick){ thickness = toNumFlex(mThick[1]); s = s.replace(mThick[0], " "); }

  // espessura às vezes vem digitada direto em mm decimal depois do "#", sem fração nem
  // parênteses — "#37,43 mm x 111 x 114 mm" (chapa de sobra/retalho, sem bitola comercial
  // padrão). Sem isso, os 3 números (espessura + 2 dimensões) ficam todos em `extras` e a
  // espessura entra no lugar da largura, além do item nunca fechar "ok" (thickness nulo).
  if(thickness==null){
    const mThickDecimal = s.match(/#\s*([\d]+(?:[.,]\d+)?)\s*mm\b/i);
    if(mThickDecimal){ thickness = toNumFlex(mThickDecimal[1]); s = s.replace(mThickDecimal[0], " "); }
  }

  // diâmetro de chapa circular às vezes vem em mm decimal (Ø482) e às vezes em fração de
  // polegada (Ø1/2") — igual a espessura, só que sem parênteses com o valor em mm do lado.
  let diametro = null;
  const mDiam = s.match(/Ø\s*(?:(\d+)\s*\/\s*(\d+)\s*"|([\d]+(?:[.,]\d+)?))/i);
  if(mDiam){
    diametro = (mDiam[1] && mDiam[2]) ? Math.round((parseInt(mDiam[1],10)/parseInt(mDiam[2],10))*25.4*100)/100 : toNumFlex(mDiam[3]);
    s = s.replace(mDiam[0], " ");
  }

  s = s.replace(/#?\d+\/\d+"/g, " ");

  let schedule = null;
  const mSch = s.match(/SCH\.?\s*(\d+)/i);
  if(mSch){ schedule = mSch[1]; s = s.replace(mSch[0], " "); }

  const nums = [...s.matchAll(/([\d]+(?:[.,]\d+)?)\s*(?:mm)?/gi)].map(m=>toNumFlex(m[1]));

  return {thickness, diametro, schedule, extras:nums};
}

// para barras/tubos/perfis: o último número da descrição original é o comprimento de corte;
// tudo antes dele é a "identidade" da seção transversal (usada para agrupar o que pode ser
// cortado de uma mesma barra comercial). Funciona para "W460 x 74,0 x 3975mm",
// "Ø114,30 SCH.40 x 329,84 mm", "Cantoneira ... 50,80 x 9,53 x 238,96 mm" etc.
function parseBarraDescricao(descricao){
  const re = /([\d]+(?:[.,]\d+)?)\s*(?:mm)?/gi;
  let m, last = null;
  while((m = re.exec(descricao))) last = m;
  if(!last) return {length:null, identidade:descricao.trim()};
  const length = toNumFlex(last[1]);
  let prefix = descricao.slice(0, last.index);
  prefix = prefix.replace(/x\s*$/i, "").trim();
  return {length, identidade: prefix || descricao.trim()};
}

// itens cuja Especificação contenha uma destas palavras são tratados como barra/perfil cortável;
// os demais (parafusos, arruelas, curvas/joelhos, grades, itens sem palavra reconhecida) entram
// como "outro" — aparecem na tabela agrupada, mas não passam pela otimização de corte.
const BARRA_KEYWORDS = ["tubo", "perfil", "cantoneira", "barra", "viga", "vergalh", "cordoalha", "trilho", "eletroduto"];

// espessura de chapa costuma vir como fração comercial em polegada — "#3/4"", "#1/2"" etc. —
// seguida do valor em mm entre parênteses, mas esse mm é digitado à mão e varia de arredondamento
// (19 mm numa linha, 19,05 mm noutra, mesma chapa de 3/4"). Quando há fração, usamos o valor
// exato dela (polegada × 25,4) como espessura "canônica" para agrupar — assim chapas do mesmo
//3/4" não viram grupos de corte separados só por causa do arredondamento de quem digitou.
// (?<!Ø\s{0,3}) evita pegar a fração do DIÂMETRO (ex: "Ø1/2\"") quando a chapa é circular
// e o diâmetro também vem em fração — essa fração é outra coisa, não a espessura.
const FRACTION_THICKNESS_RE = /(?<!Ø\s{0,3})#?\s*(\d+)\s*\/\s*(\d+)\s*"/;
function extractFractionMm(descricao){
  const m = descricao.match(FRACTION_THICKNESS_RE);
  if(!m) return null;
  const den = parseInt(m[2],10);
  if(!den) return null;
  return Math.round((parseInt(m[1],10)/den)*25.4*100)/100;
}
// texto original da fração em polegada (ex: "3/4\""), pra mostrar junto da espessura em mm
function extractFractionLabel(descricao){
  const m = descricao.match(FRACTION_THICKNESS_RE);
  return m ? `${m[1]}/${m[2]}"` : null;
}

function classifyGroup(g){
  const specLower = g.especificacao.toLowerCase();

  if(specLower.includes("chapa")){
    const parsed = parseDescricao(g.descricao);
    const fracMm = extractFractionMm(g.descricao);
    const fracLabel = extractFractionLabel(g.descricao);
    const thickness = fracMm!=null ? fracMm : parsed.thickness;
    let width=null, height=null, circular=false, diamChapa=null;
    if(parsed.extras.length>=2){ width=parsed.extras[0]; height=parsed.extras[1]; }
    else if(parsed.diametro!=null){ circular=true; diamChapa=parsed.diametro; }
    const ok = Boolean(thickness!=null && ((!circular && width && height) || (circular && diamChapa)));
    return {tipo:"chapa", thickness, width, height, circular, diametro:diamChapa, ok, fracLabel};
  }

  const isBarra = BARRA_KEYWORDS.some(k=>specLower.includes(k));
  if(isBarra){
    const {length, identidade} = parseBarraDescricao(g.descricao);
    const ok = length!=null;
    return {tipo:"barra", length, identidade, ok};
  }

  return {tipo:"outro"};
}
// chapa/barra cuja Descrição não bateu com o padrão esperado (classifyGroup ok:false) fica
// de fora de lastChapaResults/lastBarraResults (que só existem pras peças que a otimização
// de corte conseguiu ler) — sem isso, o item simplesmente some das planilhas exportadas,
// aparecendo só no painel "não interpretadas" da página. Aqui ele entra junto com os itens
// soltos ("Outro"): sem Área/Comprimento (não dá pra calcular sem entender a descrição), mas
// com a quantidade certa, pra continuar aparecendo na lista de compra.
function isLooseOrUnclassified(g){
  const cls = classifyGroup(g);
  return cls.tipo==="outro" || !cls.ok;
}

/* ---------------- encaixe 2D (chapas) — heurístico por prateleiras, com rotação ---------------- */
/* ---------------- modelos .xlsx embutidos (base64) ---------------- */
// modelo do formato "Lista Preliminar": só a moldura da planilha — título, cabeçalho e
// estilos. As linhas de material do arquivo de origem foram removidas (o gerador nunca as
// usou: ele reescreve tudo a partir da linha 3), junto com o histórico de revisão e o nome
// de quem salvou o arquivo.
const TEMPLATE_LISTA_B64 = "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";
// modelo da lista consolidada: clone de "LM Consolidada.xlsx" com a aba LLI vazia
// (só o cabeçalho + a linha 2 guardando o estilo de cada coluna). A capa vem do próprio
// arquivo de referência — sem logotipo, slogan, nome de projeto ou de empresa.
const TEMPLATE_CONSOLIDADA_B64 = "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";
/* ---------------- leitura de arquivo .xlsx (sem bibliotecas externas) ---------------- */
// .xlsx é um ZIP contendo XML. Lemos o ZIP na mão (End Of Central Directory + Central
// Directory) e usamos a API nativa DecompressionStream("deflate-raw") do navegador para
// descomprimir cada parte — não precisa de nenhuma biblioteca de terceiros.
async function readZipEntries(arrayBuffer){
  const bytes = new Uint8Array(arrayBuffer);
  const dv = new DataView(arrayBuffer);
  let eocdOffset = -1;
  const minStart = Math.max(0, bytes.length - 22 - 65557);
  for(let i = bytes.length - 22; i >= minStart; i--){
    if(dv.getUint32(i, true) === 0x06054b50){ eocdOffset = i; break; }
  }
  if(eocdOffset < 0) throw new Error("Não parece ser um arquivo .xlsx válido (ZIP não encontrado).");

  const totalEntries = dv.getUint16(eocdOffset+10, true);
  const cdOffset = dv.getUint32(eocdOffset+16, true);

  const centralEntries = [];
  let p = cdOffset;
  for(let i=0; i<totalEntries; i++){
    if(dv.getUint32(p, true) !== 0x02014b50) break;
    const compMethod = dv.getUint16(p+10, true);
    const compSize = dv.getUint32(p+20, true);
    const nameLen = dv.getUint16(p+28, true);
    const extraLen = dv.getUint16(p+30, true);
    const commentLen = dv.getUint16(p+32, true);
    const localHeaderOffset = dv.getUint32(p+42, true);
    const name = new TextDecoder("utf-8").decode(bytes.slice(p+46, p+46+nameLen));
    centralEntries.push({name, compMethod, compSize, localHeaderOffset});
    p += 46 + nameLen + extraLen + commentLen;
  }

  const result = {};
  for(const e of centralEntries){
    const lp = e.localHeaderOffset;
    if(dv.getUint32(lp, true) !== 0x04034b50) continue;
    const lNameLen = dv.getUint16(lp+26, true);
    const lExtraLen = dv.getUint16(lp+28, true);
    const dataStart = lp + 30 + lNameLen + lExtraLen;
    const raw = bytes.slice(dataStart, dataStart + e.compSize);
    let data;
    if(e.compMethod === 0){
      data = raw;
    } else if(e.compMethod === 8){
      const ds = new DecompressionStream("deflate-raw");
      const stream = new Blob([raw]).stream().pipeThrough(ds);
      data = new Uint8Array(await new Response(stream).arrayBuffer());
    } else {
      continue;
    }
    result[e.name] = data;
  }
  return result;
}

function parseSharedStrings(xmlBytes){
  if(!xmlBytes) return [];
  const xml = new TextDecoder("utf-8").decode(xmlBytes);
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return [...doc.getElementsByTagName("si")].map(si=>{
    const tNodes = [...si.getElementsByTagName("t")];
    return tNodes.map(t=>t.textContent).join("");
  });
}

function colLettersToIndex(letters){
  let n = 0;
  for(const ch of letters) n = n*26 + (ch.charCodeAt(0)-64);
  return n-1;
}

function parseWorksheetTable(xmlBytes, sharedStrings){
  const xml = new TextDecoder("utf-8").decode(xmlBytes);
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const table = [];
  for(const row of doc.getElementsByTagName("row")){
    const rIdx = parseInt(row.getAttribute("r"), 10);
    const rowData = [];
    for(const c of row.getElementsByTagName("c")){
      const ref = c.getAttribute("r") || "";
      const m = ref.match(/^[A-Z]+/);
      if(!m) continue;
      const colIdx = colLettersToIndex(m[0]);
      const type = c.getAttribute("t");
      const vNode = c.getElementsByTagName("v")[0];
      const isNode = c.getElementsByTagName("is")[0];
      let value = null;
      if(type === "s" && vNode) value = sharedStrings[parseInt(vNode.textContent,10)] ?? "";
      else if(type === "inlineStr" && isNode) value = isNode.textContent;
      else if(vNode) value = vNode.textContent;
      rowData[colIdx] = value;
    }
    if(!isNaN(rIdx)) table[rIdx-1] = rowData;
  }
  return table;
}

function findFirstSheetPath(zipData){
  try{
    const wbXml = new TextDecoder("utf-8").decode(zipData["xl/workbook.xml"]);
    const wbDoc = new DOMParser().parseFromString(wbXml, "application/xml");
    const sheetEl = wbDoc.getElementsByTagName("sheet")[0];
    let rId = sheetEl.getAttribute("r:id");
    if(!rId){
      for(const attr of sheetEl.attributes){ if(attr.name.endsWith(":id")){ rId = attr.value; break; } }
    }
    const relsXml = new TextDecoder("utf-8").decode(zipData["xl/_rels/workbook.xml.rels"]);
    const relsDoc = new DOMParser().parseFromString(relsXml, "application/xml");
    const rel = [...relsDoc.getElementsByTagName("Relationship")].find(r=>r.getAttribute("Id")===rId);
    let target = rel.getAttribute("Target");
    target = target.replace(/^\//, "");
    if(!target.startsWith("xl/")) target = "xl/" + target;
    return target;
  }catch(e){
    return null;
  }
}

async function parseXlsxFile(arrayBuffer){
  const zipData = await readZipEntries(arrayBuffer);
  if(!zipData["xl/workbook.xml"]) throw new Error("Não encontrei xl/workbook.xml — não parece ser uma planilha .xlsx válida.");
  const sharedStrings = parseSharedStrings(zipData["xl/sharedStrings.xml"]);
  let sheetPath = findFirstSheetPath(zipData);
  if(!sheetPath || !zipData[sheetPath]) sheetPath = "xl/worksheets/sheet1.xml";
  if(!zipData[sheetPath]) throw new Error("Não encontrei a planilha dentro do arquivo.");

  const table = parseWorksheetTable(zipData[sheetPath], sharedStrings);

  let startIdx = 0;
  for(let i=0; i<Math.min(table.length,6); i++){
    const row = table[i] || [];
    if(String(row[1]||"").toLowerCase().includes("qtd")){ startIdx = i+1; break; }
  }

  const parsed = [];
  const problems = [];
  for(let i=startIdx; i<table.length; i++){
    const row = table[i];
    if(!row) continue;
    const [item, qtdRaw, especificacao, descricao, material, massaRaw] = row;
    if(isJunkRow(especificacao)) continue;
    const qtd = toNumBR(qtdRaw);
    const massa = toNumBR(massaRaw);
    if(isNaN(qtd) || isNaN(massa)){ problems.push(`Linha ${i+1} da planilha: Qtd ou Massa não numérico — ignorada.`); continue; }
    parsed.push(buildRow(item, qtd, especificacao, descricao, material, massa));
  }
  return {rows:parsed, problems};
}

/* ---------------- escrita de arquivo .xlsx (sem bibliotecas externas) ---------------- */
function crc32(bytes){
  if(!crc32.table){
    const t = [];
    for(let n=0;n<256;n++){
      let c = n;
      for(let k=0;k<8;k++) c = (c&1) ? (0xEDB88320 ^ (c>>>1)) : (c>>>1);
      t[n] = c>>>0;
    }
    crc32.table = t;
  }
  let crc = 0xFFFFFFFF;
  for(let i=0;i<bytes.length;i++) crc = crc32.table[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function makeZip(files, mimeType){
  const encoder = new TextEncoder();
  const parts = [];
  const centralParts = [];
  let offset = 0;

  for(const f of files){
    const nameBytes = encoder.encode(f.name);
    const data = f.data;
    const crc = crc32(data);
    const size = data.length;

    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    parts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);

    offset += local.length + data.length;
  }

  const centralStart = offset;
  const centralSize = centralParts.reduce((s,c)=>s+c.length, 0);

  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralStart, true);

  return new Blob([...parts, ...centralParts, end], {type: mimeType});
}

function xmlEscape(s){
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
}
function colLetter(i){
  let s = "", n = i+1;
  while(n>0){ const rem=(n-1)%26; s=String.fromCharCode(65+rem)+s; n=Math.floor((n-1)/26); }
  return s;
}

// gera um .xlsx com uma ou mais planilhas. sheets: [{name, headers, rows}], rows = array de
// arrays (number -> célula numérica, qualquer outra coisa -> célula de texto).
function buildXlsxWorkbookBlob(sheets){
  const stringSet = [];
  const stringIndex = new Map();
  function sIdx(str){
    str = String(str);
    if(stringIndex.has(str)) return stringIndex.get(str);
    const idx = stringSet.length;
    stringSet.push(str);
    stringIndex.set(str, idx);
    return idx;
  }

  const sheetXmls = sheets.map(sheet=>{
    const cell = (ci, ri, val)=> typeof val === "number"
      ? `<c r="${colLetter(ci)}${ri}"><v>${val}</v></c>`
      : `<c r="${colLetter(ci)}${ri}" t="s"><v>${sIdx(val)}</v></c>`;
    let rowsXml = `<row r="1">${sheet.headers.map((h,ci)=>cell(ci,1,h)).join("")}</row>`;
    sheet.rows.forEach((r,ri)=>{
      const rn = ri+2;
      rowsXml += `<row r="${rn}">${r.map((v,ci)=>cell(ci,rn,v)).join("")}</row>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`;
  });

  const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${stringSet.length}" uniqueCount="${stringSet.length}">${stringSet.map(s=>`<si><t xml:space="preserve">${xmlEscape(s)}</t></si>`).join("")}</sst>`;
  const sheetEntriesXml = sheets.map((s,i)=>`<sheet name="${xmlEscape(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("");
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetEntriesXml}</sheets></workbook>`;

  const wbRelEntries = sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("");
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRelEntries}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;

  const contentTypesOverrides = sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${contentTypesOverrides}<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const enc = new TextEncoder();
  const files = [
    {name:"[Content_Types].xml", data:enc.encode(contentTypes)},
    {name:"_rels/.rels", data:enc.encode(rootRels)},
    {name:"xl/workbook.xml", data:enc.encode(workbookXml)},
    {name:"xl/_rels/workbook.xml.rels", data:enc.encode(workbookRels)},
    {name:"xl/sharedStrings.xml", data:enc.encode(sharedStringsXml)},
  ];
  sheetXmls.forEach((xml,i)=> files.push({name:`xl/worksheets/sheet${i+1}.xml`, data:enc.encode(xml)}));

  return makeZip(files, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function buildListaCompactadaSheet(grouped){
  const headers = ["Item","Qtd","Especificação","Descrição","Material","Massa (kg)"];
  const rows = grouped.map((g,i)=>[String(i+1), g.qtd, g.especificacao, g.descricao, g.material, Math.round(g.massa*100)/100]);
  return {name:"Lista Compactada", headers, rows};
}

// espessura de chapa com a fração em polegada junto do mm, quando disponível (ex:
// "3/4\" (19,05 mm)") — igual ao jeito que aparecia na descrição original.
function chapaThicknessLabel(r){
  return r.fracLabel ? `${r.fracLabel} (${fmt(r.thickness,2)} mm)` : `${fmt(r.thickness,2)} mm`;
}

function buildResumoCorteSheet(){
  const headers = ["Tipo","Material","Espessura / Bitola","Tamanho comercial","Qtd. comercial necessária","Aproveitamento","Peças","Massa total (kg)"];
  const rows = [];
  lastChapaResults.forEach(r=>{
    if(!r) return;
    rows.push([
      r.circular ? "Chapa circular" : "Chapa retangular", r.material,
      chapaThicknessLabel(r), r.comercialLabel, r.count,
      (r.utilization*100).toFixed(1)+"%", r.pecas, Math.round(r.massaTotal*100)/100,
    ]);
  });
  lastBarraResults.forEach(r=>{
    if(!r) return;
    rows.push([
      "Barra / Tubo", r.material, r.identidade, r.comercialLabel, r.count,
      (r.utilization*100).toFixed(1)+"%", r.pecas, Math.round(r.massaTotal*100)/100,
    ]);
  });
  return {name:"Resumo de Corte", headers, rows};
}

// quantidade de material em unidades de obra: chapas em m² (área real das peças) e
// perfis/tubos em m (comprimento total), por espessura/bitola + material. Ordenado
// alfabeticamente pela coluna descritiva (espessura / bitola), como o resto da lista
// consolidada.
function buildChapasM2Sheet(){
  const headers = ["Material","Espessura","Área total (m²)","Peças"];
  const items = lastChapaResults.filter(Boolean).map(r=>({label:(r.circular?"Ø ":"")+chapaThicknessLabel(r), r}));
  items.sort((a,b)=> a.label.localeCompare(b.label, "pt-BR", {sensitivity:"base"}));
  const rows = items.map(({label,r})=>[r.material, label, Math.round(r.areaM2*100)/100, r.pecas]);
  return {name:"Chapas (m²)", headers, rows};
}
function buildPerfisMSheet(){
  const headers = ["Material","Bitola / Perfil","Comprimento total (m)","Peças"];
  const items = [...lastBarraResults.filter(Boolean)];
  items.sort((a,b)=> a.identidade.localeCompare(b.identidade, "pt-BR", {sensitivity:"base"}));
  const rows = items.map(r=>[r.material, r.identidade, Math.round(r.lengthM*100)/100, r.pecas]);
  return {name:"Perfis (m)", headers, rows};
}

/* ---------------- exportação clonando os modelos reais (.xlsx) ---------------- */
function base64ToArrayBuffer(b64){
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
function bytesToText(bytes){ return new TextDecoder("utf-8").decode(bytes); }
function textToBytes(text){ return new TextEncoder().encode(text); }

// lê o style="s" aplicado à célula `ref` (ex: "D2") dentro de um XML de worksheet já decodificado
function extractCellStyle(xml, ref){
  const idx = xml.indexOf(`<c r="${ref}"`);
  if(idx < 0) return null;
  const tagEnd = xml.indexOf(">", idx);
  const tag = xml.slice(idx, tagEnd+1);
  const m = tag.match(/\ss="(\d+)"/);
  return m ? m[1] : null;
}

// monta uma célula OOXML reaproveitando o estilo do modelo. Usa inlineStr para texto —
// assim não precisamos tocar em xl/sharedStrings.xml do arquivo original.
function tplCell(ref, style, value, isText){
  const sAttr = style ? ` s="${style}"` : "";
  if(value===null || value===undefined || value===""){
    return `<c r="${ref}"${sAttr}/>`;
  }
  if(isText){
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }
  return `<c r="${ref}"${sAttr}><v>${value}</v></c>`;
}

// melhor esforço: atualiza docProps/app.xml para refletir abas extras adicionadas (LLI (2), (3)...)
function patchAppXmlAddSheets(appXmlText, newSheetNames){
  if(!appXmlText || !newSheetNames.length) return appXmlText;
  let xml = appXmlText;

  // a contagem de planilhas é sempre o primeiro <vt:i4> dentro de HeadingPairs (o rótulo
  // pode estar em qualquer idioma — "Worksheets", "Planilhas" etc. — não filtramos por texto)
  const hpMatch = xml.match(/<HeadingPairs><vt:vector size="(\d+)" baseType="variant">([\s\S]*?)<\/vt:vector><\/HeadingPairs>/);
  const tpMatch = xml.match(/<TitlesOfParts><vt:vector size="(\d+)" baseType="lpstr">([\s\S]*?)<\/vt:vector><\/TitlesOfParts>/);
  if(!hpMatch || !tpMatch) return xml;

  const firstI4 = hpMatch[2].match(/<vt:i4>(\d+)<\/vt:i4>/);
  if(!firstI4) return xml;
  const oldSheetCount = parseInt(firstI4[1], 10);

  const newHpInner = hpMatch[2].replace(/<vt:i4>\d+<\/vt:i4>/, `<vt:i4>${oldSheetCount+newSheetNames.length}</vt:i4>`);
  xml = xml.replace(hpMatch[0], `<HeadingPairs><vt:vector size="${hpMatch[1]}" baseType="variant">${newHpInner}</vt:vector></HeadingPairs>`);

  const titleEntries = tpMatch[2].match(/<vt:lpstr>[\s\S]*?<\/vt:lpstr>/g) || [];
  titleEntries.splice(oldSheetCount, 0, ...newSheetNames.map(s=>`<vt:lpstr>${xmlEscape(s)}</vt:lpstr>`));
  const newSize = parseInt(tpMatch[1],10) + newSheetNames.length;
  xml = xml.replace(tpMatch[0], `<TitlesOfParts><vt:vector size="${newSize}" baseType="lpstr">${titleEntries.join("")}</vt:vector></TitlesOfParts>`);

  return xml;
}

// gera um .xlsx idêntico ao modelo "LISTA PRELIMINAR" (título + cabeçalho originais
// preservados), a partir de um ou mais grupos nomeados de linhas — cada grupo vira uma
// aba própria (a primeira reaproveita a aba "Sheet1" original; as demais são clonadas
// dela, com o mesmo título/cabeçalho/estilo, só o nome da aba muda).
async function buildListaXlsxFromTemplate(groups){
  const zipData = await readZipEntries(base64ToArrayBuffer(TEMPLATE_LISTA_B64));
  const sheetPath = "xl/worksheets/sheet1.xml";
  const templateXml = bytesToText(zipData[sheetPath]);

  // estilos observados nas linhas de dados reais do modelo
  const styles = {A:"2", B:"3", C:"4", D:"5", E:"3", F:"3"};

  const sheetDataStart = templateXml.indexOf("<sheetData>") + "<sheetData>".length;
  const sheetDataEnd = templateXml.indexOf("</sheetData>");
  const head = templateXml.slice(0, sheetDataStart);
  const body = templateXml.slice(sheetDataStart, sheetDataEnd);
  const row3Idx = body.indexOf('<row r="3"');
  const keptHead = row3Idx>=0 ? body.slice(0, row3Idx) : body; // título (linha 1) + cabeçalho (linha 2)
  const tail = templateXml.slice(sheetDataEnd);

  function buildSheetXml(rows){
    const rowsXml = rows.map((g,i)=>{
      const r = i+3;
      return `<row r="${r}">` +
        tplCell(`A${r}`, styles.A, String(i+1), true) +
        tplCell(`B${r}`, styles.B, Math.round(g.qtd*100)/100, false) +
        tplCell(`C${r}`, styles.C, g.especificacao, true) +
        tplCell(`D${r}`, styles.D, g.descricao, true) +
        tplCell(`E${r}`, styles.E, g.material, true) +
        tplCell(`F${r}`, styles.F, Math.round(g.massa*100)/100, false) +
        `</row>`;
    }).join("");
    const lastRow = 2 + rows.length;
    let xml = head + keptHead + rowsXml + tail;
    xml = xml.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:F${lastRow}"/>`);
    return xml;
  }

  const wbXmlOrig = bytesToText(zipData["xl/workbook.xml"]);
  const sheet1Tag = (wbXmlOrig.match(/<sheet [^>]*r:id="[^"]*"[^>]*\/>/) || [])[0]
    || '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>';
  const sheet1IdM = sheet1Tag.match(/sheetId="(\d+)"/);
  const sheet1RidM = sheet1Tag.match(/r:id="([^"]+)"/);
  const sheet1Id = sheet1IdM ? parseInt(sheet1IdM[1],10) : 1;
  const sheet1Rid = sheet1RidM ? sheet1RidM[1] : "rId1";
  let nextSheetId = Math.max(1000, sheet1Id) + 1;

  const sheetEntries = [];
  const relEntries = [];
  const contentOverrides = [];
  const newFiles = [];
  const newSheetNames = [];

  groups.forEach((group, gIdx)=>{
    const isFirst = gIdx===0;
    const partName = isFirst ? "worksheets/sheet1.xml" : `worksheets/sheetLista${gIdx+1}.xml`;
    const rid = isFirst ? sheet1Rid : `rId${100+gIdx}`;
    const sheetId = isFirst ? sheet1Id : nextSheetId++;
    sheetEntries.push(`<sheet name="${xmlEscape(group.name)}" sheetId="${sheetId}" r:id="${rid}"/>`);
    if(!isFirst){
      relEntries.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${partName}"/>`);
      contentOverrides.push(`<Override PartName="/xl/${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      newSheetNames.push(group.name);
    }
    newFiles.push({name:`xl/${partName}`, data: textToBytes(buildSheetXml(group.rows))});
  });

  let wbXml = wbXmlOrig.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${sheetEntries.join("")}</sheets>`);
  let relsXml = bytesToText(zipData["xl/_rels/workbook.xml.rels"])
    .replace("</Relationships>", relEntries.join("") + "</Relationships>");
  let ctXml = bytesToText(zipData["[Content_Types].xml"])
    .replace("</Types>", contentOverrides.join("") + "</Types>");
  const appXml = zipData["docProps/app.xml"] ? bytesToText(zipData["docProps/app.xml"]) : null;
  const patchedAppXml = appXml ? patchAppXmlAddSheets(appXml, newSheetNames) : null;

  const skip = new Set(["xl/workbook.xml","xl/_rels/workbook.xml.rels","[Content_Types].xml","docProps/app.xml", sheetPath]);
  const files = [];
  for(const [name, data] of Object.entries(zipData)){
    if(skip.has(name)) continue;
    files.push({name, data});
  }
  files.push({name:"xl/workbook.xml", data: textToBytes(wbXml)});
  files.push({name:"xl/_rels/workbook.xml.rels", data: textToBytes(relsXml)});
  files.push({name:"[Content_Types].xml", data: textToBytes(ctXml)});
  if(patchedAppXml) files.push({name:"docProps/app.xml", data: textToBytes(patchedAppXml)});
  newFiles.forEach(f=>files.push(f));

  return makeZip(files, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// leitor do modelo "Lista consolidada": pega da aba "LLI" do modelo (clone de
// "LM Consolidada.xlsx") o estilo de cada coluna, as larguras e a configuração de página, e
// devolve o construtor do XML de uma aba. Os dois geradores que usam o modelo passam por aqui,
// então qualquer conteúdo exportado sai com exatamente o mesmo layout de 6 colunas:
// Item No. | Especificação | Área total (m²) | Comprimento (m) | Material | Total QTY.
// Cada linha preenche só a coluna que se aplica (chapa usa Área, perfil/tubo usa Comprimento,
// item solto usa Total QTY) e marca "-" nas outras.
const CONSOLIDADA_HEADERS = ["Item No.", "Especificação", "Área total (m²)", "Comprimento (m)", "Material", "Total QTY"];
const CONSOLIDADA_COLS = ["A","B","C","D","E","F"];

function makeConsolidadaSheetWriter(templateXml){
  // os fallbacks são os índices de estilo do próprio modelo embutido — só entram em ação se
  // alguém trocar o modelo por um sem a linha 2 de referência
  const FALLBACK_DATA = {A:"23", B:"24", C:"23", D:"23", E:"23", F:"23"};
  const FALLBACK_HEAD = {A:"1", B:"1", C:"19", D:"1", E:"1", F:"1"};
  const dataStyle = {}, headStyle = {};
  CONSOLIDADA_COLS.forEach(c=>{
    dataStyle[c] = extractCellStyle(templateXml, c+"2") || FALLBACK_DATA[c];
    headStyle[c] = extractCellStyle(templateXml, c+"1") || FALLBACK_HEAD[c];
  });

  const rootTag = (templateXml.match(/<worksheet[^>]*>/) || [])[0]
    || '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
  const colsXml = (templateXml.match(/<cols>[\s\S]*?<\/cols>/) || [])[0] || "";
  // o r:id do pageSetup aponta para o printerSettings da aba original; as abas extras não têm
  // esse relacionamento, então sai fora (o resto — papel A4, paisagem, margens — fica igual)
  const pageMarginsXml = (templateXml.match(/<pageMargins[^>]*\/>/) || [])[0] || "";
  const pageSetupXml = ((templateXml.match(/<pageSetup[^>]*\/>/) || [])[0] || "")
    .replace(/\sr:id="[^"]*"/, "");
  const sheetFormatXml = (templateXml.match(/<sheetFormatPr[^>]*\/>/) || [])[0] || '<sheetFormatPr defaultRowHeight="15"/>';
  const headerRowAttrs = ((templateXml.match(/<row r="1"([^>]*)>/) || [])[1] || ' ht="52.5" customHeight="1"')
    .replace(/\sspans="[^"]*"/, "");

  const headerRowXml = `<row r="1"${headerRowAttrs}>` +
    CONSOLIDADA_COLS.map((c,i)=> tplCell(`${c}1`, headStyle[c], CONSOLIDADA_HEADERS[i], true)).join("") +
    `</row>`;

  // `pageRows`: linhas no shape consolidado ({especificacao, area, comprimento, material, qty});
  // `startItemNo`: primeiro "Item No." da aba — a numeração é contínua entre as abas, não
  // reinicia a cada página nova.
  return function buildSheetXml(pageRows, startItemNo){
    let rowsXml = headerRowXml;
    pageRows.forEach((g,i)=>{
      const r = i+2;
      const lines = estimateWrappedLines(g.especificacao);
      const rowAttrs = lines>1 ? ` ht="${lines*LLI_CONSOLIDADO_LINE_HEIGHT_PT}" customHeight="1"` : "";
      rowsXml += `<row r="${r}"${rowAttrs}>` +
        tplCell(`A${r}`, dataStyle.A, startItemNo + i, false) +
        tplCell(`B${r}`, dataStyle.B, g.especificacao, true) +
        tplCell(`C${r}`, dataStyle.C, g.area!=null ? g.area : "-", g.area==null) +
        tplCell(`D${r}`, dataStyle.D, g.comprimento!=null ? g.comprimento : "-", g.comprimento==null) +
        tplCell(`E${r}`, dataStyle.E, g.material, true) +
        tplCell(`F${r}`, dataStyle.F, g.qty!=null ? g.qty : "-", g.qty==null) +
        `</row>`;
    });
    const lastRow = 1 + pageRows.length;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootTag}<dimension ref="A1:F${lastRow}"/>` +
      `${sheetFormatXml}${colsXml}<sheetData>${rowsXml}</sheetData>${pageMarginsXml}${pageSetupXml}</worksheet>`;
  };
}

// gera a lista consolidada a partir de um ou mais grupos nomeados de linhas no shape genérico
// ({especificacao, descricao, material, qtd, massa, pecas?}) — usado pelos conteúdos que não
// são a "lista consolidada" propriamente dita (resumo de corte, lista compactada, itens
// soltos). Cada grupo vira uma ou mais abas, no máximo 14 itens por aba (abas extras
// "Nome (2)", "Nome (3)"...). Especificação e descrição entram juntas na coluna B, e a
// quantidade na coluna "Total QTY" — as colunas de área e comprimento ficam com "-", porque
// esses conteúdos contam peças, não m²/m.
// `applyBuffer` soma os 5% de folga do modelo original (peças soltas/consumíveis); para
// conteúdos que já são a quantidade final a comprar (resumo de corte) ele vem desligado.
async function buildLLIXlsxFromTemplate(groups, applyBuffer=true){
  const zipData = await readZipEntries(base64ToArrayBuffer(TEMPLATE_CONSOLIDADA_B64));
  const templateXml = bytesToText(zipData["xl/worksheets/sheet2.xml"]);
  const writeSheet = makeConsolidadaSheetWriter(templateXml);

  const PAGE_SIZE = 14;
  // achata os grupos numa lista de páginas, cada uma sabendo seu nome de aba
  const pages = [];
  groups.forEach(group=>{
    const rows = group.rows.map(g=>({
      especificacao: `${g.especificacao||""} ${g.descricao||""}`.trim(),
      area: null,
      comprimento: null,
      material: g.material,
      qty: g.pecas!=null ? g.pecas
        : (applyBuffer ? Math.ceil((g.qtd||0)*1.05) : Math.round((g.qtd||0)*100)/100),
    }));
    const chunks = [];
    for(let i=0;i<rows.length;i+=PAGE_SIZE) chunks.push(rows.slice(i, i+PAGE_SIZE));
    if(chunks.length===0) chunks.push([]);
    chunks.forEach((chunk,ci)=> pages.push({rows:chunk, name: ci===0 ? group.name : `${group.name} (${ci+1})`}));
  });
  if(pages.length===0) pages.push({rows:[], name:"LLI"});

  return assembleLLIWorkbook(zipData, pages, writeSheet);
}

// monta o pacote final do .xlsx a partir do modelo "Lista consolidada" (aba Cover Page
// preservada como está no modelo) e uma lista de páginas já prontas ({rows, name}) — comum ao
// buildLLIXlsxFromTemplate (conteúdos genéricos) e ao buildConsolidadoLLIXlsx (lista
// consolidada propriamente dita). `buildSheetXml(pageRows, startItemNo)` gera o XML da
// planilha de cada página; o "Item No." é contínuo entre as abas, então cada chamada recebe o
// número em que aquela página começa.
function assembleLLIWorkbook(zipData, pages, buildSheetXml){
  const wbXmlOrig = bytesToText(zipData["xl/workbook.xml"]);
  const coverSheetTag = (wbXmlOrig.match(/<sheet [^>]*name="Cover Page"[^>]*\/>/) || [])[0]
    || '<sheet name="Cover Page" sheetId="2" r:id="rId1"/>';
  const lliSheetTag = (wbXmlOrig.match(/<sheet [^>]*name="LLI"[^>]*\/>/) || [])[0]
    || '<sheet name="LLI" sheetId="1" r:id="rId2"/>';
  // sheetId precisa ser único no workbook — usamos uma faixa alta para as páginas extras,
  // bem longe dos ids 1/2 já usados pelas abas originais (Cover Page e LLI).
  const usedSheetIds = [coverSheetTag, lliSheetTag].map(t=>{
    const m = t.match(/sheetId="(\d+)"/);
    return m ? parseInt(m[1],10) : 0;
  });
  let nextSheetId = Math.max(1000, ...usedSheetIds) + 1;

  // o r:id das abas extras também precisa ser único — o modelo já gasta rIds com as duas abas,
  // os estilos, o tema e a tabela de strings, então continuamos a partir do maior que existe
  // em vez de chutar uma faixa fixa (chutar colidia com o rId dos estilos).
  const relsXmlOrig = bytesToText(zipData["xl/_rels/workbook.xml.rels"]);
  const usedRelNums = (relsXmlOrig.match(/Id="rId(\d+)"/g) || [])
    .map(s=> parseInt(s.replace(/\D/g,""), 10));
  let nextRelNum = Math.max(0, ...usedRelNums) + 1;

  const sheetEntries = [coverSheetTag];
  const relEntries = [];
  const contentOverrides = [];
  const newFiles = [];
  const newSheetNames = [];
  // área de impressão de cada aba gerada: o modelo traz a do arquivo de referência, com o
  // número de linhas daquele arquivo — sem reescrever, a impressão cortaria (ou sobraria)
  // linhas conforme o tamanho da lista.
  const printAreas = [];
  let itemNo = 1;

  pages.forEach((page, pIdx)=>{
    const isFirst = pIdx===0;
    const partName = isFirst ? "worksheets/sheet2.xml" : `worksheets/sheetLLI${pIdx+1}.xml`;
    const sheetName = page.name;
    const rid = isFirst ? "rId2" : `rId${nextRelNum++}`;
    const sheetId = isFirst ? (usedSheetIds[1] || 1) : nextSheetId++;
    sheetEntries.push(`<sheet name="${xmlEscape(sheetName)}" sheetId="${sheetId}" r:id="${rid}"/>`);
    if(!isFirst){
      relEntries.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${partName}"/>`);
      contentOverrides.push(`<Override PartName="/xl/${partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      newSheetNames.push(sheetName);
    }
    // localSheetId é a posição da aba na lista <sheets> — a Cover Page ocupa a 0, então as
    // páginas começam na 1.
    printAreas.push(`<definedName name="_xlnm.Print_Area" localSheetId="${pIdx+1}">'${sheetName.replace(/'/g,"''")}'!$A$1:$F$${1+page.rows.length}</definedName>`);
    newFiles.push({name:`xl/${partName}`, data: textToBytes(buildSheetXml(page.rows, itemNo))});
    itemNo += page.rows.length;
  });

  let wbXml = wbXmlOrig.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${sheetEntries.join("")}</sheets>`);
  // troca as áreas de impressão das abas de lista (localSheetId >= 1) pelas recém-calculadas,
  // preservando a da Cover Page (localSheetId 0) e os demais nomes definidos do modelo
  wbXml = wbXml.replace(/<definedName name="_xlnm\.Print_Area" localSheetId="[1-9]\d*">[\s\S]*?<\/definedName>/g, "");
  wbXml = wbXml.replace(/<\/definedNames>/, printAreas.join("") + "</definedNames>");

  const relsXml = relsXmlOrig.replace("</Relationships>", relEntries.join("") + "</Relationships>");
  const ctXml = bytesToText(zipData["[Content_Types].xml"])
    .replace("</Types>", contentOverrides.join("") + "</Types>");

  const appXml = zipData["docProps/app.xml"] ? bytesToText(zipData["docProps/app.xml"]) : null;
  const patchedAppXml = appXml ? patchAppXmlAddSheets(appXml, newSheetNames) : null;

  const skip = new Set(["xl/workbook.xml","xl/_rels/workbook.xml.rels","[Content_Types].xml","docProps/app.xml","xl/worksheets/sheet2.xml"]);
  const files = [];
  for(const [name, data] of Object.entries(zipData)){
    if(skip.has(name)) continue;
    files.push({name, data});
  }
  files.push({name:"xl/workbook.xml", data: textToBytes(wbXml)});
  files.push({name:"xl/_rels/workbook.xml.rels", data: textToBytes(relsXml)});
  files.push({name:"[Content_Types].xml", data: textToBytes(ctXml)});
  if(patchedAppXml) files.push({name:"docProps/app.xml", data: textToBytes(patchedAppXml)});
  newFiles.forEach(f=>files.push(f));

  return makeZip(files, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

/* ---------------- paginação da lista consolidada ---------------- */
// medidas tiradas do arquivo de referência do usuário ("LM Consolidada.xlsx", A4 paisagem):
// largura da coluna B (Especificação, onde entra o texto mais longo) e altura de linha padrão
// — usadas só para ESTIMAR quantas linhas de texto quebrado cada item ocupa e decidir onde
// cortar a página no modo automático. Não é o motor de quebra de linha do Excel (isso exigiria
// medir a fonte de verdade), é uma estimativa por contagem de caracteres — suficiente pra
// decidir onde abrir aba nova sem estourar muito uma folha A4 impressa.
const LLI_CONSOLIDADO_CHARS_PER_LINE = 58;
const LLI_CONSOLIDADO_LINE_HEIGHT_PT = 15;
const LLI_CONSOLIDADO_PAGE_USABLE_PT = 430; // A4 paisagem, margens da referência, descontado o cabeçalho

function estimateWrappedLines(text){
  const len = String(text||"").length;
  return Math.max(1, Math.ceil(len / LLI_CONSOLIDADO_CHARS_PER_LINE));
}

// divide os itens soltos em páginas: modo "fixed" usa N itens por página (número fixo,
// escolhido pelo usuário); modo "auto" estima quantas linhas cada descrição ocupa quebrada na
// coluna B e enche cada página até o limite de altura impressa em A4 paisagem — itens com
// descrição longa cabem menos por página, igual ao ajuste manual que dá pra ver no arquivo de
// referência "LM Consolidada.xlsx" (as páginas de lá não têm todas o mesmo número de linhas).
function paginateConsolidadoOutro(rows, pageOpts){
  const pages = [];
  if(pageOpts && pageOpts.mode === "fixed"){
    const size = Math.max(1, pageOpts.size || 14);
    for(let i=0; i<rows.length; i+=size) pages.push(rows.slice(i, i+size));
  } else {
    let current = [], used = 0;
    rows.forEach(g=>{
      const h = estimateWrappedLines(g.especificacao) * LLI_CONSOLIDADO_LINE_HEIGHT_PT;
      if(current.length && used + h > LLI_CONSOLIDADO_PAGE_USABLE_PT){ pages.push(current); current = []; used = 0; }
      current.push(g); used += h;
    });
    if(current.length) pages.push(current);
  }
  if(pages.length===0) pages.push([]);
  return pages;
}

// linhas da lista consolidada, no shape usado pelo formato "Lista consolidada" (uma única coluna de texto
// combinando especificação+descrição — igual ao que vai pra coluna B do arquivo de
// referência). Funções puras e síncronas (sem precisar do template .xlsx) pra poderem ser
// reaproveitadas tanto na geração real quanto na pré-visualização, sem duplicar a lógica.
function buildConsolidadoCortavelRows(){
  return sortByDescricao([
    ...lastChapaResults.filter(Boolean).map(r=>{
      const label = "Chapa " + (r.fracLabel ? `#${r.fracLabel} (${fmt(r.thickness,2)} mm)` : `e=${fmt(r.thickness,2)} mm`);
      return {especificacao: label, descricao: label, area: Math.round(r.areaM2*100)/100, comprimento: null, material: r.material, qty: null};
    }),
    ...lastBarraResults.filter(Boolean).map(r=>{
      const label = barraEspecificacaoCombinada(r);
      return {especificacao: label, descricao: label, area: null, comprimento: Math.round(r.lengthM*100)/100, material: r.material, qty: null};
    }),
  ]);
}
function buildConsolidadoOutroRows(){
  return sortByDescricao(lastGrouped.filter(isLooseOrUnclassified).map(g=>{
    const label = `${g.especificacao} ${g.descricao}`.trim();
    return {especificacao: label, descricao: label, area: null, comprimento: null, material: g.material, qty: Math.round(g.qtd*100)/100};
  }));
}

// gera a lista consolidada, igual ao arquivo de referência do usuário ("LM Consolidada.xlsx"):
// Item No. | Especificação | Área total (m²) | Comprimento (m) | Material | Total QTY — cada
// linha preenche só a coluna que se aplica (área OU comprimento OU quantidade) e usa "-" nas
// outras duas. Os itens cortáveis (chapas + perfis/tubos) vão todos numa aba "LLI" só, sem
// paginar; os itens soltos (parafusos, arruelas, curvas etc.) vêm depois, paginados em abas
// extras "LLI (2)", "LLI (3)"... conforme `pageOpts`. As duas listas saem em ordem alfabética
// por descrição (`pageOpts`: {mode:"auto"|"fixed", size}).
// monta as abas da lista consolidada: os itens cortáveis (chapas + perfis/tubos) numa aba só,
// sem paginar, e os itens soltos paginados logo depois. As abas são nomeadas pela POSIÇÃO —
// "LLI", "LLI (2)", "LLI (3)"... — e uma lista sem nenhum item cortável (só parafusos,
// arruelas, curvas etc.) começa direto na "LLI", sem deixar uma aba vazia na frente.
// Cada página sai marcada com `tipo` ("cortavel" ou "outro") para a prévia poder rotulá-la.
// Fonte única usada pela geração do .xlsx, pela pré-visualização e pela mensagem de status,
// pra que os três concordem sobre em que aba cada item vai parar.
function buildConsolidadoPages(cortavelRows, outroRows, pageOpts){
  const pages = [];
  if(cortavelRows.length) pages.push({rows: cortavelRows, tipo: "cortavel"});
  if(outroRows.length){
    paginateConsolidadoOutro(outroRows, pageOpts).forEach(rows=> pages.push({rows, tipo: "outro"}));
  }
  if(!pages.length) pages.push({rows: [], tipo: "cortavel"});
  pages.forEach((p,i)=>{ p.name = i===0 ? "LLI" : `LLI (${i+1})`; });
  return pages;
}

async function buildConsolidadoLLIXlsx(includeOutro=true, pageOpts={mode:"auto", size:14}){
  const zipData = await readZipEntries(base64ToArrayBuffer(TEMPLATE_CONSOLIDADA_B64));
  const templateXml = bytesToText(zipData["xl/worksheets/sheet2.xml"]);
  const writeSheet = makeConsolidadaSheetWriter(templateXml);

  const cortavelRows = buildConsolidadoCortavelRows();
  const outroRows = includeOutro ? buildConsolidadoOutroRows() : [];

  return assembleLLIWorkbook(zipData, buildConsolidadoPages(cortavelRows, outroRows, pageOpts), writeSheet);
}

function nestPlates(pieces, sheetW, sheetH, kerf){
  const items = [];
  pieces.forEach(p=>{ for(let i=0;i<p.qty;i++) items.push({w:p.w, h:p.h, label:p.label, circular:!!p.circular}); });
  items.sort((a,b)=> Math.max(b.w,b.h) - Math.max(a.w,a.h));

  const sheets = [];
  function newSheet(){ const sh={shelves:[], placements:[]}; sheets.push(sh); return sh; }
  let current = newSheet();

  for(const it of items){
    const orientations = it.w===it.h ? [[it.w,it.h]] : [[it.w,it.h],[it.h,it.w]];
    let placed = false;

    for(const shelf of current.shelves){
      for(const [w,h] of orientations){
        if(h <= shelf.height + 0.001 && shelf.usedWidth + w <= sheetW + 0.001){
          current.placements.push({x:shelf.usedWidth, y:shelf.y, w, h, label:it.label, circular:it.circular});
          shelf.usedWidth += w + kerf;
          placed = true; break;
        }
      }
      if(placed) break;
    }
    if(placed) continue;

    for(const [w,h] of orientations){
      const lastShelf = current.shelves[current.shelves.length-1];
      const nextY = lastShelf ? (lastShelf.y + lastShelf.height + kerf) : 0;
      if(nextY + h <= sheetH + 0.001 && w <= sheetW + 0.001){
        current.shelves.push({y:nextY, height:h, usedWidth:w+kerf});
        current.placements.push({x:0, y:nextY, w, h, label:it.label, circular:it.circular});
        placed = true; break;
      }
    }
    if(placed) continue;

    current = newSheet();
    let fitted = false;
    for(const [w,h] of orientations){
      if(w <= sheetW + 0.001 && h <= sheetH + 0.001){
        current.shelves.push({y:0, height:h, usedWidth:w+kerf});
        current.placements.push({x:0, y:0, w, h, label:it.label, circular:it.circular});
        fitted = true; break;
      }
    }
    if(!fitted){
      current.placements.push({x:0, y:0, w:Math.min(orientations[0][0],sheetW), h:Math.min(orientations[0][1],sheetH), label:it.label, circular:it.circular, overflow:true});
    }
  }

  const usedArea = sheets.reduce((s,sh)=> s + sh.placements.reduce((s2,p)=>s2+p.w*p.h,0), 0);
  const totalArea = sheets.length * sheetW * sheetH;
  return {sheets, count:sheets.length, utilization: totalArea>0 ? usedArea/totalArea : 0};
}

/* ---------------- corte 1D (barras/tubos) — first-fit decreasing ---------------- */
function cutStock1D(pieces, barLength, kerf){
  const items = [];
  pieces.forEach(p=>{ for(let i=0;i<p.qty;i++) items.push({len:p.len, label:p.label}); });
  items.sort((a,b)=> b.len - a.len);

  const bars = [];
  for(const it of items){
    if(it.len > barLength + 0.001){
      bars.push({cuts:[{len:it.len, label:it.label, overflow:true}], used:it.len});
      continue;
    }
    let target = null;
    for(const bar of bars){
      if(!bar.cuts.some(c=>c.overflow) && bar.used + it.len + kerf <= barLength + 0.001){ target = bar; break; }
    }
    if(!target){ target = {cuts:[], used:0}; bars.push(target); }
    target.cuts.push({len:it.len, label:it.label});
    target.used += it.len + kerf;
  }

  const totalUsed = bars.reduce((s,b)=> s + b.cuts.reduce((s2,c)=>s2+c.len,0), 0);
  const totalBarLen = bars.length * barLength;
  return {bars, count:bars.length, utilization: totalBarLen>0 ? totalUsed/totalBarLen : 0};
}

// variante de cutStock1D que aceita mais de um comprimento comercial ao mesmo tempo —
// permite, por exemplo, fechar a última barra num tamanho menor em vez de desperdiçar
// quase uma barra inteira de 12m para sobrar só uma peça pequena.
function cutStock1DMulti(pieces, barSizes, kerf){
  const sizesAsc = [...barSizes].sort((a,b)=>a.len-b.len);
  const maxLen = sizesAsc[sizesAsc.length-1].len;

  let pool = [];
  pieces.forEach(p=>{ for(let i=0;i<p.qty;i++) pool.push({len:p.len, label:p.label}); });
  pool.sort((a,b)=> b.len - a.len);

  const bars = [];

  // encaixe guloso (first-fit) de uma lista de peças, em ordem, dentro de um compartimento
  // de comprimento `size` — usado tanto para decidir quanto para de fato preencher uma barra
  function fillBin(size, items){
    let used = 0;
    const taken = [];
    for(const it of items){
      if(used + it.len + kerf <= size + 0.001){ taken.push(it); used += it.len + kerf; }
    }
    return {used, taken};
  }

  while(pool.length){
    const anchor = pool[0];
    if(anchor.len > maxLen + 0.001){
      const biggest = sizesAsc[sizesAsc.length-1];
      bars.push({cuts:[{len:anchor.len, label:anchor.label, overflow:true}], used:anchor.len, barLen:biggest.len, barLabel:biggest.label});
      pool = pool.slice(1);
      continue;
    }

    // primeiro tenta aproveitar espaço já sobrando numa barra aberta (best-fit)
    let existing = null, bestRemaining = Infinity;
    for(const bar of bars){
      if(bar.cuts.some(c=>c.overflow)) continue;
      const remaining = bar.barLen - bar.used;
      if(anchor.len + kerf <= remaining + 0.001 && remaining < bestRemaining){ existing = bar; bestRemaining = remaining; }
    }
    if(existing){
      existing.cuts.push({len:anchor.len, label:anchor.label});
      existing.used += anchor.len + kerf;
      pool = pool.slice(1);
      continue;
    }

    // nenhuma barra aberta serve — escolhe o tamanho comercial que dá o melhor
    // aproveitamento simulando o encaixe guloso do restante da fila em cada opção
    // disponível (evita abrir uma barra pequena "só porque cabe" quando uma maior
    // combinaria melhor com as próximas peças da fila)
    let bestSize = null, bestFill = null, bestUtil = -1;
    for(const size of sizesAsc){
      if(anchor.len > size.len + 0.001) continue;
      const fill = fillBin(size.len, pool);
      const util = fill.used / size.len;
      if(util > bestUtil){ bestUtil = util; bestSize = size; bestFill = fill; }
    }
    bars.push({cuts:bestFill.taken.map(it=>({len:it.len, label:it.label})), used:bestFill.used, barLen:bestSize.len, barLabel:bestSize.label});
    const takenSet = new Set(bestFill.taken);
    pool = pool.filter(it=> !takenSet.has(it));
  }

  const totalUsed = bars.reduce((s,b)=> s + b.cuts.reduce((s2,c)=>s2+c.len,0), 0);
  const totalBarLen = bars.reduce((s,b)=> s + b.barLen, 0);
  return {bars, count:bars.length, utilization: totalBarLen>0 ? totalUsed/totalBarLen : 0};
}

/* ---------------- desenho SVG ---------------- */
// desenha cada peça como retângulo ou círculo (inscrito no quadrado que a peça ocupa),
// conforme a própria peça — assim uma chapa retangular e uma circular da mesma espessura
// podem aparecer juntas na mesma folha, já que passaram a ser um grupo só.
function renderSheetSvg(sheet, sheetW, sheetH){
  const scale = 620 / Math.max(sheetW, sheetH);
  const W = sheetW*scale, H = sheetH*scale;
  let s = `<rect class="sheet-outline" x="0.5" y="0.5" width="${W-1}" height="${H-1}"/>`;
  sheet.placements.forEach(p=>{
    const x=p.x*scale, y=p.y*scale, w=p.w*scale, h=p.h*scale;
    if(p.circular){
      const r = Math.min(w,h)/2 - 1;
      s += `<circle class="piece-circle" cx="${(x+w/2).toFixed(1)}" cy="${(y+h/2).toFixed(1)}" r="${Math.max(r,0).toFixed(1)}"/>`;
      if(r>18){
        s += `<text class="piece-label" x="${(x+w/2).toFixed(1)}" y="${(y+h/2+3).toFixed(1)}" text-anchor="middle" font-size="10">${p.label}</text>`;
      }
    } else {
      s += `<rect class="piece-rect" x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(w-1.5,0).toFixed(1)}" height="${Math.max(h-1.5,0).toFixed(1)}"/>`;
      if(w>34 && h>16){
        s += `<text class="piece-label" x="${(x+w/2).toFixed(1)}" y="${(y+h/2+3).toFixed(1)}" text-anchor="middle" font-size="10">${p.label}</text>`;
      }
    }
    if(p.overflow){
      s += `<text class="piece-label" x="${(x+w/2).toFixed(1)}" y="${(y+h/2+16).toFixed(1)}" text-anchor="middle" font-size="9" fill="#B23A32">excede a chapa!</text>`;
    }
  });
  return `<svg viewBox="0 0 ${W.toFixed(1)} ${H.toFixed(1)}" xmlns="http://www.w3.org/2000/svg">${s}</svg>`;
}

function renderBarSvg(bar, barLength){
  const H = 46;
  const scale = 900 / barLength;
  const W = barLength*scale;
  let s = `<rect class="bar-outline" x="0.5" y="0.5" width="${(W-1).toFixed(1)}" height="${H-1}"/>`;
  let x = 0;
  bar.cuts.forEach(c=>{
    const w = c.len*scale;
    s += `<rect class="cut-seg" x="${x.toFixed(1)}" y="0.5" width="${Math.max(w-1.5,0).toFixed(1)}" height="${H-1}"/>`;
    if(w>28){
      s += `<text class="cut-label" x="${(x+w/2).toFixed(1)}" y="${H/2+3}" text-anchor="middle">${fmt(c.len,0)}</text>`;
    }
    x += w;
  });
  const restW = W-x;
  if(restW>2 && !bar.cuts.some(c=>c.overflow)){
    s += `<rect x="${x.toFixed(1)}" y="0.5" width="${(restW-1).toFixed(1)}" height="${H-1}" fill="none" stroke="var(--border)" stroke-dasharray="3,3"/>`;
    if(restW>34) s += `<text class="waste-label" x="${(x+restW/2).toFixed(1)}" y="${H/2+3}" text-anchor="middle">sobra ${fmt(barLength-bar.used,0)}</text>`;
  }
  return `<svg viewBox="0 0 ${W.toFixed(1)} ${H}" xmlns="http://www.w3.org/2000/svg">${s}</svg>`;
}

/* ---------------- tamanhos comerciais ---------------- */
const CHAPAS_COMERCIAIS = [
  {w:1200, h:2400, label:"1,2 x 2,4 m"},
  {w:1500, h:3000, label:"1,5 x 3 m"},
  {w:2000, h:3000, label:"2 x 3 m"},
  {w:2000, h:6000, label:"2 x 6 m"},
  {w:2440, h:6000, label:"2,44 x 6 m"},
  {w:3000, h:6000, label:"3 x 6 m"},
];
const BARRAS_COMERCIAIS = [
  {len:6000, label:"6 m"},
  {len:12000, label:"12 m"},
];

/* ---------------- wiring da interface ---------------- */
const els = {
  pasteArea: document.getElementById("pasteArea"),
  processBtn: document.getElementById("processBtn"),
  clearBtn: document.getElementById("clearBtn"),
  xlsxInput: document.getElementById("xlsxInput"),
  parseStatus: document.getElementById("parseStatus"),
  groupedSection: document.getElementById("groupedSection"),
  groupedTable: document.getElementById("groupedTable"),
  groupedBody: document.getElementById("groupedBody"),
  groupedToggleBtn: document.getElementById("groupedToggleBtn"),
  addGroupedRowBtn: document.getElementById("addGroupedRowBtn"),
  reprocessBtn: document.getElementById("reprocessBtn"),
  exportXlsxBtn: document.getElementById("exportXlsxBtn"),
  exportStatus: document.getElementById("exportStatus"),
  contentSection: document.getElementById("contentSection"),
  exportSection: document.getElementById("exportSection"),
  exportContentSel: document.getElementById("exportContentSel"),
  exportFormatSel: document.getElementById("exportFormatSel"),
  consolidadaPageOpts: document.getElementById("consolidadaPageOpts"),
  pageModeSel: document.getElementById("pageModeSel"),
  pageSizeWrap: document.getElementById("pageSizeWrap"),
  pageSizeInput: document.getElementById("pageSizeInput"),
  previewBtn: document.getElementById("previewBtn"),
  previewPanel: document.getElementById("previewPanel"),
  previewBody: document.getElementById("previewBody"),
  previewCloseBtn: document.getElementById("previewCloseBtn"),
  exportFinalBtn: document.getElementById("exportFinalBtn"),
  exportFinalStatus: document.getElementById("exportFinalStatus"),
  chapasSection: document.getElementById("chapasSection"),
  chapaGroups: document.getElementById("chapaGroups"),
  chapasToggleAllBtn: document.getElementById("chapasToggleAllBtn"),
  barrasSection: document.getElementById("barrasSection"),
  barraGroups: document.getElementById("barraGroups"),
  barrasToggleAllBtn: document.getElementById("barrasToggleAllBtn"),
};

// minimiza/expande todos os cards de um contêiner de uma vez (usa o mesmo atributo
// data-*-toggle / data-*-body que cada card já tem individualmente)
function toggleAllGroupCards(container, toggleBtn, bodyAttr, toggleAttr, collapseLabel, expandLabel){
  const bodies = [...container.querySelectorAll(`[${bodyAttr}]`)];
  const shouldCollapse = bodies.some(b=>!b.hidden);
  bodies.forEach(b=> b.hidden = shouldCollapse);
  container.querySelectorAll(`[${toggleAttr}]`).forEach(btn=>{
    btn.textContent = shouldCollapse ? "▸ Expandir" : "▾ Minimizar";
  });
  toggleBtn.textContent = shouldCollapse ? expandLabel : collapseLabel;
}
els.chapasToggleAllBtn.addEventListener("click", ()=>{
  toggleAllGroupCards(els.chapaGroups, els.chapasToggleAllBtn, "data-chapa-body", "data-chapa-toggle",
    "▾ Minimizar todas as espessuras", "▸ Expandir todas as espessuras");
});
els.barrasToggleAllBtn.addEventListener("click", ()=>{
  toggleAllGroupCards(els.barraGroups, els.barrasToggleAllBtn, "data-barra-body", "data-barra-toggle",
    "▾ Minimizar todos os perfis", "▸ Expandir todos os perfis");
});

let lastGrouped = [];
let lastChapaResults = [];
let lastBarraResults = [];
let chapasHasData = false;
let barrasHasData = false;

els.clearBtn.addEventListener("click", ()=>{
  els.pasteArea.value = "";
  els.xlsxInput.value = "";
  els.parseStatus.textContent = "";
  els.groupedSection.hidden = true;
  els.contentSection.hidden = true;
  els.chapasSection.hidden = true;
  els.barrasSection.hidden = true;
  els.exportSection.hidden = true;
  els.previewPanel.hidden = true;
});

// as seções 4/5 (otimização de corte de chapas e de perfis/tubos) só fazem sentido quando o
// usuário vai gerar o "Resumo de corte" (chapa/comprimento comercial escolhidos ali embaixo
// decidem quanto comprar) — nos outros conteúdos (lista consolidada em m²/m, compactada,
// itens soltos) elas ficam escondidas, mesmo que existam grupos cortáveis na lista.
function applySectionVisibility(){
  const relevantForCorte = els.exportContentSel.value === "corte";
  els.chapasSection.hidden = !chapasHasData || !relevantForCorte;
  els.barrasSection.hidden = !barrasHasData || !relevantForCorte;
}
els.exportContentSel.addEventListener("change", ()=>{
  applySectionVisibility();
  updateFormatOptsVisibility();
  els.previewPanel.hidden = true;
});

// os controles de paginação (itens por página nos soltos) só fazem sentido pra lista
// consolidada no formato "Lista consolidada" — nos outros conteúdos/formatos não há paginação em abas.
function updateFormatOptsVisibility(){
  const showPageOpts = els.exportContentSel.value === "consolidado" && els.exportFormatSel.value === "consolidada";
  els.consolidadaPageOpts.hidden = !showPageOpts;
  els.pageSizeWrap.hidden = els.pageModeSel.value !== "fixed";
}
function currentPageOpts(){
  return {mode: els.pageModeSel.value, size: parseInt(els.pageSizeInput.value,10) || 14};
}
els.exportFormatSel.addEventListener("change", ()=>{ updateFormatOptsVisibility(); els.previewPanel.hidden = true; });
els.pageModeSel.addEventListener("change", ()=>{ updateFormatOptsVisibility(); els.previewPanel.hidden = true; });
els.pageSizeInput.addEventListener("input", ()=>{ els.previewPanel.hidden = true; });
updateFormatOptsVisibility();

function runPipeline(rows, problems, sourceLabel){
  if(rows.length===0){
    els.parseStatus.textContent = `Nenhuma linha válida encontrada em ${sourceLabel}. Confira as 6 colunas (Item, Qtd, Especificação, Descrição, Material, Massa).`;
    els.parseStatus.className = "status-msg err";
    return;
  }
  const grouped = sortByDescricao(groupRows(rows));
  lastGrouped = grouped;

  els.parseStatus.textContent = `${sourceLabel}: ${rows.length} linha(s) lida(s) → ${grouped.length} item(ns) agrupado(s).` + (problems.length ? `  ${problems.length} linha(s) ignorada(s).` : "");
  els.parseStatus.className = "status-msg ok";

  renderGroupedTable(lastGrouped);
  renderChapasSection(lastGrouped);
  renderBarrasSection(lastGrouped);
  applySectionVisibility();
  els.contentSection.hidden = false;
  els.exportSection.hidden = false;
  els.previewPanel.hidden = true;
}

// reconstrói tudo (reclassificação chapa/perfil/outro, reagrupamento, seções 4/5, export) a
// partir do estado atual da tabela da seção 2 — chamado depois que o usuário edita, adiciona
// ou remove uma linha ali. Reagrupar de novo é o que deixa dois grupos com a mesma
// especificação/descrição se juntarem quando a correção fizer o material (ou outro campo)
// bater entre eles.
function reprocessGrouped(){
  const cleaned = lastGrouped
    .map(g=>({
      especificacao: normalizeText(g.especificacao),
      descricao: normalizeText(g.descricao),
      material: normalizeMaterial(g.material),
      qtd: Number(g.qtd)||0,
      massa: Number(g.massa)||0,
    }))
    .filter(g=> g.especificacao || g.descricao);

  if(!cleaned.length){
    els.parseStatus.textContent = "Nenhum item com Especificação ou Descrição preenchida — nada para reagrupar.";
    els.parseStatus.className = "status-msg err";
    return;
  }

  const before = lastGrouped.length;
  lastGrouped = sortByDescricao(groupRows(cleaned));

  els.parseStatus.textContent = `Reagrupado: ${cleaned.length} item(ns) → ${lastGrouped.length} grupo(s)` + (before!==lastGrouped.length ? ` (era ${before}).` : ".");
  els.parseStatus.className = "status-msg ok";

  renderGroupedTable(lastGrouped);
  renderChapasSection(lastGrouped);
  renderBarrasSection(lastGrouped);
  applySectionVisibility();
  els.contentSection.hidden = false;
  els.exportSection.hidden = false;
  els.previewPanel.hidden = true;
}
els.reprocessBtn.addEventListener("click", reprocessGrouped);

els.addGroupedRowBtn.addEventListener("click", ()=>{
  lastGrouped.push({especificacao:"", descricao:"", material:"", qtd:0, massa:0});
  renderGroupedTable(lastGrouped);
  els.groupedSection.hidden = false;
  els.groupedBody.hidden = false;
  // foca a primeira célula da linha recém-criada
  const inputs = els.groupedTable.querySelectorAll(`tr[data-idx="${lastGrouped.length-1}"] input`);
  if(inputs[0]) inputs[0].focus();
});

els.processBtn.addEventListener("click", ()=>{
  const text = els.pasteArea.value;
  if(!text.trim()){
    els.parseStatus.textContent = "Cole os dados da planilha antes de processar.";
    els.parseStatus.className = "status-msg err";
    return;
  }
  const {rows, problems} = parseRows(text);
  runPipeline(rows, problems, "texto colado");
});

els.xlsxInput.addEventListener("change", async ()=>{
  const file = els.xlsxInput.files[0];
  if(!file) return;
  els.parseStatus.textContent = `Lendo ${file.name}…`;
  els.parseStatus.className = "status-msg";
  try{
    const buf = await file.arrayBuffer();
    const {rows, problems} = await parseXlsxFile(buf);
    runPipeline(rows, problems, `arquivo "${file.name}"`);
  }catch(err){
    els.parseStatus.textContent = "Não consegui ler o arquivo: " + (err && err.message ? err.message : err);
    els.parseStatus.className = "status-msg err";
  }
});

let downloadsCap = null;
async function initDownloads(){
  try{
    if(window.claude && typeof window.claude.use === "function") downloadsCap = await window.claude.use("downloads");
  }catch(e){ downloadsCap = null; }
}
initDownloads();

// tenta baixar direto no navegador (funciona em qualquer aba normal — GitHub Pages, ou o
// arquivo index.html aberto localmente). Dentro do preview do Claude isso fica bloqueado
// pelo sandbox, então nesse caso a mensagem final orienta a abrir a página "de verdade".
function nativeDownload(filename, data, statusEl){
  try{
    const blob = (data instanceof Blob) ? data : new Blob([data]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url), 2000);
    statusEl.textContent = "Baixado!"; statusEl.className = "status-msg ok";
    return true;
  }catch(e){
    return false;
  }
}

async function saveFile(filename, data, statusEl){
  const inClaudePreview = !!(window.claude && typeof window.claude.use === "function");

  if(!inClaudePreview){
    if(!nativeDownload(filename, data, statusEl)){
      statusEl.textContent = "Não foi possível baixar."; statusEl.className = "status-msg err";
    }
    return;
  }

  // dentro do preview do Claude: .xlsx não está na lista de formatos liberados para
  // download (só imagens, txt/json/md, docx/pptx/csv/html/svg/pdf) — então isso sempre
  // será recusado aqui, mesmo com a permissão de download habilitada. Avisa e orienta a
  // abrir a página publicada (GitHub Pages) ou o arquivo .html localmente, onde o
  // download funciona normalmente.
  if(downloadsCap){
    statusEl.textContent = "Salvando…"; statusEl.className = "status-msg";
    try{
      await downloadsCap.save({filename, data});
      statusEl.textContent = "Salvo!"; statusEl.className = "status-msg ok";
      return;
    }catch(err){
      const code = err && err.code;
      const msg = code==="declined" ? "Cancelado."
        : code==="rejected_extension" ? "Arquivo .xlsx não pode ser baixado neste preview do Claude — abra a página publicada (GitHub Pages) ou o arquivo .html direto no navegador para baixar."
        : code==="extension_not_enabled" ? "Formato não habilitado nesta visualização."
        : "Não foi possível salvar aqui — abra a página publicada (GitHub Pages) ou o arquivo .html direto no navegador para baixar.";
      statusEl.textContent = msg; statusEl.className = "status-msg err";
      return;
    }
  }
  statusEl.textContent = "Download indisponível neste preview — abra a página publicada (GitHub Pages) ou o arquivo .html direto no navegador para baixar.";
  statusEl.className = "status-msg err";
}

els.exportXlsxBtn.addEventListener("click", ()=>{
  if(!lastGrouped.length) return;
  const blob = buildXlsxWorkbookBlob([buildListaCompactadaSheet(lastGrouped), buildResumoCorteSheet()]);
  saveFile("lista_compactada.xlsx", blob, els.exportStatus);
});

// monta as linhas no "formato genérico" (mesmo shape de um item da lista compactada:
// especificacao/descricao/material/qtd/massa, com `pecas` opcional) a partir de cada
// conteúdo disponível para exportação, para reaproveitar os mesmos geradores de planilha
// (simples / Lista Preliminar / Lista consolidada) nos três casos.
// lista compactada ("lista simples") — todos os itens, sem otimização, em ordem alfabética
// por descrição.
function contentRowsCompactada(){
  return sortByDescricao(lastGrouped.map(g=>({especificacao:g.especificacao, descricao:g.descricao, material:g.material, qtd:g.qtd, massa:g.massa})));
}
// itens que não são chapa nem barra/tubo (parafusos, arruelas, curvas, grades etc.) — a
// categoria "Outro", que não entra na otimização de corte mas ainda precisa aparecer numa
// lista de compra própria (é o uso original da aba "LLI" do modelo — "Loose Load Items").
// Também sai em ordem alfabética por descrição — é a mesma lista usada dentro da lista
// consolidada e do resumo de corte.
function contentRowsOutro(){
  return sortByDescricao(lastGrouped
    .filter(isLooseOrUnclassified)
    .map(g=>({especificacao:g.especificacao, descricao:g.descricao, material:g.material, qtd:g.qtd, massa:g.massa})));
}
// título (curto, categoria) x especificação/spec (detalhada): título = a "Especificação"
// original que o usuário digitou na planilha (ex: "Chapa de Aço", "Tubo", "Perfil
// Estrutural") — igual ao que a "LISTA PRELIMINAR" de referência guarda na coluna TitlePT;
// spec = o detalhe técnico (espessura/bitola + tamanho comercial), igual à coluna SPECPT.
function chapaTitulo(r){
  return r.especOriginal || (r.circular ? "Chapa Circular" : "Chapa");
}
function barraTitulo(r){
  return r.especOriginal || "Perfil/Tubo";
}
// pra colunas com um texto só (sem título/spec separados): junta a especificação original
// (ex: "Tubo circular", "Perfil Estrutural") com a identidade, mas só quando a identidade
// ainda não começa com essa mesma palavra — em muitas linhas a descrição original já
// começa com "Tubo ..." e duplicar viraria "Tubo Tubo Ø12" SCH.80...".
function barraEspecificacaoCombinada(r){
  const espec = (r.especOriginal || "").trim();
  const ident = r.identidade || "";
  if(!espec) return ident;
  const firstWord = espec.split(/\s+/)[0].toLowerCase();
  if(firstWord && ident.toLowerCase().startsWith(firstWord)) return ident;
  return `${espec} ${ident}`.trim();
}
function contentRowsResumoCorte(){
  const rows = [];
  lastChapaResults.forEach(r=>{
    if(!r) return;
    rows.push({
      especificacao: chapaTitulo(r),
      descricao: chapaThicknessLabel(r) + " — " + r.comercialLabel + " (" + r.count + " un.)",
      material: r.material, qtd: r.count, massa: r.massaTotal,
    });
  });
  lastBarraResults.forEach(r=>{
    if(!r) return;
    rows.push({
      especificacao: barraTitulo(r),
      descricao: r.identidade + " — " + r.comercialLabel + " (" + r.count + " un.)",
      material: r.material, qtd: r.count, massa: r.massaTotal,
    });
  });
  // é uma lista de material do projeto — se o item faz parte da obra, ele aparece aqui
  // também, mesmo não passando pela otimização de corte (parafusos, arruelas, curvas etc.)
  rows.push(...contentRowsOutro());
  return rows;
}
// área e comprimento viram DUAS listas separadas (uma aba cada nos formatos com modelo) —
// título curto na especificação, detalhe (espessura/bitola) na descrição, e `pecas` à parte
// para não perder a contagem de peças que a quantidade principal (área/comprimento) substituiu.
function contentRowsAreaChapas(){
  return sortByDescricao(lastChapaResults.filter(Boolean).map(r=>({
    especificacao: chapaTitulo(r),
    descricao: chapaThicknessLabel(r),
    material: r.material, qtd: Math.round(r.areaM2*100)/100, pecas: r.pecas, massa: r.massaTotal,
  })));
}
function contentRowsComprimentoPerfis(){
  return sortByDescricao(lastBarraResults.filter(Boolean).map(r=>({
    especificacao: barraTitulo(r),
    descricao: r.identidade,
    material: r.material, qtd: Math.round(r.lengthM*100)/100, pecas: r.pecas, massa: r.massaTotal,
  })));
}
function buildSimpleGenericSheet(rows, sheetName){
  const headers = ["Item","Qtd","Título/Especificação","Descrição/Spec","Material","Massa (kg)"];
  const dataRows = rows.map((g,i)=>[String(i+1), g.qtd, g.especificacao, g.descricao, g.material, Math.round((g.massa||0)*100)/100]);
  return {name:sheetName, headers, rows:dataRows};
}

/* ---------------- pré-visualização (antes de gerar o Excel) ---------------- */
// monta as mesmas seções/abas que a geração real vai escrever, na mesma ordem — reaproveita
// os geradores de linha (contentRows*, buildConsolidadoCortavelRows/OutroRows) e a mesma
// paginação (paginateConsolidadoOutro) usada pelo formato "Lista consolidada", pra não ter risco da prévia
// mostrar uma coisa e o arquivo gerado sair com outra.
function buildPreviewSections(content, format){
  if(content==="consolidado"){
    if(format==="consolidada"){
      return buildConsolidadoPages(buildConsolidadoCortavelRows(), buildConsolidadoOutroRows(), currentPageOpts())
        .map(p=>({
          label: p.tipo==="cortavel"
            ? `Aba "${p.name}" — itens cortáveis (chapas + perfis/tubos)`
            : `Aba "${p.name}" — itens soltos`,
          rows: p.rows,
        }));
    }
    return [
      {label:"Chapas — Área (m²)", rows:contentRowsAreaChapas()},
      {label:"Perfis/Tubos — Comprimento (m)", rows:contentRowsComprimentoPerfis()},
      {label:"Itens soltos", rows:contentRowsOutro()},
    ];
  }
  if(content==="corte") return [{label:"Resumo de corte (ordem do plano de corte)", rows:contentRowsResumoCorte()}];
  if(content==="outro") return [{label:"Itens soltos", rows:contentRowsOutro()}];
  return [{label:"Lista compactada", rows:contentRowsCompactada()}];
}
// cada linha de conteúdo pode ter a quantidade num campo diferente (qtd, área, comprimento ou
// qty) dependendo de que gerador a produziu — mostra o que estiver preenchido.
function previewRowValue(r){
  if(r.area!=null) return fmt(r.area,2)+" m²";
  if(r.comprimento!=null) return fmt(r.comprimento,2)+" m";
  if(r.qty!=null) return fmt(r.qty,2);
  if(r.qtd!=null) return fmt(r.qtd,2);
  return "—";
}
function renderPreviewSection(label, rows){
  if(!rows.length) return "";
  let html = `<div class="preview-page"><h4>${escapeHtml(label)}<span class="preview-count">${rows.length} item(ns)</span></h4>`;
  html += `<div class="table-wrap"><table class="data-table"><thead><tr><th class="num">#</th><th>Especificação</th><th>Descrição</th><th>Material</th><th class="num">Qtd / Medida</th></tr></thead><tbody>`;
  rows.forEach((r,i)=>{
    html += `<tr><td class="num">${i+1}</td><td>${escapeHtml(r.especificacao)}</td><td>${escapeHtml(r.descricao||"")}</td><td>${escapeHtml(r.material)}</td><td class="num">${previewRowValue(r)}</td></tr>`;
  });
  html += `</tbody></table></div></div>`;
  return html;
}
els.previewBtn.addEventListener("click", ()=>{
  const content = els.exportContentSel.value;
  const format = els.exportFormatSel.value;
  const sections = buildPreviewSections(content, format);
  const total = sections.reduce((s,sec)=>s+sec.rows.length, 0);
  els.previewBody.innerHTML = total
    ? sections.map(s=>renderPreviewSection(s.label, s.rows)).join("")
    : `<p class="help">Nada para pré-visualizar com essa combinação de conteúdo.</p>`;
  els.previewPanel.hidden = false;
  if(typeof els.previewPanel.scrollIntoView === "function") els.previewPanel.scrollIntoView({behavior:"smooth", block:"start"});
});
els.previewCloseBtn.addEventListener("click", ()=>{ els.previewPanel.hidden = true; });

els.exportFinalBtn.addEventListener("click", async ()=>{
  const content = els.exportContentSel.value;
  const format = els.exportFormatSel.value;

  if(content==="consolidado"){
    const areaRows = contentRowsAreaChapas();
    const compRows = contentRowsComprimentoPerfis();
    const outroRows = contentRowsOutro();
    if(!areaRows.length && !compRows.length && !outroRows.length){
      els.exportFinalStatus.textContent = "Nada para exportar.";
      els.exportFinalStatus.className = "status-msg err";
      return;
    }
    els.exportFinalStatus.textContent = "Gerando…"; els.exportFinalStatus.className = "status-msg";
    try{
      let blob, filename;
      if(format==="consolidada"){
        blob = await buildConsolidadoLLIXlsx(true, currentPageOpts());
        filename = "lista_consolidada_LLI.xlsx";
      } else if(format==="lista"){
        blob = await buildListaXlsxFromTemplate([{name:"Área (m²)", rows:areaRows}, {name:"Comprimento (m)", rows:compRows}, {name:"Itens Soltos", rows:outroRows}]);
        filename = "lista_consolidada_lista_preliminar.xlsx";
      } else {
        blob = buildXlsxWorkbookBlob([buildChapasM2Sheet(), buildPerfisMSheet(), buildSimpleGenericSheet(outroRows, "Itens Soltos")]);
        filename = "lista_consolidada.xlsx";
      }
      await saveFile(filename, blob, els.exportFinalStatus);
      // resumo de quantas linhas de cada categoria foram geradas — pra conferir sem precisar
      // rolar a planilha até o fim (itens soltos ficam nas últimas linhas da lista)
      if(els.exportFinalStatus.className.includes("ok")){
        let onde = "";
        if(outroRows.length && format==="consolidada"){
          // pergunta às páginas de verdade em que aba cada parte caiu — sem itens cortáveis,
          // os soltos começam na própria "LLI", não na "LLI (2)"
          const pages = buildConsolidadoPages(buildConsolidadoCortavelRows(), buildConsolidadoOutroRows(), currentPageOpts());
          const abasOutro = pages.filter(p=>p.tipo==="outro").map(p=>p.name);
          const abaCortavel = (pages.find(p=>p.tipo==="cortavel" && p.rows.length) || {}).name;
          const parteCortavel = abaCortavel ? `itens cortáveis na aba "${abaCortavel}", ` : "";
          const parteOutro = abasOutro.length>1
            ? `itens soltos nas abas "${abasOutro[0]}" a "${abasOutro[abasOutro.length-1]}"`
            : `itens soltos na aba "${abasOutro[0]}"`;
          onde = ` — ${parteCortavel}${parteOutro}`;
        } else if(outroRows.length){
          onde = ` — itens soltos na aba "Itens Soltos"`;
        }
        els.exportFinalStatus.textContent += ` (${areaRows.length} chapa(s), ${compRows.length} perfil(is)/tubo(s), ${outroRows.length} item(ns) solto(s)${onde})`;
      }
    }catch(err){
      els.exportFinalStatus.textContent = "Erro ao gerar: " + (err && err.message ? err.message : err);
      els.exportFinalStatus.className = "status-msg err";
    }
    return;
  }


  let rows, sheetName, filenameBase;
  if(content==="corte"){ rows = contentRowsResumoCorte(); sheetName = "Resumo de Corte"; filenameBase = "resumo_corte"; }
  else if(content==="outro"){ rows = contentRowsOutro(); sheetName = "Itens Soltos"; filenameBase = "itens_soltos"; }
  else { rows = contentRowsCompactada(); sheetName = "Lista Compactada"; filenameBase = "lista_compactada"; }

  if(!rows.length){
    els.exportFinalStatus.textContent = "Nada para exportar com essa combinação de conteúdo.";
    els.exportFinalStatus.className = "status-msg err";
    return;
  }

  els.exportFinalStatus.textContent = "Gerando…"; els.exportFinalStatus.className = "status-msg";
  try{
    let blob, filename;
    if(format==="lista"){
      blob = await buildListaXlsxFromTemplate([{name:"Sheet1", rows}]);
      filename = `${filenameBase}_lista_preliminar.xlsx`;
    } else if(format==="consolidada"){
      blob = await buildLLIXlsxFromTemplate([{name:"LLI", rows}], content==="compactada" || content==="outro");
      filename = `${filenameBase}_lista_consolidada_LLI.xlsx`;
    } else {
      blob = buildXlsxWorkbookBlob([buildSimpleGenericSheet(rows, sheetName)]);
      filename = `${filenameBase}.xlsx`;
    }
    await saveFile(filename, blob, els.exportFinalStatus);
    if(content==="corte" && els.exportFinalStatus.className.includes("ok")){
      const nOutro = contentRowsOutro().length;
      if(nOutro){
        // linha 1 é sempre cabeçalho; no formato "Lista Preliminar" tem ainda uma linha de
        // título antes disso (dados começam na linha 3, não na 2)
        const headerRows = format==="lista" ? 2 : 1;
        const first = rows.length - nOutro + 1 + headerRows;
        els.exportFinalStatus.textContent += ` (${nOutro} item(ns) solto(s) nas linhas ${first} a ${rows.length+headerRows})`;
      }
    }
  }catch(err){
    els.exportFinalStatus.textContent = "Erro ao gerar: " + (err && err.message ? err.message : err);
    els.exportFinalStatus.className = "status-msg err";
  }
});

// tabela da seção 2 é editável: cada célula de texto/número é um <input>, identificado por
// data-idx (posição em lastGrouped) + data-field. A coluna "Interpretação" só mostra como a
// linha está sendo lida agora (não é editável) — ela é recalculada a cada tecla, mas o
// reagrupamento em si (reclassificar, juntar grupos que passaram a bater) só acontece quando
// o usuário clica em "↻ Reagrupar", pra não ficar reordenando a tabela embaixo do cursor
// enquanto a pessoa ainda está digitando.
function renderGroupedTable(grouped){
  let html = `<thead><tr>
    <th>Especificação</th><th>Descrição</th><th>Material</th>
    <th class="num">Qtd</th><th class="num">Massa (kg)</th><th>Interpretação</th><th></th>
  </tr></thead><tbody>`;
  grouped.forEach((g,idx)=>{
    const cls = classifyGroup(g);
    let tagHtml, interpretHtml;
    if(cls.tipo==="chapa"){
      tagHtml = `<span class="pill-tag chapa">Chapa</span>`;
      interpretHtml = cls.ok
        ? (cls.circular ? `e=${fmt(cls.thickness,2)}mm · Ø${fmt(cls.diametro,1)}mm` : `e=${fmt(cls.thickness,2)}mm · ${fmt(cls.width,1)}×${fmt(cls.height,1)}mm`)
        : `<span class="pill-tag warn">revisar descrição</span>`;
    } else if(cls.tipo==="barra"){
      tagHtml = `<span class="pill-tag barra">Barra/Tubo</span>`;
      interpretHtml = cls.ok
        ? `${escapeHtml(cls.identidade)} · L=${fmt(cls.length,1)}mm`
        : `<span class="pill-tag warn">revisar descrição</span>`;
    } else {
      tagHtml = `<span class="pill-tag outro">Outro</span>`;
      interpretHtml = `<span style="color:var(--muted);">não entra na otimização de corte</span>`;
    }
    html += `<tr data-idx="${idx}">
      <td>${tagHtml}<br><input class="cell-input" data-field="especificacao" value="${escapeHtml(g.especificacao)}"></td>
      <td><input class="cell-input" data-field="descricao" value="${escapeHtml(g.descricao)}"></td>
      <td><input class="cell-input" data-field="material" value="${escapeHtml(g.material)}"></td>
      <td class="num"><input class="cell-input num" data-field="qtd" value="${fmt(g.qtd,0)}"></td>
      <td class="num"><input class="cell-input num" data-field="massa" value="${fmt(g.massa,2)}"></td>
      <td data-interpret>${interpretHtml}</td>
      <td><button class="btn-icon" type="button" data-del title="Remover item">🗑</button></td>
    </tr>`;
  });
  html += "</tbody>";
  els.groupedTable.innerHTML = html;
  els.groupedSection.hidden = false;
  els.groupedBody.hidden = false;
  els.groupedToggleBtn.textContent = "▾ Minimizar";
}

// atualiza só a coluna "Interpretação" de uma linha, sem re-renderizar a tabela inteira —
// dá feedback imediato de como a edição está sendo lida, sem perder o foco/cursor do input.
function refreshRowInterpretation(idx){
  const g = lastGrouped[idx];
  if(!g) return;
  const row = els.groupedTable.querySelector(`tr[data-idx="${idx}"]`);
  if(!row) return;
  const cls = classifyGroup(g);
  let interpretHtml;
  if(cls.tipo==="chapa"){
    interpretHtml = cls.ok
      ? (cls.circular ? `e=${fmt(cls.thickness,2)}mm · Ø${fmt(cls.diametro,1)}mm` : `e=${fmt(cls.thickness,2)}mm · ${fmt(cls.width,1)}×${fmt(cls.height,1)}mm`)
      : `<span class="pill-tag warn">revisar descrição</span>`;
  } else if(cls.tipo==="barra"){
    interpretHtml = cls.ok
      ? `${escapeHtml(cls.identidade)} · L=${fmt(cls.length,1)}mm`
      : `<span class="pill-tag warn">revisar descrição</span>`;
  } else {
    interpretHtml = `<span style="color:var(--muted);">não entra na otimização de corte</span>`;
  }
  row.querySelector("[data-interpret]").innerHTML = interpretHtml;
}

els.groupedTable.addEventListener("input", (ev)=>{
  const input = ev.target.closest("input[data-field]");
  if(!input) return;
  const row = input.closest("tr[data-idx]");
  const idx = parseInt(row.dataset.idx, 10);
  const g = lastGrouped[idx];
  if(!g) return;
  const field = input.dataset.field;
  if(field==="qtd" || field==="massa") g[field] = toNumBR(input.value);
  else g[field] = input.value;
  refreshRowInterpretation(idx);
});

els.groupedTable.addEventListener("click", (ev)=>{
  const btn = ev.target.closest("button[data-del]");
  if(!btn) return;
  const row = btn.closest("tr[data-idx]");
  const idx = parseInt(row.dataset.idx, 10);
  lastGrouped.splice(idx, 1);
  renderGroupedTable(lastGrouped);
});

els.groupedToggleBtn.addEventListener("click", ()=>{
  els.groupedBody.hidden = !els.groupedBody.hidden;
  els.groupedToggleBtn.textContent = els.groupedBody.hidden ? "▸ Expandir" : "▾ Minimizar";
});

function escapeHtml(s){
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

/* ---------------- seção de chapas ---------------- */
function renderChapasSection(grouped){
  els.chapasToggleAllBtn.textContent = "▾ Minimizar todas as espessuras";
  const buckets = new Map();
  const naoClassificadas = [];

  // agrupa só por espessura + material — chapa circular e retangular da mesma espessura
  // são a mesma chapa comercial, então entram no mesmo grupo/corte (cada peça carrega seu
  // próprio formato, então o desenho mostra retângulos e círculos juntos na mesma folha).
  grouped.forEach(g=>{
    const cls = classifyGroup(g);
    if(cls.tipo !== "chapa") return;
    if(!cls.ok){ naoClassificadas.push(g); return; }
    const key = `${fmt(cls.thickness,2)}|${g.material}`;
    if(!buckets.has(key)){
      buckets.set(key, {thickness:cls.thickness, material:g.material, pieces:[], massaTotal:0, fracLabel:null, hasCircular:false, hasRect:false, especOriginal:null});
    }
    const b = buckets.get(key);
    b.massaTotal += g.massa;
    if(!b.fracLabel && cls.fracLabel) b.fracLabel = cls.fracLabel;
    // especificação original (ex: "Chapa de Aço") — usada como "título" curto nas planilhas
    // que separam título de especificação/spec detalhada
    if(!b.especOriginal) b.especOriginal = g.especificacao;
    if(cls.circular){
      b.hasCircular = true;
      b.pieces.push({w:cls.diametro, h:cls.diametro, qty:g.qtd, label:"Ø"+fmt(cls.diametro,0), diam:cls.diametro, circular:true});
    } else {
      b.hasRect = true;
      b.pieces.push({w:cls.width, h:cls.height, qty:g.qtd, label:`${fmt(cls.width,0)}×${fmt(cls.height,0)}`, circular:false});
    }
  });

  els.chapaGroups.innerHTML = "";
  chapasHasData = buckets.size>0 || naoClassificadas.length>0;
  lastChapaResults = [];
  if(!chapasHasData) return;

  let i = 0;
  for(const [key, b] of buckets){
    els.chapaGroups.appendChild(buildChapaGroupCard(b, i));
    i++;
  }
  if(naoClassificadas.length){
    const div = document.createElement("div");
    div.className = "group-card";
    div.innerHTML = `<h3>Chapas não interpretadas (${naoClassificadas.length})</h3>
      <div class="group-meta">A descrição não seguiu o padrão esperado (espessura entre parênteses + duas dimensões, ou Ø para chapa circular). Ajuste a descrição na planilha e cole novamente para entrar na otimização de corte — até lá, este item é exportado como item avulso ("Outro"), com a quantidade certa mas sem área calculada.</div>
      <div class="table-wrap"><table class="data-table"><tbody>
      ${naoClassificadas.map(g=>`<tr><td>${escapeHtml(g.especificacao)}</td><td>${escapeHtml(g.descricao)}</td><td>${escapeHtml(g.material)}</td><td class="num">${fmt(g.qtd,0)}</td></tr>`).join("")}
      </tbody></table></div>`;
    els.chapaGroups.appendChild(div);
  }
}

function buildChapaGroupCard(b, idx){
  const div = document.createElement("div");
  div.className = "group-card";
  const totalQty = b.pieces.reduce((s,p)=>s+p.qty,0);
  const shapeLabel = (b.hasCircular && b.hasRect) ? "chapas retangulares e circulares"
    : b.hasCircular ? "chapas circulares" : "chapas retangulares";

  const optionsHtml = CHAPAS_COMERCIAIS.map((c,ci)=>`<option value="${ci}">${c.label}</option>`).join("");

  div.innerHTML = `
    <div class="group-card-header">
      <h3>Espessura ${fmt(b.thickness,2)} mm — ${escapeHtml(b.material)}</h3>
      <button class="btn-toggle" type="button" data-chapa-toggle="${idx}">▾ Minimizar</button>
    </div>
    <div data-chapa-body="${idx}">
      <div class="group-meta">${totalQty} peça(s), ${b.pieces.length} tamanho(s) distinto(s) de ${shapeLabel}</div>
      <div class="group-controls">
        <div>
          <label>Chapa comercial</label><br>
          <select data-chapa-sel="${idx}">${optionsHtml}</select>
        </div>
        <div>
          <label>Perda de corte (mm)</label><br>
          <input type="number" data-chapa-kerf="${idx}" value="3" min="0" step="0.5" style="width:56px;">
        </div>
      </div>
      <div class="result-summary" data-chapa-summary="${idx}"></div>
      <div class="sheets-wrap" data-chapa-sheets="${idx}"></div>
    </div>
  `;

  const sel = div.querySelector(`[data-chapa-sel="${idx}"]`);
  const kerfInput = div.querySelector(`[data-chapa-kerf="${idx}"]`);
  const recompute = ()=> computeChapaGroup(b, idx, div);
  sel.addEventListener("change", recompute);
  kerfInput.addEventListener("input", recompute);
  // calculado na hora (não em setTimeout) — o botão "Pré-visualizar"/"Reagrupar" pode ser
  // clicado logo em seguida e precisa encontrar lastChapaResults/lastBarraResults já prontos.
  recompute();

  const toggleBtn = div.querySelector(`[data-chapa-toggle="${idx}"]`);
  const bodyEl = div.querySelector(`[data-chapa-body="${idx}"]`);
  toggleBtn.addEventListener("click", ()=>{
    bodyEl.hidden = !bodyEl.hidden;
    toggleBtn.textContent = bodyEl.hidden ? "▸ Expandir" : "▾ Minimizar";
  });
  return div;
}

function computeChapaGroup(b, idx, container){
  const sel = container.querySelector(`[data-chapa-sel="${idx}"]`);
  const kerfInput = container.querySelector(`[data-chapa-kerf="${idx}"]`);
  const comercial = CHAPAS_COMERCIAIS[Number(sel.value)];
  const kerf = parseFloat(kerfInput.value) || 0;

  const result = nestPlates(b.pieces, comercial.w, comercial.h, kerf);

  const summaryEl = container.querySelector(`[data-chapa-summary="${idx}"]`);
  summaryEl.innerHTML = `
    <div class="result-stat"><div class="n">${result.count}</div><div class="lbl">chapa(s) comercial(is)</div></div>
    <div class="result-stat"><div class="n">${(result.utilization*100).toFixed(0)}%</div><div class="lbl">aproveitamento</div></div>
    <div class="result-stat"><div class="n">${comercial.label}</div><div class="lbl">tamanho escolhido</div></div>
  `;

  const sheetsEl = container.querySelector(`[data-chapa-sheets="${idx}"]`);
  sheetsEl.innerHTML = "";
  result.sheets.forEach((sh,si)=>{
    const fig = document.createElement("div");
    fig.className = "sheet-fig";
    const svg = renderSheetSvg(sh, comercial.w, comercial.h);
    fig.innerHTML = `<div class="sheet-label">Chapa ${si+1} de ${result.count} — ${comercial.label} (${sh.placements.length} peça(s))</div>${svg}`;
    sheetsEl.appendChild(fig);
  });

  const totalPecas = b.pieces.reduce((s,p)=>s+p.qty,0);
  // área real das peças (não da chapa comercial) — usada no resumo em m²
  const areaMm2 = b.pieces.reduce((s,p)=> s + (p.circular ? Math.PI*Math.pow(p.diam/2,2) : p.w*p.h) * p.qty, 0);
  lastChapaResults[idx] = {
    thickness:b.thickness, material:b.material, circular:b.hasCircular && !b.hasRect, fracLabel:b.fracLabel,
    comercialLabel:comercial.label, count:result.count, utilization:result.utilization,
    pecas:totalPecas, massaTotal:b.massaTotal, areaM2: areaMm2/1e6, especOriginal:b.especOriginal,
  };
}

/* ---------------- seção de barras / tubos ---------------- */
function renderBarrasSection(grouped){
  els.barrasToggleAllBtn.textContent = "▾ Minimizar todos os perfis";
  const buckets = new Map();
  const naoClassificadas = [];

  grouped.forEach(g=>{
    const cls = classifyGroup(g);
    if(cls.tipo !== "barra") return;
    if(!cls.ok){ naoClassificadas.push(g); return; }
    const key = `${cls.identidade}|${g.material}`;
    if(!buckets.has(key)){
      buckets.set(key, {identidade:cls.identidade, material:g.material, pieces:[], massaTotal:0, especOriginal:null});
    }
    const b = buckets.get(key);
    b.massaTotal += g.massa;
    // especificação original (ex: "Tubo", "Perfil Estrutural", "Cantoneira") — usada como
    // "título" curto nas planilhas que separam título de especificação/spec detalhada
    if(!b.especOriginal) b.especOriginal = g.especificacao;
    b.pieces.push({len:cls.length, qty:g.qtd, label:fmt(cls.length,0)});
  });

  els.barraGroups.innerHTML = "";
  barrasHasData = buckets.size>0 || naoClassificadas.length>0;
  lastBarraResults = [];
  if(!barrasHasData) return;

  let i = 0;
  for(const [key, b] of buckets){
    els.barraGroups.appendChild(buildBarraGroupCard(b, i));
    i++;
  }
  if(naoClassificadas.length){
    const div = document.createElement("div");
    div.className = "group-card";
    div.innerHTML = `<h3>Perfis/tubos não interpretados (${naoClassificadas.length})</h3>
      <div class="group-meta">Não foi possível identificar o comprimento de corte na descrição. Ajuste a descrição na planilha e cole novamente para entrar na otimização de corte — até lá, este item é exportado como item avulso ("Outro"), com a quantidade certa mas sem comprimento calculado.</div>
      <div class="table-wrap"><table class="data-table"><tbody>
      ${naoClassificadas.map(g=>`<tr><td>${escapeHtml(g.especificacao)}</td><td>${escapeHtml(g.descricao)}</td><td>${escapeHtml(g.material)}</td><td class="num">${fmt(g.qtd,0)}</td></tr>`).join("")}
      </tbody></table></div>`;
    els.barraGroups.appendChild(div);
  }
}

function buildBarraGroupCard(b, idx){
  const div = document.createElement("div");
  div.className = "group-card";
  const totalQty = b.pieces.reduce((s,p)=>s+p.qty,0);
  // por padrão, todos os comprimentos comerciais ficam habilitados — permite fechar a
  // última barra num tamanho menor em vez de desperdiçar quase uma barra inteira maior.
  const checksHtml = BARRAS_COMERCIAIS.map((c,ci)=>
    `<label style="font-weight:400; text-transform:none; letter-spacing:0; font-size:13px; display:inline-flex; align-items:center; gap:4px; margin-right:12px;">
      <input type="checkbox" data-barra-size="${idx}" value="${ci}" checked> ${c.label}
    </label>`).join("");

  div.innerHTML = `
    <div class="group-card-header">
      <h3>${escapeHtml(b.identidade)} — ${escapeHtml(b.material)}</h3>
      <button class="btn-toggle" type="button" data-barra-toggle="${idx}">▾ Minimizar</button>
    </div>
    <div data-barra-body="${idx}">
      <div class="group-meta">${totalQty} peça(s), ${b.pieces.length} comprimento(s) distinto(s)</div>
      <div class="group-controls">
        <div>
          <label>Barras comerciais permitidas</label><br>
          ${checksHtml}
        </div>
        <div>
          <label>Perda de corte (mm)</label><br>
          <input type="number" data-barra-kerf="${idx}" value="3" min="0" step="0.5" style="width:56px;">
        </div>
      </div>
      <div class="result-summary" data-barra-summary="${idx}"></div>
      <div class="bars-wrap" data-barra-bars="${idx}"></div>
    </div>
  `;

  const sizeChecks = [...div.querySelectorAll(`[data-barra-size="${idx}"]`)];
  const kerfInput = div.querySelector(`[data-barra-kerf="${idx}"]`);
  const recompute = ()=> computeBarraGroup(b, idx, div);
  sizeChecks.forEach(chk=> chk.addEventListener("change", recompute));
  kerfInput.addEventListener("input", recompute);
  // calculado na hora (não em setTimeout) — o botão "Pré-visualizar"/"Reagrupar" pode ser
  // clicado logo em seguida e precisa encontrar lastChapaResults/lastBarraResults já prontos.
  recompute();

  const toggleBtn = div.querySelector(`[data-barra-toggle="${idx}"]`);
  const bodyEl = div.querySelector(`[data-barra-body="${idx}"]`);
  toggleBtn.addEventListener("click", ()=>{
    bodyEl.hidden = !bodyEl.hidden;
    toggleBtn.textContent = bodyEl.hidden ? "▸ Expandir" : "▾ Minimizar";
  });
  return div;
}

function computeBarraGroup(b, idx, container){
  const sizeChecks = [...container.querySelectorAll(`[data-barra-size="${idx}"]`)];
  const kerfInput = container.querySelector(`[data-barra-kerf="${idx}"]`);
  let checked = sizeChecks.filter(c=>c.checked).map(c=>BARRAS_COMERCIAIS[Number(c.value)]);
  if(!checked.length){ checked = BARRAS_COMERCIAIS; sizeChecks.forEach(c=>c.checked=true); }
  const kerf = parseFloat(kerfInput.value) || 0;

  const result = cutStock1DMulti(b.pieces, checked, kerf);

  // resume os tamanhos comerciais realmente usados, ex: "7× 12 m + 1× 6 m"
  const counts = new Map();
  result.bars.forEach(bar=> counts.set(bar.barLabel, (counts.get(bar.barLabel)||0)+1));
  const comercialSummary = [...counts.entries()].map(([label,n])=>`${n}× ${label}`).join(" + ");

  const summaryEl = container.querySelector(`[data-barra-summary="${idx}"]`);
  summaryEl.innerHTML = `
    <div class="result-stat"><div class="n">${result.count}</div><div class="lbl">barra(s) comercial(is)</div></div>
    <div class="result-stat"><div class="n">${(result.utilization*100).toFixed(0)}%</div><div class="lbl">aproveitamento</div></div>
    <div class="result-stat"><div class="n">${comercialSummary}</div><div class="lbl">combinação usada</div></div>
  `;

  const barsEl = container.querySelector(`[data-barra-bars="${idx}"]`);
  barsEl.innerHTML = "";
  result.bars.forEach((bar,bi)=>{
    const fig = document.createElement("div");
    fig.className = "bar-fig";
    const overflow = bar.cuts.some(c=>c.overflow);
    fig.innerHTML = `<div class="bar-label">Barra ${bi+1} de ${result.count} — ${bar.barLabel} (${bar.cuts.length} corte(s))${overflow ? ' <span class="pill-tag warn">peça maior que a barra!</span>' : ""}</div>${renderBarSvg(bar, bar.barLen)}`;
    barsEl.appendChild(fig);
  });

  const totalPecas = b.pieces.reduce((s,p)=>s+p.qty,0);
  // comprimento real das peças (não das barras comerciais) — usado no resumo em m
  const lengthMm = b.pieces.reduce((s,p)=> s + p.len*p.qty, 0);
  lastBarraResults[idx] = {
    identidade:b.identidade, material:b.material,
    comercialLabel:comercialSummary, count:result.count, utilization:result.utilization,
    pecas:totalPecas, massaTotal:b.massaTotal, lengthM: lengthMm/1000, especOriginal:b.especOriginal,
  };
}
