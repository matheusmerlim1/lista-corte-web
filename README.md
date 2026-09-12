# Lista & Corte

Cole uma lista de materiais copiada do Excel, agrupe itens iguais e descubra quantas chapas e barras comerciais comprar — com o desenho de como distribuir o corte em cada uma.

## Como usar

1. Duas formas de entrada:
   - **Colar**: copie as colunas **Item, Qtd, Especificação, Descrição, Material, Massa** (nessa ordem) de uma planilha, cole no campo de texto e clique em **Processar lista colada**.
   - **Importar arquivo**: clique em **📂 Importar arquivo .xlsx** e escolha o arquivo Excel direto — mesmas 6 colunas, cabeçalho na primeira ou segunda linha (detecta sozinho).
2. Linhas sem especificação (separadores de seção) ou marcadas como "Esboço Remover" são descartadas automaticamente.
3. A ferramenta agrupa itens idênticos (mesma especificação + descrição + material), somando quantidade e massa.
4. Itens cuja "Especificação" contém **Chapa** viram grupos de corte 2D — chapas retangulares e circulares da mesma espessura e material entram no MESMO grupo (podem ser cortadas da mesma chapa comercial); itens com **Tubo, Perfil, Cantoneira, Barra, Viga** etc. viram grupos de corte 1D; o resto (parafusos, arruelas, curvas/joelhos, grades) aparece na tabela marcado como "Outro" e não entra na otimização — não são itens cortados de chapa/barra, mas podem ser exportados à parte (conteúdo "Itens soltos").
5. A tabela da seção **2 · Itens agrupados** é editável: qualquer célula (Especificação, Descrição, Material, Qtd, Massa) pode ser corrigida direto ali, **+ Adicionar item** cria uma linha em branco pra um item que não veio na lista colada/importada, e o 🗑 de cada linha remove um item. Depois de editar, clique em **↻ Reagrupar** — ele relê a tabela inteira e reagrupa/reclassifica do zero, então dois grupos que só existiam separados por causa de um dado errado (ex.: mesma descrição, material digitado diferente) se juntam automaticamente quando a correção fizer os dois baterem. A lista sai sempre em ordem alfabética por Descrição.
6. Seção **3 · O que você vai gerar?** — escolha o **Conteúdo** aqui antes de descer a página: as seções de otimização de corte (**4 · Chapas** e **5 · Perfis/Tubos**) só aparecem quando o conteúdo escolhido é o **"Resumo de corte"** (o único que depende de escolher chapa/comprimento comercial); nos outros conteúdos (lista consolidada em m²/m, compactada, itens soltos) essas seções ficam escondidas, mesmo que existam grupos cortáveis na lista.
7. Nas seções 4/5 (quando aparecem), marque os tamanhos comerciais permitidos de cada grupo cortável (chapa: um por vez; barra: pode marcar mais de um ao mesmo tempo — veja "Barras com mais de um comprimento" abaixo) — o resultado (quantas comprar, aproveitamento, desenho do corte) atualiza na hora. O botão **▾ Minimizar** de cada card recolhe só aquele grupo; **▾ Minimizar todas as espessuras / todos os perfis**, no topo de cada seção, recolhe (ou expande) todos de uma vez.
8. **⇩ Excel compactado + resumo de corte**, na seção 2 — planilha própria com a lista agrupada e um resumo (chapas/barras comerciais necessárias, aproveitamento, massa).
9. Seção **6 · Exportar**, no fim da página — escolha o **Formato** e clique em **👁 Pré-visualizar lista** pra ver exatamente como as linhas vão ficar organizadas (nas mesmas abas/ordem que o arquivo real vai sair) antes de gastar tempo gerando e baixando; quando o resultado parecer certo, clique em **⇩ Gerar Excel**. É uma lista de material do projeto: se o item faz parte da obra, ele aparece na exportação, mesmo os que não passam pela otimização de corte (parafusos, arruelas, curvas etc.) — a única coisa que muda entre "consolidado" e "resumo de corte" é como as chapas e os perfis/tubos são quantificados (por m²/m ou por chapa/barra comercial a comprar); os itens soltos entram do mesmo jeito nos dois.
   - **Conteúdo** (seção 3): lista consolidada (chapas em m² + perfis em m + itens soltos) / resumo de corte (chapas/barras comerciais a comprar + itens soltos) / lista de itens compactada (tudo, sem otimização) / itens soltos apenas (parafusos, arruelas, curvas... a categoria "Outro" — é o uso original da aba "LLI" do modelo, "Loose Load Items"). **A lista consolidada e a lista compactada saem em ordem alfabética por Descrição**; o resumo de corte mantém a ordem do plano de corte (mais útil ali do que ordem alfabética).
   - **Formato** (seção 6): Excel simples (planilha própria) / formato "Lista Preliminar" / **Lista consolidada (LLI)** — os dois últimos saem com o mesmo título, cabeçalho e formatação dos arquivos de referência, só os dados mudam. O formato **Lista consolidada** é um clone do arquivo de referência "LM Consolidada.xlsx": capa **sem logotipo, sem slogan e sem qualquer identificação de empresa, projeto ou pessoa**, com o título "LISTA DE MATERIAL CONSOLIDADA", e a lista na aba "LLI" com as 6 colunas Item No. | Especificação | Área total (m²) | Comprimento (m) | Material | Total QTY. A cada N itens é criada uma aba extra (LLI (2), LLI (3)...), com a área de impressão de cada aba ajustada ao número real de linhas.
   - No formato com **TitlePT + Spec/SPECPT separados** (Lista Preliminar, conteúdo "Resumo de corte"), TitlePT recebe a Especificação original que o usuário digitou na planilha (ex: "Chapa de Aço", "Tubo", "Perfil Estrutural") e Spec/SPECPT recebe o detalhe técnico (espessura/bitola + tamanho comercial escolhido) — igual à divisão da "LISTA PRELIMINAR" de referência original.
   - **Consolidado + Lista consolidada** gera a lista no layout do arquivo de referência do usuário ("LM Consolidada.xlsx"): Item No. | Especificação | Área total (m²) | Comprimento (m) | Material | Total QTY — cada linha preenche só a coluna que se aplica (chapa usa Área, perfil/tubo usa Comprimento, item solto usa Total QTY) e marca "-" nas outras. Os itens cortáveis (chapas + perfis/tubos, ordenados alfabeticamente) vão todos numa aba **"LLI"** só, sem paginar; os itens soltos (parafusos, arruelas, curvas etc., também ordenados alfabeticamente) vêm depois, em abas extras **"LLI (2)", "LLI (3)"...** — o Item No. continua contando sem reiniciar a cada aba nova. As abas são nomeadas pela posição, então uma lista **sem nenhum item cortável** (só parafusos, arruelas, curvas...) começa direto na **"LLI"**, sem deixar uma aba vazia na frente. A seção 6 tem um controle de **paginação dos itens soltos**: modo **automático** (estima quantas linhas cada descrição ocupa quebrada na largura da coluna e enche cada página até o limite de uma folha A4 impressa — itens com descrição mais longa cabem menos por página) ou **quantidade fixa** por página (você escolhe o número). O nome da especificação original (ex: "Tubo circular", "Perfil Estrutural") é adicionado antes da identidade do perfil/tubo quando ela ainda não menciona isso — assim "Ø48,30 SCH.40" vira "Tubo circular Ø48,30 SCH.40", mas "Tubo Ø12" SCH.80..." (que já começa com "Tubo") não fica duplicado.

### Barras com mais de um comprimento comercial

Cada grupo de barra/tubo tem uma caixa de seleção por comprimento comercial (6 m, 12 m) — os dois vêm marcados por padrão. Se só um estiver marcado, o corte usa só aquele tamanho; com mais de um marcado, a ferramenta escolhe o melhor comprimento pra cada barra nova (simulando o encaixe do restante da fila em cada opção antes de abrir a barra), então uma barra de 12m só é usada quando ela realmente aproveita bem o espaço — sobras pequenas fecham numa barra de 6m em vez de desperdiçar quase uma barra de 12m inteira.

## Como a descrição é interpretada

A ferramenta lê a coluna Descrição procurando por padrões como:

- `#1/2" (12,7 mm) x 301 x 130 mm` → chapa retangular, espessura 12,7 mm, 301×130 mm
- `Ø114,30 SCH.40 x 329,84 mm` → tubo circular Ø114,30 SCH.40, comprimento de corte 329,84 mm
- `#3/8" (9,3 mm) x Ø323,9 x 1422,50 mm` → tubo com parede 9,3 mm, Ø323,9 mm, comprimento 1422,50 mm
- `3/4" (19,05 mm) x Ø 482 mm` → chapa circular (disco), espessura 19,05 mm, Ø482 mm
- `3/4" (19,05 mm) x Ø1/2"` → chapa circular, espessura 19,05 mm (3/4"), Ø12,7 mm (1/2" — o diâmetro também pode vir em fração de polegada, não só em mm)

Itens que não seguem esses padrões aparecem em uma lista separada de "não interpretados", para ajuste manual da descrição na planilha — mas continuam saindo em todas as exportações como item avulso ("Outro"), com a quantidade certa, só sem área/comprimento calculado; nenhum item é descartado silenciosamente.

## Algoritmos

- **Chapas (2D):** heurístico de encaixe por prateleiras (shelf packing), com rotação de 90° permitida. Chapas circulares são encaixadas pelo quadrado que as envolve.
- **Barras/tubos (1D):** *first-fit decreasing* — ordena os cortes do maior para o menor e preenche as barras na ordem, considerando a perda de corte (serra) informada. Com mais de um comprimento comercial marcado, antes de abrir cada barra nova a ferramenta simula o encaixe guloso do restante da fila em cada tamanho disponível e escolhe o que dá o melhor aproveitamento — assim uma sobra pequena fecha numa barra menor em vez de abrir uma barra grande quase vazia.

Nenhum dos dois é o ótimo matemático absoluto (isso é um problema NP-difícil), mas são heurísticas padrão de mercado, com resultado tipicamente próximo do ótimo — suficiente para estimar quantidade de compra. Sempre confira com o encarregado de corte antes de fechar o pedido.

## Leitura e escrita de .xlsx

Não usa nenhuma biblioteca externa (SheetJS, ExcelJS etc.) — o `.xlsx` é lido e escrito na mão em JavaScript puro: o formato é um ZIP com XML dentro, e a descompressão usa a API nativa do navegador `DecompressionStream`. Isso funciona em qualquer navegador atual (Chrome, Edge, Firefox, Safari recentes); navegadores muito antigos podem não suportar.

Os formatos "Lista Preliminar" e "Lista consolidada", na seção **6 · Exportar**, funcionam clonando os arquivos de referência reais: o `.xlsx` original fica embutido na página em base64, é lido pelo mesmo leitor de ZIP, e só a área de dados da planilha é reescrita (título, cabeçalho, cores e estilos do modelo são preservados).

Os dois modelos embutidos foram preparados uma vez e **limpos antes de entrar no repositório** — eles guardam só a moldura da planilha (título, cabeçalho, estilos, largura de coluna e configuração de página), nunca dados de projeto:

- **Lista consolidada** — sobraram as abas "Cover Page" e "LLI" (essa vazia: o cabeçalho e uma linha guardando o estilo de cada coluna). Saíram as fórmulas (`calcChain.xml`), a lista de material do projeto de origem e os campos de autor. O arquivo de referência já não tinha logotipo nem menção a empresa, projeto ou pessoa, então **nada disso existe no modelo nem nos arquivos gerados** — não há sequer imagem dentro do pacote.
- **Lista Preliminar** — sobraram a linha de título e a linha de cabeçalho. As linhas de material do projeto de origem (que o gerador nunca usou: ele reescreve tudo a partir da linha 3), o histórico de revisão e o nome de quem salvou o arquivo foram removidos.

Os `.xlsx` de referência que deram origem aos modelos **não fazem parte deste repositório** (estão no `.gitignore`): eles carregam a lista de material completa de um projeto real. A ferramenta não precisa deles — os modelos já vão embutidos no `app.js`.

## Publicar no GitHub Pages

1. Suba esta pasta para um repositório no GitHub (o `index.html` na raiz do repositório, ou na pasta escolhida no passo 3).
2. **Settings → Pages** → escolha a branch e a pasta.
3. A página fica em `https://<seu-usuário>.github.io/<nome-do-repo>/`.

## Estrutura dos arquivos

- `index.html` — só a estrutura da página (marcação HTML).
- `styles.css` — todo o CSS (tema claro/escuro, layout, tabela editável, cartões de grupo, desenhos de corte).
- `app.js` — toda a lógica (parsing/agrupamento, interpretação da descrição, algoritmos de corte, leitura/escrita de `.xlsx`, geração das planilhas, wiring da interface) e os dois modelos de planilha embutidos em base64. Comentado por seção no topo do arquivo.
- `.gitignore` — mantém fora do repositório as planilhas de projeto real que deram origem aos modelos.

Os três arquivos precisam ficar na mesma pasta (é assim que o `index.html` os referencia). Nenhum dos dois (CSS/JS) usa bundler nem módulos — são carregados direto pela tag `<link>`/`<script>` do `index.html`.

## Rodar localmente

Abra `index.html` direto no navegador — não precisa de servidor nem instalação.

## Sobre o preview do Claude x a página publicada

O link de preview do Claude (o gerado ao conversar com o Claude sobre esse projeto) roda dentro de um sandbox que só libera download de alguns formatos (imagens, txt/json/md, docx/pptx/csv/html/svg/pdf) — **`.xlsx` não está nessa lista**, então os botões de exportar Excel não conseguem baixar nada ali, mesmo com a permissão concedida. Isso é uma limitação do preview, não um bug da ferramenta.

Para os downloads de `.xlsx` funcionarem, abra a ferramenta de uma dessas formas:
- direto o `index.html` no navegador (duplo clique no arquivo), ou
- pela página publicada no GitHub Pages, depois de seguir os passos acima.

Em qualquer uma dessas duas, o download funciona normalmente (é um `<a download>` comum, sem sandbox no meio).
