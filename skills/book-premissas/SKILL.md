---
name: book-premissas
description: Use when Raquel (Lançamentos/Interiores, Seazone) asks to migrate, update or validate a Spot's "Book de Premissas de Fachada e Interiores" — "migra o book do <Spot>", "atualiza o book de premissas", "deixa tudo DEFINIDO", "troca os MDFs do book pelo estilo", "manda os books pra Rachel", or mentions Rachel Souto's books, template v3/Campeche or "interiores das áreas comuns".
---

# Book de Premissas (migração para as regras v3)

## Overview
O book vai para o administrador de obras assinar: tudo DEFINIDO, estilo do Spot, MDFs do catálogo do estilo,
sem unidades. Mantém o layout do book da Rachel (cópia v2 no Google Slides). Regras e decisões-padrão:
`references/regras.md` — **ler antes de começar**.

## Fluxo (um Spot por vez)
1. **Ler** o book original: `py -3 scripts/deck.py get <ID> deck.json` e `deck.py text deck.json "MDF|metal|marcenaria|validar|pendente|definir|confirmar|unidade"`.
2. **Perguntar o estilo** (Bruma, Industrial, Clean, Biofílico). Nunca supor.
3. **Propor MDFs com imagens**: `swatches.py fetch` → baixar amostras originais do Drive → montar `compare` (book × catálogo, sugestão em verde) → mostrar o painel (Read no PNG para conferir e dar à Raquel o caminho absoluto clicável, ou abrir com `cmd //c start "" "<png>"`) → esperar aprovação. Nome solto não serve.
4. **Copiar** o original na mesma pasta (Drive `mcp__drive-raquel__GOOGLEDRIVE_COPY_FILE_ADVANCED`, nome "...Interiores das Áreas Comuns v2").
5. **Spec + fase1**: escrever `spec.json` (ver `migrate.py` doc) com as trocas de texto exatas (copiar do `deck.py text`), linhas/slide de unidades e slide de materiais. Rodar `migrate.py fase1 spec.json --dry`, conferir, rodar sem `--dry`. Trocas "sem ocorrência" = texto não bateu: corrigir e refazer só elas.
6. **Amostras**: `swatches.py swatch/color` + `upload` (reusar IDs de `regras.md`). Preencher `slide_mdf` e `trocar_imagens`; rodar `migrate.py fase2`.
7. **Conferir**: `deck.py thumbs` em capa, slide 2, divisor 02, texturas, slide de MDFs e comparativos; `deck.py text ... "A VALIDAR|PENDENTE|CRÍTICO|a definir|confirmar"` tem de voltar vazio (exceto nomes como "pendente" de luminária).
8. **Entregar** à Raquel: link + tabela das decisões tomadas pela regra (o que ela precisa revisar, com conflitos imagem × briefing em destaque). Registrar no `memory.md` do projeto.
9. **Slack** só com OK explícito (texto mostrado antes, salvo se ela ditou o conteúdo). Slack pelo Composio do Seazone Hub: ToolSearch "+seazone-hub composio" → `mcp__seazone-hub__composio-COMPOSIO_SEARCH_TOOLS` (toolkit SLACK, "send message") → `mcp__seazone-hub__composio-COMPOSIO_MULTI_EXECUTE_TOOL` com `tool_slug` `SLACK_SEND_MESSAGE`, `account` `slack_rial-kilim`, `arguments` {`channel`: `D07E8QJV751` (DM Rachel) ou `C0BD06X3KK6`, `markdown_text`}; nunca `mcp__slack-raquel`.

## Modo V03 (layout do Campeche v3 — padrão desde 14/10/2026)
Rachel pediu todos os books no layout igual ao Campeche v3. Gera um pptx novo (python-pptx + `scripts/v3/tplkit.py`)
a partir do template do repo dela e sobe convertido em Slides. Exemplos prontos: `C:\Users\Seazone\Claude\projetos\book-premissas-template\scripts\books_v3\build_<spot>.py`
(Batel, Cachoeira, Rosa Norte, Imbassaí, Itacaré, Ilha, Marista, Meireles, Salvador, Santo). Funções comuns em `scripts/v3/bookv3.py`.
1. Conteúdo = texto final da v2 (`deck.py get <v2> v2.json` + `deck.py text`). Nada novo inventado.
2. Imagens: legado `~/Claude/projetos/book-premissas-template/books/legado/<SPOT>_...pptx` (extrair amostras e marcadores, `old_markers(OLD, slide)`)
   + renders do Drive (`08/05 - Imagens 3D` e `08/10 - Projeto de Interiores/.../02 - Imagens`), reduzidos a 2400 px.
3. Copiar um `build_*.py` e trocar textos/imagens. Mapa: capa 1 · contexto 2 · como usar 3 · divisor 4 · moodboard 5 ·
   paleta 6 · leitura 7 · texturas 8 ×3 · vistas 11/12/13 · atenção 24 · divisor 14 · conceito 15 · fichas 27/30/32 ·
   moodboard interiores 28 (9 img) ou 31 (11 img) · atenção 24 · divisor 33 · materiais 34 ×2 · MDF 37 · fim 38.
4. `py -3 build_x.py DADOS out.pptx` → `C:\Users\Seazone\Claude\projetos\book-premissas-template\scripts\render.ps1` (PowerPoint COM) → folhas 2×3 → conferir marcadores, amostras
   (recorte com `focus,hfrac`), linhas vazias de ficha (preencher as 9) e títulos em 2 linhas.
5. Subir: `POST /api/v3/files/upload/request` + PUT → MCP `mcp__drive-raquel__GOOGLEDRIVE_RESUMABLE_UPLOAD` com `metadata.mimeType`
   = `application/vnd.google-apps.presentation`, na pasta do book, nome "... Áreas Comuns v3". Conferir `deck.py thumbs`.

## Common Mistakes
| Erro | Correção |
|---|---|
| Propor MDF só pelo nome | Sempre painel de imagens; tom pelo nome engana |
| `replaceAllText` deixa o card todo em negrito | A fase2 refaz rótulo em negrito + texto normal |
| Rótulo da capa quebra em 2 linhas | fase2 põe 9 pt |
| Título do divisor 02 empurra o subtítulo | fase2 desce o subtítulo |
| Sobra "unidades" em subtítulo de divisor/índice | Incluir nas `trocas` |
| Batch rejeitado por revisão | Alguém editou: reler e refazer (o script sempre lê antes) |
| V03: marcador "fora do quadro" | Imagem mais larga que o quadro: usar `hfrac` menor/focus, ou a imagem inteira e converter v |
| V03: amostra do legado ruim (céu, guarda-sol) | Recortar do render com `card(..., img, focus, hfrac)` |
| Inventar produto/cor para fechar "à definir" | Usar fonte existente (imagem/briefing/LOG/memorial) ou regra fixa (B208); senão perguntar |

Scripts usam a conta raquel.tavares no Composio (`BOOK_USER` muda). Chave lida da variável de usuário `COMPOSIO_API_KEY` (fallback legado: `.claude.json` do Claude Code), nunca impressa; nunca do config do app desktop nem de nota de memória.
