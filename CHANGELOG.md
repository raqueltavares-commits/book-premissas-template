# Changelog

## Brises por região + 10 books migrados + skill — 15/10/2026

- **Regra nova (reunião Rachel × Raquel):** brise de fachada nunca em PVC. Nordeste = madeira no mesmo tom, com a ressalva "pode ser substituído por alumínio com o mesmo tom e acabamento"; demais regiões = alumínio com o acabamento que já estava. Painéis ripados de interiores não mudam.
- Regra aplicada nos 10 books v3 (Batel, Cachoeira, Rosa Norte, Marista 144, Ilha do Campeche II e Santo Antônio em alumínio amadeirado; Imbassaí, Itacaré, Salvador e Meireles em madeira com a ressalva). Links do Drive mantidos.
- **Pendente:** o Canasvieiras v3 (`books/`) ainda cita brises em PVC preto e Oak; pela regra nova, vira alumínio preto e Oak.
- Novos: `skills/book-premissas/` (skill do Claude) e `scripts/books_v3/` (geradores dos 10 books).

## Regra geral: Papel de Presente no lugar de Branco Puro — 13/10/2026

- Nenhum empreendimento usa Branco Puro na fachada externa (suja muito). O claro da fachada passa a ser **Papel de Presente Suvinil (B208)**, como no Cachoeira Spot.
- Canasvieiras: REV-07 alterado na paleta, na textura T03, na leitura geral e na vista posterior. A legenda da leitura geral foi alinhada às decisões validadas (brises em PVC, guarda-corpo de vidro com perfil "U" inox, pergolado metálico preto, ripado executado).

## Canasvieiras · versão validada — 13/10/2026

- Todos os itens passam a DEFINIDO. Não há mais A VALIDAR nem PENDENTE, porque o book vai para a assinatura do administrador de obras.
- Regra aplicada: prevalecem o Briefing e a imagem de lançamento, e o executivo replica as texturas do lançamento. Decisões:
  - **Fachada:** brises em PVC preto e Oak, guarda-corpo de vidro com perfil "U" inox e pergolado metálico preto. EV3/EV4 com peças pretas e EV1 com a borda "pixelada" da imagem. O ripado horizontal das janelas é executado. Muro na cor Seazone com logo e nome iluminados. REV-05 e REV-08 seguem o tom da imagem.
  - **Interiores das áreas comuns:**
    - Paredes: recepção em cimento queimado tom areia, corredores "Outono Gelado" e rooftop cinza médio.
    - Portas e rodapés PRETOS.
    - Rooftop externo: piscina esverdeada 20×20, balanço fixado na laje, deck em porcelanato amadeirado 120×20 antiderrapante.
    - Marcenaria em MDF Savana e Cinza Sagrado.
- Páginas de pendências trocadas por "Pontos de atenção na obra".

## Canasvieiras Spot migrado para o template v3.1 — 13/10/2026

- Novo book com 27 slides: capa, contextualização, como usar, **01 Fachada** (moodboard, paleta, leitura geral, texturas T01–T12, vistas de esquina, posterior e térreo, pendências), **02 Interiores das áreas comuns** (conceito × situação, recepção e corredores, rooftop interno e externo, pendências), **03 Materiais de interiores das áreas comuns** (T13–T24) e encerramento.
- Saíram do book antigo: interiores e materiais de unidades (Cetim Bianco, MDF Cinza Urban, reforço de TV, jacuzzi, copa), áreas em m² e quadros de quantidade. As tabelas de comparação lançamento × executivo e briefing × executivo agora aparecem como status por item nas vistas, nas fichas e nas páginas de pendências.
- Gerador em `scripts/` para replicar nos próximos Spots.

## v3.1 — 13/10/2026 (comentário Bianca)

- **Capa:** imagem interna de unidade substituída por imagem externa (render de lançamento 002 · acesso), mesmo enquadramento.
- **Slide 2:** objetivo passa a falar em "premissas de fachada e de **interiores das áreas comuns**"; rodapé: "Foco: imagens externas, texturas e interiores das áreas comuns. **Interiores das unidades**, índices urbanísticos, áreas e custos não fazem parte deste Book."
- **Slide 3:** seção 02 renomeada de "Áreas comuns" para "**Interiores das áreas comuns**" (03 já era "Materiais de interiores das áreas comuns").
- Alterações feitas direto no Google Slides da v3 (mesmo link) e re-exportadas para este repositório.

## v3 — 13/10/2026 (Campeche Spot, rev. Raquel Tavares)

Template passa a ser o **Book do Campeche Spot** (`CAMPECHE SPOT - Book de Premissas_MOODBOARD_rev13-10-2026_v3`).

- MDFs ajustados ao catálogo Industrial Berneck:
  - Jequitibá (Arauco) → **Carvalho Treviso Berneck**
  - Beton Matt (Arauco) → **Ópera Berneck**
- Amostras de material atualizadas.
- Slide novo **"Materiais · MDFs e metais (4/4)"**: T25 MDF Carvalho Treviso, T26 MDF Ópera, T27 Metal preto RAL 9005, T28 Metal grafite RAL 7024.
- **Seção de interiores de unidades removida** — a partir desta versão nenhum book inclui interiores de unidades.

### Pontos ainda citando unidades na v3 (revisar ao replicar)
Alguns cards de materiais ainda mencionam unidades no texto: T13 ("parede da cama das unidades"), T17 ("e unidades"), T22 (marcenaria das unidades), T23 (bancadas das unidades), T24 (guarda-corpo do mezanino das unidades), T25/T26 ("e das unidades"). Ao gerar books novos, restringir esses itens às áreas comuns ou removê-los.

## v2 — 07/10/2026 (Campeche Spot, moodboard)

- Formato "moodboard" com nova recepção (rev. 07/10/2026), portas e rodapés brancos, forro do corredor V1 × V2.

## v1 — 23/09/2026

- Template original: Novo Campeche Spot III — Book de Premissas de Projeto (Projetos para Obras). Ver `template/legacy/`.
