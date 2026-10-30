# Regras e decisões-padrão dos books de premissas

Fonte das regras: Raquel Tavares (Lançamentos/Interiores) + README do repositório da Rachel Souto
(github.com/rachelsouto-source/book-premissas-template). Dona do formato: Rachel Souto.

## Regras fixas
- O book vai para o **administrador de obras assinar**: nenhum item fica como A VALIDAR, PENDENTE, A COTAR,
  "à definir", "confirmar" ou CRÍTICO. Todo item vira **DEFINIDO**, com a decisão escrita.
- **Prevalecem a imagem de lançamento, o conceito Seazone, o Briefing e o LOG validado pela diretoria.** Onde o
  executivo/memorial/projeto legal diverge, escrever a decisão + "compatibilizar o executivo/memorial".
- Fachada: claro de fachada **nunca Branco Puro** → **Papel de Presente Suvinil (B208)**. Cor "à definir" de parede
  clara externa também vira B208.
- **"Interiores" → "Interiores das áreas comuns"** na capa, títulos, cabeçalhos de seção, rodapé e nome do arquivo.
  Não mexer em nome de documento-fonte ("memorial de interiores", "10 - Projeto de Interiores", "renders de interiores").
- **Estilo do Spot** na capa ("· Estilo X") e no rodapé do slide 2 ("Estilo de interiores: X."). Só o nome do
  estilo: não mencionar exceção de catálogo.
- **Unidades fora:** apagar o slide "Unidades: acabamentos e escopo" e as linhas de unidades nos comparativos
  (piso/paredes/banheiro/marcenaria/nicho/hidro/portas das unidades). O slide de **materiais das unidades fica**
  (como no Campeche), renomeado "Materiais das unidades (2/3)".
- **MDF/metal** → equivalentes do **catálogo do estilo** (moodboards REV.02, Berneck). Proposta SEMPRE com o
  painel de imagens (`swatches.py compare`) e aprovação da Raquel ANTES de aplicar. Se a Raquel escolher fora do
  catálogo, aceitar (ela decide) e registrar como exceção só na memória do projeto.
- Novo slide **"Materiais · MDFs e metais (3/3)"** (duplicando o slide de materiais das unidades) com um card por
  MDF/metal aprovado; cada "Obra:" diz "compatibilizar memorial de interiores (<MDF antigo>)".
- Sempre **Google Slides**, cópia na **mesma pasta** do book da Rachel, nome
  "<SPOT> - Book de Premissas de Fachada e Interiores das Áreas Comuns v2" (v3, v4... se já existir). Original intacto.
- Slack para a Rachel só com OK explícito da Raquel, como r.tavares, assinado "_Enviado por Claude_". Slack pelo Composio do Seazone Hub: ToolSearch "+seazone-hub composio" → `mcp__seazone-hub__composio-COMPOSIO_SEARCH_TOOLS` (toolkit SLACK, "send message") → `mcp__seazone-hub__composio-COMPOSIO_MULTI_EXECUTE_TOOL` com `tool_slug` `SLACK_SEND_MESSAGE`, `account` `slack_rial-kilim`, `arguments` {`channel`: `D07E8QJV751` (DM Rachel) ou `C0BD06X3KK6`, `markdown_text`}; nunca `mcp__slack-raquel`.

## Catálogo por estilo (MDF Berneck · metal)
| Estilo | MDF | Metal |
|---|---|---|
| Bruma | Nude, Ópera, Argento, Ruggine, Nero, Cacau, Louro Preto (Premium Mostrato) | chumbo, preto, bronze |
| Industrial | Cinza Cristal, Cinza Cobalto, Crômio, Argento, Linen Grigio, Chumbo, Ópera, Nero, Solanum, Parquet, Carvalho Treviso (Premium Azul) | preto RAL 9005, grafite RAL 7024 |
| Clean | Cinza Cobalto, Crômio, Gelo, Latte, Desert, Falésia, Lana, Carvalho Treviso, Roble Catedral (Premium Verti) | branco 9003, bege 1013, pearl 1036, cinza claro 7035 |
| Biofílico | Lana, Cinza Argila, Desert, Cinza Cobalto, Nogal Sevilha, Carvalho Treviso, Roble Catedral, Verti (Premium Ceramik) | branco 9003, bege 1013, cinza claro 7035, pearl 1036 |
Acabamentos (para o texto da especificação): Carvalho Treviso (Design), Roble Catedral (Grann), Ópera (Micro),
Chumbo (Micro), Crômio (Linha Smart, Vel), Cinza Cobalto (Tx), Cinza Cristal (Tx), Desert (Linha Smart, Vel).

Cor chapada dos metais (para `swatches.py color`, não há foto no moodboard): preto RAL 9005 `28,28,30` ·
grafite RAL 7024 `56,62,66` · branco RAL 9003 `236,236,231` · bege RAL 1013 `234,230,202` · pearl RAL 1036
`112,106,96` · cinza claro RAL 7035 `203,208,204` · Bruma (sem RAL): preto `28,28,30`, chumbo `70,72,74`, bronze `140,105,70`.
MDFs Premium (Mostrato, Azul, Verti, Ceramik) não têm caixa de recorte em `swatches.py`: se forem candidatos,
medir a caixa na miniatura do moodboard e acrescentar em `CROPS`.

**Metal ou MDF sem equivalente no estilo** (ex.: metal dourado): mostrar no painel o mais próximo do catálogo
e perguntar. Se a Raquel mantiver o original, registrar como exceção na memória (no book, só "Estilo X").

**MDFs das unidades também trocam**: o slide de materiais das unidades fica no book, então os MDFs dele entram
no painel e no slide de MDFs (como Rícori/Cinza Cristal no Batel). Só os slides de acabamento/escopo das unidades saem.

**Amostra original do MDF antigo**: buscar no Drive inteiro (`name contains '<NOME>'` e `mimeType contains 'image/'`).
Preferir arquivo "MDF <NOME>.jpg" ou "<NOME>.jpg"; vários resultados → o de nome mais limpo. Nenhum → linha do
painel "sem amostra" e dizer à Raquel que a referência é só o render (o painel continua obrigatório).

## Trocas já aprovadas (referência de tom, não regra)
Rícori Nativo → Carvalho Treviso · Grafito Chess → Ópera (Batel) ou Chumbo (Cachoeira, Rosa Norte) ·
Frapê → Crômio · Nogueira Persa → Roble Catedral · Cinza Cristal → Cinza Cobalto (Bruma) ou manter (Industrial) ·
Tauari → Carvalho Treviso · Savana → Roble Catedral · marcenaria grafite sem nome → Chumbo · metal preto fosco → RAL 9005.
A Raquel recusou opções escuras (Louro Preto, Cacau, Nero) para substituir madeiras claras.

## Brises de fachada: regra por região (Rachel + Raquel, reunião 15/10/2026)
O brise **nunca é PVC**. A cor e o acabamento ficam os mesmos que o book, o executivo e a imagem já descrevem
(amadeirado tom amêndoa, escuro, etc.); só o material muda conforme a região:
- **Nordeste** (BA, CE, PE, RN… ex.: Salvador, Imbassaí, Itacaré, Trancoso, Meireles): **madeira**, no tom da imagem/executivo,
  com a ressalva escrita no book: "pode ser substituído por alumínio com o mesmo tom e acabamento". Não cravar madeira:
  quem decide na obra é o administrador.
- **Demais regiões** (Sul, Sudeste, Centro-Oeste): **alumínio com o acabamento que já estava** (ex.: brise amadeirado →
  alumínio com acabamento amadeirado no mesmo tom).
Vale só para brise/ripado de **fachada**. Painéis ripados de interiores (recepção, rooftop) não mudam.
**Ripado de interiores (Raquel 22/10/2026):** se a fonte diz PVC amadeirado, **mantém PVC amadeirado** (não entra no
painel de MDF); se é ripado de MDF, vira **Carvalho Treviso**.
Texturas de fachada (cimento queimado, concreto ripado) ficam como estão no book: a restrição de área é da skill
`conceito-fachada`, não desta.

## Decisões-padrão quando as fontes divergem (aplicar e listar para a Raquel revisar)
Ordem: a **imagem/conceito decide o visual** (forma, tom, o que existe). O **Briefing/LOG decide material e
especificação** que a imagem não mostra (perfil inox, solar, formato do revestimento; brise segue a regra por região). Executivo, memorial
e projeto legal só compatibilizam. Quando imagem e Briefing mostram **cores diferentes** para o mesmo item
(ex.: piscina), seguir o Briefing e destacar o conflito na entrega para a Raquel decidir.
| Conflito | Decisão |
|---|---|
| Brise (executivo × Briefing × imagem) | **Regra por região** (ver "Brises de fachada" abaixo). Nunca PVC. |
| Guarda-corpo perfil Infinity/alumínio × Briefing | "U" inox (áreas comuns) · alumínio anodizado com montante aparente (privativo) |
| Esquadria só "alumínio" | Linha 32 (gold) na cor do Briefing |
| Letreiro/logo/totem/marquise "sem detalhe" | Como na imagem; Projetos detalha fixação e ponto elétrico |
| Muro: imagem × Briefing (cor Seazone, logo e nome) | Visual da imagem + logo e nome |
| Prévia/anteprojeto de interiores × conceito Seazone | Vale o conceito (imagem) |
| Piscina: Briefing esverdeado 20×20 × imagem de outra cor | Briefing; avisar a Raquel (destacar na revisão) |
| Aquecimento: bomba de calor (memorial) × solar (Briefing) | Solar |
| Forro: gesso (projeto) × modular (Briefing) | Briefing |
| Item "sem projeto" (ducha, casa de máquinas, escada) | "Detalhar no executivo" / "incluir no hidrossanitário/elétrico" |
| Item fora da cobertura aprovada mas na imagem e no LOG | Manter + "compatibilizar a aprovação"; destacar na revisão |

## IDs úteis
- Moodboards REV.02 (Slides): Bruma `1dEvH_cjlHUFMOzQaRgrh212D5WAW6Q1bfy8hPbTyeGU` · Biofílico `1Ti480HgXwdH3piK3vEZ6qmSSCrS0lvP9NwdqvXtustU` ·
  Clean `1l5nn89FSNder2IFYC835e9aZvVPm6k6WWx9kltbIGjo` · Industrial `1l_S4Yyq5teNGXgxDMwGcwE2dq8g9IpRTGfZy3Fo0bsI`.
- Pasta de amostras no Drive (subir aqui): `1Y3ZYMQZIMRZ54Jm879o_Gd613xk0Rowg`. Já existem (link público, leitor):
  Carvalho Treviso `10PVvMsW0g7FGbvL-_hOp1Y0ZSzqU_yTF` · Roble Catedral `1GTCbMjMG78bplqNd1iWy7iWrHeC2kY9Y` ·
  Crômio `1DyP8ZfNrHHcEMMju2PILjCW-FRDPzFt4` · Ópera `14e5HOLo4M2MklwP4mF2VcgomC_TN9byt` · Cinza Cobalto `1o7RF8mHxZvuMLRCYNCxgk-NQ7AWCOIzS` ·
  Chumbo `1t4tphQzEVtpWq_uVgep3cZmIP1AWQEAP` · Cinza Cristal `1C0chlbnOJcN6_Jtw9rEqWAArVnEfh9R1` · Desert `1XEi1zXkn3GI-ctoBjQ8rV5ZycusMqgUN` ·
  metal preto `1-APZ90Rl8-OrrjjG1VE-tBmciU6sl1Re` · metal grafite RAL 7024 `15y4sOpM_35-hK6KkFVHHpYOJpd7oYIHZ` ·
  metal bege RAL 1013 `1WypfI9eTCf_0UJ7qBBMcqrJ-PlwLRLM6`.
- Amostras originais dos MDFs antigos: buscar no Drive pelo nome (ex. `name contains 'TAUARI'`) e baixar.
- Slack: DM Rachel `D07E8QJV751` · canal #szni-projetos-estruturação `C0BD06X3KK6` · Rachel Souto `U07DXLFP1GT`.
- Layout V03 (template Campeche, script `build_canasvieiras.py` + `tplkit.py`) no repositório da Rachel: só usar se ela pedir.
