# Seção 04 · Comunicação visual (padrão desde 30/10/2026)

Pedido da Rachel Souto: os books passam a ter a especificação de comunicação visual do prédio, começando
pela placa de número das unidades. Primeiro book feito: Campeche Spot (seção 04, págs. 38–45).
Exemplo pronto: `C:\Users\Seazone\Claude\projetos\book-premissas-template\scripts\books_v3\build_comunicacao_visual_campeche.py`
(usa `tplkit.py`; copiar e trocar textos, imagens e medidas).

## Onde entra
Seção nova **04 · Comunicação visual**, depois de 03 Materiais e antes do encerramento. Numeração do rodapé
continua a do book. Índice do slide 3 ganha a linha "04 Comunicação visual".

## Regras de conteúdo
- Mesmas do book: tudo **DEFINIDO**, sem custo, nada de "a confirmar". O que a fonte não diz, **propor**
  com base no projeto (estilo, metal do catálogo, renders) e listar na entrega o que foi proposta sua.
- **Medidas tiradas em escala**: primeiro a elevação/vista ortogonal do Revit; se só houver render, medir só
  elementos no **mesmo plano da porta** (mesma parede, vista de frente), usando a folha da porta (80 × 210 cm) como
  régua: altura do dígito, distância à guarnição, altura do piso, posição da arandela. Em perspectiva não comparar
  peças de paredes diferentes. Nunca medida genérica.
- Desenho de elevação **em escala real** (cotas do piso acabado + escala gráfica). Número fora de escala em
  relação à porta foi o erro que a Raquel apontou.
- Quantidades do **SpotSys** (`mcp__spotsys-prod__decor_list_typologies(projeto, incluir_unidades=True)`):
  nº de unidades por pavimento, lojas. Não listar os números das unidades um a um (regra do book).
- Material e cor seguem o **metal do estilo** (ex.: Industrial = preto fosco RAL 9005). Sobre fundo escuro
  (pilastra, empena), a mesma peça vai em branco fosco RAL 9003.
- Itens obrigatórios por norma (PPCI/CBMSC, piscina, elevador) entram como "conforme norma/PPCI aprovado".

## Fontes (nesta ordem)
1. Book v2/v3 do Spot: letreiro CV-xx da fachada, logo-casinha, logo na empena, faixa das lojas, frase do rooftop,
   capachos, quadros, arandelas (já estão nas texturas, vistas e itens de recepção/corredor).
2. Memorial de interiores / nova recepção / corredor (medidas de letreiro, quadros, capachos).
3. Renders de corredor e hall (posição e tamanho do número, arandela).
4. Diretrizes Spot (logo wall, quadros de ativação, kit de marca do deck).

## Slides (8, layout v3; nº = slide do template a duplicar)
| # | Slide | Template |
|---|---|---|
| 1 | Divisor "04 Comunicação visual" com render do corredor/hall | 33 |
| 2 | Fachada e acesso: imagem de acesso com marcadores (logo, letreiro, LED, nº predial, lojas, capacho, logo da empena) | 13 |
| 3 | Recepção e corredores: 8 cards com QTD (letreiro, quadros, capacho de entrada, número, arandela, capacho de porta, pavimento/direcional) | 19 |
| 4 | Número da unidade: referência e posição (texto CONCEITO/ESPECIFICAÇÃO/SITUAÇÃO + render grande + 3 detalhes) | 26 |
| 5 | Número da unidade: elevação cotada (PNG gerado) + ficha de 9 linhas | 27 |
| 6 | Família de números: 4 cards com dígitos na mesma escala (unidades, escada, predial, hall) | 21 |
| 7 | Rooftop: frase, logo wall, mini market, pavimento, lavabo, regras/profundidade da piscina, kit de marca | 19 |
| 8 | Segurança, acessibilidade e regras gerais (ficha de 9 linhas) | 27 |

## Padrão proposto no Campeche (ponto de partida, ajustar ao Spot)
- Número da unidade: dígitos avulsos em aço inox 2 mm, corte a laser, pintura eletrostática na cor do metal do
  estilo, fonte Montserrat SemiBold, afastados 10 mm (2 pinos por dígito + adesivo PU, gabarito de furação).
- Posição medida no Campeche: dígito de 100 mm, base 40 mm acima da guarnição (2,21 m do piso), alinhado à direita
  da guarnição, do lado da arandela (arandela a 160 mm, base 1,82 m). Remedir em cada Spot.
- Família: 100 mm unidades/lojas · 150 mm escada (fotoluminescente) · 200 mm número predial · 250 mm hall do elevador.
- Placas de lavabo/áreas técnicas 150 × 150 mm; regras da piscina em ACM 600 × 400 mm; letreiros e logos em LED 3000 K.

## Entrega
PDF + pptx na **mesma pasta do book** no Drive (upload: `POST /api/v3/files/upload/request` + PUT →
`GOOGLEDRIVE_UPLOAD_FILE`, ≤ 5 MB), conferir com `FIND_FILE "'<pasta>' in parents"`. Slack para a Rachel só com
OK da Raquel, assinado "_Enviado por Claude._".
