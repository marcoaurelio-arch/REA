# REA — Modelo Padronizado de Viabilidade de Strip Mall

Template reutilizável em Excel para análise econômico-financeira de strip malls, com
três cenários (Pessimista / Realista / Otimista), fórmulas vivas e parecer automático.

**Arquivo principal:** [`modelo/Viabilidade_Strip_Mall.xlsx`](modelo/Viabilidade_Strip_Mall.xlsx)

> Os valores de exemplo do modelo são **ilustrativos**: usam as faixas de referência
> internas da Nexa e não são dados de mercado verificados. Troque pelos dados reais do
> terreno, de uma pesquisa de aluguel local e do orçamento de obra antes de tomar
> qualquer decisão.

## Abas

| Aba | Função |
|---|---|
| **Instruções** | Como usar, legenda de cores, lógica do modelo e limitações |
| **Resumo** | Indicadores lado a lado por cenário, limites de viabilidade, parecer e gráfico do fluxo |
| **Premissas** | Todas as entradas. Amarelo = premissa que varia por cenário |
| **Mix** | Segmentos, lojas, área e aluguel/m², ABL, quadro de áreas e checagens de produto (CA, TO, vagas, âncora, concentração) |
| **Fluxo** | Fluxo de caixa anual nominal e desalavancado (Ano 0 a 11), um bloco por cenário |
| **Sensibilidade** | Yield on Cost (aluguel × custo de obra) e margem de desenvolvimento (aluguel × cap rate) |

Cores: **azul** = input · **preto** = fórmula · **verde** = link entre abas · **amarelo** = premissa-chave.

## Indicadores calculados

- Custo total do projeto e custo por m² de ABL
- NOI estabilizado (Ano 4, a preços do Ano 0) e margem NOI
- Yield on Cost, cap rate de saída e spread de desenvolvimento
- Valor do ativo estabilizado, criação de valor e margem de desenvolvimento
- Payback simples e ano de recuperação do capital
- TIR, VPL à TMA, MOIC e exposição máxima de caixa
- **Limites de viabilidade:** aluguel médio mínimo e custo de obra máximo que levam o YoC até o alvo (cap rate de saída + spread)
- **Parecer:** VIÁVEL (TIR ≥ TMA **e** YoC ≥ alvo) · ATENÇÃO (só um dos critérios) · INVIÁVEL (nenhum)

## Fórmulas principais

```
Receita potencial  = ABL × aluguel/m² × fator do cenário × 12 × índice de reajuste
Receita efetiva    = potencial − vacância − carência + outras receitas − inadimplência
NOI                = receita efetiva − administração − não repassáveis − encargos de lojas vagas − reserva técnica
Yield on Cost      = NOI estabilizado / custo total
Cap rate           = NOI / valor do ativo
Payback simples    = custo total / NOI estabilizado
Fluxo livre        = NOI − tributos − CAPEX + (venda − custos de venda − IR sobre o ganho)
Valor de saída     = NOI do ano seguinte / cap rate de saída
```

## Leitura do exemplo (apenas para ilustrar o modelo)

O exemplo tem terreno de 7.000 m², ABL de 2.870 m², 18 lojas e aluguel médio de R$ 62/m².
O modelo dá **INVIÁVEL nos três cenários**. Na prática, isso mostra a função do template:
expor números que não fecham.

| Cenário | YoC | TIR | Aluguel mínimo | Obra máxima |
|---|---|---|---|---|
| Pessimista | 5,3% | 3,9% | R$ 137/m² | R$ 1.270/m² |
| Realista | 7,8% | 9,2% | R$ 88/m² | R$ 2.300/m² |
| Otimista | 10,1% | 12,9% | R$ 65/m² | R$ 3.300/m² |

Com obra de R$ 3.500 a 5.000/m² e aluguel de R$ 50 a 70/m², o custo por m² de ABL
(cerca de R$ 6.900) não se paga ao cap rate de mercado adotado. As alavancas são:
âncora pré-locada com aluguel maior, permuta do terreno, obra mais enxuta ou ABL maior
no mesmo terreno.

## Como regenerar e validar

O `.xlsx` é gerado por código, o que garante que cada versão seja reprodutível e auditável.

```bash
python scripts/gerar_modelo.py
```

```bash
python scripts/validar_modelo.py
```

O gerador cria o arquivo com fórmulas. O Excel recalcula tudo ao abrir o arquivo. Para
gravar os valores sem abrir o Excel, recalcule com o LibreOffice (`soffice`). O validador
refaz o fluxo dos três cenários em Python puro e compara TIR, VPL, NOI, custo, YoC, valor e
MOIC com os números da planilha (tolerância de 0,01%).

## Premissas e classificação

Cada premissa tem, na aba Premissas, uma nota com a fonte ou classificação:

- **Referência Nexa:** faixas internas (aluguel R$ 50–90/m², obra R$ 3.500–5.000/m², vacância de 10–20%, administração de 10–15%, não repassáveis de 3–7%, reserva de 0,5–2%).
- **Premissa:** valor assumido; precisa ser validado em cada estudo.
- **Hipótese:** o cap rate de saída deve ser validado com transações comparáveis.
- **Tributação:** lucro presumido simplificado (11,33% sobre a receita; 34% sobre o ganho de capital). Validar com o contador, inclusive os efeitos da reforma tributária (IBS/CBS).

## Limitações (v1.0) e evolução prevista

1. Não há financiamento (análise desalavancada). Próximo passo: estrutura de dívida e TIR do equity.
2. Não há aluguel percentual, CDU/luvas nem contratos com reajuste escalonado.
3. O fluxo é anual: não captura o descasamento mensal entre a obra e as receitas.
4. A tributação é simplificada e não modela FII nem pessoa física.
