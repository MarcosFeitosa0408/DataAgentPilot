# DataAgentPilot

[README.md](https://github.com/user-attachments/files/34209626/README.md)
<div align="center">

# Data Agent Pilot

**Transformando dados brutos em decisões confiáveis.**

Análise de qualidade de dados, diagnósticos rastreáveis e dashboards a partir de um CSV ou Excel — tudo processado **localmente no navegador**, sem backend e sem enviar seus dados para lugar nenhum.

![HTML](https://img.shields.io/badge/HTML-CSS-JavaScript-2563EB)
![Sem backend](https://img.shields.io/badge/backend-nenhum-00B8FF)
![Testes](https://img.shields.io/badge/testes-automatizados-10B981)

</div>

---

## Sobre o projeto

O **Data Agent Pilot** é uma plataforma web de qualidade de dados e BI. Você envia uma planilha, e ele devolve indicadores de qualidade, diagnósticos linha a linha, insights automáticos e relatórios exportáveis, sempre com **rastreabilidade**: cada análise registra o arquivo, o hash SHA-256 dos bytes originais, os horários e o tempo de execução.

Dois princípios guiam o projeto:

- **Nada é inventado.** Quando uma regra não se aplica ao arquivo (por exemplo, não há colunas financeiras), ela aparece como *Não avaliada*, nunca como aprovada ou com erro.
- **Os dados originais ficam intactos.** Nada é removido, corrigido ou sobrescrito; duplicados e erros são apenas apontados.

## Funcionalidades

- **Entrada:** CSV (separador `,` ou `;`, números no formato brasileiro como `1.234,56`) e Excel `.xlsx`, com seletor de abas.
- **Mapeamento de colunas:** detecção automática por nome (sinônimos em português e inglês, como Receita/Revenue/Faturamento, Custo/Cost, Lucro/Profit). Se a planilha usa outros nomes, você escolhe a coluna e a análise é refeita.
- **Perfil inteligente do esquema:** tipo de cada coluna, completude, tipos mistos e estatísticas numéricas.
- **Insights automáticos**, gráfico de receita por categoria e evolução mensal da receita.
- **Consistência financeira** com tolerância configurável em R$ e detalhamento das divergências.
- **Diagnóstico de registros** com filtros por regra e por severidade.
- **Execution Trace & Data Lineage:** ID da análise, SHA-256, método de cálculo do hash e janela de execução (UTC e horário de Brasília).
- **Exportações:** PDF (impressão), dashboard HTML autônomo, descrição da análise em Markdown, relatório JSON e diagnósticos em CSV.

## Data Quality Contract (v1.0)

| Regra | O que verifica | Severidade dos diagnósticos |
|---|---|---|
| **R00** Estrutura das linhas | Linhas com número de campos diferente do cabeçalho | Alto |
| **R01** Completude | Valores ausentes (vazio, null, nulo, nan, n/a) | Baixo |
| **R02** Unicidade | Linhas inteiras duplicadas (nada é removido) | Médio |
| **R03** Tipos de dados | Colunas com números e textos misturados | Médio |
| **R04** Calendário | Mês fora de 1–12, dia fora de 1–31, datas inexistentes (ex.: 30/02) e formatos não reconhecidos (ISO e DD/MM/AAAA) | Alto / Médio |
| **R05** Lucro = Receita − Custo | Divergência financeira, calculada em centavos para evitar erro de ponto flutuante | Crítico |
| **R06** Negativos/zeros financeiros | Valores negativos ou zero (podem ser estornos legítimos) | Informativo |

Cada regra termina como **Aprovada**, **Atenção**, **Falha** ou **Não avaliada**. O projeto **não calcula um índice único de qualidade**: prefere mostrar regra a regra o que foi verificado e o que não foi.

## Privacidade e segurança

- Tudo roda no navegador. O arquivo não é enviado a nenhum servidor.
- O conteúdo da planilha é sempre exibido como texto: tags HTML dentro das células **não são interpretadas** (proteção contra XSS).
- O CSV exportado é protegido contra injeção de fórmula (células iniciadas por `=` recebem aspa).
- O hash SHA-256 usa a Web Crypto API e, na falta dela, uma implementação própria validada contra a nativa; o método usado aparece no relatório.
- Arquivos exportados podem conter dados confidenciais. Nada é publicado automaticamente.

## Como usar

1. Baixe ou clone o repositório e abra o `index.html` no navegador (não precisa de instalação).
2. Clique em **Selecionar arquivo** (ou arraste a planilha) ou use **Carregar exemplo fictício**.
3. Ajuste a tolerância financeira, se quiser, e clique em **Analisar dados**.
4. Explore os resultados e exporte o que precisar.

```bash
git clone https://github.com/MarcosFeitosa0408/DataAgentPilot.git
cd DataAgentPilot
# abra o index.html no navegador
```

> Mantenha a pasta `libs/` ao lado do `index.html`. Ela contém a biblioteca SheetJS usada para ler `.xlsx` offline.

## Estrutura do repositório

```
DataAgentPilot/
├── index.html              # aplicação completa (HTML, CSS e JavaScript)
├── libs/
│   └── xlsx.full.min.js    # SheetJS 0.18.5, leitura de Excel offline
├── tests/
│   └── run-tests.js        # suíte de testes automatizados (Node + jsdom)
├── package.json
└── .gitignore
```

## Testes

A suíte em `tests/run-tests.js` abre o `index.html` em um DOM simulado (jsdom) e verifica mais de 50 comportamentos, entre eles:

- leitura de CSV brasileiro, Excel com várias abas, arquivos vazios, corrompidos ou de formato incompatível;
- divergências financeiras (R$ 0,01 e R$ 1,00), tolerância e ausência de falso erro de ponto flutuante;
- datas inválidas em colunas ano/mês/dia e em coluna única de data;
- SHA-256 próprio idêntico ao nativo, de 0 a 100 mil bytes;
- proteção contra XSS e contra injeção de fórmula no CSV;
- filtros, exportações (MD, JSON, CSV, PDF, dashboard HTML) e abertura do dashboard sem erros de script;
- desempenho com 30.000 registros.

```bash
npm install
npm test
```

## Tecnologias

HTML5, CSS3 e JavaScript puro · [SheetJS](https://sheetjs.com/) para Excel · Node.js, jsdom para testes.

## Autor

**Marcos Feitosa** — analista de dados júnior (Power BI/DAX, SQL, Python, Excel).
GitHub: [@MarcosFeitosa0408](https://github.com/MarcosFeitosa0408)
