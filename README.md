# Dashboard de Vendas - Assinaturas Xbox Game Pass

> Projeto desenvolvido para o módulo **"Criando Dashboards com Excel"** do curso **Santander - Excel com IA e Claude**.

---

## Descrição do Desafio

O objetivo deste desafio é criar um **dashboard de vendas**, com foco na organização e visualização de dados. A meta central é transformar dados brutos em informações visuais claras e úteis, permitindo uma análise eficaz do desempenho de vendas e a tomada de decisões orientadas por dados (*data-driven decision making*).

Neste projeto prático, foi analisada uma base transacional completa de assinaturas do ecossistema **Xbox Game Pass**, incluindo planos base (`Core`, `Standard`, `Ultimate`), modelos de cobrança (`Mensal`, `Trimestral`, `Anual`), adesão a passes de temporada adicionais (**EA Play Season Pass** e **Minecraft Season Pass**) e a política de descontos por cupons.

---

## Funcionalidades da Solução

- **Identidade Visual e Padronização (`A̳ssets`):**
  - Guia de estilo documentado com a paleta oficial da marca Xbox (`#9BC848`, `#22C55E`), tons secundários de menu (`#2AE6B1`, `#5BF6A8`) e zonas neutras de fundo (`#E8E6E9`).
  - Biblioteca de ícones e logotipos aplicados diretamente nos componentes visuais.
- **Base Transacional Estruturada (`B̳ases`):**
  - Registro de 295 clientes com 13 campos analíticos.
  - Regra de composição de receita com abatimento de cupons de desconto:
    $$\text{Total Value} = \text{Subscription Price} + \text{EA Play Price} + \text{Minecraft Price} - \text{Coupon Value}$$
- **Mecanismo Analítico e Resolução de Problemas de Negócio (`C̳álculos`):**
  - Modelagem analítica via Tabelas Dinâmicas (*DataPilot*) para isolamento e agregação de métricas por tipo de assinatura, autorenovação e planos.
- **Painel Executivo Interativo (`D̳ashboard`):**
  - Cabeçalho executivo com período de apuração e saudação ao usuário.
  - Cards de KPIs para monitoramento de receita dos passes adicionais.
  - Gráfico analítico de barras para comparação de faturamento por renovação automática.

---

## Estrutura da Planilha

O projeto foi construído e disponibilizado nos formatos [`Curso Excel Santander - Aula 7 PROJETO.ods`] e [`Curso Excel Santander - Aula 7 PROJETO.xlsx`], organizado em quatro abas integradas:

```
 Projeto Dashboard Xbox
 ┣ 🎨 A̳ssets (Identidade Visual)
 ┃   ┣ 🔹 Paleta de cores oficial Xbox e menus
 ┃   ┗ 🔹 Repositório de logos e ícones
 ┣ 🗃️ B̳ases (Base de Dados Bruta)
 ┃   ┣ 🔹 295 registros transacionais de assinantes
 ┃   ┣ 🔹 Planos, periodicidades e adesão a Season Passes
 ┃   ┗ 🔹 Preços, cupons e totalização líquida
 ┣ ⚙️ C̳álculos (Motor Analítico)
 ┃   ┣ 🔹 Tabela Dinâmica 1: Faturamento por Renovação e Periodicidade
 ┃   ┣ 🔹 Tabela Dinâmica 2: Faturamento do EA Play Season Pass por Plano
 ┃   ┣ 🔹 Tabela Dinâmica 3: Faturamento do Minecraft Season Pass por Plano
 ┃   ┗ 🔹 Respostas às 4 Perguntas de Negócio
 ┗ 📊 D̳ashboard (Apresentação Visual)
     ┣ 🔹 Cabeçalho executivo e período de apuração
     ┣ 🔹 Cards de KPI (EA Play Season Pass e Minecraft Season Pass)
     ┗ 🔹 Gráfico de vendas Xbox Game Pass por renovação automática
```

---

## Perguntas de Negócio Respondidas

A aba **`C̳álculos`** transforma os dados brutos em respostas diretas para a gestão do produto:

| # | Pergunta de Negócio | Resposta / Métrica Identificada |
| :-: | :--- | :--- |
| **1** | *Qual o faturamento total de vendas de planos anuais (contendo todas as assinaturas agregadas)?* | **R$ 1.754,00** |
| **2** | *Qual o faturamento total de vendas de planos anuais, separado por autorenovação e não autorenovação?* | • Com autorenovação (`Yes`): **R$ 1.537,00**<br>• Sem autorenovação (`No`): **R$ 217,00** |
| **3** | *Qual o total de vendas de assinaturas do EA Play?* | • No recorte Anual (`Ultimate`): **R$ 600,00**<br>• Faturamento Geral Consolidado: **R$ 2.940,00** |
| **4** | *Qual o total de vendas de assinaturas do Minecraft Season Pass?* | • No recorte Anual (`Standard` + `Ultimate`): **R$ 940,00**<br>• Faturamento Geral Consolidado: **R$ 3.880,00** |

*Faturamento Total Geral da Base (todos os planos e periodicidades):* **R$ 7.633,00**.

---

## Fórmulas e Técnicas Empregadas

### 1. Vínculo Dinâmico entre Abas
Os cartões de KPI do painel executivo consomem automaticamente os resultados consolidados no motor de cálculos:
```excel
='C̳álculos'!F23   ' Alimenta o Card de Total Subscriptions EA Play
='C̳álculos'!F35   ' Alimenta o Card de Total Subscriptions Minecraft
```

### 2. Tabelas Dinâmicas para Análise Multidimensional
Construção de resumos cruzando:
- **Linhas:** Status de autorenovação (`Auto Renewal`) e tipo de plano (`Plan`).
- **Valores:** Soma dos valores agregados (`Sum - Total Value`, `Sum - EA Play Season Pass Price`, `Sum - Minecraft Season Pass Price`).
- **Páginas / Filtros:** Filtro por periodicidade (`Subscription Type`: `Annual`, `Quarterly`, `Monthly`).

### 3. Visualização Gráfica Integrada
- Criação de gráfico de barras no painel executivo apontando para o intervalo dinâmico `'C̳álculos'!$C$11:$D$14`, permitindo visualização direta do impacto da autorenovação na receita da empresa.

---

## Nota Técnica de Compatibilidade e Ferramental

> [!NOTE]
> **Ambiente de Desenvolvimento e Interoperabilidade:**
> 
> - Este projeto foi concebido e estruturado utilizando o **Google Planilhas** (*Google Sheets*) e exportado para os formatos **.ODS** (*OpenDocument Spreadsheet*) e **.XLSX** (*Microsoft Excel OpenXML*).
> - Por limitações de licença local (ausência do aplicativo de desktop proprietário do Microsoft Office na máquina de desenvolvimento), a manipulação foi conduzida via **Google Sheets** e **Excel Online**.
> - **Comportamento das Segmentações de Dados (*Slicers*):** No Excel desktop moderno, os componentes de Segmentação de Dados (*Table Slicers*) operam de forma nativa. Contudo, ao transitar entre Google Sheets, Excel Online e LibreOffice Calc, esse recurso específico de segmentação gráfica sofre limitações de suporte entre plataformas (apresentando a notificação técnica de que o *shape* de *table slicer* não é suportado pelo visualizador online). 
> - A integridade de todos os dados, tabelas dinâmicas, fórmulas de vínculo e gráficos permanece 100% funcional e preservada em ambas as versões disponibilizadas.

---

## Como Utilizar

1. Abra o arquivo [`Curso Excel Santander - Aula 7 PROJETO.ods`] no LibreOffice Calc ou [`Curso Excel Santander - Aula 7 PROJETO.xlsx`] no Microsoft Excel.
2. Navegue até a aba **`D̳ashboard`** para visualizar o painel executivo de vendas com os KPIs de Season Passes e o gráfico de distribuição de receita.
3. Acesse a aba **`C̳álculos`** para verificar as tabelas dinâmicas configuradas e as respostas detalhadas para cada pergunta de negócio.
4. Na aba **`B̳ases`**, você encontrará a matriz completa com os 295 registros de clientes que alimentam todo o ecossistema.
5. A aba **`A̳ssets`** serve como referência de design system, documentando a paleta de cores aplicada.

---

## Competências Desenvolvidas

- Transformação de dados brutos (*raw data*) em inteligência acionável e relatórios executivos.
- Estruturação arquitetural em camadas: Dados (`Bases`) $\rightarrow$ Processamento (`Cálculos`) $\rightarrow$ Apresentação (`Dashboard`).
- Formulação e resolução de perguntas de negócio com Tabelas Dinâmicas.
- Aplicação de boas práticas de UI/UX para planilhas utilizando Design System e paleta de cores padronizada.
- Compreensão prática de interoperabilidade entre suítes de produtividade (Google Sheets, LibreOffice Calc e Microsoft Excel).

---
