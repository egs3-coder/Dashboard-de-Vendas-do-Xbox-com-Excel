# Desafio de Dashboard de Vendas - Xbox Game Pass

## 🎯 Objetivo do Projeto

O objetivo deste desafio é transformar dados brutos de assinaturas do Xbox Game Pass em um **Dashboard de Vendas** claro e útil no Microsoft Excel. O foco é na organização e visualização de dados para permitir uma análise eficaz do desempenho de vendas e auxiliar na tomada de decisões estratégicas.

## 📊 Dados Utilizados

O dashboard foi construído com base no arquivo `base.xlsx`, que contém a aba principal de dados (`Bases`) com informações detalhadas sobre as assinaturas.

**Estrutura da Tabela de Dados (`Bases`):**

| Coluna | Descrição |
| :--- | :--- |
| `Subscriber ID` | Identificador único do assinante. |
| `Name` | Nome do assinante. |
| `Plan` | Tipo de plano de assinatura (Ultimate, Standard, Core). |
| `Start Date` | Data de início da assinatura. |
| `Auto Renewal` | Indica se a assinatura possui renovação automática (Yes/No). |
| `Subscription Price` | Preço base da assinatura. |
| `Subscription Type` | Frequência de pagamento (Monthly, Quarterly, Annual). |
| `EA Play Season Pass` | Indica se o EA Play Season Pass foi adicionado (Yes/No). |
| `EA Play Season Pass Price` | Preço do EA Play Season Pass. |
| `Minecraft Season Pass` | Indica se o Minecraft Season Pass foi adicionado (Yes/No). |
| `Minecraft Season Pass Price` | Preço do Minecraft Season Pass. |
| `Coupon Value` | Valor do cupom de desconto aplicado. |
| `Total Value` | Valor total da transação (Receita). |

## 📈 Análises e Métricas Chave

O dashboard final (`dashboard_vendas_final.xlsx`) apresenta as seguintes métricas e visualizações:

### Métricas Chave (Cards)

*   **Total de Assinantes:** Contagem total de IDs de assinantes únicos.
*   **Faturamento Anual Total:** Soma total da receita (`Total Value`) gerada apenas por planos com `Subscription Type` igual a 'Annual'.
*   **Receita Média por Assinante (ARPU):** Receita total dividida pelo número total de assinantes.

### Visualizações (Gráficos)

1.  **Faturamento Anual por Auto Renovação:** Gráfico de barras mostrando a distribuição do faturamento anual total entre assinaturas com e sem renovação automática.
2.  **Distribuição de Assinantes por Plano:** Gráfico de barras mostrando a contagem de assinantes por tipo de plano (Ultimate, Standard, Core).

### Aba de Cálculos

A aba `C̳álculos` foi populada com as tabelas dinâmicas e cálculos intermediários utilizados para gerar as métricas e os dados dos gráficos, incluindo:

*   Faturamento Anual Total.
*   Faturamento Anual por Auto Renovação.
*   Faturamento EA Play por Plano.
*   Faturamento Minecraft por Plano.
*   Distribuição de Assinantes por Plano.

## 🛠️ Instruções para Reprodução

O dashboard foi gerado programaticamente usando Python e as bibliotecas `pandas` e `openpyxl`.

### Pré-requisitos

*   Python 3.x
*   Bibliotecas Python: `pandas`, `openpyxl`

### Passos

1.  **Instalar as dependências:**
    ```bash
    pip install pandas openpyxl
    ```

2.  **Baixar os arquivos:**
    Certifique-se de que os arquivos `base.xlsx` e `generate_dashboard.py` estejam no mesmo diretório.

3.  **Executar o script de geração:**
    ```bash
    python generate_dashboard.py
    ```

O script irá gerar o arquivo final `dashboard_vendas_final.xlsx` no mesmo diretório, contendo as abas `B̳ases`, `C̳álculos` e `D̳ashboard` preenchidas.

## 📦 Entrega

O repositório contém:

*   `README.md`: Este arquivo.
*   `base.xlsx`: O arquivo de dados original.
*   `dashboard_vendas_final.xlsx`: O arquivo Excel com o dashboard concluído.
*   `generate_dashboard.py`: O script Python utilizado para gerar o dashboard.
