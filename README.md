# 📊 Simulador de Investimento em Fundos Imobiliários (FIIs) | Rize Nexus

Ferramenta interativa desenvolvida em Microsoft Excel para simulação de acúmulo de património, projeção de dividendos e gestão estratégica de alocação de carteira por perfil de investidor.

O projeto foi construído como parte do desafio prático da **DIO (Digital Innovation One)**, sob orientação do Expert Felipe Aguiar (Felipão), aplicando boas práticas de modelagem financeira, arquitetura de dados e design de interface estilo aplicação.

---

## 🎯 As 5 Perguntas de Negócio Respondidas

A ferramenta foi projetada para responder prontamente às cinco questões centrais de planeamento financeiro na aba principal `Simulador`:

| # | Pergunta de Negócio | Localização na Planilha | Mecanismo / Fórmula Utilizada |
| :-: | :--- | :--- | :--- |
| **1** | **Quanto investir por mês?** | Painel 2 - Célula `D10` | Entrada direta do utilizador (Intervalo Nomeado `aporte`) |
| **2** | **Por quantos anos?** | Painel 2 - Célula `D11` | Entrada direta do utilizador (Intervalo Nomeado `qtd_anos`) |
| **3** | **Qual a taxa de rendimento mensal?** | Painel 2 - Célula `D12` | Entrada direta do utilizador (Intervalo Nomeado `taxa_mensal`) |
| **4** | **Quanto de património vai acumular?** | Painel 2 - Célula `D13` | `=FV(taxa_mensal; qtd_anos*12; aporte*-1)` |
| **5** | **Quanto vai receber de dividendos por mês?** | Painel 2 - Célula `D14` | `=patrimonio * rendimento_carteira` |

---

## 🧮 Lógica dos Cálculos e Fórmulas

### 1. Função `VF` (`FV` - Valor Futuro)
A função de Valor Futuro é utilizada para calcular o montante acumulado ao fim do período desejado, considerando juros compostos e aportes mensais recorrentes.

$$\text{Património} = \text{VF}(\text{Taxa Mensal}; \text{Anos} \times 12; -\text{Aporte})$$

* **Ajuste de Período (`nper`)**: O número de anos (`qtd_anos`) é multiplicado por 12 para converter o prazo para meses, correspondendo à frequência mensal dos aportes.
* **Inversão de Sinal (`-aporte`)**: Por convenção financeira no Excel, saídas de caixa (investimentos) são tratadas como valores negativos. Multiplicar o aporte por `-1` garante que o património final seja exibido como um valor positivo.

### 2. Busca Dinâmica com `PROCV` (`VLOOKUP`) e Chave Composta
Para determinar o percentual recomendado de cada categoria de FII em função de duas variáveis simultâneas (**Perfil do Investidor** e **Tipo de FII**), implementou-se uma **Chave Composta** na aba `Apoio`:

* **Concatenação de Chave na Aba `Apoio`**:
  ```excel
  =B3&"-"&C3  -->  Resultado: "Conservador-PAPEL"
