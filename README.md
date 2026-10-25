# 💰 Controle Financeiro — App Python + Excel sincronizados

Um app de desktop em Python para **controlar seus gastos, planejar quanto guardar e acompanhar uma meta financeira**, ligado a uma **planilha Excel**: o que você muda no app aparece na planilha, e o que você muda na planilha aparece no app.

> Feito com `tkinter`, `matplotlib` e `openpyxl`.

---

## ✨ O que ele faz

| Recurso | Descrição |
|---|---|
| 📊 **Painel** | Salário, quanto guardar no mês, gastos, **quanto ainda pode gastar** (e quanto isso dá por dia) e quanto falta guardar no mês |
| 🎯 **Meta financeira** | Barra de progresso, quanto falta, **quanto guardar por mês para bater no prazo** e em quantos meses você chega no ritmo atual |
| 📈 **Gráficos** | Gastos por categoria (pizza), salário × guardar × gastos (barras) e evolução do que você guardou, com e sem rendimento |
| 💸 **Gastos** | Lançar, editar e excluir gastos com data, descrição, categoria e valor |
| 🏦 **Guardado** | Registrar o dinheiro que você guarda e ver o rendimento acumulado |
| 🏆 **Conquistas** | 24 conquistas por valor guardado, rendimento acumulado, progresso da meta e hábitos (meses seguidos guardando, meses no azul) |
| 📉 **Investimentos** | Ranking de opções de renda fixa por rendimento líquido estimado, dicas de por onde começar e **simulador** |
| 🔄 **Sincronização com Excel** | A planilha `controle_financeiro.xlsx` é a fonte dos dados e é atualizada nos dois sentidos |

---

## 📁 Arquivos do projeto

```
📂 sua-pasta/
├── app_financas.py          ← o app (é este que você abre)
├── financas_core.py         ← cálculos e leitura/gravação da planilha
└── controle_financeiro.xlsx ← a planilha (criada automaticamente se não existir)
```

Os três arquivos precisam ficar **na mesma pasta**.

---

## 🚀 Instalação

### 1. Requisitos
- **Python 3.10 ou superior** (baixe em [python.org](https://www.python.org/downloads/) e, no Windows, marque **"Add Python to PATH"** durante a instalação)
- O `tkinter` já vem junto com o Python no Windows

### 2. Instale as bibliotecas
No Prompt de Comando (Windows) ou terminal:

```bash
pip install openpyxl matplotlib
```

Se `pip` não for reconhecido, use `python -m pip install openpyxl matplotlib`.

### 3. Rode o app

```bash
python app_financas.py
```

Ou dê **dois cliques** em `app_financas.py`. Na primeira execução a planilha é criada sozinha.

---

## 🧭 Como usar (passo a passo)

### 1️⃣ Comece pela aba **Configurações**
Preencha e clique em **Salvar configurações**:

| Campo | Para que serve |
|---|---|
| Salário mensal líquido | Quanto entra por mês, já com descontos |
| % do salário para guardar | Ex.: `20` = guardar 20% todo mês |
| Nome e valor da meta | Ex.: "Reserva de emergência", R$ 10.000 |
| Prazo da meta (meses) | Em quanto tempo você quer bater a meta |
| Já guardado antes de usar o app | Saldo que você já tinha |
| Rendimento anual líquido | Estimativa do que seu dinheiro rende por ano, já com IR |
| Selic, CDI e IPCA | Taxas de referência usadas no ranking de investimentos |

> 💡 Porcentagens: digite só o número (`20` para 20%). Datas: `dd/mm/aaaa`.

### 2️⃣ Lance seus gastos na aba **Gastos**
Informe data, descrição, categoria e valor e clique em **Adicionar**. Para corrigir, clique na linha, altere os campos e use **Atualizar selecionado**. Para apagar, **Excluir selecionado**.

### 3️⃣ Registre o que você guardou na aba **Guardado**
Cada vez que separar dinheiro, lance a data e o valor. O app calcula o rendimento com juros compostos pela taxa anual que você configurou.

### 4️⃣ Acompanhe no **Painel**
- 🟢 **Verde** em "Ainda posso gastar": você está dentro do planejado.
- 🔴 **Vermelho**: você passou do que sobra depois de guardar. O app avisa quanto.
- Use a caixa **Mês** no topo para ver meses anteriores.

### 5️⃣ Veja as **Conquistas** e os **Investimentos**
Desbloqueie conquistas guardando e deixando o dinheiro render (aparece um aviso a cada nova conquista). Na aba Investimentos, compare opções e use o simulador: "e se eu guardar R$ X por mês durante Y meses?".

---

## 🔄 Sincronização com o Excel

O app confere a planilha a cada ~1,5 segundo.

- **Excel → app:** edite a planilha e **salve (Ctrl+S)**. O app se atualiza sozinho.
- **App → Excel:** toda alteração feita no app é gravada na hora.

### ⚠️ Importante sobre o Excel aberto
No Windows, o Excel **trava o arquivo enquanto está aberto**, e o app não consegue gravar nele. Quando isso acontece:

- o app guarda suas alterações e mostra um aviso na barra de status;
- assim que você **fechar o Excel**, as alterações são gravadas automaticamente;
- se a planilha e o app forem alterados ao mesmo tempo, o app pergunta qual versão manter.

**Dica:** para o dia a dia, use o app. Abra a planilha só quando quiser conferir ou editar em lote (o botão **Abrir planilha no Excel** já grava tudo antes de abrir).

### Abas da planilha

| Aba | Conteúdo | Editável? |
|---|---|---|
| **Resumo** | Cálculos do mês atual com fórmulas e gráficos do Excel | Não (é recalculada) |
| **Config** | Salário, meta, taxas | ✅ Sim |
| **Gastos** | Data, descrição, categoria, valor | ✅ Sim |
| **Guardado** | Data, valor, observação | ✅ Sim |
| **Investimentos** | Lista de investimentos e taxas | ✅ Sim |
| **Conquistas** | Status das conquistas | Não (é recalculada) |

> Mantenha as colunas das abas editáveis na mesma ordem. Abas extras criadas por você são preservadas.

---

## 📉 Sobre os investimentos

O ranking compara opções de renda fixa pelo **rendimento líquido estimado** (depois do IR), usando Selic, CDI e IPCA da aba Configurações. A lista inicial inclui Tesouro Selic, CDB (liquidez diária e prazo fixo), LCI/LCA, Tesouro IPCA+ e poupança.

- As taxas iniciais são **valores de referência de outubro de 2026** (Selic em 14,00% a.a.). As taxas reais variam por banco e corretora: **atualize a aba Investimentos** com as da sua instituição.
- O indexador de cada linha pode ser `%CDI`, `%Selic`, `IPCA+` ou `Fixo`.
- ⚠️ Isto é **informação educativa, não recomendação de investimento**. Confira sempre as condições reais antes de aplicar.

---

## 🏆 Conquistas

| Grupo | Exemplos |
|---|---|
| **Guardar** | Primeiro passo (R$ 100), Mil guardados, Cinco dígitos (R$ 10 mil), Seis dígitos (R$ 100 mil) |
| **Render** | Dinheiro trabalhando (R$ 1), Cem de renda passiva, Mil em rendimentos |
| **Meta** | 25%, 50%, 75% e Meta batida! |
| **Hábito** | 3, 6 e 12 meses seguidos guardando a meta mensal; mês no azul; trimestre no controle |

---

## 🛠️ Solução de problemas

| Problema | Solução |
|---|---|
| `pip` não é reconhecido | Use `python -m pip install openpyxl matplotlib` ou reinstale o Python marcando "Add Python to PATH" |
| `ModuleNotFoundError: No module named 'openpyxl'` | Rode `pip install openpyxl matplotlib` |
| App abre e fecha rápido | Abra pelo Prompt de Comando (`python app_financas.py`) para ler a mensagem de erro |
| "Planilha aberta no Excel" na barra de status | Salve e feche o Excel; o app grava sozinho em seguida |
| Alterei no Excel e não apareceu no app | Salve a planilha (Ctrl+S) e aguarde 1 ou 2 segundos |

---

## 🧱 Como é por dentro

- `financas_core.py`: sem interface; contém a leitura/gravação do Excel (`openpyxl`), os cálculos (rendimento por juros compostos, meta, evolução), as conquistas e o ranking de investimentos.
- `app_financas.py`: interface `tkinter` com abas, gráficos `matplotlib` e o monitoramento da planilha.

Quer adaptar? Categorias de gasto ficam em `CATEGORIAS`, os investimentos iniciais em `INVESTIMENTOS_PADRAO` e as conquistas na função `conquistas()`, todos em `financas_core.py`.

---

## 🔒 Privacidade (se for publicar no GitHub)

A planilha `controle_financeiro.xlsx` guarda **seus dados financeiros reais**. Para não enviá-la ao GitHub, crie um arquivo `.gitignore` com:

```
controle_financeiro.xlsx
*.tmp
__pycache__/
```

---

## 📄 Licença

Projeto de estudo. Defina a licença que preferir (por exemplo, MIT) ao publicar.
