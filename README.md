

# 🚀 Auto Conf - Reconciliação Bancária Automática

Este projeto automatiza a conferência entre uma lista de compras (PDF) e o extrato bancário do Itaú (PDF). O programa identifica os matches por valor e nome, gerando um relatório colorido em Excel.

---

## 🛠️ Requisitos Prévios (Instalação do Python)

Antes de rodar o programa, você precisa ter o Python instalado na sua máquina.

### **No Windows:**
1. Acesse [python.org](https://www.python.org/downloads/).
2. Clique no botão **Download Python 3.x**.
3. **IMPORTANTE:** Ao abrir o instalador, marque a caixa **"Add Python to PATH"** antes de clicar em "Install Now".
4. Após finalizar, abra o Prompt de Comando (CMD) e digite `python --version` para confirmar.

### **No macOS:**
1. O Mac costuma vir com uma versão antiga. Recomenda-se instalar a mais atual via [python.org](https://www.python.org/downloads/) ou usando o Homebrew:
   ```bash
   brew install python
   ```
2. No terminal, verifique digitando `python3 --version`.

---

## 📦 Instalação das Dependências

Com o Python instalado, abra o terminal (ou CMD) dentro da pasta `auto_conf` e execute o comando abaixo para instalar as bibliotecas necessárias:

```bash
pip install -r requirements.txt
```

---

## 📂 Como Utilizar

Para que o programa funcione corretamente, siga exatamente estes passos:

1. **Preparar os arquivos:**

### 1. Planilha "Financeiro Compras" (PDF de Compras)
O programa precisa de uma tabela limpa para evitar erros de leitura.
1.  Abra a planilha **Financeiro Compras**.
2.  Selecione com o mouse apenas os pagamentos que deseja conferir.
3.  Certifique-se de selecionar **apenas as colunas "Nome" e "Valor"**.
4.  Vá em **Arquivo > Imprimir** (ou `Ctrl + P`).
5.  Nas configurações de impressão, escolha **"Células selecionadas"**.
6.  Salve como PDF com o nome exato: `compras.pdf`.
7.  Mova o arquivo para a pasta `auto_conf`.

### 2. Itaú Web (PDF de Extrato)
O extrato deve vir direto do Internet Banking para manter o padrão das colunas.
1.  Acesse sua conta no **Itaú Web** pelo navegador.
2.  Vá na área de extratos e **filtre a data** para o período exato que deseja conferir.
3.  Clique no ícone de exportar/imprimir e escolha a opção **Salvar como PDF**.
4.  Salve o arquivo com o nome exato: `extrato.pdf`.
5.  Mova o arquivo para a pasta `auto_conf`.

---

2. **Executar o programa:**
   Após colocar os dois arquivos (`compras.pdf` e `extrato.pdf`) dentro da pasta do programa:

- * **No Windows:** Dê dois cliques no arquivo `executar.bat`.
- * **No macOS:** Dê dois cliques no arquivo `executar.command`.

O terminal abrirá, processará os dados em poucos segundos e gerará o arquivo **`Reconciliacao_Final.xlsx`** pronto para análise!

3. **Resultado:**
   - O programa criará (ou substituirá) o arquivo **`Reconciliacao_Final.xlsx`** na pasta.
   - Na aba **Extrato Completo**, as linhas marcadas em **verde** são os matches confirmados.
   - Na aba **Nao Localizados**, estarão as compras que não foram encontradas no extrato.

---

## ⚠️ Observações Importantes
- **Feche o Excel:** Se o arquivo `Reconciliacao_Final.xlsx` estiver aberto, o programa não conseguirá salvar os novos dados e apresentará um erro.
- **Nomes de Arquivo:** Se o nome for diferente de `compras` ou `extrato`, o programa não encontrará os dados.
```

---

### Dica extra de organização:
Para o comando `pip install -r requirements.txt` funcionar, você **precisa** criar o arquivo `requirements.txt` na mesma pasta com este conteúdo:

```text
pdfplumber
pandas
xlsxwriter
```