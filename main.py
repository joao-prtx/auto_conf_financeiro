import pdfplumber
import pandas as pd
import re
import sys
from pathlib import Path

def clean_val(v):
    if not v: return 0.0
    v = str(v).replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        return abs(float(v))
    except:
        return 0.0

def extract_purchases(path):
    items = []
    if not path.exists():
        print(f"Erro: Arquivo {path.name} não encontrado.")
        return items
    
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split('\n'):
                match_valor = re.search(r'(\d+[\d.,]*)$', line.strip())
                if match_valor:
                    valor_raw = match_valor.group(1)
                    valor_final = clean_val(valor_raw)
                    if valor_final > 1.0:
                        texto_bruto = line.replace(valor_raw, '').strip()
                        nome_com_lixo = re.sub(r'^\d{1,5}\s+', '', texto_bruto)
                        nome_limpo = re.sub(r'\d{8,}', '', nome_com_lixo)
                        nome_limpo = nome_limpo.replace('R$', '').strip().upper()
                        nome_limpo = " ".join(nome_limpo.split())
                        if len(nome_limpo) > 2:
                            items.append({"Nome Pix": nome_limpo, "Valor Pagamento": valor_final})
    return items

def extract_itau(path):
    items = []
    if not path.exists():
        print(f"Erro: Arquivo {path.name} não encontrado.")
        return items
        
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split('\n'):
                if re.match(r'^\d{2}/\d{2}', line):
                    partes = line.split()
                    if len(partes) >= 3:
                        data = partes[0]
                        valor_encontrado = 0.0
                        for p in reversed(partes):
                            v = clean_val(p)
                            if v != 0.0:
                                valor_encontrado = v
                                break
                        desc = " ".join(partes[1:-1]).upper()
                        nome_busca = desc.replace("PIX TRANSF", "").replace("PIX PAGTO", "").strip()
                        nome_busca = re.sub(r'\d{2}/\d{2}', '', nome_busca).strip()
                        if valor_encontrado != 0:
                            items.append({
                                "Data": data,
                                "Lançamento Original": desc,
                                "Valor (Extrato)": valor_encontrado,
                                "Nome_Busca": nome_busca
                            })
    return items

def main():
    # Define a pasta onde o script está rodando (Universal)
    base_path = Path(__file__).parent.resolve()
    
    arq_itau = base_path / "extrato.pdf"
    arq_compras = base_path / "compras.pdf"
    output_excel = base_path / "Reconciliacao_Final.xlsx"

    print(f"Pasta de trabalho: {base_path}")
    
    if not arq_itau.exists() or not arq_compras.exists():
        print("!!! ERRO: Certifique-se de que os arquivos 'extrato.pdf' e 'compras.pdf' estão na mesma pasta do script.")
        input("Pressione Enter para sair...")
        return

    print("Extraindo dados...")
    lista_compras = extract_purchases(arq_compras)
    lista_extrato = extract_itau(arq_itau)

    df_final = pd.DataFrame(lista_extrato).astype(object)
    df_final["Status"] = "Nao Encontrado"
    df_final["Nome na Lista"] = ""
    df_final["Valor na Lista"] = None

    sobras = []
    for compra in lista_compras:
        achou = False
        for idx, row in df_final.iterrows():
            if row["Status"] == "Nao Encontrado" and row["Valor (Extrato)"] == compra["Valor Pagamento"]:
                if str(row["Nome_Busca"]) in str(compra["Nome Pix"]):
                    df_final.at[idx, "Status"] = "OK"
                    df_final.at[idx, "Nome na Lista"] = compra["Nome Pix"]
                    df_final.at[idx, "Valor na Lista"] = compra["Valor Pagamento"]
                    achou = True
                    break
        if not achou:
            sobras.append(compra)

    df_final = df_final.drop(columns=["Nome_Busca"])

    try:
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name='Extrato Completo', index=False)
            pd.DataFrame(sobras).to_excel(writer, sheet_name='Nao Localizados', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Extrato Completo']
            format_verde = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            for i, val in enumerate(df_final["Status"]):
                if val == "OK":
                    worksheet.set_row(i + 1, None, format_verde)
        print(f"SUCESSO! Arquivo gerado em: {output_excel}")
    except Exception as e:
        print(f"Erro ao salvar Excel (ele pode estar aberto): {e}")
    
    input("\nProcesso finalizado. Pressione Enter para fechar...")

if __name__ == "__main__":
    main()