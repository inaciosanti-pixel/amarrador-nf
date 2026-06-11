import streamlit as st
import pandas as pd
import pdfplumber
import re
from rapidfuzz import fuzz
from io import BytesIO

st.set_page_config(
    page_title="Amarrador NF",
    layout="wide"
)

st.title("🔗 Amarrador Automático de NF x Pedido")

st.markdown("""
Faça upload:
- da Nota Fiscal PDF
- do Pedido de Compra PDF

O sistema irá:
- extrair os itens
- comparar os produtos
- realizar a amarração automática
- gerar tabela final
""")

nf_file = st.file_uploader(
    "Upload Nota Fiscal PDF",
    type=["pdf"]
)

pedido_file = st.file_uploader(
    "Upload Pedido PDF",
    type=["pdf"]
)

if nf_file and pedido_file:

    st.success("Arquivos carregados com sucesso.")

    texto_nf = ""
    texto_pedido = ""

    with pdfplumber.open(nf_file) as pdf:
        for p in pdf.pages:
            txt = p.extract_text()
            if txt:
                texto_nf += txt

    with pdfplumber.open(pedido_file) as pdf:
        for p in pdf.pages:
            txt = p.extract_text()
            if txt:
                texto_pedido += txt

    linhas_nf = texto_nf.split("\n")

    itens_nf = []

    for linha in linhas_nf:

        match = re.search(
            r'(\d{5})\s+(.+?)\s+PC\s+([\d,]+)\s+([\d,]+)\s+[\d,]+\s+([\d,]+)',
            linha
        )

        if match:

            itens_nf.append({
                "cod_nf": match.group(1),
                "descricao_nf": match.group(2).strip(),
                "qtd": float(match.group(3).replace(",", ".")),
                "unit": float(match.group(4).replace(",", ".")),
                "total": float(match.group(5).replace(",", "."))
            })

    df_nf = pd.DataFrame(itens_nf)

    st.write("Itens encontrados na NF:", len(df_nf))
    st.dataframe(df_nf)

    linhas_pedido = texto_pedido.split("\n")

    for linha in linhas_pedido:
        if "EMV-" in linha:
            st.write(linha)

    itens_pedido = []

    for i, linha in enumerate(linhas_pedido):

        if "EMV-" in linha and "Centro de Custo" not in linha:

            try:

                cod_match = re.search(r'(EMV-\d+)', linha)

                if not cod_match:
                    continue

                cod = cod_match.group(1)

                descricao = linha.split(cod)[1].strip()

                if i + 1 < len(linhas_pedido):

                    prox = linhas_pedido[i + 1]

                    if (
                        "Marca" not in prox
                        and "Informações" not in prox
                        and "Centro de Custo" not in prox
                    ):
                        descricao += " " + prox.strip()

                bloco = "\n".join(linhas_pedido[i:i+20])

                padrao_item = re.search(
                    r'(\d+,\d+)UN\s+(\d+,\d+)',
                    bloco
                )

                if padrao_item:

                    qtd = float(
                        padrao_item.group(1).replace(",", ".")
                    )

                    unit = float(
                        padrao_item.group(2).replace(",", ".")
                    )

                    total = round(qtd * unit, 2)

                else:

                    qtd = 0
                    unit = 0
                    total = 0

                itens_pedido.append({
                    "cod_pedido": cod,
                    "descricao_pedido": descricao,
                    "qtd": qtd,
                    "unit": unit,
                    "total": total
                })

            except Exception as e:
                st.error(f"Erro: {e}")

    df_pedido = pd.DataFrame(itens_pedido)

    st.subheader("Pedido Extraído")

    for item in itens_pedido:
        st.json(item)

    st.write("Itens encontrados no Pedido:", len(df_pedido))
    st.dataframe(df_pedido)

    resultado = []

    pedidos_usados = set()

    st.write("Iniciando amarração...")

    pedidos_usados = set()    
       
    st.write("Iniciando amarração...")

    for _, nf in df_nf.iterrows():

        melhor = None
        maior_score = 0

        for _, ped in df_pedido.iterrows():
            
            if ped["cod_pedido"] in pedidos_usados:
                continue

            score = 0

            # QUANTIDADE
            if abs(nf["qtd"] - ped["qtd"]) < 0.01:
                score += 25

            # UNITÁRIO
            if abs(nf["unit"] - ped["unit"]) < 0.05:
                score += 25

            # TOTAL
            if abs(nf["total"] - ped["total"]) < 0.05:
                score += 30

            # SIMILARIDADE DA DESCRIÇÃO
            similaridade = fuzz.token_sort_ratio(
                str(nf["descricao_nf"]),
                str(ped["descricao_pedido"])
                st.write(
                    nf["descricao_nf"],
                    " --> ",
                    ped["descricao_pedido"],
                    " = ",
                    similaridade
                )
            )

            score += similaridade * 0.2

            if score > maior_score:
                maior_score = score
                melhor = ped

        if melhor is not None and maior_score > 0:

            pedidos_usados.add(melhor["cod_pedido"])
            
            resultado.append({
                "COD NF": nf["cod_nf"],
                "COD PEDIDO": melhor["cod_pedido"],
                "DESCRIÇÃO": melhor["descricao_pedido"],
                "QTD": nf["qtd"],
                "UNITÁRIO": nf["unit"],
                "TOTAL": nf["total"],
                "STATUS": f"{round(maior_score)}%"
            })

    df_resultado = pd.DataFrame(resultado)

    st.write("Amarrações encontradas:", len(df_resultado))

    st.subheader("Resultado da Amarração")

    st.dataframe(
        df_resultado,
        use_container_width=True
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(
            writer,
            index=False,
            sheet_name="Amarracao"
        )

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Baixar Excel",
        data=excel_data,
        file_name="amarracao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
