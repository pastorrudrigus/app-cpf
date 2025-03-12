import streamlit as st
import pandas as pd
import requests
from time import sleep
from io import BytesIO

st.title("Processamento de Inscrições para CPF")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    if "inscrição" not in df.columns:
        st.error("O arquivo deve conter uma coluna chamada 'inscrição'.")
    else:
        df['cpf'] = None
        total_rows = len(df)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        if st.button("Iniciar Processamento"):
            for i, row in df.iterrows():
                inscricao_valor = str(row["inscrição"])
                try:
                    url = f'https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00201w0.asp?txt_nr_iptu={inscricao_valor}&txt_captcha='
                    response = requests.post(url)
                    # Extraindo o CPF/CNPJ da resposta
                    cpf_extraido = response.text.split('''<td align="left" style="height: 23px">CPF/CNPJ</td>''')[1].split('<td style="height: 23px">')[2].split('  ')[0][1:]
                    df.at[i, 'cpf'] = cpf_extraido
                except Exception as e:
                    # Em caso de erro, podemos deixar o campo em branco ou registrar o erro
                    df.at[i, 'cpf'] = None
                    st.error(f"Erro ao processar inscrição {inscricao_valor}: {e}")
                sleep(0.5)
                progress_bar.progress((i + 1) / total_rows)
                status_text.text(f"Processado {i + 1} de {total_rows} inscrições")
            
            st.success("Processamento concluído!")
            
            # Gera o arquivo Excel processado em memória para download
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            processed_file = output.getvalue()
            
            st.download_button(
                label="Download do arquivo processado",
                data=processed_file,
                file_name="arquivo_processado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
