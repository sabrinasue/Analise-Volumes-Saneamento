import streamlit as st

#Configurações da página
st.set_page_config(
    layout="centered",
    page_title="Análise de Volumes"

)
#side bar
st.sidebar.image("images/GEE.png", caption="Data Analytics")
st.sidebar.markdown('Desenvolvido por [Sabrina Bilio](sabrina.bilio@aegea.com.br)')
#Configurações da página
st.markdown('# Bem vindo ao Analisador de Volumes por Unidade de Negócio')

st.divider()
st.markdown(
        '''
    Esse projeto foi desenvolvido como projeto final ***para usuários do Excel***.
    Utilizadas três principais bibliotecas para o seu desenvolvimento
    - `pandas`: para manipulação dos dados
    - `ploty`: para geração de gráficos
    - `streamlit`: para criação do webApp interativo
    Os dados utilizados foram gerados pelo Viridis.
    Sugestões podem ser enviadas para o e-mail sabrina.bilio@aegea.com.br
        '''
)