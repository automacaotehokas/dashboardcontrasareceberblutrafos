import os
import pandas as pd
from io import BytesIO
import streamlit as st
from dotenv import load_dotenv
from shareplum import Site, Office365
from shareplum.site import Version
import plotly.express as px

# Carregar as variáveis do arquivo .env
load_dotenv()

class SharePoint:
    def __init__(self):
        self.USERNAME = os.getenv("SHAREPOINT_USER")
        self.PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
        self.SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")
        self.SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
        self.SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC_LIBRARY")
        self.FOLDER_NAME = os.getenv("SHAREPOINT_FOLDER_NAME")

    def auth(self):
        self.authcookie = Office365(self.SHAREPOINT_URL, username=self.USERNAME, password=self.PASSWORD).GetCookies()
        self.site = Site(self.SHAREPOINT_SITE, version=Version.v365, authcookie=self.authcookie)
        return self.site

    def connect_folder(self):
        self.auth_site = self.auth()
        self.sharepoint_dir = '/'.join([self.SHAREPOINT_DOC, self.FOLDER_NAME])
        self.folder = self.auth_site.Folder(self.sharepoint_dir)
        return self.folder

    def get_file_content(self, file_name):
        self._folder = self.connect_folder()
        file_content = self._folder.get_file(file_name)
        return file_content  # Retorna o conteúdo bruto do arquivo

# Nome do arquivo a ser lido
file_name = "Blutrafos_GC Eventos.xlsx"

# Função para obter os dados do SharePoint e exibir no Streamlit
@st.cache_data
def load_data():
    sharepoint = SharePoint()
    try:
        file_content = sharepoint.get_file_content(file_name)

        # Carregar o conteúdo do arquivo Excel diretamente na memória
        df = pd.read_excel(BytesIO(file_content))
        
        return df
    except Exception as e:
        st.error(f"Erro ao obter os dados: {e}")
        return None

# Função principal para exibir o gráfico no Streamlit
def main():
    st.set_page_config(initial_sidebar_state="collapsed", layout="wide")
    
    # Cria colunas para posicionar a imagem no canto esquerdo
    col1, col2 = st.columns([4, 0.1])
    with col1:
        st.title("Análise Contas a Receber - Blutrafos")
    with col2:
        st.image("B (1).png", width=75)

    st.write("Abaixo estão apresentados gráficos que analisam os valores previstos para eventos de pagamento, organizados por mês e classificados por status.")
    df = load_data()
    if df is not None:
        # Adiciona um filtro para "Status Evento" na barra lateral com a opção de selecionar múltiplos valores, excluindo "Concluído"
        status_evento_options = [status for status in df['Status Evento'].unique().tolist() if status != "Concluído"]
        status_evento_selecionados = st.sidebar.multiselect("Selecione os Status Evento", status_evento_options, default=status_evento_options)
        
        camporef_options = [camporef for camporef in df['Campo Ref'].unique().tolist()]
        # Remove "Faturamento" e "Entrega" da lista padrão
        default_camporef_options = [camporef for camporef in camporef_options if camporef not in ["Faturamento", "Entrega"]]
        camporef_selecionados = st.sidebar.multiselect("Selecione os Campos de referência para eventos:", camporef_options, default=default_camporef_options)

        # Aplica os filtros de forma eficiente
        if status_evento_selecionados:
            df = df[df['Status Evento'].isin(status_evento_selecionados)]
        if camporef_selecionados:
            df = df[df['Campo Ref'].isin(camporef_selecionados)]

        # Filtra os dados para que "Dt Real Pagto" seja None
        df = df[df['Dt Real Pagto'].isna()]

        # Converte as colunas de data para datetime
        df['Dt Prev Pagto'] = pd.to_datetime(df['Dt Prev Pagto'], errors='coerce')

        # Extrai o mês/ano das datas
        df['Mes Prev Pagto'] = df['Dt Prev Pagto'].dt.to_period('M')

        # Agrupa e soma os valores por mês/ano
        previsto = df.groupby('Mes Prev Pagto')['Valor Prev'].sum()

        # Calcula os totais previstos acumulados
        previsto_acumulado = previsto.cumsum()

        # Cria o DataFrame para o gráfico
        df_plot = pd.DataFrame({
            'Mes': previsto.index.astype(str),
            'Previsto Acumulado': previsto_acumulado.values
        })

        # Formata os valores para exibição no gráfico em milhões
        df_plot['Previsto Acumulado MI'] = df_plot['Previsto Acumulado'] / 1_000_000
        df_plot['Previsto Acumulado MI'] = df_plot['Previsto Acumulado MI'].map('{:,.1f} MI'.format)

        st.markdown("---")
        # Calcula o total de Valor Prev
        total_valor_prev = df['Valor Prev'].sum()

        # Exibe o cartão com o total de Valor Prev
        st.metric(label="Total de Valor Previsto", value=f"{total_valor_prev:,.2f}")

        st.markdown("---")
        # Cria o gráfico de linha acumulado
        fig = px.line(df_plot, x='Mes', y='Previsto Acumulado', labels={
            'value': 'Valores Acumulados',
            'Mes': 'Mês/Ano'
        }, title='Previsto Acumulado', text=df_plot['Previsto Acumulado MI'])

        # Adiciona rótulos de dados e define a cor da linha
        fig.update_traces(
            mode="lines+markers+text",  # Inclui linhas, marcadores e rótulos de texto
            textposition='top center',  # Define a posição dos rótulos
            line=dict(color='#245d46')  # Define a cor da linha
        )

        # Atualiza o layout para mostrar todos os meses no eixo X e ajustar a rotação dos rótulos
        fig.update_layout(
            xaxis=dict(
                tickmode='array',
                tickvals=df_plot['Mes'],
                ticktext=df_plot['Mes'],
                tickangle=-45
            )
        )

        # Exibe o gráfico no Streamlit
        st.plotly_chart(fig)

        st.markdown("---")

        # Cria colunas para exibir as tabelas lado a lado
        col1, col2 = st.columns([1, 2.5])

        with col1:
            # Exibe a tabela com os valores previstos acumulados
            st.write("A tabela de valores previstos acumulados:")

            # Filtra os dados para o status "Atraso"
            df_atraso = df[df['Status Evento'] == 'Atraso']

            # Agrupa e soma os valores por mês/ano
            atraso = df_atraso.groupby('Mes Prev Pagto')['Valor Prev'].sum()

            # Calcula os totais acumulados para o status "Atraso"
            atraso_acumulado = atraso.cumsum()

            # Adiciona a coluna "Atraso Acumulado" ao DataFrame df_plot
            df_plot['Atraso Acumulado'] = atraso_acumulado.reindex(df_plot['Mes']).fillna(0).values

            st.dataframe(df_plot.reset_index(drop=True).style.format({
                'Previsto Acumulado': '{:,.2f}',
                'Atraso Acumulado': '{:,.2f}'
            }))

        with col2:
            st.write("A tabela abaixo mostra uma listagem ordens de venda onde o evento associado ao pagamento foi realizado:")

            # Filtra os dados para "Status Cobrança" = "Cobrar Cliente"
            df_filtered = df[df['Status Cobrança'] == 'Cobrar Cliente']

            # Seleciona as colunas desejadas e formata a data
            df_filtered = df_filtered[['Divisão', 'Cliente', 'OV', 'Nome da Obra', 'Gestor', 'Evento', 'Campo Ref', 'Dt Efetiva Ref', 'Status Evento', 'Dt Prev Pagto', 'Valor Prev', 'Observações Financeiro']]

            # Remove separadores do campo "OV"
            df_filtered['OV'] = df_filtered['OV'].astype(str).str.replace(',', '')

            # Converte a coluna 'Dt Prev Pagto' para datetime
            df_filtered['Dt Prev Pagto'] = pd.to_datetime(df_filtered['Dt Prev Pagto'], errors='coerce')
            
            # Ordena os dados pela data mais antiga de 'Dt Prev Pagto'
            df_filtered = df_filtered.sort_values(by='Dt Prev Pagto', ascending=True)
            
            # Formata a data para exibição
            df_filtered['Dt Prev Pagto'] = df_filtered['Dt Prev Pagto'].dt.strftime('%d/%m/%Y')

            # Converte a coluna 'Dt Efetiva Ref' para datetime
            df_filtered['Dt Efetiva Ref'] = pd.to_datetime(df_filtered['Dt Efetiva Ref'], errors='coerce')

            # Formata a data para exibição
            df_filtered['Dt Efetiva Ref'] = df_filtered['Dt Efetiva Ref'].dt.strftime('%d/%m/%Y')

            # Exibe a tabela filtrada no Streamlit
            st.dataframe(df_filtered.reset_index(drop=True).style.format({
                'Valor Prev': '{:,.2f}'
            }))

# Executa a função principal
if __name__ == "__main__":
    main()