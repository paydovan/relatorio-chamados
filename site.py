import streamlit as st
import pandas as pd
import plotly.express as px
from collections import Counter
import re

# ---------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------
st.set_page_config(page_title="Dashboard TI", page_icon="💻", layout="wide")

# Título Principal
st.title("📊 Relatório Executivo de Chamados T.I.")
st.markdown("Visão interativa e analítica dos tickets de suporte.")

# ---------------------------------------------------------
# FUNÇÃO: ANÁLISE DE TEXTO (NOVA FEATURE)
# ---------------------------------------------------------
def analisar_texto(df_alvo, coluna_texto):
    """
    Pega uma coluna de texto, limpa palavras comuns (stopwords)
    e conta a frequência das palavras restantes.
    """
    if coluna_texto not in df_alvo.columns:
        return pd.DataFrame()

    # 1. Juntar todo o texto em uma única string
    texto_completo = " ".join(df_alvo[coluna_texto].dropna().astype(str).tolist())
    
    # 2. Limpeza (deixar minúsculo e remover pontuação/números)
    texto_limpo = texto_completo.lower()
    texto_limpo = re.sub(r'[^\w\s]', '', texto_limpo) # remove pontuação
    texto_limpo = re.sub(r'\d+', '', texto_limpo)     # remove números
    
    # 3. Lista de Stopwords (Palavras para ignorar)
    # Adicione ou remova palavras aqui conforme a necessidade da sua empresa
    stopwords = [
        'de', 'a', 'o', 'que', 'e', 'do', 'da', 'em', 'um', 'para', 'é', 'com', 'não', 'uma', 'os', 'no',
        'se', 'na', 'por', 'mais', 'as', 'dos', 'como', 'mas', 'ao', 'ele', 'das', 'tem', 'à', 'seu', 'sua',
        'ou', 'ser', 'quando', 'muito', 'nos', 'já', 'está', 'eu', 'também', 'só', 'pelo', 'pela', 'até',
        'isso', 'ela', 'entre', 'depois', 'sem', 'mesmo', 'aos', 'ter', 'seus', 'quem', 'nas', 'me', 'esse',
        'eles', 'estão', 'você', 'tinha', 'foram', 'essa', 'num', 'nem', 'suas', 'meu', 'minha', 'têm', 
        'numa', 'pelos', 'elas', 'havia', 'seja', 'qual', 'será', 'nós', 'tenho', 'lhe', 'deles', 'essas', 
        'esses', 'pelas', 'este', 'fosse', 'dele', 
        # Palavras de "educação" e comuns em emails que não agregam análise técnica
        'bom', 'dia', 'tarde', 'noite', 'favor', 'att', 'grato', 'obrigado', 'obrigada',
        'ola', 'olá', 'prezados', 'caro', 'cara',
        # Palavras genéricas de chamado que não indicam a causa raiz
        'chamado', 'solicito', 'verificar', 'erro', 'problema', 'ticket', 'abertura', 'gentileza'
    ]
    
    # 4. Separar palavras e filtrar
    palavras = texto_limpo.split()
    # Filtra stopwords e palavras muito curtas (menos de 2 letras)
    palavras_filtradas = [p for p in palavras if p not in stopwords and len(p) > 2]
    
    # 5. Contar frequência
    contagem = Counter(palavras_filtradas)
    
    # Transformar em DataFrame para o gráfico (Top 30 palavras)
    df_palavras = pd.DataFrame(contagem.most_common(30), columns=['Palavra', 'Frequência'])
    return df_palavras

# ---------------------------------------------------------
# CARREGAMENTO E PROCESSAMENTO DE DADOS
# ---------------------------------------------------------
st.sidebar.header("📁 Carregar Dados")
uploaded_file = st.sidebar.file_uploader("Faça upload do Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Função com cache para não recarregar o Excel a cada clique
    @st.cache_data
    def load_data(file):
        try:
            df = pd.read_excel(file)
            # Limpeza básica de nomes de colunas
            df.columns = df.columns.str.strip()
            
            # Tratamento de Datas
            if 'Data Abertura' in df.columns:
                df['Data Abertura'] = pd.to_datetime(df['Data Abertura'], dayfirst=True, errors='coerce')
                df['Data_Dia'] = df['Data Abertura'].dt.date
            
            # Tratamento de Strings (remove espaços extras)
            cols_texto = ['Status', 'Subcategoria', 'Prioridade', 'PDV', 'Assunto', 'Categoria']
            for col in cols_texto:
                 if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()
            
            return df
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")
            return None

    df = load_data(uploaded_file)

    if df is not None:
        # ---------------------------------------------------------
        # BARRA LATERAL (FILTROS)
        # ---------------------------------------------------------
        st.sidebar.header("🔍 Filtros")
        
        # Filtro de Data
        if 'Data_Dia' in df.columns:
            min_date = df['Data_Dia'].dropna().min()
            max_date = df['Data_Dia'].dropna().max()
            try:
                date_range = st.sidebar.date_input("Período", [min_date, max_date])
            except:
                st.sidebar.warning("Verifique as datas no Excel.")
                date_range = [min_date, max_date]
        else:
            date_range = []

        # Filtro de Prioridade
        if 'Prioridade' in df.columns:
            all_priorities = list(df['Prioridade'].unique())
            selected_priorities = st.sidebar.multiselect("Prioridade", all_priorities, default=all_priorities)
        else:
            selected_priorities = []

        # Filtro de Status
        if 'Status' in df.columns:
            all_status = list(df['Status'].unique())
            selected_status = st.sidebar.multiselect("Status", all_status, default=all_status)
        else:
            selected_status = []

        # APLICAR FILTROS
        # Inicia com todos os dados
        df_filtered = df.copy()

        # Aplica máscaras se as colunas existirem
        if 'Data_Dia' in df.columns and len(date_range) == 2:
            mask_date = (df['Data_Dia'] >= date_range[0]) & (df['Data_Dia'] <= date_range[1])
            df_filtered = df_filtered.loc[mask_date]
        
        if 'Prioridade' in df.columns and selected_priorities:
            df_filtered = df_filtered[df_filtered['Prioridade'].isin(selected_priorities)]
            
        if 'Status' in df.columns and selected_status:
            df_filtered = df_filtered[df_filtered['Status'].isin(selected_status)]

        # ---------------------------------------------------------
        # DASHBOARD - KPIs
        # ---------------------------------------------------------
        st.markdown("### Visão Geral")
        col1, col2, col3, col4 = st.columns(4)
        
        total_chamados = len(df_filtered)
        
        # Tenta calcular métricas se as colunas existirem
        abertos = len(df_filtered[df_filtered['Status'] == 'Aberto']) if 'Status' in df_filtered.columns else 0
        andamento = len(df_filtered[df_filtered['Status'] == 'Andamento']) if 'Status' in df_filtered.columns else 0
        finalizados = len(df_filtered[df_filtered['Status'] == 'Finalizado']) if 'Status' in df_filtered.columns else 0

        col1.metric("Total Selecionado", total_chamados)
        col2.metric("Em Aberto", abertos, delta_color="inverse")
        col3.metric("Em Andamento", andamento)
        col4.metric("Finalizados", finalizados)

        st.markdown("---")

        # ---------------------------------------------------------
        # DASHBOARD - GRÁFICOS LINHA 1
        # ---------------------------------------------------------
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            st.subheader("Onde dói mais? (Top 10 Subcategorias)")
            if 'Subcategoria' in df_filtered.columns:
                top_subs = df_filtered['Subcategoria'].value_counts().head(10).reset_index()
                top_subs.columns = ['Subcategoria', 'Qtd']
                fig_bar = px.bar(top_subs, x='Qtd', y='Subcategoria', orientation='h', 
                                 text='Qtd', color='Qtd', color_continuous_scale='Bluered')
                fig_bar.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Coluna 'Subcategoria' não encontrada.")

        with col_g2:
            st.subheader("Status dos Chamados")
            if 'Status' in df_filtered.columns:
                status_counts = df_filtered['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Qtd']
                fig_pie = px.pie(status_counts, values='Qtd', names='Status', hole=0.4, 
                                 color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Coluna 'Status' não encontrada.")

        # ---------------------------------------------------------
        # DASHBOARD - GRÁFICOS LINHA 2
        # ---------------------------------------------------------
        col_g3, col_g4 = st.columns(2)

        with col_g3:
            st.subheader("Evolução Diária")
            if 'Data_Dia' in df_filtered.columns:
                daily_counts = df_filtered.groupby('Data_Dia').size().reset_index(name='Qtd')
                fig_line = px.line(daily_counts, x='Data_Dia', y='Qtd', markers=True, line_shape='spline')
                st.plotly_chart(fig_line, use_container_width=True)
            else:
                st.info("Coluna de data não encontrada para montar a linha do tempo.")

        with col_g4:
            st.subheader("Volume por Prioridade")
            if 'Prioridade' in df_filtered.columns:
                # Define ordem lógica se possível
                ordem = ["Baixa", "Média", "Alta", "Crítica"]
                fig_col = px.histogram(df_filtered, x='Prioridade', color='Prioridade', 
                                       category_orders={"Prioridade": ordem})
                st.plotly_chart(fig_col, use_container_width=True)
            else:
                st.info("Coluna 'Prioridade' não encontrada.")

        # ---------------------------------------------------------
        # DASHBOARD - ANÁLISE DE TEXTO (NOVO)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader("🕵️ Análise de Termos Recorrentes (Mineração de Texto)")
        st.markdown("Palavras mais frequentes encontradas no campo **'Assunto'** (excluindo palavras comuns).")
        
        if 'Assunto' in df_filtered.columns:
            df_palavras = analisar_texto(df_filtered, 'Assunto')
            
            if not df_palavras.empty:
                # Gráfico de barras verticais para as palavras
                fig_word = px.bar(df_palavras, x='Palavra', y='Frequência', 
                                  text='Frequência', color='Frequência',
                                  color_continuous_scale='Tealgrn')
                fig_word.update_layout(xaxis_tickangle=-45) # Inclina o texto para caber melhor
                st.plotly_chart(fig_word, use_container_width=True)
            else:
                st.warning("Não foram encontradas palavras suficientes para análise após a limpeza.")
        else:
            st.error("Coluna 'Assunto' não encontrada no arquivo.")

        # ---------------------------------------------------------
        # DADOS BRUTOS
        # ---------------------------------------------------------
        with st.expander("Ver Tabela de Dados Completa"):
            st.dataframe(df_filtered)

else:
    st.info("👈 Aguardando upload do arquivo Excel na barra lateral.")