import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ---------------------------------------------------------
# CONFIGURAÇÃO
# ---------------------------------------------------------
# Coloque o nome exato do seu arquivo aqui (certifique-se de que é .xlsx)
nome_do_arquivo = "chamados.xlsx" 

try:
    # ---------------------------------------------------------
    # 1. CARREGAMENTO DOS DADOS (MODO REAL)
    # ---------------------------------------------------------
    print(f"Lendo o arquivo: {nome_do_arquivo}...")
    df = pd.read_excel(nome_do_arquivo)
    
    # Limpeza preventiva dos nomes das colunas (remove espaços extras antes/depois)
    # Ex: Transforma " Status " em "Status"
    df.columns = df.columns.str.strip()

    # Verifica se as colunas essenciais existem
    colunas_necessarias = ['Status', 'Data Abertura', 'Subcategoria', 'Prioridade']
    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"A coluna '{col}' não foi encontrada no Excel. Verifique se o nome está exato.")

    # ---------------------------------------------------------
    # 2. LIMPEZA E TRATAMENTO DE DADOS
    # ---------------------------------------------------------
    
    # Converter 'Data Abertura' para data (trata erros como '-' transformando em NaT)
    # dayfirst=True força a leitura como Dia/Mês/Ano (padrão Brasil)
    df['Data Abertura'] = pd.to_datetime(df['Data Abertura'], dayfirst=True, errors='coerce')
    
    # Criar coluna apenas com a Data (para o gráfico de linha)
    df['Data_Dia'] = df['Data Abertura'].dt.date

    # Limpar espaços em branco nos textos das colunas categóricas
    cols_texto = ['Status', 'Subcategoria', 'Prioridade']
    for col in cols_texto:
        # Converte para string e remove espaços
        df[col] = df[col].astype(str).str.strip()

    # ---------------------------------------------------------
    # 3. CRIAÇÃO DO VISUAL (DASHBOARD)
    # ---------------------------------------------------------
    print("Gerando gráficos...")
    
    # Definir estilo visual
    sns.set_theme(style="whitegrid")
    plt.rcParams['font.family'] = 'sans-serif'

    # Criar figura 2x2
    fig, axes = plt.subplots(2, 2, figsize=(18, 12))
    fig.suptitle(f'Relatório de Chamados T.I. - {nome_do_arquivo}', fontsize=20, fontweight='bold', y=0.96)

    # --- GRÁFICO 1: Status (Pizza/Rosca) ---
    status_counts = df['Status'].value_counts()
    # Pega cores suficientes para a quantidade de status
    colors = sns.color_palette('pastel')[0:len(status_counts)]
    
    axes[0, 0].pie(status_counts, labels=status_counts.index, autopct='%1.1f%%', startangle=140, colors=colors, wedgeprops=dict(width=0.4))
    axes[0, 0].set_title('Distribuição por Status', fontsize=14, fontweight='bold')

    # --- GRÁFICO 2: Top 10 Subcategorias (Barras Horizontais) ---
    # Aumentei para Top 10 para dar mais detalhe se tiver muitos tipos
    top_problems = df['Subcategoria'].value_counts().head(10)
    sns.barplot(x=top_problems.values, y=top_problems.index, ax=axes[0, 1], palette="viridis", hue=top_problems.index, legend=False)
    axes[0, 1].set_title('Top 10 Assuntos/Subcategorias', fontsize=14, fontweight='bold')
    axes[0, 1].set_xlabel('Quantidade de Chamados')
    axes[0, 1].set_ylabel('')

    # --- GRÁFICO 3: Prioridade ---
    # Define a ordem lógica das prioridades
    ordem_prioridade = ['Baixa', 'Média', 'Alta', 'Crítica']
    # Filtra apenas as prioridades que existem nos dados atuais para não dar erro
    ordem_existente = [p for p in ordem_prioridade if p in df['Prioridade'].unique()]
    # Se houver prioridades fora do padrão, adiciona elas ao final
    outras = [p for p in df['Prioridade'].unique() if p not in ordem_prioridade]
    ordem_final = ordem_existente + outras

    sns.countplot(x='Prioridade', data=df, ax=axes[1, 0], palette="magma", order=ordem_final, hue='Prioridade', legend=False)
    axes[1, 0].set_title('Volume por Prioridade', fontsize=14, fontweight='bold')
    axes[1, 0].set_xlabel('')
    axes[1, 0].set_ylabel('Quantidade')

    # --- GRÁFICO 4: Evolução Diária (Linha) ---
    chamados_por_dia = df.groupby('Data_Dia').size()
    
    if not chamados_por_dia.empty:
        chamados_por_dia.plot(kind='line', marker='o', ax=axes[1, 1], color='#2980b9', linewidth=2)
        axes[1, 1].set_title('Volume de Abertura (Dia a Dia)', fontsize=14, fontweight='bold')
        axes[1, 1].set_xlabel('Data')
        axes[1, 1].grid(True, linestyle='--', alpha=0.7)
    else:
        axes[1, 1].text(0.5, 0.5, "Sem dados de data válidos", ha='center')

    # Ajuste fino do layout
    plt.tight_layout(rect=[0, 0.03, 1, 0.95])

    # Salvar a imagem final
    nome_imagem = "Relatorio_TI_Visual.png"
    plt.savefig(nome_imagem, dpi=300)
    print(f"Sucesso! Relatório salvo como '{nome_imagem}' e exibido na tela.")
    
    plt.show()

except FileNotFoundError:
    print(f"ERRO: O arquivo '{nome_do_arquivo}' não foi encontrado.")
    print("Verifique se o nome está correto e se ele está na mesma pasta do script.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")