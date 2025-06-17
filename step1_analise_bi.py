import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
import base64
from matplotlib.ticker import FuncFormatter # Para formatar o eixo Y do gráfico

# === 0. Configurações de visualização no terminal e gráfico ===
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_rows', None)
pd.options.display.float_format = '{:,.2f}'.format
sns.set_theme(style="whitegrid") # Define um tema padrão para os gráficos Seaborn

# === 1. Carregar a planilha BI ===
arquivo_bi = 'BI.xlsx'
df_bi = pd.read_excel(arquivo_bi)

# === 2. Selecionar as 3 maiores O.S com base em "Fat Total" ===
maiores = df_bi.nlargest(3, 'Fat Total')

# === 3. Selecionar as 3 menores O.S com base em "Fat Total" ===
menores = df_bi.nsmallest(3, 'Fat Total')

# === 4. Remover as O.S já selecionadas para evitar repetição ===
ids_excluidos = maiores.index.union(menores.index)
restantes = df_bi.drop(index=ids_excluidos)

# === 5. Selecionar 4 O.S aleatórias entre as restantes ===
n_aleatorias = min(4, len(restantes))
aleatorias = restantes.sample(n=n_aleatorias, random_state=42)

# === 6. Juntar todas ===
selecionadas = pd.concat([maiores, menores, aleatorias])

# === 7. Ordenar por "Fat Total" (decrescente) ===
selecionadas = selecionadas.sort_values(by='Fat Total', ascending=False).reset_index(drop=True)

# === 8. Calcular a Soma Total do "Fat Total" das OS selecionadas ===
soma_fat_total_selecionadas = selecionadas['Fat Total'].sum()
soma_formatada = f"R$ {soma_fat_total_selecionadas:,.2f}" # Formata para exibição

# === 9. Exibir no terminal (incluindo a soma) ===
print("\n📊 O.S Selecionadas (ordenadas por Fat Total):\n")
print(selecionadas.to_string(index=False))
print(f"\n💰 Soma do 'Fat Total' das O.S. selecionadas: {soma_formatada}")


# === 10. Salvar o resultado em um novo arquivo Excel ===
arquivo_excel_saida = 'OS_selecionadas_BI.xlsx'
selecionadas.to_excel(arquivo_excel_saida, index=False)
print(f"\n✅ Arquivo '{arquivo_excel_saida}' salvo com sucesso.")

# === 11. Gerar Gráfico de Barras do "Fat Total" por OS ===
plt.figure(figsize=(10, 6)) # Define o tamanho da figura do gráfico
# Converter a coluna 'OS' para string para garantir que seja tratada como categórica no gráfico
selecionadas_grafico = selecionadas.copy()
selecionadas_grafico['OS_str'] = selecionadas_grafico['OS'].astype(str)

barplot = sns.barplot(x='OS_str', y='Fat Total', data=selecionadas_grafico, palette="viridis", hue='OS_str', dodge=False, legend=False)
plt.title('Faturamento Total por O.S. Selecionada', fontsize=16)
plt.xlabel('Ordem de Serviço (OS)', fontsize=12)
plt.ylabel('Faturamento Total (R$)', fontsize=12)
plt.xticks(rotation=45, ha='right') # Rotaciona os labels do eixo X para melhor visualização
plt.yticks(fontsize=10)

# Formatar o eixo Y para mostrar como moeda
formatter = FuncFormatter(lambda y, _: f'R$ {y:,.0f}') # Mostra sem centavos para não poluir
barplot.yaxis.set_major_formatter(formatter)

plt.tight_layout() # Ajusta o layout para evitar cortes

# Salvar gráfico em um buffer de bytes para embutir no HTML
img_buffer = io.BytesIO()
plt.savefig(img_buffer, format='png', dpi=100) # dpi pode ser ajustado
img_buffer.seek(0)
imagem_base64 = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
plt.close() # Fecha a figura para liberar memória

# === 12. Salvar o resultado em um arquivo HTML com CSS aprimorado, soma e gráfico ===
arquivo_html_saida = 'OS_selecionadas_BI_completo.html'

colunas_numericas = selecionadas.select_dtypes(include=['number']).columns

html_table = (selecionadas.style
               .set_caption("O.S. Selecionadas do BI (Ordenadas por Fat Total)")
               .format("{:,.2f}", subset=pd.IndexSlice[:, colunas_numericas], na_rep="-")
               .set_table_styles([
                   {'selector': 'th', 'props': [('background-color', '#2c3e50'), ('color', 'white'), ('font-weight', 'bold'), ('padding', '10px 8px'), ('text-align', 'left'), ('border-bottom', '2px solid #1abc9c')]},
                   {'selector': 'td', 'props': [('padding', '8px'), ('text-align', 'left'), ('border', '1px solid #ddd')]},
                   {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#f8f9fa')]},
                   {'selector': 'tr:hover', 'props': [('background-color', '#e9ecef')]},
                   {'selector': 'caption', 'props': [('caption-side', 'top'), ('font-size', '1.2em'), ('font-weight', 'bold'), ('color', '#34495e'), ('margin-bottom', '15px'), ('text-align', 'left')]}
               ])
               .to_html(index=False, escape=False))

full_html_page = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Análise de O.S. Selecionadas do BI</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600&display=swap');
        body {{ font-family: 'Open Sans', sans-serif; margin: 0; padding: 20px; background-color: #f4f7f6; color: #333; line-height: 1.6; }}
        .container {{ max-width: 90%; margin: 20px auto; padding: 25px; background-color: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-radius: 8px; }}
        h1 {{ color: #2c3e50; text-align: center; margin-bottom: 25px; border-bottom: 2px solid #1abc9c; padding-bottom: 15px; font-size: 1.8em; }}
        h2 {{ color: #34495e; margin-top: 30px; margin-bottom: 15px; font-size: 1.4em; border-bottom: 1px solid #eee; padding-bottom: 5px;}}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 25px; }}
        .summary-box {{ background-color: #e9ecef; padding: 15px; border-radius: 5px; margin-bottom: 25px; text-align: center; border: 1px solid #ced4da; }}
        .summary-box p {{ margin: 5px 0; font-size: 1.1em; }}
        .summary-box strong {{ color: #1abc9c; font-size: 1.3em; }}
        .chart-container {{ text-align: center; margin-bottom: 25px; padding:15px; border: 1px solid #eee; border-radius: 5px; background-color: #fdfdfd; }}
        .chart-container img {{ max-width: 100%; height: auto; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
        .footer {{ text-align: center; margin-top: 40px; font-size: 0.9em; color: #7f8c8d; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Análise Detalhada das Ordens de Serviço Selecionadas (BI)</h1>

        <h2>Sumário Geral</h2>
        <div class="summary-box">
            <p>Soma Total do 'Fat Total' das 10 O.S. Selecionadas: <strong>{soma_formatada}</strong></p>
        </div>

        <h2>Gráfico de Faturamento por O.S.</h2>
        <div class="chart-container">
            <img src="data:image/png;base64,{imagem_base64}" alt="Gráfico de Fat Total por OS">
        </div>

        <h2>Tabela Detalhada das O.S. Selecionadas</h2>
        {html_table}
    </div>
    <div class="footer">
        Relatório gerado automaticamente via Python/Pandas & Matplotlib.
    </div>
</body>
</html>
"""

with open(arquivo_html_saida, 'w', encoding='utf-8') as f:
    f.write(full_html_page)
print(f"\n📄 Arquivo HTML completo '{arquivo_html_saida}' salvo com sucesso. Abra-o em seu navegador.")