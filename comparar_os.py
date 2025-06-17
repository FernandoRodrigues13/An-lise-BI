import pandas as pd
import io
import base64

# --- Configurações Iniciais ---
arquivo_bi_selecionadas = 'OS_selecionadas_BI.xlsx'
arquivo_producao = 'planilha_producao_teste.xlsx' # << COLOQUE O NOME DA SUA PLANILHA DE PRODUÇÃO COM DADOS ALEATÓRIOS
arquivo_saida_html = 'relatorio_comparativo_BI_x_Producao.html'
arquivo_saida_excel = 'comparativo_BI_x_Producao_debug.xlsx' # Para debug, se precisar conferir

coluna_chave = 'OS'
coluna_valor_comparar = 'Fat Total'

pd.options.display.float_format = '{:,.2f}'.format # Formatação para print no console

# --- 1. Carregar os DataFrames ---
print(f"Carregando '{arquivo_bi_selecionadas}'...")
df_bi = pd.read_excel(arquivo_bi_selecionadas)
print(f"'{arquivo_bi_selecionadas}' carregado ({len(df_bi)} linhas).")

print(f"Carregando '{arquivo_producao}'...")
df_producao = pd.read_excel(arquivo_producao)
print(f"'{arquivo_producao}' carregado ({len(df_producao)} linhas).")


# --- 2. Preparar a coluna chave (garantir que seja string) ---
df_bi[coluna_chave] = df_bi[coluna_chave].astype(str)
df_producao[coluna_chave] = df_producao[coluna_chave].astype(str)

# --- 3. Filtrar a planilha de produção (opcional, mas foca nas OSs do BI) ---
os_do_bi = df_bi[coluna_chave].unique()
df_producao_filtrada = df_producao[df_producao[coluna_chave].isin(os_do_bi)].copy()

if df_producao_filtrada.empty and not df_producao.empty:
    print(f"AVISO: Nenhuma das OSs de '{arquivo_bi_selecionadas}' foi encontrada em '{arquivo_producao}'.")
elif df_producao_filtrada.empty and df_producao.empty:
     print(f"AVISO: A planilha de produção '{arquivo_producao}' está vazia.")


# --- 4. Realizar o merge ---
df_comparativo = pd.merge(
    df_bi,
    df_producao_filtrada,
    on=coluna_chave,
    how='left',
    suffixes=('_BI', '_Producao')
)

# --- 5. Comparar as colunas e calcular diferenças ---
coluna_valor_bi_renomeada = f"{coluna_valor_comparar}_BI"
coluna_valor_producao_renomeada = f"{coluna_valor_comparar}_Producao"

# Se a coluna de produção não foi criada pelo merge (nenhum match), adicionar com NaN
if coluna_valor_producao_renomeada not in df_comparativo.columns and not df_producao_filtrada.empty:
    df_comparativo[coluna_valor_producao_renomeada] = pd.NA

# Garantir que a coluna do BI tenha o sufixo se não houve conflito
if coluna_valor_bi_renomeada not in df_comparativo.columns and coluna_valor_comparar in df_comparativo.columns:
    coluna_valor_bi_renomeada = coluna_valor_comparar


df_comparativo['Encontrado_Producao'] = ~df_comparativo[coluna_valor_producao_renomeada].isna()

# Para a comparação e diferença, preenchemos NaNs com 0 para evitar erros em operações,
# mas a coluna 'Encontrado_Producao' já nos diz se o dado original era NaN.
# Ou podemos ser mais explícitos: se um for NaN, a igualdade é False e a diferença pode ser o valor do outro.
# Para simplificar, se um dos valores não existe, a igualdade é falsa.
val_bi = df_comparativo[coluna_valor_bi_renomeada]
val_prod = df_comparativo[coluna_valor_producao_renomeada]

df_comparativo[f'{coluna_valor_comparar}_Igual'] = (val_bi == val_prod) & (df_comparativo['Encontrado_Producao'])

# Diferença: BI - Produção. Se Produção não encontrado, diferença é o valor do BI.
# Ou, se Produção não encontrado, pode ser NaN ou o valor do BI.
# Vamos fazer: BI - Produção (se ambos existem), senão NaN se um deles não existe.
df_comparativo[f'{coluna_valor_comparar}_Diferenca'] = val_bi.fillna(0) - val_prod.fillna(0)
# Ajustar a diferença para ser NaN se a produção não foi encontrada, para não mostrar 0 falsamente.
df_comparativo.loc[~df_comparativo['Encontrado_Producao'], f'{coluna_valor_comparar}_Diferenca'] = pd.NA


# --- 6. Reordenar colunas para o relatório ---
colunas_relatorio = [
    coluna_chave,
    coluna_valor_bi_renomeada,
    coluna_valor_producao_renomeada,
    f'{coluna_valor_comparar}_Igual',
    f'{coluna_valor_comparar}_Diferenca',
    'Encontrado_Producao'
]
# Adicionar outras colunas do BI que não sejam a chave ou o valor já listado
outras_colunas_bi_originais = [col for col in df_bi.columns if col not in [coluna_chave, coluna_valor_comparar]]
colunas_finais_ordenadas = colunas_relatorio + outras_colunas_bi_originais
# Garantir que todas as colunas selecionadas existem no df_comparativo
df_relatorio_final = df_comparativo[[col for col in colunas_finais_ordenadas if col in df_comparativo.columns]].copy()

print("\nDataFrame Comparativo para Relatório:")
print(df_relatorio_final.to_string())

# Salvar um Excel para debug, caso necessário
df_relatorio_final.to_excel(arquivo_saida_excel, index=False)
print(f"\nArquivo Excel de debug '{arquivo_saida_excel}' salvo.")

# --- 7. Gerar Relatório HTML ---

# Definir como formatar e estilizar as células no HTML
def highlight_diff(row):
    style = [''] * len(row) # Estilo padrão vazio
    is_equal_col_name = f'{coluna_valor_comparar}_Igual'
    diff_col_name = f'{coluna_valor_comparar}_Diferenca'
    encontrado_col_name = 'Encontrado_Producao'

    # Obter os índices das colunas (mais robusto que hardcoding)
    try:
        idx_igual = row.index.get_loc(is_equal_col_name)
        idx_diff = row.index.get_loc(diff_col_name)
        idx_encontrado = row.index.get_loc(encontrado_col_name)
        idx_val_prod = row.index.get_loc(coluna_valor_producao_renomeada)
        idx_val_bi = row.index.get_loc(coluna_valor_bi_renomeada)

        if not row[encontrado_col_name]: # Não encontrado na produção
            style[idx_val_prod] = 'background-color: #ffe0b2; color: #8d6e63;' # Laranja claro para não encontrado
            style[idx_diff] = 'background-color: #ffe0b2; color: #8d6e63;'
        elif row[is_equal_col_name]: # Encontrado e igual
            style[idx_diff] = 'background-color: #c8e6c9; color: #2e7d32;' # Verde para igual
            style[idx_val_prod] = 'background-color: #c8e6c9; color: #2e7d32;'
        else: # Encontrado e diferente
            style[idx_diff] = 'background-color: #ffcdd2; color: #c62828;' # Vermelho para diferente
            style[idx_val_prod] = 'background-color: #ffcdd2; color: #c62828;'
    except KeyError:
        # Se alguma coluna não existir, não aplica o estilo específico
        pass
    return style

# Colunas numéricas para formatação de float
colunas_numericas_relatorio = df_relatorio_final.select_dtypes(include=['number']).columns
format_dict = {col: '{:,.2f}' for col in colunas_numericas_relatorio}
format_dict[f'{coluna_valor_comparar}_Diferenca'] = lambda x: f'{x:,.2f}' if pd.notna(x) else 'N/A (Não encontrado)'


styled_df = (df_relatorio_final.style
             .set_caption("Relatório Comparativo BI vs. Produção")
             .format(format_dict, na_rep="N/A") # Formata números e trata NaNs
             .apply(highlight_diff, axis=1) # Aplica o estilo por linha
             .set_table_styles([
                 {'selector': 'th', 'props': [('background-color', '#2c3e50'), ('color', 'white'), ('font-weight', 'bold'), ('padding', '10px 8px'), ('text-align', 'left'), ('border-bottom', '2px solid #1abc9c')]},
                 {'selector': 'td', 'props': [('padding', '8px'), ('text-align', 'left'), ('border', '1px solid #ddd')]},
                 {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#f8f9fa')]},
                 # Não usar hover aqui pois pode conflitar com os backgrounds já aplicados
                 # {'selector': 'tr:hover', 'props': [('background-color', '#e9ecef')]},
                 {'selector': 'caption', 'props': [('caption-side', 'top'), ('font-size', '1.5em'), ('font-weight', 'bold'), ('color', '#34495e'), ('margin-bottom', '20px'), ('text-align', 'center')]}
             ])
            )

html_table_output = styled_df.to_html(index=False, escape=False)

full_html_page = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Comparativo BI vs. Produção</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600&display=swap');
        body {{ font-family: 'Open Sans', sans-serif; margin: 0; padding: 20px; background-color: #f4f7f6; color: #333; line-height: 1.6; }}
        .container {{ max-width: 95%; margin: 20px auto; padding: 25px; background-color: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-radius: 8px; }}
        h1 {{ color: #2c3e50; text-align: center; margin-bottom: 30px; border-bottom: 2px solid #1abc9c; padding-bottom: 15px; font-size: 1.8em; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 25px; font-size: 0.9em; }} /* Reduzir um pouco a fonte da tabela */
        /* Estilos da tabela são primariamente definidos pelo Styler do Pandas acima */
        .legend {{ margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-radius: 5px; border: 1px solid #eee; }}
        .legend h3 {{ margin-top: 0; color: #34495e; }}
        .legend span {{ display: inline-block; width: 20px; height: 20px; margin-right: 8px; vertical-align: middle; border: 1px solid #ccc; }}
        .legend-item {{ margin-bottom: 8px; }}
        .legend-igual {{ background-color: #c8e6c9; }}
        .legend-diferente {{ background-color: #ffcdd2; }}
        .legend-nao-encontrado {{ background-color: #ffe0b2; }}
        .footer {{ text-align: center; margin-top: 40px; font-size: 0.9em; color: #7f8c8d; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Relatório Comparativo: Dados do BI vs. Dados da Produção</h1>
        
        {html_table_output}

        <div class="legend">
            <h3>Legenda de Cores (Valor Produção / Diferença):</h3>
            <div class="legend-item"><span class="legend-igual"></span> Valores Iguais</div>
            <div class="legend-item"><span class="legend-diferente"></span> Valores Diferentes</div>
            <div class="legend-item"><span class="legend-nao-encontrado"></span> OS Não Encontrada na Produção / Valor Ausente</div>
        </div>
    </div>
    <div class="footer">
        Relatório gerado automaticamente via Python/Pandas.
    </div>
</body>
</html>
"""

with open(arquivo_saida_html, 'w', encoding='utf-8') as f:
    f.write(full_html_page)

print(f"\n📄 Relatório HTML '{arquivo_saida_html}' salvo com sucesso. Abra-o em seu navegador.")