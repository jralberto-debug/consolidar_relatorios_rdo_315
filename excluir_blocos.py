import pandas as pd

# Caminhos
input_file = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO\etapa2_excluir_vazio\consolidado_filtrado.xlsx"
output_file = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO\etapa3_excluir_blocos\consolidado_filtrado_sem_blocos.xlsx"

# Carregar a planilha
df = pd.read_excel(input_file)

# Coluna O é a 15ª coluna (índice 14)
coluna = df.columns[14]  # Certifique-se que essa coluna é "ACUM. REAL"

# Filtrar: mantém apenas as linhas onde a coluna não é nula, diferente de 0 e diferente de "-"
df_filtrado = df[(df[coluna].notna()) & (df[coluna] != 0) & (df[coluna] != "-")]

# Salvar nova planilha
df_filtrado.to_excel(output_file, index=False)

print("✅ Arquivo salvo em:", output_file)

