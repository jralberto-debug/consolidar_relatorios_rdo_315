import pandas as pd
import os

# Caminhos
entrada = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO\consolidado_relatorios.xlsx"
saida = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO\etapa2_excluir_vazio\consolidado_filtrado.xlsx"

# Carregar planilha
df = pd.read_excel(entrada)

# Remover linhas onde a coluna "SERVIÇO" está vazia ou só com espaços
df_filtrado = df[df["SERVIÇO"].notna() & (df["SERVIÇO"].astype(str).str.strip() != "")]

# Salvar nova planilha
df_filtrado.to_excel(saida, index=False)

print(f"✅ Linhas com 'SERVIÇO' vazio foram removidas.")
print(f"📄 Novo arquivo salvo em: {saida}") 