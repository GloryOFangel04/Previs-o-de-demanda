import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


df_2024 = pd.read_excel('Vendas_24.xlsx')
df_2025 = pd.read_excel('Pedidos_Jan_Maio_2025.xlsx')


df_total = pd.concat([df_2024, df_2025], ignore_index=True)


df_total['DataPedido'] = pd.to_datetime(df_total['DataPedido'], errors='coerce', dayfirst=True)


df_total['Ano'] = df_total['DataPedido'].dt.year
df_total['Mês'] = df_total['DataPedido'].dt.month


df_total['Status'] = df_total['Status'].fillna('Desconhecido').astype(str)


df_2024_5m = df_total[(df_total['Ano'] == 2024) & (df_total['Mês'] <= 5)]
df_2025_5m = df_total[(df_total['Ano'] == 2025) & (df_total['Mês'] <= 5)]



# Pedidos totais
media_pedidos_2024 = df_2024_5m.groupby('Mês')['PedidoID'].nunique().mean()
media_pedidos_2025 = df_2025_5m.groupby('Mês')['PedidoID'].nunique().mean()

# Cancelamentos 
cancel_2024 = df_2024_5m[df_2024_5m['Status'].str.lower().str.contains('cancelado')]
cancel_2025 = df_2025_5m[df_2025_5m['Status'].str.lower().str.contains('cancelado')]

media_cancel_2024 = cancel_2024.groupby('Mês')['PedidoID'].nunique().mean()
media_cancel_2025 = cancel_2025.groupby('Mês')['PedidoID'].nunique().mean()

#  GRÁFICO MÉDIA DE PEDIDOS 
plt.figure(figsize=(6,4))
sns.barplot(x=['2024', '2025'], y=[media_pedidos_2024, media_pedidos_2025], palette='Greens')
plt.title('Média de Pedidos por Mês (Jan-Mai)')
plt.ylabel('Média de Pedidos')
plt.grid(True, axis='y', linestyle='--')
plt.tight_layout()
plt.show()

# GRÁFICO MÉDIA DE CANCELAMENTOS 
plt.figure(figsize=(6,4))
sns.barplot(x=['2024', '2025'], y=[media_cancel_2024, media_cancel_2025], palette='Reds')
plt.title('Média de Cancelamentos por Mês (Jan-Mai)')
plt.ylabel('Média de Cancelamentos')
plt.grid(True, axis='y', linestyle='--')
plt.tight_layout()
plt.show()
