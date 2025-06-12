import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# as duas planilhas
df_2024 = pd.read_excel('Vendas_24.xlsx')
df_2025 = pd.read_excel('Pedidos_Jan_Maio_2025.xlsx')

# Converter as colunas de data 
df_2024['DataPedido'] = pd.to_datetime(df_2024['DataPedido'], errors='coerce', dayfirst=True)
df_2024['DataEntregaPrometida'] = pd.to_datetime(df_2024['DataEntregaPrometida'], errors='coerce', dayfirst=True)
df_2024['DataEntregaReal'] = pd.to_datetime(df_2024['DataEntregaReal'], errors='coerce', dayfirst=True)

df_2025['DataPedido'] = pd.to_datetime(df_2025['DataPedido'], errors='coerce', dayfirst=True)
df_2025['DataEntregaPrometida'] = pd.to_datetime(df_2025['DataEntregaPrometida'], errors='coerce', dayfirst=True)
df_2025['DataEntregaReal'] = pd.to_datetime(df_2025['DataEntregaReal'], errors='coerce', dayfirst=True)

# juntando as planilhas 
df_total = pd.concat([df_2024, df_2025], ignore_index=True)

# Verificar colunas
print("Linhas no total:", df_total.shape[0])
print("Colunas:", df_total.columns.tolist())


# Extrair mês e ano
df_2024['Ano'] = df_2024['DataPedido'].dt.year
df_2024['Mês'] = df_2024['DataPedido'].dt.month
df_2025['Ano'] = df_2025['DataPedido'].dt.year
df_2025['Mês'] = df_2025['DataPedido'].dt.month

# filtro apenas os meses de janeiro a maio
df_2024_5m = df_2024[df_2024['Mês'] <= 5]
df_2025_5m = df_2025[df_2025['Mês'] <= 5]

# grupo por mês
vendas_2024 = df_2024_5m.groupby('Mês')['Quantidade'].sum()
vendas_2025 = df_2025_5m.groupby('Mês')['Quantidade'].sum()

# Garantir que todos os meses de 1 a 5 estejam presentes
meses = [1, 2, 3, 4, 5]
vendas_2024 = vendas_2024.reindex(meses, fill_value=0)
vendas_2025 = vendas_2025.reindex(meses, fill_value=0)

# Comparativo
vendas_comparativo = pd.DataFrame({
    '2024': vendas_2024,
    '2025': vendas_2025
})

# Cálculo da variação percentual (evitando divisão por zero)
vendas_comparativo['% Variação'] = (
    (vendas_comparativo['2025'] - vendas_comparativo['2024']) / 
    vendas_comparativo['2024'].replace(0, pd.NA)
) * 100

# Renomear os índices de mês
nomes_meses = {1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio'}
vendas_comparativo.index = vendas_comparativo.index.map(nomes_meses)

# imprimindo no console
print("\n Comparativo de Vendas Jan-Mai:")
print(vendas_comparativo)

# Plotar gráfico
variacao_plot = vendas_comparativo['% Variação'].dropna()

if not variacao_plot.empty:
    variacao_plot.plot(
        kind='bar',
        figsize=(10, 5),
        title='Variação % nas Vendas (Jan-Mai 2025 vs 2024)',
        ylabel='% Variação',
        color='steelblue'
    )
    plt.xticks(rotation=45)
    plt.grid(True, axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.show()
else:
    print(" Nenhuma variação percentual válida para plotar.")

# Gráfico de linhas - evolução dos pedidos
plt.figure(figsize=(10, 5))
plt.plot(vendas_2024.index, vendas_2024.values, marker='o', label='2024', linestyle='-', color='tab:blue')
plt.plot(vendas_2025.index, vendas_2025.values, marker='o', label='2025', linestyle='--', color='tab:orange')
plt.title('Evolução dos Pedidos - Janeiro a Maio')
plt.xlabel('Mês')
plt.ylabel('Quantidade de Pedidos')
plt.xticks([1, 2, 3, 4, 5], ['Jan', 'Fev', 'Mar', 'Abr', 'Mai'])
plt.grid(True, linestyle='--', alpha=0.5)
plt.legend()
plt.tight_layout()
plt.show()


#  UF QUE MAIS RECEBEU VENDAS


print("\n UF com maior quantidade de pedidos:")

uf_2024 = df_2024_5m.groupby('UF')['Quantidade'].sum().sort_values(ascending=False)
uf_2025 = df_2025_5m.groupby('UF')['Quantidade'].sum().sort_values(ascending=False)

uf_2024_top = uf_2024.idxmax()
uf_2025_top = uf_2025.idxmax()

print(f"🔹 2024: {uf_2024_top} ({uf_2024.max()} pedidos)")
print(f"🔸 2025: {uf_2025_top} ({uf_2025.max()} pedidos)")

# Gráfico comparativo das 5 UFs que mais receberam pedidos em cada ano
top_ufs = pd.DataFrame({
    '2024': uf_2024.head(5),
    '2025': uf_2025.head(5)
}).fillna(0)

top_ufs.plot(kind='bar', figsize=(10, 5), title='Top 5 UFs por Quantidade de Pedidos (Jan-Mai)')
plt.ylabel('Quantidade de Pedidos')
plt.xticks(rotation=45)
plt.grid(True, axis='y', linestyle='--', alpha=0.6)
plt.tight_layout()
plt.show()

# Status possíveis
todos_status = sorted(list(set(df_2024_5m['Status'].unique()) | set(df_2025_5m['Status'].unique())))

# Contagem por status
status_2024 = df_2024_5m['Status'].value_counts().reindex(todos_status, fill_value=0)
status_2025 = df_2025_5m['Status'].value_counts().reindex(todos_status, fill_value=0)

# Comparativo
status_comparativo = pd.DataFrame({
    '2024': status_2024,
    '2025': status_2025
})

print("\n Comparativo de Status dos Pedidos (Jan-Mai):")
print(status_comparativo)

# Gráfico de barras
status_comparativo.plot(kind='bar', figsize=(10, 5), title='Comparativo de Status dos Pedidos (Jan-Mai)')
plt.ylabel('Quantidade de Pedidos')
plt.xticks(rotation=45)
plt.grid(True, axis='y', linestyle='--', alpha=0.6)
plt.tight_layout()
plt.show()


#UF que mais recebe pedidos


uf_2024 = df_2024_5m['UF'].value_counts()
uf_2025 = df_2025_5m['UF'].value_counts()

uf_mais_2024 = uf_2024.idxmax()
uf_mais_2025 = uf_2025.idxmax()

print(f" UF com mais entregas em 2024: {uf_mais_2024} ({uf_2024.max()} pedidos)")
print(f" UF com mais entregas em 2025: {uf_mais_2025} ({uf_2025.max()} pedidos)")

if uf_mais_2024 == uf_mais_2025:
    print(" É a mesma UF nos dois anos.\n")
else:
    print(" São UFs diferentes nos dois anos.\n")

# Gráfico comparativo de UFs
uf_comparativo = pd.DataFrame({'2024': uf_2024, '2025': uf_2025}).fillna(0)
uf_comparativo.plot(kind='bar', figsize=(12, 5), title='Pedidos por UF (Jan-Mai)')
plt.ylabel('Total de Pedidos')
plt.xticks(rotation=45)
plt.tight_layout()
plt.grid(True, axis='y', linestyle='--')
plt.show()

#SKU mais e menos vendido

sku_2024 = df_2024_5m.groupby('SKU')['Quantidade'].sum()
sku_2025 = df_2025_5m.groupby('SKU')['Quantidade'].sum()

sku_mais_2024 = sku_2024.idxmax()
sku_menos_2024 = sku_2024.idxmin()

sku_mais_2025 = sku_2025.idxmax()
sku_menos_2025 = sku_2025.idxmin()

print(f" SKU mais vendido em 2024: {sku_mais_2024} ({sku_2024.max()} unidades)")
print(f"SKU menos vendido em 2024: {sku_menos_2024} ({sku_2024.min()} unidades)\n")

print(f" SKU mais vendido em 2025: {sku_mais_2025} ({sku_2025.max()} unidades)")
print(f" SKU menos vendido em 2025: {sku_menos_2025} ({sku_2025.min()} unidades)\n")



#  Maior Lead Time e comparação


lead_2024 = df_2024_5m['LeadTime'].max()
lead_2025 = df_2025_5m['LeadTime'].max()

print(f" Maior Lead Time em 2024: {lead_2024} dias")
print(f" Maior Lead Time em 2025: {lead_2025} dias")

if lead_2025 > lead_2024:
    print(" O Lead Time aumentou em 2025.\n")
elif lead_2025 < lead_2024:
    print(" O Lead Time diminuiu em 2025.\n")
else:
    print(" O Lead Time se manteve igual nos dois anos.\n")


# LEAD TIME MÉDIO COMPARATIVO

lead_mean_2024 = df_2024_5m['LeadTime'].mean()
lead_mean_2025 = df_2025_5m['LeadTime'].mean()

plt.figure(figsize=(6,4))
sns.barplot(x=['2024', '2025'], y=[lead_mean_2024, lead_mean_2025], palette='Blues_d')
plt.title('Lead Time Médio (em dias) - Jan a Mai')
plt.ylabel('Dias')
plt.ylim(0, max(lead_mean_2024, lead_mean_2025) + 2)
plt.grid(True, axis='y', linestyle='--')
plt.tight_layout()
plt.show()


# MOTIVOS DE CANCELAMENTO



cancelados_2024 = df_2024_5m[df_2024_5m['Status'].str.lower().str.contains('cancelado')]
cancelados_2025 = df_2025_5m[df_2025_5m['Status'].str.lower().str.contains('cancelado')]


motivos_2024 = cancelados_2024['MotivoCancelamento'].value_counts()
motivos_2025 = cancelados_2025['MotivoCancelamento'].value_counts()


df_motivos = pd.DataFrame({'2024': motivos_2024, '2025': motivos_2025}).fillna(0)

# Gráfico comparativo
df_motivos.plot(kind='barh', figsize=(10, 6), color=['#1f77b4', '#ff7f0e'])
plt.title('Motivos de Cancelamento (Jan a Mai)')
plt.xlabel('Quantidade de Pedidos')
plt.ylabel('Motivo')
plt.grid(True, axis='x', linestyle='--')
plt.tight_layout()
plt.show()

# . Dia da semana com mais pedidos


dia_2024 = df_2024_5m['Dia da Semana'].value_counts()
dia_2025 = df_2025_5m['Dia da Semana'].value_counts()

dia_mais_2024 = dia_2024.idxmax()
dia_mais_2025 = dia_2025.idxmax()

print(f" Dia mais cheio de pedidos em 2024: {dia_mais_2024} ({dia_2024.max()} pedidos)")
print(f" Dia mais cheio de pedidos em 2025: {dia_mais_2025} ({dia_2025.max()} pedidos)")

if dia_mais_2024 == dia_mais_2025:
    print(" Mesmo dia nos dois anos.\n")
else:
    print(" Dias diferentes entre 2024 e 2025.\n")

# Gráfico de pedidos por dia da semana
dias_comparativo = pd.DataFrame({'2024': dia_2024, '2025': dia_2025}).fillna(0)
dias_comparativo = dias_comparativo.reindex(['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado', 'domingo'])

dias_comparativo.plot(kind='bar', figsize=(10, 5), title='Pedidos por Dia da Semana (Jan-Mai)')
plt.ylabel('Total de Pedidos')
plt.xticks(rotation=45)
plt.tight_layout()
plt.grid(True, axis='y', linestyle='--')
plt.show()
