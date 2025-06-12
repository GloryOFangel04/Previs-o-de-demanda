import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Importando a planilha 
df = pd.read_excel('Pedidos_Jan_Maio_2025.xlsx')
df['DataPedido'] = pd.to_datetime(df['DataPedido'], errors='coerce', dayfirst=True)
df['DataEntregaPrometida'] = pd.to_datetime(df['DataEntregaPrometida'], errors='coerce', dayfirst=True)
df['DataEntregaReal'] = pd.to_datetime(df['DataEntregaReal'], errors='coerce', dayfirst=True)



# Mostrar números de linhas e colunas
print(f"total de linhas: {df.shape[0]}")
print(f"total de colunas: {df.shape[1]}")

# Mês e dia da semana do pedido
df['Mês'] = df['DataPedido'].dt.month_name(locale='pt_BR')
df['DiaSemana'] = df['DataPedido'].dt.day_name(locale='pt_BR')

# Ordenar corretamente os meses
ordem_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio']
df['Mês'] = pd.Categorical(df['Mês'], categories=ordem_meses, ordered=True)

# Vendas por mês 
vendas_por_mes = df.groupby('Mês')['Quantidade'].sum().reindex(ordem_meses)
mes_mais_vendido = vendas_por_mes.idxmax()
mes_menos_vendido = vendas_por_mes.idxmin()

print(f"Mês com mais vendas: {mes_mais_vendido}")
print(f"Mês com menos vendas: {mes_menos_vendido}")

# UF 
uf_mais_compra = df.groupby('UF')['Quantidade'].sum().idxmax()
print(f"UF que mais Comprou: {uf_mais_compra}")

# Item SKU mais ou menos vendidos
sku_vendas = df.groupby('SKU')['Quantidade'].sum()
sku_mais_vendido = sku_vendas.idxmax()
sku_menos_vendido = sku_vendas.idxmin()

print(f"SKU mais vendido: {sku_mais_vendido}")
print(f"SKU menos vendido: {sku_menos_vendido}")

# Lead time
maior_lead_time = df['LeadTime'].max()
pedido_maior_lead = df[df['LeadTime'] == maior_lead_time]

print(f"Maior Lead Time: {maior_lead_time}")
print("Pedido com maior lead time:")
print(pedido_maior_lead)

# Cancelados
cancelamentos = df[df['Status'] == 'Cancelado']
total_cancelados = cancelamentos.shape[0]

print(f"Total de pedidos cancelados: {total_cancelados}")

total_pedidos = df.shape[0]
porcentagem_cancelados = (total_cancelados / total_pedidos) * 100

print(f"Cancelamentos representam {porcentagem_cancelados:.2f}% do total de pedidos.")

# Motivos com mais votos
motivo_mais_frequente = cancelamentos['MotivoCancelamento'].value_counts().idxmax()

print(f"Motivo com mais cancelamentos: {motivo_mais_frequente}")

motivos = cancelamentos['MotivoCancelamento'].value_counts()
print("Motivos mais comuns para cancelamento:")
print(motivos)

# Gráfico de vendas por mês
vendas_por_mes.plot(kind='line', marker='o', title='Vendas por Mês', ylabel='Quantidade')
plt.grid()
plt.tight_layout()
plt.show()

# Gráfico de vendas por UF
vendas_por_uf = df.groupby('UF')['Quantidade'].sum().sort_values()
vendas_por_uf.plot(kind='barh', title='Vendas por UF', xlabel='Quantidade de Vendas')
plt.tight_layout()
plt.show()

# SKU top 10
sku_top10 = sku_vendas.sort_values(ascending=False).head(10)
sku_top10.plot(kind='bar', title='Top 10 SKUs Mais Vendidos', ylabel='Quantidade')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Motivo cancelamento
motivos.plot(kind='bar', title='Motivos de Cancelamento', ylabel='Total')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Cálculo do atraso 
df['AtrasoDias'] = (df['DataEntregaReal'] - df['DataEntregaPrometida']).dt.days

atrasos = df[df['AtrasoDias'] > 0]
pontuais = df[df['AtrasoDias'] <= 0]

print(f"Total de pedidos atrasados: {atrasos.shape[0]}")
print(f"Total de pedidos pontuais ou adiantados: {pontuais.shape[0]}")

# Gráfico de distribuição dos atrasos
sns.histplot(df['AtrasoDias'], bins=30, kde=True)
plt.title("Distribuição dos Dias de Atraso")
plt.xlabel("Dias de Atraso")
plt.ylabel("Quantidade de Pedidos")
plt.show()

# Atraso médio por UF
atraso_por_uf = df.groupby('UF')['AtrasoDias'].mean().sort_values()
atraso_por_uf.plot(kind='barh', title='Atraso Médio por UF', xlabel='Dias')
plt.tight_layout()
plt.show()

# Pedidos por dia da semana
vendas_dia_semana = df['DiaSemana'].value_counts().reindex([
    'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo'])

vendas_dia_semana.plot(kind='bar', title='Pedidos por Dia da Semana')
plt.ylabel('Total de Pedidos')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
