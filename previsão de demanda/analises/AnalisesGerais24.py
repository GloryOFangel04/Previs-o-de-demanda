import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

#importando a planilha 

df= pd.read_excel('Vendas_24.xlsx')


#mostrar numeros de linha e de colunas

print(f"total de linhas: {df.shape[0]}")
print(f"total de colunas: {df.shape[1]}")

# Ver nomes das colunas
# print("\nColunas disponíveis:")
# print(df.columns.tolist())

#  mês e dia da semana em português
df['DataPedido'] = pd.to_datetime(df['DataPedido'], errors='coerce', dayfirst=True)
df['Mês'] = df['DataPedido'].dt.month_name(locale='pt_BR').str.lower()
df['DiaSemana'] = df['DataPedido'].dt.day_name()

# Definir ordem correta dos meses em português pq se n da erro
ordem_meses_pt = [
    'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
]



# Garantir ordem correta
df['Mês'] = pd.Categorical(df['Mês'], categories=ordem_meses_pt, ordered=True)

#vendas por mes 

vendas_por_mes = df.groupby('Mês')['Quantidade'].sum()
mes_mais_vendido = vendas_por_mes.idxmax()
mes_menos_vendido = vendas_por_mes.idxmin()

print(f"Mês com mais vendas: {mes_mais_vendido}")
print(f"Mês com menos vendas: {mes_menos_vendido}")

#uf 

uf_mais_compra = df.groupby('UF')['Quantidade'].sum().idxmax()
print(f"UF que mais Comprou: {uf_mais_compra}")

#item sku mais ou menos vendidos

sku_vendas = df.groupby('SKU')['Quantidade'].sum()
sku_mais_vendido = sku_vendas.idxmax()
sku_menos_vendido = sku_vendas.idxmin()

print(f"SKU mais vendido: {sku_mais_vendido}")
print(f"SKU menos vendido: {sku_menos_vendido}")

#lead time

maior_lead_time = df['LeadTime'].max()
pedido_maior_lead = df[df['LeadTime'] == maior_lead_time]

print(f"Maior Lead Time: {maior_lead_time}")
print("Pedido com maior lead time:")
print(pedido_maior_lead)

# cancelados
cancelamentos = df[df['Status'] == 'Cancelado']
total_cancelados = cancelamentos.shape[0]

print(f"Total de pedidos cancelados: {total_cancelados}")


total_pedidos = df.shape[0]
porcentagem_cancelados = (total_cancelados / total_pedidos) * 100

print(f"Cancelamentos representam {porcentagem_cancelados:.2f}% do total de pedidos.")



#motivos com mais votos

cancelados = df[df['Status'] == 'Cancelado']
motivo_mais_frequente = cancelados['MotivoCancelamento'].value_counts().idxmax()

print(f"Motivo com mais cancelamentos: {motivo_mais_frequente}")

motivos = cancelamentos['MotivoCancelamento'].value_counts()
print("Motivos mais comuns para cancelamento:")
print(motivos)


#etapas dois 

# Dicionário de nome completo para abreviação
abreviacoes_meses = {
    'janeiro': 'jan', 'fevereiro': 'fev', 'março': 'mar', 'abril': 'abr',
    'maio': 'mai', 'junho': 'jun', 'julho': 'jul', 'agosto': 'ago',
    'setembro': 'set', 'outubro': 'out', 'novembro': 'nov', 'dezembro': 'dez'
}

# Gráfico
ax = vendas_por_mes.plot(kind='line', marker='o', title='Vendas por Mês', ylabel='Quantidade')
ax.set_xticks(range(len(vendas_por_mes)))
ax.set_xticklabels([abreviacoes_meses[mes] for mes in vendas_por_mes.index], rotation=0)

plt.grid()
plt.tight_layout()
plt.show()

# Gráfico de vendas por uf
vendas_por_uf = df.groupby('UF')['Quantidade'].sum().sort_values()

vendas_por_uf.plot(kind='barh', title='Vendas por UF', xlabel='Quantidade de Vendas')
plt.tight_layout()
plt.show()

#sku top 10

sku_top10 = sku_vendas.sort_values(ascending=False).head(10)

sku_top10.plot(kind='bar', title='Top 10 SKUs Mais Vendidos', ylabel='Quantidade')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


# motivo cancelamento

motivos.plot(kind='bar', title='Motivos de Cancelamento', ylabel='Total')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


# mês e dia da semana do pedido
df['Mês'] = df['DataPedido'].dt.month_name()
df['DiaSemana'] = df['DataPedido'].dt.day_name()

# Cálculo do atraso 
df['AtrasoDias'] = (df['DataEntregaReal'] - df['DataEntregaPrometida']).dt.days

atrasos = df[df['AtrasoDias'] > 0]
pontuais = df[df['AtrasoDias'] <= 0]

print(f"Total de pedidos atrasados: {atrasos.shape[0]}")
print(f"Total de pedidos pontuais ou adiantados: {pontuais.shape[0]}")


#grafico distribuição dos atrasos

sns.histplot(df['AtrasoDias'], bins=30, kde=True)
plt.title("Distribuição dos Dias de Atraso")
plt.xlabel("Dias de Atraso")
plt.ylabel("Quantidade de Pedidos")
plt.show()


#Atraso médio por UF  gráfico de barras

atraso_por_uf = df.groupby('UF')['AtrasoDias'].mean().sort_values()

atraso_por_uf.plot(kind='barh', title='Atraso Médio por UF', xlabel='Dias')
plt.tight_layout()
plt.show()

# Pedidos por dia da semana
vendas_dia_semana = df['DiaSemana'].value_counts().reindex([
    'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])

vendas_dia_semana.plot(kind='bar', title='Pedidos por Dia da Semana')
plt.ylabel('Total de Pedidos')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
