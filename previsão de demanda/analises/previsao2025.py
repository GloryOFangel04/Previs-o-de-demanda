import pandas as pd
from prophet import Prophet
import matplotlib.pyplot as plt
import seaborn as sns


df_2024 = pd.read_excel('Vendas_24.xlsx')
df_2025 = pd.read_excel('Pedidos_Jan_Maio_2025.xlsx')
df_total = pd.concat([df_2024, df_2025], ignore_index=True)

#arrumando e padronizando as datas
df_total['DataPedido'] = pd.to_datetime(df_total['DataPedido'], errors='coerce', dayfirst=True)


df_total['AnoMes'] = df_total['DataPedido'].dt.to_period('M')
df_mensal = df_total.groupby('AnoMes').size().reset_index(name='Pedidos')
df_mensal['ds'] = df_mensal['AnoMes'].dt.to_timestamp()  # Prophet exige coluna 'ds' com datas
df_mensal = df_mensal[['ds', 'Pedidos']].rename(columns={'Pedidos': 'y'})  # 'y' é o valor previsto

# Criando o modelo Prophet
modelo = Prophet()
modelo.fit(df_mensal)

# Criando datas futuras até setembro de 2025 
futuro = modelo.make_future_dataframe(periods=9, freq='M')  # Prever 9 meses além do último mês existente
previsao = modelo.predict(futuro)

# previsões  de jun/jul/ago/setembro de 2025
previsoes_2025 = previsao[previsao['ds'].dt.strftime('%Y-%m').isin(['2025-06', '2025-07', '2025-08', '2025-09'])]
previsoes_2025 = previsoes_2025[['ds', 'yhat']].copy()
previsoes_2025['mes'] = previsoes_2025['ds'].dt.strftime('%b')

# Obtendo valores reais de 2024 (jun/jul/ago/setembro) de 2024
reais_2024 = df_mensal[df_mensal['ds'].dt.strftime('%Y-%m').isin(['2024-06', '2024-07', '2024-08','2024-09'])]
reais_2024 = reais_2024[['ds', 'y']].copy()
reais_2024['mes'] = reais_2024['ds'].dt.strftime('%b')

# união dos dados reais com previstos para comparação
comparativo = pd.merge(reais_2024, previsoes_2025, on='mes', how='inner')
comparativo = comparativo.rename(columns={'y': '2024', 'yhat': '2025'})

# Plotar gráficos de barras para cada mês
sns.set(style="whitegrid")

for i, row in comparativo.iterrows():
    mes = row['mes']
    valores = [row['2024'], row['2025']]
    
    plt.figure(figsize=(4, 5))
    sns.barplot(x=['2024', '2025'], y=valores, palette='pastel')
    plt.title(f'Previsão de pedidos em {mes}')
    plt.ylabel('Quantidade de Pedidos')
    plt.ylim(0, max(valores) * 1.2)
    for idx, valor in enumerate(valores):
        plt.text(idx, valor + 5, f'{int(valor)}', ha='center', va='bottom')
    plt.tight_layout()
    plt.show()
