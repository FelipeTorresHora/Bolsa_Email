#!/usr/bin/env python
# coding: utf-8

# In[ ]:


Esse projeto tem como finalidade buscar dados da API do yahoo e a partir disso tratei os dados com o 
objetivo anual, mensal e diário tanto do dolar como da ibovespa. Após tratar esses dados separadamente
são enviadas para o email que você desejar só com um clique.


# In[ ]:


# import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as win32


# In[ ]:


#Trazendo o conteúdo do projeto e tratando os dados
codneg = ['^BVSP', 'BRL=X']
hoje = datetime.datetime.now()
ano_passado = hoje - datetime.timedelta(days = 365)
merdados = yf.download(codneg, ano_passado, hoje)
dadosprin = merdados['Adj Close']
dadosprin.columns = ['dolar', 'ibovespa']
dadosprin


# In[4]:


#Pegando os dodos anuais e mensais do projeto
dados_anuais = dadosprin.resample('Y').last()
dados_mensais = dadosprin.resample('M').last()
dados_anuais


# In[5]:


#Visualizando o retorno do investimento em determinado tempo
retorno_anual = dados_anuais.pct_change().dropna()
retorno_mensal = dados_mensais.pct_change().dropna()
retorno_diario = dadosprin.pct_change().dropna()
retorno_diario


# In[6]:


#Buscanndo na lista os valores mais atuais para envio dos emails
retorno_diario_dolar = retorno_diario.iloc [-1,0]
retorno_diario_ibov = retorno_diario.iloc [-1,1]
retorno_mensal_dolar = retorno_mensal.iloc [-1,0]
retorno_mensal_ibov = retorno_mensal.iloc [-1,1]
retorno_anual_dolar = retorno_anual.iloc [-1,0]
retorno_anual_ibov = retorno_anual.iloc [-1,1]
print (retorno_anual_ibov)


# In[7]:


#Arredondando os valores para visualização correta
retorno_diario_dolar = round(retorno_diario_dolar *100,2)
retorno_diario_ibov = round(retorno_diario_ibov *100,2)
retorno_mensal_dolar = round(retorno_mensal_dolar *100,2)
retorno_mensal_ibov = round(retorno_mensal_ibov *100,2)
retorno_anual_dolar = round(retorno_anual_dolar *100,2)
retorno_anual_ibov = round(retorno_anual_ibov *100,2)


# In[ ]:


#Criando os gráficos das perfomances dos ativos


# In[8]:


#Criando os gráficos das perfomances dos ativos
plt.style.use ("cyberpunk")
dados_mensais.plot(y = "ibovespa", use_index = True, legend = False)
plt.title("IBOVESPA")
plt.savefig('ibovespa.png', dpi = 300)
plt.show()


# In[9]:


#Criando os gráficos das perfomances dos ativos
plt.style.use ("cyberpunk")
dados_mensais.plot(y = "dolar", use_index = True, legend = False)
plt.title("DOLAR")
plt.savefig('dolar.png', dpi = 300)
plt.show()


# In[4]:


outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'fhora93@hotmail.com'
email.Subject = 'Cotação do Mercado Financeiro'
email.Body = f'''Prezado Diretor, segue o relatório diário:

Bolsa:

No ano o Ibovespa está tendo uma rentabilidade de {retorno_anual_ibov}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_ibov}%.

No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

Dólar:

No ano o Dólar está tendo uma rentabilidade de {retorno_anual_dolar}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_dolar}%.

No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.


Abs,

O melhor estagiário do mundo, Felipe Hora.

'''
anexo_ibovespa = r'C:\Users\lsiqu\botcamp202302\ibovespa.png'
anexo_dolar = r'C:\Users\lsiqu\botcamp202302\dolar.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)
email.Send()


# In[ ]:





# In[ ]:




