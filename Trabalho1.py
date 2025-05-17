import pandas as pd
import random
import matplotlib.pyplot as plt
import numpy as np
import scipy.stats as stats
import matplotlib.lines as mlines


def calcular_multiplas_medias(arquivo_excel, qtde_medias, qtde_amostras):
    try:
        df = pd.read_excel(arquivo_excel, skiprows=2)

        if "Rentabilidade diaria" not in df.columns:
            return "Erro: A coluna 'Rentabilidade diaria' não foi encontrada no arquivo."

        valores = df["Rentabilidade diaria"].dropna().tolist()

        if len(valores) < 60:
            return "Erro: A coluna 'Rentabilidade diaria' não tem valores suficientes para a amostragem."

        medias_aleatorias = []

        # Executar o processo num_vezes vezes
        for _ in range(qtde_medias):
            valores_selecionados = random.sample(valores, qtde_amostras)
            media = sum(valores_selecionados) / len(valores_selecionados)
            medias_aleatorias.append(media)
        
        # Criar DataFrame e salvar no Excel
        df_medias = pd.DataFrame({"Médias Aleatórias": medias_aleatorias})
        df_medias.to_excel("medias_aleatorias3.xlsx", index=False)
        return medias_aleatorias

    except FileNotFoundError:
        return "Erro: O arquivo Excel não foi encontrado."
    except ValueError:
        return "Erro: Problema no formato do arquivo ou da coluna."
    except Exception as e:
        return f"Erro inesperado: {e}"

# Calcular retorno médio e desvio padrão dos últimos três meses (~60 dias úteis)
arquivo = "PSSA3_3.xlsx"
resultado = calcular_multiplas_medias(arquivo, 60, 45)
#print(resultado)

media = np.mean(resultado)
desvio_padrao = np.std(resultado, ddof=1)

# Criar valores para a curva normal
#x = np.linspace(min(resultado), max(resultado), 100)
#y = stats.norm.pdf(x, media, desvio_padrao)

# Calcular intervalo de confiança de 95%
intervalo = stats.norm.interval(0.95, loc=media, scale=desvio_padrao)
#print(intervalo)

# Criar gráfico
plt.figure(figsize=(8, 5))

# Histograma dos retornos
plt.hist(resultado ,bins='auto', alpha=0.6, color="blue", edgecolor="black", label="Histograma")

# Curva da distribuição normal
#plt.plot(x, y, color="red", linewidth=2, label="Curva Normal")

# Linhas do intervalo de confiança
plt.axvline(intervalo[0], color="red", linestyle="dashed", label=f"Limite Inferior (95%): {intervalo[0]:.5f}%")
plt.axvline(intervalo[1], color="red", linestyle="dashed", label=f"Limite Superior (95%): {intervalo[1]:.5f}%")

# Linha da média
plt.axvline(media, color="black", linestyle="solid", label=f"Média: {media:.5f}%")

desvio_padrao_marker = mlines.Line2D([], [], color="yellow", linestyle="None", marker="o", label=f"Desvio Padrão: {desvio_padrao:.5f}%")

"""legendas = [
    "Histograma",
    f"Limite Inferior (95%): {intervalo[0]:.5f}%",
    f"Limite Superior (95%): {intervalo[1]:.5f}%",
    f"Média: {media:.5f}%",
    f"Desvio Padrão: {desvio_padrao:.5f}%"  # Adicionando manualmente
]"""

# Adicionar títulos e legenda
plt.title("Histograma dos Retornos e Intervalo de Confiança")
plt.xlabel("Retorno (%)")
plt.ylabel("Frequência")
plt.legend(handles=[plt.Rectangle((0,0),1,1, color="blue", label="Histograma"),
                    plt.axvline(intervalo[0], color="red", linestyle="dashed", label=f"Limite Inferior (95%): {intervalo[0]:.5f}%"),
                    plt.axvline(intervalo[1], color="red", linestyle="dashed", label=f"Limite Superior (95%): {intervalo[1]:.5f}%"),
                    plt.axvline(media, color="black", linestyle="solid", label=f"Média: {media:.5f}%"),
                    desvio_padrao_marker])


# Mostrar gráfico
plt.show()