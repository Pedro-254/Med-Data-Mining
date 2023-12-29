import pandas as pd
from unidecode import unidecode

def remover_duplicatas(lista):
    # Usar um conjunto para armazenar versões normalizadas das strings
    conjunto_normalizado = set()

    # Lista para armazenar resultados sem duplicatas
    resultado_sem_duplicatas = []

    for nome in lista:
        # Normalizar a string removendo acentos e convertendo para minúsculas
        nome_normalizado = unidecode(nome).lower()

        # Verificar se a versão normalizada já está no conjunto
        if nome_normalizado not in conjunto_normalizado:
            # Adicionar à lista de resultados e ao conjunto normalizado
            resultado_sem_duplicatas.append(nome)
            conjunto_normalizado.add(nome_normalizado)

    return resultado_sem_duplicatas

Lista_de_Nomes = []
for i in range(1,12):
    arquivo = str(i) + ".xlsx"
    # Leia o arquivo Excel
    df = pd.read_excel('./Consultas/' + arquivo)

    # Selecione a sexta coluna usando a notação de colchetes
    coluna_5 = df[df.columns[5]]

    for i in range(len(coluna_5)):
        if isinstance(coluna_5[i], str):
            if ("prontuario" or "prontuário") in coluna_5[i].lower():
                posicao = coluna_5[i].find('-')
                nome = coluna_5[i][:posicao]
                posicao = nome.find(',')
                nome = nome[:posicao]
                Lista_de_Nomes.append(nome)
    
Lista_de_Nomes = remover_duplicatas(Lista_de_Nomes)
Lista_de_Nomes = sorted(Lista_de_Nomes)
print(len(Lista_de_Nomes))
# Cria um DataFrame do pandas a partir da lista
df = pd.DataFrame({'Coluna': Lista_de_Nomes})

# Especifique o nome do arquivo Excel
nome_do_arquivo = 'lista_de_nomes.xlsx'

# Salva o DataFrame no arquivo Excel
df.to_excel(nome_do_arquivo, index=False)