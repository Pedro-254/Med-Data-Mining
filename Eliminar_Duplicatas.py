import pandas as pd

def remover_linhas_duplicadas(arquivo_entrada, arquivo_saida):
    # Carrega o arquivo do Excel em um DataFrame do pandas
    df = pd.read_excel(arquivo_entrada)

    # Remove linhas duplicadas com base em todas as colunas
    df_sem_duplicatas = df.drop_duplicates()

    # Salva o DataFrame resultante de volta em um novo arquivo Excel
    df_sem_duplicatas.to_excel(arquivo_saida, index=False)

if __name__ == "__main__":
    arquivo_entrada = "./Duplicados.xlsx"
    arquivo_saida = "./Sem_Duplicados.xlsx"

    remover_linhas_duplicadas(arquivo_entrada, arquivo_saida)

    print(f"Linhas duplicadas removidas. Resultado salvo em: {arquivo_saida}")
