from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time
navegador = webdriver.Chrome()

#___________PROCESSO DE LOGIN_________________
navegador.get("https://v65.medx.med.br/login_unificado/loginunificado.html")
navegador.implicitly_wait(5)
elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-secondary")
elemento.click()

pag1_elemento_email = navegador.find_element(By.ID, "emailAssinante")
pag1_elemento_senha = navegador.find_element(By.ID, "senhaAssinante")

pag1_elemento_email.send_keys("espacolisbsbadm@gmail.com")
pag1_elemento_senha.send_keys("FKhx765@")

elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-block.btn-primary")
elemento.click()

pag2_elemento_email = navegador.find_element(By.ID, "usuario")
pag2_elemento_senha = navegador.find_element(By.ID, "senhaUsuario")

pag2_elemento_email.send_keys("Dra. Carolina")
pag2_elemento_senha.send_keys("dracarol2021")

#___________PAGINA DE CONTATOS_________________
elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-block.btn-primary")
elemento.click()

time.sleep(1)
wait = WebDriverWait(navegador, 10)
elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".btn.btn-primary.ng-binding")))
elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-primary.ng-binding")
elemento.click()

navegador.get("https://v65.medx.med.br/pages_front_Desk/contatos.html")


# Carregar o arquivo Excel existente
wb = openpyxl.load_workbook('./Resultados/lista_de_nomes.xlsx')

# Selecionar a planilha ativa (por padrão, a primeira planilha)
sheet = wb.active

# Especificar a letra da coluna desejada (por exemplo, 'A' para a coluna A)
letra_coluna = 'A'

# Obter todos os valores da coluna, excluindo a primeira linha
coluna = sheet[letra_coluna][1:]

# Criar uma lista com os valores da coluna
lista_de_nomes = [celula.value for celula in coluna if celula.value is not None]

#__________Criar planilha de Excel_____________
excel = openpyxl.Workbook()
sheet = excel.active
# Adicionar dados à planilha
sheet['A1'] = 'Nome'
sheet['B1'] = 'Celular'
sheet['C1'] = 'Email'
sheet['D1'] = 'Endereço'
sheet['E1'] = 'Observações'

#___________PESQUISAR CONTATO_________________
for i in range(len(lista_de_nomes)):
    nome = lista_de_nomes[i]
    barra_de_pesquisa = navegador.find_element(By.ID, "inputBuscaContatos")
    barra_de_pesquisa.clear()
    barra_de_pesquisa.send_keys(nome)

    elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-primary.float-right")
    elemento.click()
    
    # Tente encontrar o elemento
    try:
        elemento_td = navegador.find_element(By.CSS_SELECTOR, 'td[role="gridcell"]')
        elemento_td.click()
    except:
        # Código para lidar com a situação em que o elemento não é encontrado
        print("Nome: " + nome + "Não encontrado!")
        continue
    

    #____________Extrair dados________________
    lista_dados = navegador.find_element(By.XPATH, "//ul[@style='float: left; list-style: none; padding: 22px;']")
    itens = lista_dados.find_elements(By.XPATH, "li[@class='tituloFichaPaciente ng-scope']")


    # Nome / Celular / Email / Endereço / Observações
    Lista_de_valores = ["","","","",""]
    Lista_de_valores[0] = nome

    time.sleep(1)
    print("Paciente: " + nome)
    for item in itens:
        try:
            titulo = item.find_element(By.XPATH, "span[@class='textosFicha']").text
            valor = item.find_element(By.XPATH, "span[@class='ng-binding']").text

            # IDENTIFICAR QUAL O DADO E SALVAR ELE NO ARQUIVO EXCEL
            if(titulo == "Celular:"):
                print("Celular: " + valor)
                Lista_de_valores[1] = valor
            elif(titulo == "Email:"):
                print("Email: " + valor)
                Lista_de_valores[2] = valor
            elif(titulo == "Endereço:"):
                print("Endereço: " + valor)
                Lista_de_valores[3] = valor
            elif(titulo == "Observações:"):
                print("Observações: " + valor)
                Lista_de_valores[4] = valor
        except:
            print("erro no " + nome)

    # Inserir valores na primeira linha da planilha
    for col_num, valor in enumerate(Lista_de_valores, 1):
        sheet.cell(row=i+2, column=col_num, value=valor)
    
excel.save('./Resultados/exemplo.xlsx')