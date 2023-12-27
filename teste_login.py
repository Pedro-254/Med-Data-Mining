from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

wait = WebDriverWait(navegador, 10)
elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".btn.btn-primary.ng-binding")))
elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-primary.ng-binding")
elemento.click()

navegador.get("https://v65.medx.med.br/pages_front_Desk/contatos.html")

#___________PESQUISAR CONTATO_________________
nome = "Adriana Carneiro Frota" 
barra_de_pesquisa = navegador.find_element(By.ID, "inputBuscaContatos")
barra_de_pesquisa.send_keys(nome)

elemento = navegador.find_element(By.CSS_SELECTOR, ".btn.btn-primary.float-right")
elemento.click()

elemento_td = navegador.find_element(By.XPATH, "//td[text()='" + nome + "']")
elemento_td.click()

#____________Extrair dados________________
lista_dados = navegador.find_element(By.XPATH, "//ul[@style='float: left; list-style: none; padding: 22px;']")

itens = lista_dados.find_elements(By.XPATH, "li[@class='tituloFichaPaciente ng-scope']")
for item in itens:
    titulo = item.find_element(By.XPATH, "span[@class='textosFicha']").text
    valor = item.find_element(By.XPATH, "span[@class='ng-binding']").text

    # IDENTIFICAR QUAL O DADO E SALVAR ELE NO ARQUIVO EXCEL
    print(f"{titulo} {valor}")

# Aguardar uma entrada do usuário antes de fechar o navegador
# input("Pressione Enter para fechar o navegador...")
# navegador.quit()