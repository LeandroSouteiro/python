import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Inicializa o WebDriver
driver = webdriver.Chrome()

# Acessa a URL de login
driver.get('https://www.xxxxx.com.br')

# Insere o nome de usuário
usuario = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@id='user']"))
)
usuario.send_keys('xxxxx')

# Insere a senha
senha = driver.find_element(By.XPATH, "//input[@id='pass']")
senha.send_keys('xxxxxx')

# Clica no botão de login
botao_entrar = driver.find_element(By.XPATH, "//input[@id='wp-submit']")
botao_entrar.click()

# Aguarda o carregamento da página inicial
WebDriverWait(driver, 10).until(
    EC.url_contains("https://www.xxxxx.com.br")
)

# Carrega a planilha e seleciona a aba correta
planilha_clientes = openpyxl.load_workbook('intranet_.xlsx') # Não alterar o Excel, ajustar as informações para esse excel
pagina_clientes = planilha_clientes['Planilha']

# Itera sobre as linhas da planilha, começando na segunda linha
for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
    try:
        usuario_excel, nome, cpf, email = linha

        # Acessa a página de cadastro para cada novo usuário
        driver.get('https://www.xxxxx.com.br/wp-admin/user-new.php')

        # Preenche o campo de login
        campo_usuario = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@id='user_login']"))
        )
        campo_usuario.clear()
        campo_usuario.send_keys(usuario_excel)

        # Preenche o campo de e-mail
        campo_email = driver.find_element(By.XPATH, "//input[@id='email']")
        campo_email.clear()
        campo_email.send_keys(email)

        # Preenche o campo de nome
        campo_nome = driver.find_element(By.XPATH, "//input[@id='first_name']")
        campo_nome.clear()
        campo_nome.send_keys(nome)

        # Gera a senha automaticamente
        botao_senha = driver.find_element(By.XPATH, "//button[@class='button wp-generate-pw hide-if-no-js']")
        botao_senha.click()

        # Preenche o CPF como senha
        campo_senha = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@id='pass1-text']"))
        )
        campo_senha.clear()
        campo_senha.send_keys(cpf)

        # Marca a opção de senha fraca, se disponível
        try:
            botao_senha_fraca = driver.find_element(By.XPATH, "//input[@name='pw_weak']")
            botao_senha_fraca.click()
        except:
            print(f"Opção de senha fraca não encontrada para {usuario_excel}.")

        # Desmarca a notificação por e-mail (opcional)
        botao_notificacao = driver.find_element(By.XPATH, "//input[@id='send_user_notification']")
        botao_notificacao.click()

        # Clica no botão de cadastrar
        botao_cadastrar = driver.find_element(By.XPATH, "//input[@id='createusersub']")
        botao_cadastrar.click()

        # Aguarda até que o aviso de sucesso seja exibido
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "updated"))
        )

        print(f"Usuário {usuario_excel} cadastrado com sucesso.")

    except Exception as e:
        print(f"Erro ao cadastrar {usuario_excel}: {str(e)}")

# Encerra o navegador após todos os cadastros
print("Todos os cadastros foram concluídos.")
driver.quit()
