import os
import time
import shutil
import datetime
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, InvalidElementStateException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

# Prevenir execução múltipla
import sys
if hasattr(sys, '_tangara_running'):
    sys.exit(0)
sys._tangara_running = True

# Configurações de diretórios adaptadas para Docker
BASE_DIR = '/app'
DOWNLOAD_DIR = '/app/downloads'
LOG_DIR = '/app/logs'
RELATORIOS_DIR = '/app/relatorios'
ENGENHARIA_DIR = '/app/relatorios/engenharia'
SUPRIMENTOS_DIR = '/app/relatorios/suprimentos'
SUPRIMENTOS_TANGARA_DIR = '/app/relatorios/suprimentos/tangara'
ADMINISTRATIVO_DIR = '/app/relatorios/administrativo'

# Credenciais - Usar variáveis de ambiente
EMAIL = os.getenv('TANGARA_EMAIL')
EMAIL_PASSWORD = os.getenv('TANGARA_EMAIL_PASSWORD')

# Criar diretórios se não existirem
for directory in [LOG_DIR, DOWNLOAD_DIR, ENGENHARIA_DIR, SUPRIMENTOS_DIR, 
                  SUPRIMENTOS_TANGARA_DIR, ADMINISTRATIVO_DIR]:
    os.makedirs(directory, exist_ok=True)

# Configuração do log
nome_do_arquivo_de_log = f"log_tangara_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
caminho_do_arquivo_de_log = os.path.join(LOG_DIR, nome_do_arquivo_de_log)

def adicionar_ao_log(mensagem, caminho_log=caminho_do_arquivo_de_log):
    """Adiciona mensagem ao arquivo de log com timestamp"""
    with open(caminho_log, "a", encoding="utf-8") as log_file:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"{timestamp} - {mensagem}\n")
    print(f"{timestamp} - {mensagem}")  # Também imprimir no console

def mostrar_mensagem_conclusao():
    """Mostra mensagem de conclusão"""
    adicionar_ao_log("Programa concluído com sucesso")

def mostrar_mensagem_erro():
    """Mostra mensagem de erro"""
    adicionar_ao_log("Erro na plataforma")

def criar_driver():
    """Cria o driver do Chrome otimizado para Docker"""
    try:
        chrome_options = Options()
        
        # Opções essenciais para rodar no Docker
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Opções anti-detecção
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # Configurações de download
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safeBrowse.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=chrome_options)
        
        driver.execute_cdp_cmd('Network.enable', {})
        driver.execute_cdp_cmd('Network.setBlockedURLs', {
            "urls": [
                "*beamer*", 
                "*novidades.sienge.com.br*" # Adicionei este da iframe que também vi na tua imagem
            ]
        })
        
        driver.set_page_load_timeout(60)
        
        adicionar_ao_log("Driver Chrome criado com sucesso")
        return driver
        
    except Exception as e:
        adicionar_ao_log(f"Erro ao criar driver: {str(e)}")
        raise

def esperar_download_e_renomear(novo_nome_arquivo, diretorio_destino, wait_time=60):
    """Espera um novo arquivo ser baixado e o renomeia."""
    adicionar_ao_log(f"Aguardando download do arquivo '{novo_nome_arquivo}'...")
    arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
    
    fim_espera = time.time() + wait_time
    arquivo_baixado = None

    while time.time() < fim_espera:
        arquivos_depois = set(os.listdir(DOWNLOAD_DIR))
        novos_arquivos = arquivos_depois - arquivos_antes

        # Filtra arquivos temporários
        arquivos_completos = [f for f in novos_arquivos if not f.endswith(('.tmp', '.crdownload'))]

        if arquivos_completos:
            # Pega o arquivo mais recente
            arquivo_baixado = max([os.path.join(DOWNLOAD_DIR, f) for f in arquivos_completos], key=os.path.getctime)
            # Verifica se o arquivo parou de ser modificado
            tamanho_inicial = os.path.getsize(arquivo_baixado)
            time.sleep(2) # Espera 2s para ver se o tamanho muda
            if tamanho_inicial == os.path.getsize(arquivo_baixado):
                adicionar_ao_log(f"Download concluído: {os.path.basename(arquivo_baixado)}")
                break # Sai do loop
            else:
                arquivo_baixado = None # Continua esperando

        time.sleep(1) # Pausa antes de verificar novamente

    if arquivo_baixado:
        extensao = os.path.splitext(arquivo_baixado)[1]
        caminho_destino_final = os.path.join(diretorio_destino, f"{novo_nome_arquivo}{extensao}")
        
        if os.path.exists(caminho_destino_final):
            os.remove(caminho_destino_final)
            adicionar_ao_log(f"Arquivo existente removido: {caminho_destino_final}")
            
        shutil.move(arquivo_baixado, caminho_destino_final)
        adicionar_ao_log(f"Arquivo '{os.path.basename(caminho_destino_final)}' salvo em '{diretorio_destino}'")
        return True
    else:
        adicionar_ao_log("Nenhum arquivo novo foi encontrado no tempo esperado.")
        return False

def converter_xls_para_xlsx_alternativo(arquivo_entrada):
    """Conversão alternativa de XLS para XLSX usando pandas"""
    try:
        import pandas as pd
        if not os.path.exists(arquivo_entrada):
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_entrada}")
        
        df = pd.read_excel(arquivo_entrada, engine='xlrd')
        arquivo_saida = arquivo_entrada.replace('.xls', '.xlsx')
        df.to_excel(arquivo_saida, index=False, engine='openpyxl')
        
        adicionar_ao_log(f"Arquivo convertido: {arquivo_saida}")
        if os.path.exists(arquivo_saida):
            os.remove(arquivo_entrada)
            
    except Exception as e:
        adicionar_ao_log(f"Aviso: Não foi possível converter XLS para XLSX: {str(e)}")

def fechar_janela(driver, janela_original):
    """Fecha janela popup e retorna à janela original"""
    try:
        WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
        nova_janela = [janela for janela in driver.window_handles if janela != janela_original][0]
        driver.switch_to.window(nova_janela)
        driver.close()
        driver.switch_to.window(janela_original)
        adicionar_ao_log("Janela popup fechada")
    except TimeoutException:
        adicionar_ao_log("Nenhuma janela popup para fechar.")

def marcar_obras(driver, wait, valor):
    """Marca obra específica no formulário"""
    wait.until(EC.element_to_be_clickable((By.XPATH, "//td[img[@title='Abre a consulta']]"))).click()
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "layerFormConsulta")))
    
    try:
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, f"//input[@type='radio' and @name='rowSelect' and @value='{valor}']")))
        elemento.click()
        wait.until(EC.element_to_be_clickable((By.ID, 'pbSelecionar'))).click()
        adicionar_ao_log(f"Obra marcada com sucesso")
    except Exception as e:
        adicionar_ao_log(f"Erro ao tentar marcar o input: {e}")
    finally:
        driver.switch_to.parent_frame()

def configurar_datas_js(driver, id_inicio, id_fim, data_inicio="01/01/2000", data_fim="01/01/2050"):
    """Configura datas usando JavaScript"""
    driver.execute_script(f"""
        document.getElementById('{id_inicio}').value = '{data_inicio}';
        document.getElementById('{id_fim}').value = '{data_fim}';
    """)
    adicionar_ao_log(f"Datas configuradas via JS: {data_inicio} a {data_fim}")

def capturar_screenshot(driver, nome_arquivo=None, pasta_log=None):
    # Definir diretório de logs
    if pasta_log is None:
        pasta_log = os.getenv('LOG_DIR')
    
    # Gerar nome do arquivo com timestamp
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f'screenshot_{timestamp}.png'
    elif not nome_arquivo.endswith('.png'):
        nome_arquivo += '.png'
    
    # Caminho completo
    caminho_completo = os.path.join(pasta_log, nome_arquivo)
    
    try:
        # Capturar screenshot
        driver.save_screenshot(caminho_completo)
        print(f"Screenshot salvo em: {caminho_completo}")
        return caminho_completo
    except Exception as e:
        print(f"Erro ao capturar screenshot: {e}")
        return None

# -----------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------- MAIN ----------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------------------------

try:
    adicionar_ao_log("Iniciando automação TANGARA no Docker")
    
    driver = criar_driver()
    wait = WebDriverWait(driver, 30)
    
    adicionar_ao_log("Acessando página do SIENGE TANGARA...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/")
        
    wait.until(EC.element_to_be_clickable((By.ID, "btnEntrarComSiengeID"))).click()
    adicionar_ao_log("Botão de login clicado")
    
    adicionar_ao_log("Verificando tela de login adicional...")
    email_input_ms = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@name='username']"))
    )
    email_input_ms.send_keys(EMAIL)
    email_input_ms.send_keys(Keys.ENTER)
    adicionar_ao_log("E-mail inserido.")

    password_input_ms = wait.until(
        EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))
    )
    password_input_ms.send_keys(EMAIL_PASSWORD)
    password_input_ms.send_keys(Keys.ENTER)
    adicionar_ao_log("Senha inserida na tela.")
    
    try:
        alerta = driver.find_element(By.XPATH, "//div[contains(@class,'spwAlertaAviso')]")
        if alerta.is_displayed():
            driver.find_element(By.CLASS_NAME, "Button-prim").click()
            adicionar_ao_log("Alerta de aviso fechado.")
    except:
        adicionar_ao_log("Nenhum alerta de aviso encontrado.")

    # Espera o carregamento da página principal pós-login
    wait.until(EC.title_contains("Sienge"))
    adicionar_ao_log("Login realizado com sucesso, página principal carregada.")
    
    janela_original = driver.current_window_handle
    
    # ------------------------------------------------- CADASTRO DE CONTRATOS -----------------------------------------------------------
    adicionar_ao_log("Iniciando extração de Cadastro de Contratos...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/suprimentos/contratos-e-medicoes/contratos/cadastros")
    time.sleep(5)
    
    # Use ActionChains to send ESCAPE key to the active element
    actions = ActionChains(driver)
    actions.send_keys(Keys.ESCAPE)
    actions.perform()
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    # Configurar relatório
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Colunas']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Mostrar todas']"))).click()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[4]/main/div[1]/div[3]/div[2]/div/div[3]/div[2]/div/div[2]/div"))).click()
    optTodos = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Todos']")))
    driver.execute_script("arguments[0].click();", optTodos)
    
    data_inicial = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='dtContratoInicial']")))
    data_inicial.click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiPickersCalendarHeader-switchViewButton')]"))).click()
    ano1 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[text()='2000']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", ano1)
    ano1.click()
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    
    data_final = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='dtContratoFinal']")))
    data_final.click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiPickersCalendarHeader-switchViewButton')]"))).click()
    ano2 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[text()='2050']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", ano2)
    ano2.click()
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    
    btConsultar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Consultar']")))
    driver.execute_script("arguments[0].click();", btConsultar)
    time.sleep(5)
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Gerar Relatório']"))).click()
    
    # Exportar para Excel
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='presentation']//div[@role='combobox']"))).click()
    excel_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-value='excel']")))
    driver.execute_script("arguments[0].click();", excel_option)
    
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Exportar']")))
    driver.execute_script("arguments[0].click();", export_button)
    
    esperar_download_e_renomear("cadastro de contratos", ADMINISTRATIVO_DIR)
    
    # ------------------------------------------------- RELATÓRIOS ENGENHARIA -----------------------------------------------------------
    # Analítico de Apropriações por Obra
    adicionar_ao_log("Acessando Analítico de Apropriações por Obra...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/588")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))
    
    actions = ActionChains(driver)

    configurar_datas_js(driver, "analise.periodoInicio", "analise.periodoFim")
    time.sleep(2)
    
    Select(wait.until(EC.visibility_of_element_located((By.NAME, 'analise.selecao')))).select_by_value("emissao")

    wait.until(EC.element_to_be_clickable((By.XPATH, "//td[img[@title='Abre a consulta']]"))).click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "layerFormConsulta")))
    wait.until(EC.element_to_be_clickable((By.ID, 'pbMarcarTodos'))).click()
    wait.until(EC.element_to_be_clickable((By.ID, 'pbSelecionar'))).click()
    driver.switch_to.parent_frame()

    wait.until(EC.element_to_be_clickable((By.ID, 'analise.imprimirObservacaoTitulo'))).click()
    wait.until(EC.element_to_be_clickable((By.ID, 'analise.imprimirDadosEmColunasNaoMescladas'))).click()
    capturar_screenshot(driver, "analitico_de_apropriacoes.png", LOG_DIR)
    
    wait.until(EC.element_to_be_clickable((By.ID, 'visualizarButton'))).click()
    esperar_download_e_renomear("Analítico de Apropriações por Obra EMISSAO - TOM BUENO - IN531 OBRA", ENGENHARIA_DIR, wait_time=120)

    # Gerar relatório VENCIMENTO
    Select(driver.find_element(By.NAME, 'analise.selecao')).select_by_value("vencimento")
    capturar_screenshot(driver, "analitico_de_apropriacoes_vencimento.png", LOG_DIR)
    wait.until(EC.element_to_be_clickable((By.ID, 'visualizarButton'))).click()
    esperar_download_e_renomear("Analítico de Apropriações por Obra VENCIMENTO - TOM BUENO - IN531 OBRA", ENGENHARIA_DIR, wait_time=120)

    driver.switch_to.default_content()

    # ------------------------------------------------- COMPARATIVO ORÇADO X COMPROMETIDO -----------------------------------------------
    adicionar_ao_log("Acessando Comparativo Orçado x Comprometido...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/627")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))
    
    marcar_obras(driver, wait, "0")
    
    configurar_datas_js(driver, "analise.periodoInicio", "analise.periodoFim")
    
    Select(driver.find_element(By.NAME, 'analise.selecao')).select_by_value("emissao")
    Select(driver.find_element(By.ID, "analise.nivelDetalhamento")).select_by_value("4")
    Select(driver.find_element(By.ID, "analise.bdi")).select_by_value("N")
    Select(driver.find_element(By.ID, "analise.leiSocial")).select_by_value("N")
    
    for checkbox_id in ['analise.consDocPrev',
                        'analise.impPercRealiItensOrc',
                        'analise.impVlEstoqAtualObra',
                        'analise.impVlEstoqServico',
                        'analise.apreDifCompOrcEmVl',
                        'analise.ocultarRegistroSemMovimentacao'
                        ]:
        wait.until(EC.element_to_be_clickable((By.ID, checkbox_id))).click()
    
    wait.until(EC.element_to_be_clickable((By.ID, 'btOpcoesRelatorio'))).click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "layerFormConsulta")))
    Select(driver.find_element(By.ID, 'formatoSaidaDocumento')).select_by_value("XLSX")
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr[3]/td/table/tbody/tr/td[1]/span[1]/span/input'))).click()
    driver.switch_to.parent_frame()
    
    capturar_screenshot(driver, "comparativo_orcado_x_comprometido.png", LOG_DIR)
    wait.until(EC.element_to_be_clickable((By.ID, 'visualizarButton'))).click()
    esperar_download_e_renomear("OrcCom-TOM BUENO - IN531 OBRA", ENGENHARIA_DIR, wait_time=120)
    driver.switch_to.default_content()

    # ------------------------------------------------- COMPARATIVO MEDIDO X COMPROMETIDO -----------------------------------------------
    adicionar_ao_log("Acessando Comparativo Medido x Comprometido...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/623")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))
    
    marcar_obras(driver, wait, "0")

    configurar_datas_js(driver, "analise.periodoInicio", "analise.periodoFim")
    Select(driver.find_element(By.NAME, 'analise.selecao')).select_by_value("emissao")
    Select(driver.find_element(By.ID, "analise.nivelDetalhamento")).select_by_value("0")
    Select(driver.find_element(By.ID, "analise.bdi")).select_by_value("N")
    Select(driver.find_element(By.ID, "analise.leiSocial")).select_by_value("N")
    
    for checkbox_id in ['analise.consDocPrev',
                        'analise.impPercRealiItensOrc',
                        'analise.impVlEstoqAtualObra',
                        'analise.impVlEstoqTarefa',
                        'analise.impCodOrc'
                        ]:
        wait.until(EC.element_to_be_clickable((By.ID, checkbox_id))).click()
    
    capturar_screenshot(driver, "comparativo_medido_x_comprometido.png", LOG_DIR)
    
    wait.until(EC.element_to_be_clickable((By.ID, 'visualizarButton'))).click()
    esperar_download_e_renomear("MedCom-TOM BUENO - IN531 OBRA", ENGENHARIA_DIR, wait_time=120)
    driver.switch_to.default_content()
    
    # ------------------------------------------------- APROPRIAÇÕES DE INSUMOS ---------------------------------------------------------

    adicionar_ao_log("Acessando Apropriações de Insumos...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/2138")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))
    marcar_obras(driver, wait, "1")
    configurar_datas_js(driver, "filter.dataInicialPeriodo", "filter.dataFinalPeriodo")
    Select(driver.find_element(By.ID, 'tpBDI')).select_by_value("N")
    Select(driver.find_element(By.ID, 'tpEncargosSociais')).select_by_value("N")
    
    wait.until(EC.element_to_be_clickable((By.ID, "filter.imprimirSemQuantidades"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "imprimirPedidosPendentes"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, "imprimirContratosPendentes"))).click()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @name='btFiltrar']"))).click()
    esperar_download_e_renomear("ApIns-TOM BUENO - IN531 OBRA", ENGENHARIA_DIR, wait_time=120)
    driver.switch_to.default_content()

    # ------------------------------------------------- RELATÓRIOS SUPRIMENTOS ----------------------------------------------------------
    adicionar_ao_log("Acessando Relatórios de Suprimentos...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/suprimentos/compras/pedidos-de-compra/cadastros")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass

    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[4]/main/div[1]/div[2]/form/div[2]/div[3]/div/div/input"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiPickersCalendarHeader-switchViewButton')]"))).click()
    ano1 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[text()='2000']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", ano1)
    ano1.click()
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[4]/main/div[1]/div[2]/form/div[2]/div[4]/div/div/input"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiPickersCalendarHeader-switchViewButton')]"))).click()
    ano2 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[text()='2050']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", ano2)
    ano2.click()
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Colunas']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Mostrar todas']"))).click()
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    
    filtro = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div/div[4]/main/div[1]/div[2]/div[2]/div/div[3]/div[2]/div/div[2]/div")))
    driver.execute_script("arguments[0].scrollIntoView(true);", filtro)
    filtro.click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Todas']"))).click()
    adicionar_ao_log("Exibição de linhas por página alterada para 'Todas'.")
    time.sleep(15)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Gerar Relatório']"))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='presentation']//div[@role='combobox']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-value='excel']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Exportar']"))).click()
    
    esperar_download_e_renomear("RELAÇÃO DE PEDIDOS DE COMPRAS - TANGARA", SUPRIMENTOS_TANGARA_DIR, wait_time=180)

    # ------------------------------------------------- RELATÓRIO DE RELAÇÃO DE SOLICITAÇÕES ----------------------------------------------
    adicionar_ao_log("Acessando Relatório de Solicitações...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/1328")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Abre a consulta']"))).click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "layerFormConsulta")))
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and @value='0']"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, 'pbSelecionar'))).click()
    driver.switch_to.parent_frame()

    driver.execute_script("document.getElementById('dataInicialSolicitacao').value = '01/01/2000'; document.getElementById('dataFinalSolicitacao').value = '01/01/2050';")
    
    Select(wait.until(EC.visibility_of_element_located((By.NAME, 'filterRelacao.printCotadosReservas')))).select_by_value("S")
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Visualizar']"))).click()
    esperar_download_e_renomear("RELATORIO DE RELACAO DE SOLICITACOES - TANGARA", SUPRIMENTOS_TANGARA_DIR)
    driver.switch_to.default_content()

    # ------------------------------------------------- MAPA DE CONTROLE -----------------------------------------------------------------
    adicionar_ao_log("Acessando Mapa de Controle...")
    driver.get("https://guzattizompero.sienge.com.br/sienge/8/index.html#/common/page/4905")
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div[4]/button'))).click()
        time.sleep(2)
    except:
        pass
    
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entendi')]"))).click()
    except:
        pass
    
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'iFramePage')))

    wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Abre a consulta']"))).click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "layerFormConsulta")))
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and @value='0']"))).click()
    wait.until(EC.element_to_be_clickable((By.ID, 'pbSelecionar'))).click()
    driver.switch_to.parent_frame()
    
    driver.execute_script("document.querySelector(\"input[name*='inicioPeriodo']\").value = '01/01/2000'; document.querySelector(\"input[name*='finalPeriodo']\").value = '01/01/2050';")

    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/form/table/tbody/tr[5]/td[2]/input"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Visualizar']"))).click()
    
    if esperar_download_e_renomear("MAPA DE CONTROLE - TANGARA", SUPRIMENTOS_TANGARA_DIR):
        arquivo_xls = os.path.join(SUPRIMENTOS_TANGARA_DIR, "MAPA DE CONTROLE - TANGARA.xls")
        if os.path.exists(arquivo_xls):
             converter_xls_para_xlsx_alternativo(arquivo_xls)
    
    mostrar_mensagem_conclusao()
    
except Exception as e:
    adicionar_ao_log(f"Erro crítico na automação: {str(e)}")
    # Tenta tirar um screenshot do erro
    if 'driver' in locals():
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        screenshot_path = os.path.join(LOG_DIR, f'error_screenshot_{timestamp}.png')
        driver.save_screenshot(screenshot_path)
        adicionar_ao_log(f"Screenshot do erro salvo em: {screenshot_path}")
    mostrar_mensagem_erro()
    raise
finally:
    if 'driver' in locals():
        driver.quit()
    adicionar_ao_log("Driver fechado. Automação finalizada.")