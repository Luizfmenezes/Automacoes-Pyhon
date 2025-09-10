import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import win32com.client
import pyautogui
import shutil

# --- CONFIGURAÇÕES GERAIS ---
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5
# <<< VOLTAMOS A USAR O CAMINHO MANUAL DO DRIVER >>>
# Certifique-se de que o driver nesta pasta é compatível com seu Edge (versão 140)
CAMINHO_DRIVER = r"C:\edge\msedgedriver.exe"
PASTA_DOWNLOAD_TEMP = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\TEMP_DOWNLOADS"
CAMINHO_PLANILHA_MACRO = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\BASE_ICV_E_ICVFH.xlsm"
MAX_TENTATIVAS = 4

# --- CONFIGURAÇÕES TAREFA 1: ICVFH (Resumo) ---
DESTINO_DIR_ICVFH = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\ICVFHTEMP"
NOME_MACRO_ICVFH = "ImportarICVFHsFuncional"
BOTAO_EXPORTAR_ICVFH_XY = (416, 302)

# --- CONFIGURAÇÕES TAREFA 2: ICV (Detalhado) ---
DESTINO_DIR_ICV = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\ICVTEMP"
NOME_MACRO_ICV = "ImportarICVsFuncional"
BOTAO_EXPORTAR_ICV_XY = (457, 303)

# --- COORDENADAS DE NAVEGAÇÃO ---
CLIQUE_MENU_CONSULTAS_XY = (511, 142)
CLIQUE_SUBMENU_VIAGEM_MONITORADA_XY = (508, 306)
CLIQUE_ITEM_PROGRAMADA_XY = (816, 366)
CAMPO_DATA_XY = (191, 307)
BOTAO_PESQUISAR_XY = (390, 307)

# Carrega variáveis de ambiente
load_dotenv()
USUARIO = os.getenv("SPTRANS_USER", "luiz.fonseca")
SENHA = os.getenv("SPTRANS_PASS", "Felipe5191")

def iniciar_driver():
    """
    Inicia o driver do Selenium (Edge) usando o caminho manual,
    mas com as correções para limpar os logs do console.
    """
    print(f"INFO: Usando o driver manual em: {CAMINHO_DRIVER}")

    if not os.path.exists(CAMINHO_DRIVER):
        print("="*80)
        print(f"❌ ERRO CRÍTICO: O arquivo msedgedriver.exe não foi encontrado em '{CAMINHO_DRIVER}'.")
        print("Por favor, baixe o driver correto para sua versão do Edge (140) e coloque-o nesta pasta.")
        print("="*80)
        raise FileNotFoundError(f"Driver não encontrado em {CAMINHO_DRIVER}")

    # --- Configura o serviço com o caminho manual ---
    servico = Service(executable_path=CAMINHO_DRIVER)

    # --- Configura as opções do navegador para limpar os logs ---
    options = webdriver.EdgeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging']) # Remove "DevTools listening"
    options.add_argument("--disable-features=msImplicitSignin") # Remove "EDGE_IDENTITY"

    # Mantém a configuração para o diretório de download
    prefs = {
        "download.default_directory": PASTA_DOWNLOAD_TEMP,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": False,
    }
    options.add_experimental_option("prefs", prefs)

    # Inicia o driver com o serviço e as opções configuradas
    driver = webdriver.Edge(service=servico, options=options)
    driver.maximize_window()
    return driver

def fazer_login(driver, wait, usuario, senha):
    """Executa o login de forma confiável com Selenium."""
    try:
        print("INFO: Preenchendo credenciais de login...")
        wait.until(EC.presence_of_element_located((By.ID, "txtLogin"))).send_keys(usuario)
        driver.find_element(By.ID, "txtSenha").send_keys(senha)
        driver.find_element(By.ID, "entrar").click()
        time.sleep(2)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "entrar"))).click()
        except: pass
        WebDriverWait(driver, 10).until_not(EC.url_contains("frmLogin.aspx"))
        print("✅ Login bem-sucedido.")
        return True
    except Exception as e:
        print(f"❌ ERRO durante o login: {e}")
        return False

def esperar_e_validar_download(files_before):
    """Espera um novo arquivo aparecer na pasta temporária e o retorna."""
    print(f"INFO: Aguardando novo download em '{PASTA_DOWNLOAD_TEMP}'...")
    tempo_espera = 120
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_espera:
        time.sleep(2)
        current_files = set(os.listdir(PASTA_DOWNLOAD_TEMP))
        new_files = current_files - files_before
        final_files = [f for f in new_files if not f.endswith(('.crdownload', '.tmp')) and f.startswith("sptrans_ViagensMonitoradas")]
        if final_files:
            arquivo_baixado = final_files.pop()
            caminho_completo = os.path.join(PASTA_DOWNLOAD_TEMP, arquivo_baixado)
            time.sleep(3)
            if os.path.exists(caminho_completo) and os.path.getsize(caminho_completo) > 0:
                print(f"✅ Download detectado: '{arquivo_baixado}'")
                return caminho_completo
    raise Exception(f"Download não concluído após {tempo_espera} segundos.")

def executar_macro_excel(caminho_excel, nome_macro):
    """Executa uma macro em um arquivo Excel."""
    excel = None
    try:
        if not os.path.exists(caminho_excel):
            raise FileNotFoundError(f"Arquivo da macro '{caminho_excel}' não encontrado.")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(caminho_excel)
        excel.Application.Run(f"'{os.path.basename(caminho_excel)}'!{nome_macro}")
        wb.Close(SaveChanges=True)
        print(f"✅ Macro '{nome_macro}' executada com sucesso.")
    finally:
        if excel:
            excel.Quit()

def executar_downloads_para_data(data_selecionada_str):
    """Realiza o processo completo de downloads e macros para uma única data."""
    driver = None
    try:
        os.makedirs(PASTA_DOWNLOAD_TEMP, exist_ok=True)
        os.makedirs(DESTINO_DIR_ICVFH, exist_ok=True)
        os.makedirs(DESTINO_DIR_ICV, exist_ok=True)
        
        driver = iniciar_driver()
        wait = WebDriverWait(driver, 20)
        driver.get("http://v1132.webfarm.sim.sptrans.com.br/secure/frmLogin.aspx")
        if not fazer_login(driver, wait, USUARIO, SENHA):
            raise Exception("Falha crítica no login.")

        print("INFO: Navegando para a tela de consulta...")
        pyautogui.click(CLIQUE_MENU_CONSULTAS_XY); time.sleep(5)
        pyautogui.click(CLIQUE_SUBMENU_VIAGEM_MONITORADA_XY); time.sleep(5)
        pyautogui.click(CLIQUE_ITEM_PROGRAMADA_XY); time.sleep(5)

        print("INFO: Preenchendo formulário...")
        pyautogui.click(CAMPO_DATA_XY); time.sleep(2)
        pyautogui.hotkey('ctrl', 'a'); pyautogui.press('delete'); time.sleep(1)
        pyautogui.write(data_selecionada_str, interval=0.1); time.sleep(2)
        pyautogui.click(BOTAO_PESQUISAR_XY)
        print("INFO: Pesquisa realizada. Aguardando 15s..."); time.sleep(15)

        # --- TAREFA 1: PROCESSAR ICVFH (Resumo) ---
        print("\n--- TAREFA 1: DOWNLOAD DO ICVFH (Resumo) ---")
        files_before = set(os.listdir(PASTA_DOWNLOAD_TEMP))
        pyautogui.click(BOTAO_EXPORTAR_ICVFH_XY); time.sleep(1); pyautogui.click(BOTAO_EXPORTAR_ICVFH_XY)
        caminho_temp_icvfh = esperar_e_validar_download(files_before)
        data_formatada = data_selecionada_str.replace('/', '-')
        nome_final_icvfh = f"ICVFH_{data_formatada}{os.path.splitext(caminho_temp_icvfh)[1]}"
        caminho_final_icvfh = os.path.join(DESTINO_DIR_ICVFH, nome_final_icvfh)
        shutil.move(caminho_temp_icvfh, caminho_final_icvfh)
        print(f"INFO: Arquivo movido para '{caminho_final_icvfh}'")
        executar_macro_excel(CAMINHO_PLANILHA_MACRO, NOME_MACRO_ICVFH)
        print("--- TAREFA 1 (ICVFH) CONCLUÍDA ---\n")

        # --- TAREFA 2: PROCESSAR ICV (Detalhado) ---
        print("--- TAREFA 2: DOWNLOAD DO ICV (Detalhado) ---")
        files_before = set(os.listdir(PASTA_DOWNLOAD_TEMP))
        pyautogui.click(BOTAO_EXPORTAR_ICV_XY); time.sleep(1); pyautogui.click(BOTAO_EXPORTAR_ICV_XY)
        caminho_temp_icv = esperar_e_validar_download(files_before)
        nome_final_icv = f"ICV_{data_formatada}{os.path.splitext(caminho_temp_icv)[1]}"
        caminho_final_icv = os.path.join(DESTINO_DIR_ICV, nome_final_icv)
        shutil.move(caminho_temp_icv, caminho_final_icv)
        print(f"INFO: Arquivo movido para '{caminho_final_icv}'")
        executar_macro_excel(CAMINHO_PLANILHA_MACRO, NOME_MACRO_ICV)
        print("--- TAREFA 2 (ICV) CONCLUÍDA ---")
        return True
    finally:
        if driver:
            driver.quit()

def executar_processo_icv_e_icvfh(datas_a_processar):
    """
    Função principal que orquestra todo o processo de download e
    execução de macros para a automação ICV e ICVFH, para uma lista de datas.
    """
    if not datas_a_processar:
        print("⚠️ Nenhuma data fornecida para a automação ICV/ICVFH.")
        return

    print(f"🚀 Iniciando automação ICV/ICVFH para {len(datas_a_processar)} data(s): {', '.join(datas_a_processar)}")
    
    datas_sucesso = []
    datas_falha = []

    for data_para_rodar in datas_a_processar:
        processo_bem_sucedido_para_data = False
        print("\n" + "#"*70 + f"\n## PROCESSANDO DATA (ICV/FH): {data_para_rodar}\n" + "#"*70)
        
        for tentativa in range(1, MAX_TENTATIVAS + 1):
            print(f"\n🚀 Tentativa {tentativa}/{MAX_TENTATIVAS} para a data {data_para_rodar}")
            try:
                if executar_downloads_para_data(data_para_rodar):
                    print(f"\n🎉 SUCESSO PARA A DATA {data_para_rodar}! 🎉")
                    processo_bem_sucedido_para_data = True
                    break
                else:
                    raise Exception("A função de download retornou 'False'.")
            except Exception as e:
                print(f"❌ ERRO na tentativa {tentativa}: {e}")
                if tentativa < MAX_TENTATIVAS:
                    time.sleep(30)
                else:
                    print(f"🚫 FALHA FINAL para a data {data_para_rodar}.")
        
        if processo_bem_sucedido_para_data:
            datas_sucesso.append(data_para_rodar)
        else:
            datas_falha.append(data_para_rodar)

    print("\n" + "="*70 + "\n RESUMO DO PROCESSAMENTO ICV/ICVFH\n" + "="*70)
    if datas_sucesso:
        print(f"✅ Sucesso ({len(datas_sucesso)}): {', '.join(datas_sucesso)}")
    if datas_falha:
        print(f"❌ Falha ({len(datas_falha)}): {', '.join(datas_falha)}")
    print("="*70)

# Bloco para permitir execução direta do script para testes
if __name__ == "__main__":
    datas_para_teste = ["05/08/2025"]
    executar_processo_icv_e_icvfh(datas_para_teste)

