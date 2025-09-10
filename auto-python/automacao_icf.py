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
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.5
# <<< VOLTAMOS A USAR O CAMINHO MANUAL DO DRIVER >>>
# Certifique-se de que o driver nesta pasta é compatível com seu Edge (versão 140)
CAMINHO_DRIVER = r"C:\edge\msedgedriver.exe"
DESTINO_DIR = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\ICFTEMP"
CAMINHO_PLANILHA_PRINCIPAL = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS"
NOME_ARQUIVO_MACRO = "BASE_ICF.xlsm"
NOME_MACRO_IMPORTACAO = "ImportarArquivosICF"
NOME_MACRO_FINAL = "SubstituirTexto_Na_Aba_ICF"
MAX_TENTATIVAS = 4

# Coordenadas - Verifique se ainda são válidas para sua tela
CLIQUE_MENU_PRINCIPAL_XY = (188, 277)
CLIQUE_BUSCA_SUBMENU_XY = (186, 216)
CLIQUE_ITEM_FROTA_XY = (159, 308)
CAMPO_ASSUNTO_XY = (212, 310)
CAMPO_DATA_XY = (473, 308)
BOTAO_PESQUISAR_XY = (838, 300)
BOTAO_EXPORTAR_XY = (117, 384)

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
        "download.default_directory": DESTINO_DIR,
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
    """Executa a lógica de login no sistema."""
    try:
        print("INFO: Preenchendo credenciais de login...")
        wait.until(EC.presence_of_element_located((By.ID, "txtLogin"))).send_keys(usuario)
        driver.find_element(By.ID, "txtSenha").send_keys(senha)
        driver.find_element(By.ID, "entrar").click()
        time.sleep(2)
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "entrar"))).click()
        except:
            pass
        WebDriverWait(driver, 10).until_not(EC.url_contains("frmLogin.aspx"))
        print("✅ Login bem-sucedido.")
        return True
    except Exception as e:
        print(f"❌ ERRO durante o processo de login: {e}")
        return False

def fazer_download(data_selecionada_str):
    """Realiza o processo completo de login e download para uma data."""
    driver = None
    try:
        driver = iniciar_driver()
        wait = WebDriverWait(driver, 15)
        driver.get("http://v1132.webfarm.sim.sptrans.com.br/secure/frmLogin.aspx")
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='http://sim.sptrans.com.br/']"))).click()
            time.sleep(3)
        except Exception:
            pass
        if not fazer_login(driver, wait, USUARIO, SENHA):
            raise Exception("Falha crítica no login inicial.")
        
        print("INFO: Navegando via PyAutoGUI...")
        pyautogui.click(CLIQUE_MENU_PRINCIPAL_XY); time.sleep(3)
        pyautogui.click(CLIQUE_BUSCA_SUBMENU_XY); pyautogui.write("frota", interval=0.1); time.sleep(3)
        pyautogui.click(CLIQUE_ITEM_FROTA_XY); time.sleep(6)

        pyautogui.click(CAMPO_ASSUNTO_XY); time.sleep(1)
        pyautogui.write("Resumo Linha", interval=0.1); time.sleep(2)
        pyautogui.press('enter'); time.sleep(2)
        
        pyautogui.click(CAMPO_DATA_XY)
        pyautogui.hotkey('ctrl', 'a'); pyautogui.press('delete')
        pyautogui.write(data_selecionada_str, interval=0.1)
        pyautogui.press('tab'); time.sleep(2)
        print(f"INFO: Data preenchida: {data_selecionada_str}.")

        pyautogui.click(BOTAO_PESQUISAR_XY)
        print("INFO: Pesquisa realizada. Aguardando resultados..."); time.sleep(10)

        files_before_download = set(os.listdir(DESTINO_DIR))
        pyautogui.click(BOTAO_EXPORTAR_XY)

        download_timeout = 120
        download_start_time = time.time()
        downloaded_file_path = None
        while time.time() - download_start_time < download_timeout:
            current_files = set(os.listdir(DESTINO_DIR))
            new_files = current_files - files_before_download
            final_files = [f for f in new_files if not f.endswith(('.crdownload', '.tmp'))]

            if final_files:
                downloaded_file = final_files.pop()
                temp_path = os.path.join(DESTINO_DIR, downloaded_file)
                time.sleep(5) # Espera extra para garantir que o arquivo foi escrito
                if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                    downloaded_file_path = temp_path
                    print(f"✅ Download concluído: '{downloaded_file}'")
                    break
            time.sleep(2)

        if not downloaded_file_path:
            raise Exception(f"Download não concluído após {download_timeout} segundos.")
        return downloaded_file_path
    finally:
        if driver:
            driver.quit()

def executar_macro_excel(caminho_excel_com_macro, nome_macro):
    """Abre um arquivo Excel, executa uma macro e o fecha."""
    excel = None
    try:
        if not os.path.exists(caminho_excel_com_macro):
            raise FileNotFoundError(f"Arquivo da macro não encontrado: '{caminho_excel_com_macro}'")
        print(f"INFO: Executando a macro '{nome_macro}'...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(caminho_excel_com_macro)
        excel.Application.Run(f"'{os.path.basename(caminho_excel_com_macro)}'!{nome_macro}")
        time.sleep(5)
        wb.Close(SaveChanges=True)
        print(f"✅ Macro '{nome_macro}' executada.")
    finally:
        if excel:
            excel.Quit()

def executar_processo_icf(datas_a_processar):
    """
    Função principal que orquestra todo o processo de download e
    execução de macros para a automação ICF, para uma lista de datas.
    """
    if not datas_a_processar:
        print("⚠️ Nenhuma data fornecida para a automação ICF.")
        return

    print(f"🚀 Iniciando automação ICF para {len(datas_a_processar)} data(s): {', '.join(datas_a_processar)}")
    
    datas_sucesso = []
    datas_falha = []
    caminho_planilha_com_macro = os.path.join(CAMINHO_PLANILHA_PRINCIPAL, NOME_ARQUIVO_MACRO)

    for data_para_download in datas_a_processar:
        processo_bem_sucedido_para_data = False
        print("\n" + "#"*70 + f"\n## PROCESSANDO DATA (ICF): {data_para_download}\n" + "#"*70)

        for tentativa in range(1, MAX_TENTATIVAS + 1):
            print(f"\n🚀 Tentativa {tentativa}/{MAX_TENTATIVAS} para a data {data_para_download}")
            try:
                os.makedirs(DESTINO_DIR, exist_ok=True)
                caminho_arquivo_baixado = fazer_download(data_para_download)

                if caminho_arquivo_baixado:
                    data_obj = datetime.strptime(data_para_download, "%d/%m/%Y")
                    data_para_nome = data_obj.strftime("%d-%m-%Y")
                    _, extensao = os.path.splitext(caminho_arquivo_baixado)
                    novo_nome_arquivo = f"ICF_{data_para_nome}{extensao}"
                    novo_caminho_arquivo = os.path.join(DESTINO_DIR, novo_nome_arquivo)

                    if os.path.exists(novo_caminho_arquivo):
                        os.remove(novo_caminho_arquivo)
                    os.rename(caminho_arquivo_baixado, novo_caminho_arquivo)
                    print(f"✅ Arquivo renomeado para: {novo_caminho_arquivo}")
                    
                    executar_macro_excel(caminho_planilha_com_macro, NOME_MACRO_IMPORTACAO)
                    
                    processo_bem_sucedido_para_data = True
                    datas_sucesso.append(data_para_download)
                    print(f"\n🎉 SUCESSO PARA A DATA {data_para_download}! 🎉")
                    break
                else:
                    raise Exception("Função de download não retornou um caminho de arquivo.")
            except Exception as e:
                print(f"❌ ERRO na tentativa {tentativa}: {e}")
                if tentativa < MAX_TENTATIVAS:
                    time.sleep(15)
                else:
                    print(f"🚫 FALHA FINAL para a data {data_para_download}.")
                    datas_falha.append(data_para_download)
    
    if datas_sucesso:
        print("\n" + "#"*70 + "\n## EXECUTANDO MACRO DE FINALIZAÇÃO (ICF)\n" + "#"*70)
        try:
            executar_macro_excel(caminho_planilha_com_macro, NOME_MACRO_FINAL)
        except Exception as e:
            print(f"❌ Erro crítico ao executar a macro final: {e}")
    else:
        print("\nINFO: Nenhuma data foi processada com sucesso. Macro final não será executada.")

    print("\n" + "="*70 + "\n RESUMO DO PROCESSAMENTO ICF\n" + "="*70)
    if datas_sucesso:
        print(f"✅ Sucesso ({len(datas_sucesso)}): {', '.join(datas_sucesso)}")
    if datas_falha:
        print(f"❌ Falha ({len(datas_falha)}): {', '.join(datas_falha)}")
    print("="*70)

# Bloco para permitir execução direta do script para testes
if __name__ == "__main__":
    datas_para_teste = ["07/09/2025"]
    executar_processo_icf(datas_para_teste)

