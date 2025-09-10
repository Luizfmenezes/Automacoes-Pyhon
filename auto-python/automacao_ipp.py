import os
import time
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import win32com.client
import pyautogui

# --- CONFIGURAÇÕES DA AUTOMAÇÃO ---
# <<< VOLTAMOS A USAR O CAMINHO MANUAL DO DRIVER >>>
# Certifique-se de que o driver nesta pasta é compatível com seu Edge (versão 140)
CAMINHO_DRIVER = r"C:\edge\msedgedriver.exe"
CAMINHO_PLANILHA_PONTUALIDADE = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\BASE_IPP.xlsm"
NOME_MACRO_IPP = "ImportarPontualidade"
NOME_MACRO_FINAL = "RemoverEspacosColunasG_H"
LOTES_A_PROCESSAR = ["Spencer (D1)", "Spencer (D2)"]

# --- CONFIGURAÇÕES PARA O RELATÓRIO IPP ---
DESTINO_DIR_PONTUALIDADE = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\IPPTEMP"
PREFIXO_ARQUIVO_IPP = "PontualidadeLinha"

# --- CONFIGURAÇÕES PARA O RELATÓRIO IPPFH ---
DESTINO_DIR_IPPFH = r"C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\IPPFHTEMP"
PREFIXO_ARQUIVO_IPPFH = "LinhaEmpresaFaixa" 

# --- COORDENADAS GERAIS ---
NAV_MENU_PRINCIPAL_XY = (182, 309)
NAV_CAMPO_BUSCA_XY = (248, 216)
DATA_INICIO_XY = (57, 361)
DATA_FIM_XY = (181, 363)
CAMPO_LOTE_XY = (294, 360)
BOTAO_PESQUISAR_XY = (602, 356)
BOTAO_EXPORTAR_IPP_XY = (632, 356)
BOTAO_EXPORTAR_IPPFH_XY = (667, 356)

# Carrega variáveis de ambiente
load_dotenv()
USUARIO = os.getenv("SPTRANS_USER", "luiz.fonseca")
SENHA = os.getenv("SPTRANS_PASS", "Felipe5191")

def limpar_pastas_de_download(*pastas):
    """Limpa os diretórios de download antes de iniciar."""
    print("INFO: Limpando diretórios de download...")
    for pasta in pastas:
        if os.path.exists(pasta):
            for nome_item in os.listdir(pasta):
                caminho_item = os.path.join(pasta, nome_item)
                try:
                    if os.path.isfile(caminho_item) or os.path.islink(caminho_item):
                        os.unlink(caminho_item)
                    elif os.path.isdir(caminho_item):
                        shutil.rmtree(caminho_item)
                except Exception as e:
                    print(f"❌ ATENÇÃO: Falha ao remover '{caminho_item}'. Erro: {e}")
        else:
            os.makedirs(pasta)
    print("✅ Pastas de download prontas.")

def executar_macro_excel(caminho_excel, nome_macro):
    """Abre uma instância do Excel, executa uma macro e fecha."""
    excel = None
    try:
        print(f"INFO: Executando macro '{nome_macro}'...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(caminho_excel)
        excel.Application.Run(f"'{os.path.basename(caminho_excel)}'!{nome_macro}")
        wb.Close(SaveChanges=True)
        print(f"✅ Macro '{nome_macro}' executada com sucesso.")
    finally:
        if excel:
            excel.Quit()

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

    # Inicia o driver com o serviço e as opções configuradas
    driver = webdriver.Edge(service=servico, options=options)
    driver.maximize_window()
    return driver

def aguardar_e_renomear_arquivo(lote, tipo_relatorio, data_str, files_before, destino_dir, prefixo_arquivo_site):
    """Espera o download e renomeia o arquivo."""
    print(f"INFO: Aguardando download do '{tipo_relatorio}' para o lote '{lote}'...")
    tempo_espera = 180
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_espera:
        time.sleep(2)
        current_files = set(os.listdir(destino_dir))
        new_files = current_files - files_before
        final_files = [f for f in new_files if f.lower().startswith(prefixo_arquivo_site.lower()) and not f.endswith('.crdownload')]
        if final_files:
            arquivo_recente = final_files.pop()
            caminho_antigo = os.path.join(destino_dir, arquivo_recente)
            time.sleep(3)
            nome_lote_safe = lote.replace(" ", "").replace("(", "").replace(")", "")
            data_safe = data_str.replace('/', '-')
            novo_nome = f"{tipo_relatorio}_{nome_lote_safe}_{data_safe}{os.path.splitext(caminho_antigo)[1]}"
            novo_caminho = os.path.join(destino_dir, novo_nome)
            if os.path.exists(novo_caminho): os.remove(novo_caminho)
            shutil.move(caminho_antigo, novo_caminho)
            print(f"✅ Arquivo renomeado para: {novo_nome}\n")
            return
    raise ValueError(f"Download do '{tipo_relatorio}' para o lote '{lote}' não concluído a tempo.")

def fazer_downloads_pontualidade(data_selecionada, lotes):
    """Realiza o processo de login, navegação e downloads para uma data."""
    driver = None
    try:
        driver = iniciar_driver()
        wait = WebDriverWait(driver, 20)
        
        print("INFO: Acessando o portal e realizando login...")
        driver.get("http://v1132.webfarm.sim.sptrans.com.br/secure/frmLogin.aspx")
        try: wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='http://sim.sptrans.com.br/']"))).click()
        except: pass 
        wait.until(EC.presence_of_element_located((By.ID, "txtLogin"))).send_keys(USUARIO)
        driver.find_element(By.ID, "txtSenha").send_keys(SENHA)
        driver.find_element(By.ID, "entrar").click(); time.sleep(1)
        try: driver.find_element(By.ID, "entrar").click()
        except: pass
        time.sleep(3)
        print("✅ Login concluído.")

        print("INFO: Navegando para 'Pontualidade Partidas'...")
        pyautogui.click(NAV_MENU_PRINCIPAL_XY); time.sleep(1)
        pyautogui.click(NAV_CAMPO_BUSCA_XY)
        pyautogui.write("Pontualidade Partidas", interval=0.1); time.sleep(3)
        pyautogui.click(NAV_MENU_PRINCIPAL_XY); time.sleep(5) 
        pyautogui.click(x=509 , y=285); time.sleep(4)
        pyautogui.click(x=183 , y=275); time.sleep(4)
        pyautogui.click(x=509 , y=285); time.sleep(4)
        pyautogui.click(x=509 , y=285); time.sleep(4)
        pyautogui.click(x=183 , y=275); time.sleep(2)
        pyautogui.click(x=509 , y=285); time.sleep(4)
        print("✅ Navegação concluída.")

        # LOOP IPP
        print("\n" + "="*20 + " INICIANDO DOWNLOADS IPP " + "="*20)
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DESTINO_DIR_PONTUALIDADE})
        for i, lote_atual in enumerate(lotes):
            print("-" * 60 + f"\nINFO: Iniciando IPP para o Lote: {lote_atual} ({i+1}/{len(lotes)})")
            pyautogui.click(DATA_INICIO_XY); time.sleep(0.5); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(data_selecionada, interval=0.1); time.sleep(1)
            pyautogui.click(DATA_FIM_XY); time.sleep(0.5); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(data_selecionada, interval=0.1); time.sleep(1)
            pyautogui.click(CAMPO_LOTE_XY); time.sleep(1); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            time.sleep(1); pyautogui.write(lote_atual, interval=0.1); time.sleep(2)
            pyautogui.press("enter"); time.sleep(2)
            pyautogui.click(BOTAO_PESQUISAR_XY); time.sleep(10)
            files_before_ipp = set(os.listdir(DESTINO_DIR_PONTUALIDADE))
            pyautogui.click(BOTAO_EXPORTAR_IPP_XY)
            aguardar_e_renomear_arquivo(lote_atual, "IPP", data_selecionada, files_before_ipp, DESTINO_DIR_PONTUALIDADE, PREFIXO_ARQUIVO_IPP)
                
        # LOOP IPPFH
        print("\n" + "="*20 + " INICIANDO DOWNLOADS IPPFH " + "="*20)
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DESTINO_DIR_IPPFH})
        for i, lote_atual in enumerate(lotes):
            print("-" * 60 + f"\nINFO: Iniciando IPPFH para o Lote: {lote_atual} ({i+1}/{len(lotes)})")
            pyautogui.click(DATA_INICIO_XY); time.sleep(0.5); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(data_selecionada, interval=0.1); time.sleep(1)
            pyautogui.click(DATA_FIM_XY); time.sleep(0.5); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            pyautogui.write(data_selecionada, interval=0.1); time.sleep(1)
            pyautogui.click(CAMPO_LOTE_XY); time.sleep(1); pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
            time.sleep(1); pyautogui.write(lote_atual, interval=0.1); time.sleep(2)
            pyautogui.press("enter"); time.sleep(2)
            pyautogui.click(BOTAO_PESQUISAR_XY); time.sleep(10)
            files_before_ippfh = set(os.listdir(DESTINO_DIR_IPPFH))
            pyautogui.click(BOTAO_EXPORTAR_IPPFH_XY)
            aguardar_e_renomear_arquivo(lote_atual, "IPPFH", data_selecionada, files_before_ippfh, DESTINO_DIR_IPPFH, PREFIXO_ARQUIVO_IPPFH)
        return True
    finally:
        if driver:
            driver.quit()

def executar_processo_ipp(datas_a_processar):
    """
    Função principal que orquestra todo o processo de download e
    execução de macros para a automação IPP e IPPFH, para uma lista de datas.
    """
    if not datas_a_processar:
        print("⚠️ Nenhuma data fornecida para a automação IPP/IPPFH.")
        return

    print(f"🚀 Iniciando automação IPP/IPPFH para {len(datas_a_processar)} data(s): {', '.join(datas_a_processar)}")
    
    limpar_pastas_de_download(DESTINO_DIR_PONTUALIDADE, DESTINO_DIR_IPPFH)
    
    sucesso_geral = True
    for data_alvo in datas_a_processar:
        print("\n" + "#"*70 + f"\n## PROCESSANDO DATA (IPP/FH): {data_alvo}\n" + "#"*70)
        print(f"Lotes a serem processados: {', '.join(LOTES_A_PROCESSAR)}")
        
        sucesso_data_atual = False
        tentativas = 3
        for tentativa in range(1, tentativas + 1):
            print(f"\n🚀 Tentativa {tentativa}/{tentativas} para a data {data_alvo}")
            try:
                if fazer_downloads_pontualidade(data_alvo, LOTES_A_PROCESSAR):
                    sucesso_data_atual = True
                    break
            except Exception as e:
                print(f"❌ ERRO na tentativa {tentativa}: {e}")
                if tentativa < tentativas:
                    time.sleep(30)
        
        if not sucesso_data_atual:
            print(f"🚫 FALHA FINAL para a data {data_alvo}.")
            sucesso_geral = False

    if sucesso_geral:
        print("\n" + "#"*70 + "\n## EXECUTANDO MACROS DE FINALIZAÇÃO (IPP/IPPFH)\n" + "#"*70)
        executar_macro_excel(CAMINHO_PLANILHA_PONTUALIDADE, NOME_MACRO_IPP)
        executar_macro_excel(CAMINHO_PLANILHA_PONTUALIDADE, NOME_MACRO_FINAL)
        print("\n==== PROCESSO IPP/IPPFH FINALIZADO COM SUCESSO ====")
    else:
        print("\n==== PROCESSO IPP/IPPFH FALHOU. Macros não foram executadas. ====")

# Bloco para permitir execução direta do script para testes
if __name__ == "__main__":
    datas_para_teste = ["05/08/2025", "06/08/2025"]
    executar_processo_ipp(datas_para_teste)

