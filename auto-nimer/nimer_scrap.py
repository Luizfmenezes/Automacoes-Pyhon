import os
import sys
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import matplotlib.pyplot as plt

# --- BIBLIOTECAS PARA A JANELA (GUI) ---
import tkinter as tk
from tkinter import messagebox, Frame

# --- FUNÇÕES PARA GERENCIAMENTO DE CAMINHOS ---
def resource_path(relative_path):
    """ Obtém o caminho absoluto para o recurso, funciona para dev e para PyInstaller """
    try:
        # PyInstaller cria uma pasta temp e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- NOVA FUNÇÃO ADICIONADA ---
def get_application_path():
    """ Obtém o caminho do diretório onde o script ou o executável está localizado. """
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como um executável (.exe) compilado pelo PyInstaller
        application_path = os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como um script (.py)
        # __file__ é o caminho para o script atual
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path

# --- CONFIGURAÇÕES GERAIS ---
CAMINHO_DRIVER = resource_path("msedgedriver.exe")
URL_LOGIN = "https://sistema.nimer.com.br/Identity/Account/Login?ReturnUrl=/"
URL_DASHBOARD = "https://sistema.nimer.com.br/Dashboard/Lines"

LINHAS_ALVO = [
    "1017-10", "1020-10", "1024-10", "1025-10", "1026-10",
    "8015-10", "8016-10", "848L-10", "9784-10"
]

USUARIO = "spc"
SENHA = "5191"

# --- JANELA PARA SELECIONAR INTERVALO DE DATAS (GUI) ---
def solicitar_intervalo_gui():
    """
    Cria uma janela com campos de texto para o usuário inserir as datas de início e fim.
    """
    intervalo_selecionado = None

    def on_confirm():
        nonlocal intervalo_selecionado
        data_inicio_str = entry_inicio.get()
        data_fim_str = entry_fim.get()

        try:
            data_inicio_obj = datetime.strptime(data_inicio_str, '%d/%m/%Y')
            data_fim_obj = datetime.strptime(data_fim_str, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Formato Inválido", "Por favor, insira as datas no formato DD/MM/AAAA.")
            return

        if data_inicio_obj > data_fim_obj:
            messagebox.showwarning("Data Inválida", "A 'Data Início' não pode ser posterior à 'Data Fim'.")
            return

        intervalo_selecionado = (data_inicio_str, data_fim_str)
        root.destroy()

    root = tk.Tk()
    root.title("Selecionar Intervalo de Datas")
    
    window_width = 400
    window_height = 220
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    root.resizable(False, False)
    
    main_frame = Frame(root, padx=20, pady=20)
    main_frame.pack(fill="both", expand=True)

    label_inicio = tk.Label(main_frame, text="Data Início (DD/MM/AAAA):", font=("Arial", 12))
    label_inicio.pack(pady=(0, 5))
    entry_inicio = tk.Entry(main_frame, font=("Arial", 12), width=15, justify='center')
    entry_inicio.pack()

    label_fim = tk.Label(main_frame, text="Data Fim (DD/MM/AAAA):", font=("Arial", 12))
    label_fim.pack(pady=(10, 5))
    entry_fim = tk.Entry(main_frame, font=("Arial", 12), width=15, justify='center')
    entry_fim.pack()
    
    btn = tk.Button(root, text="Confirmar e Executar", command=on_confirm, height=2, font=("Arial", 10, "bold"))
    btn.pack(pady=10, padx=20, fill="x")
    
    root.mainloop()
    return intervalo_selecionado

# --- FUNÇÕES DA AUTOMAÇÃO ---
def iniciar_driver():
    """Inicia o driver do Edge."""
    print("INFO: Iniciando o WebDriver do Edge...")
    if not os.path.exists(CAMINHO_DRIVER):
        messagebox.showerror("Erro Crítico", f"Driver não encontrado em {CAMINHO_DRIVER}\n\nA automação não pode continuar.")
        raise FileNotFoundError(f"Driver não encontrado em {CAMINHO_DRIVER}")
    service = Service(executable_path=CAMINHO_DRIVER)
    options = webdriver.EdgeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Edge(service=service, options=options)
    driver.maximize_window()
    return driver

def fazer_login(driver, wait, usuario, senha):
    """Executa o login no sistema Nimer."""
    try:
        driver.get(URL_LOGIN)
        wait.until(EC.visibility_of_element_located((By.ID, "Input_UserName"))).send_keys(usuario)
        driver.find_element(By.ID, "Input_Password").send_keys(senha)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/Identity/Account/Logout')]")))
        print("✅ Login bem-sucedido.")
        return True
    except Exception as e:
        print(f"❌ ERRO inesperado durante o login: {e}")
        return False

def filtrar_por_data(driver, wait, data_filtro):
    """Filtra os dados pela data no dashboard."""
    try:
        driver.get(URL_DASHBOARD)
        date_input = wait.until(EC.visibility_of_element_located((By.ID, "Date")))
        date_input.clear()
        date_input.send_keys(data_filtro)
        update_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'main-button') and text()='Atualizar']")))
        update_button.click()
        time.sleep(5)
        print(f"✅ Filtro para {data_filtro} aplicado e dados carregados.")
        return True
    except Exception as e:
        print(f"❌ ERRO inesperado ao filtrar pela data {data_filtro}: {e}")
        return False

def _extrair_valor_progresso(style_attribute):
    """Função auxiliar para extrair o valor '--value' do atributo de estilo."""
    try:
        parts = style_attribute.split(';')
        for part in parts:
            if '--value:' in part:
                return part.split(':')[1].strip()
        return "0"
    except:
        return "0"

def extrair_dados_das_linhas(driver):
    """Extrai as informações das linhas alvo."""
    dados_extraidos = []
    try:
        cards = driver.find_elements(By.CLASS_NAME, "searchable-card")
        if not cards:
            return []
        for card in cards:
            try:
                nome_linha = card.find_element(By.TAG_NAME, "h2").text
                if nome_linha in LINHAS_ALVO:
                    barras_progresso = card.find_elements(By.TAG_NAME, "progress")
                    pct_fotos = "0"
                    pct_pendencias = "0"
                    if len(barras_progresso) >= 2:
                        pct_fotos = _extrair_valor_progresso(barras_progresso[0].get_attribute("style"))
                        pct_pendencias = _extrair_valor_progresso(barras_progresso[1].get_attribute("style"))
                    dados_extraidos.append({"linha": nome_linha, "fotos_pct": pct_fotos, "pendencias_pct": pct_pendencias})
            except NoSuchElementException:
                continue
        return dados_extraidos
    except Exception as e:
        print(f"❌ ERRO GERAL durante a extração dos dados: {e}")
        return None

# --- FUNÇÃO ÚNICA PARA GERAR O GRÁFICO DE TABELA (MODIFICADA) ---
def gerar_grafico_resultados(dados, data_pesquisa):
    """
    Gera e salva um 'gráfico' que simula uma tabela com barras de progresso.
    O arquivo é salvo no mesmo diretório do script/executável.
    """
    if not dados:
        print(f"INFO: Nenhum dado para gerar o gráfico-tabela do dia {data_pesquisa}.")
        return

    print(f"INFO: Gerando gráfico estilo tabela para o dia {data_pesquisa}...")
    df = pd.DataFrame(dados)
    df['fotos_pct'] = pd.to_numeric(df['fotos_pct'], errors='coerce').fillna(0)
    df['pendencias_pct'] = pd.to_numeric(df['pendencias_pct'], errors='coerce').fillna(0)
    df = df.sort_values(by='linha').reset_index(drop=True)

    BG_COLOR = '#FFFFFF'
    HEADER_BG = '#CCCCCC'
    ROW_BG = '#F2F2F2'
    BORDER_COLOR = '#666666'
    TEXT_COLOR = '#333333'
    BAR_COLOR_ACTIVE = "#3EFF03"
    BAR_COLOR_INACTIVE = '#DDDDDD'
    
    fig, ax = plt.subplots(figsize=(10, len(df) * 0.7 + 2))
    fig.set_facecolor(BG_COLOR)
    ax.set_facecolor(BG_COLOR)
    ax.set_axis_off()
    ax.set_xlim(0, 300)
    ax.set_ylim(-1, len(df) + 2)

    ax.text(0, len(df) + 1.2, 'Nimer fotos e pendências por linha', fontsize=16, fontweight='bold', color=TEXT_COLOR, va='center')
    ax.text(0, len(df) + 0.7, f'Data: {data_pesquisa}', fontsize=12, color=TEXT_COLOR, va='center')

    ax.add_patch(plt.Rectangle((0, len(df) - 0.5), 300, 0.5, facecolor=HEADER_BG, edgecolor=BORDER_COLOR, linewidth=1))
    ax.text(50, len(df) - 0.25, 'Linhas', fontsize=12, fontweight='bold', color=TEXT_COLOR, ha='center', va='center')
    ax.text(150, len(df) - 0.25, 'FOTOS %', fontsize=12, fontweight='bold', color=TEXT_COLOR, ha='center', va='center')
    ax.text(250, len(df) - 0.25, 'PEND %', fontsize=12, fontweight='bold', color=TEXT_COLOR, ha='center', va='center')
    
    for i, row in df.iterrows():
        y_pos = len(df) - i - 1
        ax.add_patch(plt.Rectangle((0, y_pos - 0.5), 300, 1, facecolor=ROW_BG if i % 2 == 0 else BG_COLOR, edgecolor=BORDER_COLOR, linewidth=0.5))
        ax.text(50, y_pos, row['linha'], fontsize=11, color=TEXT_COLOR, ha='center', va='center')

        # Barra de Fotos
        ax.add_patch(plt.Rectangle((100, y_pos - 0.3), 100, 0.6, facecolor=BAR_COLOR_INACTIVE, edgecolor='none'))
        ax.add_patch(plt.Rectangle((100, y_pos - 0.3), row['fotos_pct'], 0.6, facecolor=BAR_COLOR_ACTIVE, edgecolor='none'))
        ax.text(150, y_pos, f"{row['fotos_pct']:.0f}%", fontsize=11, color='black', ha='center', va='center', fontweight='bold')

        # Barra de Pendências
        ax.add_patch(plt.Rectangle((200, y_pos - 0.3), 100, 0.6, facecolor=BAR_COLOR_INACTIVE, edgecolor='none'))
        ax.add_patch(plt.Rectangle((200, y_pos - 0.3), row['pendencias_pct'], 0.6, facecolor=BAR_COLOR_ACTIVE, edgecolor='none'))
        ax.text(250, y_pos, f"{row['pendencias_pct']:.0f}%", fontsize=11, color='black', ha='center', va='center', fontweight='bold')

    plt.tight_layout(pad=1.5)
    
    # --- LÓGICA DE SALVAMENTO MODIFICADA ---
    # 1. Cria o nome do arquivo
    nome_simples = f"Relatorio_Linhas_Tabela_{data_pesquisa.replace('/', '-')}.png"
    # 2. Obtém o caminho do diretório da aplicação
    caminho_base = get_application_path()
    # 3. Junta o caminho base com o nome do arquivo para ter o caminho completo
    caminho_completo = os.path.join(caminho_base, nome_simples)
    
    # 4. Salva o arquivo no caminho completo
    plt.savefig(caminho_completo, dpi=300, facecolor=BG_COLOR, bbox_inches='tight')
    
    print(f"✅ Gráfico estilo tabela salvo com sucesso em: '{caminho_completo}'")
    plt.close(fig)

# --- FUNÇÃO PRINCIPAL QUE EXECUTA TUDO ---
def main():
    intervalo_datas = solicitar_intervalo_gui()
    if not intervalo_datas:
        print("INFO: Nenhuma data selecionada. Operação cancelada.")
        return
    
    data_inicio_str, data_fim_str = intervalo_datas
    data_inicio_obj = datetime.strptime(data_inicio_str, "%d/%m/%Y")
    data_fim_obj = datetime.strptime(data_fim_str, "%d/%m/%Y")
    
    driver = None
    try:
        driver = iniciar_driver()
        wait = WebDriverWait(driver, 20)
        
        if not fazer_login(driver, wait, USUARIO, SENHA):
            raise Exception("Falha crítica no login.")
        
        data_atual = data_inicio_obj
        while data_atual <= data_fim_obj:
            data_pesquisa_str = data_atual.strftime("%d/%m/%Y")
            print(f"\n{'='*70}\n## PROCESSANDO DADOS PARA: {data_pesquisa_str}\n{'='*70}")
            
            if filtrar_por_data(driver, wait, data_pesquisa_str):
                dados = extrair_dados_das_linhas(driver)
                if dados:
                    gerar_grafico_resultados(dados, data_pesquisa_str)
                else:
                    print(f"⚠️ Nenhum dado foi extraído para as linhas alvo no dia {data_pesquisa_str}.")
            else:
                print(f"⚠️ Falha ao filtrar os dados para {data_pesquisa_str}. Pulando para o próximo dia.")

            data_atual += timedelta(days=1)
            
    except Exception as e:
        print(f"❌ Ocorreu um erro fatal: {e}")
        messagebox.showerror("Erro na Automação", f"Ocorreu um erro fatal:\n\n{e}")
    finally:
        if driver:
            driver.quit()
        print("\nINFO: Automação finalizada.")
        
if __name__ == "__main__":
    main()