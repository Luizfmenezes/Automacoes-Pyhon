import customtkinter
import pandas as pd
import warnings
from functools import partial

# --- CONFIGURAÇÃO INICIAL ---
warnings.simplefilter("ignore") 
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# --- CONSTANTES E CONFIGURAÇÕES DAS PLANILHAS ---
CAMINHO_ICV = r'C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\BASE_ICV_E_ICVFH.xlsm'
CAMINHO_ICF = r'C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\BASE_ICF.xlsm'
CAMINHO_IPP = r'C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\BASE DE DADOS\BASE_IPP.xlsm'
# CAMINHO_IPP = r'C:\Users\luiz.menezes\Documents\BASE_IPP.xlsm'
CAMINHO_SOS = r'C:\Users\luiz.menezes\OneDrive\Planejamento\ANÁLISES\APRESENTAÇÃO\SOS.xlsx'

ABA_ICV = 'ICV'
ABA_ICF = 'ICF'
ABA_IPP = 'IPP'
ABA_SOS = 'S.O.S'

COLUNAS = {
    'icv': {'linha': 'Linha', 'data': 'DATA', 'sentido': 'Sentido', 'prog': 'Prog.', 'real': 'Monit.', 'perdas': 'PERDAS REAL'},
    'icf': {'linha': 'LINHA2', 'data': 'DATA', 'prog_pm': 'PROG PM', 'prog_ep': 'PROG EP', 'prog_pt': 'PROG PT', 'real_pm': 'REAL PM', 'real_ep': 'REAL EP', 'real_pt': 'REAL PT'},
    'sos': {'linha': 'LINHA', 'data': 'DATA'},
    'ipp': {'data': 'Data', 'linha': 'Linha', 'sentido': 'Sentido', 'perc': '% Pontualidade'}
}

LINHAS_D1 = ["1017-10", "1020-10", "1024-10", "1025-10", "1026-10", "8015-10", "8016-10", "848L-10", "9784-10", "8015-21", "N137-11"]
LINHAS_D2 = ["1206-10", "1702-10", "172K-10", "172U-10", "179X-10", "2013-10", "2014-10"]

# --- VARIÁVEIS GLOBAIS PARA CONTROLE DE ESTADO ---
dados_carregados = {}
checkboxes_linhas = {}
data_alvo_global = None

# --- FUNÇÕES DE PROCESSAMENTO DE DADOS ---
def ler_planilha(caminho, aba, cols_map, nome_indicador):
    """
    Lê uma planilha, renomeia colunas, trata tipos de dados e adiciona
    um diagnóstico para datas que não puderam ser convertidas.
    """
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        colunas_necessarias = list(cols_map.values())
        if not all(col in df.columns for col in colunas_necessarias):
            print(f"ERRO: Colunas faltando no arquivo {nome_indicador}. Colunas esperadas: {colunas_necessarias}, Colunas encontradas: {list(df.columns)}")
            return pd.DataFrame()

        # Guarda a coluna de data original para depuração antes de renomear
        coluna_data_original_nome = cols_map.get('data')
        if coluna_data_original_nome and coluna_data_original_nome in df.columns:
            df['data_original_para_debug'] = df[coluna_data_original_nome]

        df = df.rename(columns={v: k for k, v in cols_map.items()})
        
        if 'linha' in df.columns:
            df['linha'] = df['linha'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        # Converte a coluna de data
        if 'data' in df.columns:
            df['data'] = pd.to_datetime(df['data'], errors='coerce')

            # --- BLOCO DE DIAGNÓSTICO DE DATAS ---
            # Verifica se a coluna de debug existe
            if 'data_original_para_debug' in df.columns:
                # Filtra as linhas onde a conversão resultou em NaT (Not a Time), mas o valor original não era vazio
                datas_invalidas = df[df['data'].isna() & df['data_original_para_debug'].notna()]
                if not datas_invalidas.empty:
                    print(f"\n--- AVISO: Datas não reconhecidas em '{nome_indicador}' ---")
                    print("Os seguintes valores na coluna de data não puderam ser convertidos e serão ignorados:")
                    # Usamos .unique() para não repetir o mesmo valor inválido várias vezes
                    for valor_invalido in datas_invalidas['data_original_para_debug'].unique():
                        print(f" -> '{valor_invalido}'")
                    print("---------------------------------------------------\n")
                # Remove a coluna de debug que não é mais necessária
                df = df.drop(columns=['data_original_para_debug'])
        
        return df
    except Exception as e:
        print(f"ERRO ao ler {nome_indicador}: {e}")
        return pd.DataFrame()


# --- FUNÇÕES DE LÓGICA DA APLICAÇÃO ---

def carregar_linhas_disponiveis():
    """Lê a data, carrega os dados de todas as planilhas e popula a lista de checkboxes."""
    global dados_carregados, checkboxes_linhas, data_alvo_global
    
    # Limpa estado anterior
    for widget in scrollable_frame_linhas.winfo_children():
        widget.destroy()
    checkboxes_linhas.clear()
    dados_carregados.clear()
    caixa_resultado.configure(state="normal")
    caixa_resultado.delete("1.0", "end")
    caixa_resultado.configure(state="disabled")
    botao_gerar.configure(state="disabled")
    botao_copiar.configure(state="disabled")
    frame_botoes_selecao.grid_remove() 

    try:
        data_alvo_global = pd.to_datetime(entry_data.get(), format='%d/%m/%Y')
    except ValueError:
        caixa_resultado.configure(state="normal")
        caixa_resultado.insert("end", "ERRO: Formato de data inválido. Use DD/MM/AAAA.")
        caixa_resultado.configure(state="disabled")
        return

    print(f"Carregando dados para: {data_alvo_global.strftime('%d/%m/%Y')}")
    df_icv = ler_planilha(CAMINHO_ICV, ABA_ICV, COLUNAS['icv'], "ICV")
    df_icf = ler_planilha(CAMINHO_ICF, ABA_ICF, COLUNAS['icf'], "ICF")
    df_ipp = ler_planilha(CAMINHO_IPP, ABA_IPP, COLUNAS['ipp'], "IPP")
    df_sos = ler_planilha(CAMINHO_SOS, ABA_SOS, COLUNAS['sos'], "SOS")

    # --- NOVO BLOCO DE DEBUG DETALHADO PARA IPP ---
    print("\n--- Análise de Debug da Planilha IPP (antes do filtro) ---")
    if not df_ipp.empty:
        print("Primeiras 5 linhas lidas do arquivo IPP:")
        print(df_ipp.head().to_string()) # .to_string() para garantir que tudo seja impresso
        
        if 'data' in df_ipp.columns:
            datas_validas_ipp = df_ipp.dropna(subset=['data'])
            if not datas_validas_ipp.empty:
                datas_unicas = datas_validas_ipp['data'].dt.date.unique()
                print("\nDatas únicas encontradas no arquivo IPP (após conversão):")
                for d in sorted(list(datas_unicas)):
                    print(f"- {d.strftime('%d/%m/%Y')}")
            else:
                print("\nAVISO: Nenhuma data válida foi encontrada na coluna 'data' do arquivo IPP.")
        else:
            print("\nERRO DE DEBUG: A coluna 'data' não foi encontrada no DataFrame IPP após a leitura.")
    else:
        print("AVISO: O DataFrame do IPP está completamente vazio após a leitura inicial.")
    print("----------------------------------------------------------\n")
    # --- FIM DO BLOCO DE DEBUG ---

    dados_carregados['icv'] = df_icv[df_icv['data'].dt.date == data_alvo_global.date()].copy()
    dados_carregados['icf'] = df_icf[df_icf['data'].dt.date == data_alvo_global.date()].copy()
    dados_carregados['ipp'] = df_ipp[df_ipp['data'].dt.date == data_alvo_global.date()].copy()
    dados_carregados['sos'] = df_sos[df_sos['data'].dt.date == data_alvo_global.date()].copy()
    
    # Debug: Verificar se algum dado foi carregado para a data
    if dados_carregados['ipp'].empty:
        print(f"AVISO: Nenhum dado de IPP encontrado para {data_alvo_global.strftime('%d/%m/%Y')} após o filtro.")

    lista_series = [df['linha'] for df in dados_carregados.values() if 'linha' in df.columns and not df.empty]
    if not lista_series:
        caixa_resultado.configure(state="normal")
        caixa_resultado.insert("end", f"Nenhum dado de linha encontrado para o dia {data_alvo_global.strftime('%d/%m/%Y')}.")
        caixa_resultado.configure(state="disabled")
        return
        
    todas_as_linhas = pd.concat(lista_series).dropna().unique()
    todas_as_linhas.sort()

    for linha in todas_as_linhas:
        checkbox = customtkinter.CTkCheckBox(master=scrollable_frame_linhas, text=str(linha))
        checkbox.pack(padx=10, pady=2, anchor="w")
        checkboxes_linhas[str(linha)] = checkbox
    
    frame_botoes_selecao.grid() 
    botao_gerar.configure(state="normal")


def selecionar_grupo(grupo):
    """Marca ou desmarca os checkboxes baseado no grupo."""
    for linha, checkbox in checkboxes_linhas.items():
        if grupo == 'Todas':
            checkbox.select()
        elif grupo == 'Nenhuma':
            checkbox.deselect()
        elif grupo == 'D1':
            if linha in LINHAS_D1: checkbox.select()
            else: checkbox.deselect()
        elif grupo == 'D2':
            if linha in LINHAS_D2: checkbox.select()
            else: checkbox.deselect()

def gerar_resumo():
    """Gera o texto do resumo baseado nos checkboxes selecionados."""
    caixa_resultado.configure(state="normal")
    caixa_resultado.delete("1.0", "end")
    
    linhas_selecionadas = [linha for linha, cb in checkboxes_linhas.items() if cb.get() == 1]
    
    if not linhas_selecionadas:
        caixa_resultado.insert("end", "Nenhuma linha foi selecionada. Marque as linhas desejadas e tente novamente.")
        caixa_resultado.configure(state="disabled")
        return

    df_icv, df_icf, df_ipp, df_sos = dados_carregados['icv'], dados_carregados['icf'], dados_carregados['ipp'], dados_carregados['sos']

    resumo_completo = f"📊 *Resumo do Dia: {data_alvo_global.strftime('%d/%m/%Y')}*"
    for linha_atual in sorted(linhas_selecionadas):
        resumo_completo += f"\n\n--- 🚍 *LINHA: {linha_atual}* ---\n"
        
        # Lógica de cálculo (mantida)
        dados_icv_linha = df_icv[df_icv['linha'] == linha_atual]
        if not dados_icv_linha.empty:
            tp_prog = dados_icv_linha[dados_icv_linha['sentido'] == 'TPTS']['prog'].sum(); tp_real = dados_icv_linha[dados_icv_linha['sentido'] == 'TPTS']['real'].sum()
            perc_tp = (tp_real / tp_prog * 100) if tp_prog > 0 else 0
            ts_prog = dados_icv_linha[dados_icv_linha['sentido'] == 'TSTP']['prog'].sum(); ts_real = dados_icv_linha[dados_icv_linha['sentido'] == 'TSTP']['real'].sum()
            perc_ts = (ts_real / ts_prog * 100) if ts_prog > 0 else 0
            total_perdas = pd.to_numeric(dados_icv_linha['perdas'], errors='coerce').sum()
            resumo_completo += f"  - *ICV TP*: Prog {tp_prog:.0f}, Real {tp_real:.0f} (*{perc_tp:.1f}%*)\n"
            resumo_completo += f"  - *ICV TS*: Prog {ts_prog:.0f}, Real {ts_real:.0f} (*{perc_ts:.1f}%*)\n"
            if total_perdas > 0: resumo_completo += f"  - *Perdas ICV*: {total_perdas:.0f}\n"

        dados_icf_linha = df_icf[df_icf['linha'] == linha_atual]
        if not dados_icf_linha.empty:
            prog_pm = pd.to_numeric(dados_icf_linha['prog_pm'], errors='coerce').sum(); real_pm = pd.to_numeric(dados_icf_linha['real_pm'], errors='coerce').sum()
            prog_ep = pd.to_numeric(dados_icf_linha['prog_ep'], errors='coerce').sum(); real_ep = pd.to_numeric(dados_icf_linha['real_ep'], errors='coerce').sum()
            prog_pt = pd.to_numeric(dados_icf_linha['prog_pt'], errors='coerce').sum(); real_pt = pd.to_numeric(dados_icf_linha['real_pt'], errors='coerce').sum()
            resumo_completo += f"  - *ICF Prog*: PM({prog_pm:.0f}), EP({prog_ep:.0f}), PT({prog_pt:.0f})\n"
            resumo_completo += f"  - *ICF Real*: PM({real_pm:.0f}), EP({real_ep:.0f}), PT({real_pt:.0f})\n"

        dados_ipp_linha = df_ipp[df_ipp['linha'] == linha_atual]
        if not dados_ipp_linha.empty:
            valor_tp = pd.to_numeric(dados_ipp_linha[dados_ipp_linha['sentido'] == 'TP-TS']['perc'], errors='coerce').mean()
            valor_ts = pd.to_numeric(dados_ipp_linha[dados_ipp_linha['sentido'] == 'TS-TP']['perc'], errors='coerce').mean()
            partes_ipp = []
            if not pd.isna(valor_tp): partes_ipp.append(f"TP ({valor_tp:.0f}%)")
            if not pd.isna(valor_ts): partes_ipp.append(f"TS ({valor_ts:.0f}%)")
            if partes_ipp: resumo_completo += "  - *IPP*: " + " ".join(partes_ipp) + "\n"

        ocorrencias_sos = len(df_sos[df_sos['linha'] == str(linha_atual)])
        if ocorrencias_sos > 0: resumo_completo += f"  - *S.O.S*: {ocorrencias_sos} ocorrência(s)\n"
    
    caixa_resultado.insert("end", resumo_completo)
    caixa_resultado.configure(state="disabled")
    botao_copiar.configure(state="normal")

def copiar_texto():
    app.clipboard_clear()
    app.clipboard_append(caixa_resultado.get("1.0", "end-1c"))
    botao_copiar.configure(text="Copiado!")
    app.after(1500, lambda: botao_copiar.configure(text="Copiar Resumo"))

# --- MONTAGEM DA APLICAÇÃO (UI) ---
app = customtkinter.CTk()
app.title("Gerador de Resumo Diário v2.3")
app.minsize(700, 650)

app.grid_columnconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=2)
app.grid_rowconfigure(0, weight=1)

frame_controles = customtkinter.CTkFrame(app)
frame_controles.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
frame_controles.grid_columnconfigure(0, weight=1)
frame_controles.grid_rowconfigure(4, weight=1)

# 1. Entrada de Data
label_data = customtkinter.CTkLabel(frame_controles, text="1. Data da Análise (DD/MM/AAAA):", font=("Arial", 14, "bold"))
label_data.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")
entry_data = customtkinter.CTkEntry(frame_controles, placeholder_text="dd/mm/aaaa")
entry_data.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

# 2. Botão para Carregar
botao_carregar = customtkinter.CTkButton(frame_controles, text="Carregar Linhas do Dia", command=carregar_linhas_disponiveis)
botao_carregar.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

# 3. Lista de Linhas Selecionáveis
label_selecao = customtkinter.CTkLabel(frame_controles, text="2. Selecione as Linhas:", font=("Arial", 14, "bold"))
label_selecao.grid(row=3, column=0, padx=10, pady=(10,0), sticky="w")
scrollable_frame_linhas = customtkinter.CTkScrollableFrame(frame_controles, label_text="")
scrollable_frame_linhas.grid(row=4, column=0, padx=10, pady=5, sticky="nsew")

# 4. Botões de ação rápida para seleção
frame_botoes_selecao = customtkinter.CTkFrame(frame_controles, fg_color="transparent")
frame_botoes_selecao.grid(row=5, column=0, padx=10, pady=5, sticky="ew")
frame_botoes_selecao.grid_columnconfigure((0,1,2,3), weight=1)
btn_d1 = customtkinter.CTkButton(frame_botoes_selecao, text="D1", command=partial(selecionar_grupo, 'D1')); btn_d1.grid(row=0, column=0, padx=2, pady=5)
btn_d2 = customtkinter.CTkButton(frame_botoes_selecao, text="D2", command=partial(selecionar_grupo, 'D2')); btn_d2.grid(row=0, column=1, padx=2, pady=5)
btn_todas = customtkinter.CTkButton(frame_botoes_selecao, text="Todas", command=partial(selecionar_grupo, 'Todas')); btn_todas.grid(row=0, column=2, padx=2, pady=5)
btn_limpar = customtkinter.CTkButton(frame_botoes_selecao, text="Limpar", command=partial(selecionar_grupo, 'Nenhuma')); btn_limpar.grid(row=0, column=3, padx=2, pady=5)
frame_botoes_selecao.grid_remove()

# 5. Botão de Gerar Resumo
botao_gerar = customtkinter.CTkButton(frame_controles, text="3. Gerar Resumo", font=("Arial", 14, "bold"), state="disabled", command=gerar_resumo)
botao_gerar.grid(row=6, column=0, padx=10, pady=(10,10), sticky="ew")

# --- Frame da Direita (Resultado) ---
frame_resultado = customtkinter.CTkFrame(app)
frame_resultado.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
frame_resultado.grid_rowconfigure(0, weight=1)
frame_resultado.grid_columnconfigure(0, weight=1)

caixa_resultado = customtkinter.CTkTextbox(frame_resultado, font=("Courier New", 12), state="disabled")
caixa_resultado.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

botao_copiar = customtkinter.CTkButton(frame_resultado, text="Copiar Resumo", state="disabled", command=copiar_texto)
botao_copiar.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

app.mainloop()
