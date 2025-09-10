import customtkinter
import threading
from datetime import datetime, timedelta
import sys
import time
from tkinter import messagebox
import pythoncom # Importação necessária para a correção

# ===================================================================
#           IMPORTAÇÃO DAS FUNÇÕES DE AUTOMAÇÃO
# ===================================================================
# Garanta que os outros 3 arquivos .py estejam na mesma pasta
try:
    # As funções que preparamos nos passos anteriores
    from automacao_icf import executar_processo_icf
    from automacao_icv_icvfh import executar_processo_icv_e_icvfh
    from automacao_ipp import executar_processo_ipp
except ImportError as e:
    messagebox.showerror("Erro de Importação",
                         f"Não foi possível encontrar os scripts de automação.\n\n"
                         f"Verifique se os arquivos 'automacao_icf.py', 'automacao_icv_icvfh.py' e 'automacao_ipp.py' "
                         f"estão na mesma pasta que este orquestrador.\n\nDetalhe do erro: {e}")
    sys.exit()
# ===================================================================


class TextboxRedirector:
    """Uma classe para redirecionar o output do console (stdout) para um widget Textbox."""
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        """Escreve o texto no Textbox e rola para o final."""
        # Garante que a escrita seja feita na thread principal da GUI
        self.textbox.after(0, self._insert_text, text)

    def _insert_text(self, text):
        """Método auxiliar para inserir texto de forma segura na thread da GUI."""
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text)
        self.textbox.see("end") # Auto-scroll
        self.textbox.configure(state="disabled")

    def flush(self):
        """Método necessário para a interface de stdout."""
        pass


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # --- CONFIGURAÇÃO DA JANELA PRINCIPAL ---
        self.title("Orquestrador de Automações")
        self.geometry("800x650")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1) # Linha do log agora é a 3

        customtkinter.set_appearance_mode("Dark")
        customtkinter.set_default_color_theme("blue")

        # --- FRAME DE CONTROLES (INPUTS) ---
        controls_frame = customtkinter.CTkFrame(self)
        controls_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")
        controls_frame.grid_columnconfigure((0, 1, 2), weight=1)

        title_label = customtkinter.CTkLabel(controls_frame, text="Configuração da Execução", font=customtkinter.CTkFont(size=18, weight="bold"))
        title_label.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 15))

        # --- ENTRADAS DE DATA ---
        start_date_label = customtkinter.CTkLabel(controls_frame, text="Data de Início:")
        start_date_label.grid(row=1, column=0, padx=(20, 5), pady=5, sticky="e")
        self.start_date_entry = customtkinter.CTkEntry(controls_frame, placeholder_text="dd/mm/aaaa")
        self.start_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        end_date_label = customtkinter.CTkLabel(controls_frame, text="Data de Fim:")
        end_date_label.grid(row=2, column=0, padx=(20, 5), pady=5, sticky="e")
        self.end_date_entry = customtkinter.CTkEntry(controls_frame, placeholder_text="dd/mm/aaaa")
        self.end_date_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # --- FRAME DE SELEÇÃO DE AUTOMAÇÕES (CHECKBOXES) ---
        selection_frame = customtkinter.CTkFrame(self)
        selection_frame.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        selection_frame.grid_columnconfigure((0, 1, 2), weight=1)

        selection_label = customtkinter.CTkLabel(selection_frame, text="Selecione as automações para executar:", font=customtkinter.CTkFont(size=14, weight="bold"))
        selection_label.grid(row=0, column=0, columnspan=3, pady=(10, 5))

        self.check_var_icf = customtkinter.StringVar(value="on")
        self.checkbox_icf = customtkinter.CTkCheckBox(selection_frame, text="1. Relatório ICF", variable=self.check_var_icf, onvalue="on", offvalue="off")
        self.checkbox_icf.grid(row=1, column=0, padx=10, pady=10)

        self.check_var_icv_fh = customtkinter.StringVar(value="on")
        self.checkbox_icv_fh = customtkinter.CTkCheckBox(selection_frame, text="2. Relatórios ICV + ICVFH", variable=self.check_var_icv_fh, onvalue="on", offvalue="off")
        self.checkbox_icv_fh.grid(row=1, column=1, padx=10, pady=10)

        self.check_var_ipp = customtkinter.StringVar(value="on")
        self.checkbox_ipp = customtkinter.CTkCheckBox(selection_frame, text="3. Relatórios IPP + IPPFH", variable=self.check_var_ipp, onvalue="on", offvalue="off")
        self.checkbox_ipp.grid(row=1, column=2, padx=10, pady=10)

        # --- BOTÃO DE EXECUÇÃO ---
        self.run_button = customtkinter.CTkButton(self, text="Executar Automações Selecionadas", command=self.start_automation_thread, height=40, font=customtkinter.CTkFont(size=16, weight="bold"))
        self.run_button.grid(row=2, column=0, padx=15, pady=(5, 15), sticky="ew")

        # --- FRAME DO LOG (OUTPUT) ---
        log_frame = customtkinter.CTkFrame(self)
        log_frame.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.log_textbox = customtkinter.CTkTextbox(log_frame, wrap="word", state="disabled")
        self.log_textbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- REDIRECIONAMENTO DO STDOUT ---
        sys.stdout = TextboxRedirector(self.log_textbox)
        sys.stderr = TextboxRedirector(self.log_textbox)
        
        print("Bem-vindo! Preencha as datas, selecione as automações e clique em 'Executar'.\n" + "="*80 + "\n")

    def start_automation_thread(self):
        """Inicia a automação em uma thread separada para não travar a GUI."""
        self.run_button.configure(state="disabled", text="Executando...")
        
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        thread = threading.Thread(target=self.run_all_automations)
        thread.daemon = True
        thread.start()

    def run_all_automations(self):
        """Lógica principal que valida as datas e chama os scripts selecionados."""
        try:
            # CORREÇÃO: Inicializa a biblioteca COM para esta thread
            pythoncom.CoInitialize()

            start_date_str = self.start_date_entry.get()
            end_date_str = self.end_date_entry.get()

            if not start_date_str or not end_date_str:
                messagebox.showerror("Erro", "Por favor, preencha as datas de início e fim.")
                return

            date_list = []
            try:
                start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
                end_date = datetime.strptime(end_date_str, "%d/%m/%Y")
                
                if start_date > end_date:
                    messagebox.showerror("Erro de Data", "A data de início não pode ser maior que a data de fim.")
                    return

                current_date = start_date
                while current_date <= end_date:
                    date_list.append(current_date.strftime("%d/%m/%Y"))
                    current_date += timedelta(days=1)
                
                print(f"Datas a serem processadas: {date_list}\n")
            except ValueError:
                messagebox.showerror("Erro de Formato", "Formato de data inválido. Por favor, use dd/mm/aaaa.")
                return
            
            start_time = time.time()
            
            run_icf = self.check_var_icf.get() == "on"
            run_icv_fh = self.check_var_icv_fh.get() == "on"
            run_ipp = self.check_var_ipp.get() == "on"
            
            if not any([run_icf, run_icv_fh, run_ipp]):
                messagebox.showwarning("Atenção", "Nenhuma automação foi selecionada para execução.")
                return

            # ETAPA 1: ICF
            if run_icf:
                print("="*80 + "\n>>> INICIANDO ETAPA 1: Automação ICF...\n")
                executar_processo_icf(date_list)
                print("\n>>> ETAPA 1 (ICF) CONCLUÍDA <<<\n")
            
            # ETAPA 2: ICV e ICVFH
            if run_icv_fh:
                print("\n" + "="*80 + "\n>>> INICIANDO ETAPA 2: Automação ICV e ICVFH...\n")
                executar_processo_icv_e_icvfh(date_list)
                print("\n>>> ETAPA 2 (ICV/ICVFH) CONCLUÍDA <<<\n")

            # ETAPA 3: IPP e IPPFH
            if run_ipp:
                print("\n" + "="*80 + "\n>>> INICIANDO ETAPA 3: Automação IPP e IPPFH...\n")
                executar_processo_ipp(date_list)
                print("\n>>> ETAPA 3 (IPP/IPPFH) CONCLUÍDA <<<\n")

            end_time = time.time()
            total_time = end_time - start_time
            
            print("\n" + "="*80)
            print("🏁 ORQUESTRADOR FINALIZADO! TAREFAS SELECIONADAS FORAM CONCLUÍDAS. 🏁")
            print(f"Tempo total de execução: {time.strftime('%H:%M:%S', time.gmtime(total_time))}")
            print("="*80 + "\n")
            messagebox.showinfo("Sucesso", "As automações selecionadas foram concluídas com sucesso!")

        except Exception as e:
            print(f"\n!!!!!! ERRO CRÍTICO INESPERADO NO ORQUESTRADOR !!!!!!\n")
            print(f"Ocorreu um erro que interrompeu a execução: {e}\n")
            import traceback
            print("Traceback completo:\n", traceback.format_exc())
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado no orquestrador: {e}")
        finally:
            # CORREÇÃO: Libera a biblioteca COM ao final da thread
            pythoncom.CoUninitialize()
            self.after(0, self.enable_run_button)

    def enable_run_button(self):
        """Reabilita o botão de execução."""
        self.run_button.configure(state="normal", text="Executar Automações Selecionadas")


if __name__ == "__main__":
    app = App()
    app.mainloop()
