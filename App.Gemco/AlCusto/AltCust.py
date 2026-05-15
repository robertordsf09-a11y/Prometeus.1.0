import pandas as pd
import pyautogui
import customtkinter as ctk
from tkinter import messagebox, filedialog
import time
from datetime import datetime
import os
import traceback

# --- CONFIGURAÇÕES DE CAMINHO ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_IMAGENS = BASE_DIR 

def obter_caminho(nome_arquivo):
    return os.path.join(PASTA_IMAGENS, nome_arquivo)

# --- CONFIGURAÇÕES VISUAIS ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

USUARIOS_AUTORIZADOS = ["ROBERTO", "JOEDSON", "TAISSA"]

# --- FUNÇÕES UTILITÁRIAS ---

def registrar_log(usuario, mensagem, coluna="", valor=""):
    """Registra eventos em uma única linha conforme preferência do usuário."""
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log_entry = f"[{data_hora}] Usuário: {usuario} | Coluna: {coluna} | Valor: {valor} | Evento: {mensagem}\n"
    with open("log_automacao.txt", "a", encoding="utf-8") as f:
        f.write(log_entry)

def formatar_numero_limpo(valor):
    """Formata valores para o padrão decimal brasileiro (vírgula)."""
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        v_str = str(valor).replace(',', '.').strip()
        num = float(v_str)
        return f"{num:.2f}".replace('.', ',')
    except:
        return str(valor).strip()

# --- CLASSE PRINCIPAL DA APLICAÇÃO ---

class AutomacaoApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Automação ERP - Casa Mais Fácil")
        self.master.geometry("440x480")
        self.master.resizable(False, False)

        self.usuario_logado = ""
        self.df_banco = None
        self.caminho_arquivo = ""
        self.coluna_sel = ""

        self.tela_login()

    def limpar_tela(self):
        for widget in self.master.winfo_children():
            widget.destroy()

    def tela_login(self):
        self.limpar_tela()
        frame = ctk.CTkFrame(self.master, corner_radius=14)
        frame.pack(fill="both", expand=True, padx=24, pady=24)

        ctk.CTkLabel(frame, text="⚙  Automação ERP", font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(28, 4))
        ctk.CTkLabel(frame, text="Selecione seu usuário", font=ctk.CTkFont(size=12), text_color="gray").pack(pady=(0, 28))

        self.combo_usuario = ctk.CTkComboBox(frame, values=USUARIOS_AUTORIZADOS, state="readonly", width=320)
        self.combo_usuario.set("Selecione...")
        self.combo_usuario.pack(pady=(6, 24), padx=36)

        ctk.CTkButton(frame, text="Entrar  →", command=self.validar_login, width=320, height=42).pack(padx=36)
        
        # Créditos do Desenvolvedor (Alinhado à esquerda conforme histórico)
        self.label_creditos = ctk.CTkLabel(frame, text="Desenvolvido por Roberto Santos", font=ctk.CTkFont(size=10))
        self.label_creditos.pack(side="bottom", anchor="w", padx=10, pady=5)

    def validar_login(self):
        nome = self.combo_usuario.get()
        if nome == "Selecione..." or nome not in USUARIOS_AUTORIZADOS:
            messagebox.showwarning("Aviso", "Selecione um usuário válido.")
            return
        self.usuario_logado = nome
        registrar_log(nome, "Login realizado")
        self.tela_pre_selecao()

    def tela_pre_selecao(self):
        self.limpar_tela()
        frame = ctk.CTkFrame(self.master, corner_radius=14)
        frame.pack(fill="both", expand=True, padx=24, pady=24)

        ctk.CTkLabel(frame, text=f"👋 Bem-vinda(o), {self.usuario_logado}!", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(28, 6))
        ctk.CTkButton(frame, text="📁 Procurar Arquivo Excel", command=self.abrir_seletor, width=320, height=48, fg_color="#E65100").pack(padx=36, pady=20)

    def abrir_seletor(self):
        caminho = filedialog.askopenfilename(filetypes=(("Arquivos Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")))
        if caminho:
            self.caminho_arquivo = caminho
            try:
                self.df_banco = pd.read_excel(self.caminho_arquivo)
                self.tela_seletor_colunas()
            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Erro ao abrir Excel: {e}")

    def tela_seletor_colunas(self):
        self.limpar_tela()
        frame = ctk.CTkFrame(self.master, corner_radius=14)
        frame.pack(fill="both", expand=True, padx=24, pady=24)

        ctk.CTkLabel(frame, text="✅ Arquivo carregado!", font=ctk.CTkFont(size=16, weight="bold"), text_color="#2E7D32").pack(pady=(24, 4))
        ctk.CTkLabel(frame, text="Escolha a coluna de CUSTO:").pack()
        
        self.combo_colunas = ctk.CTkComboBox(frame, values=list(self.df_banco.columns), state="readonly", width=320)
        self.combo_colunas.set("Selecione a coluna...")
        self.combo_colunas.pack(pady=(6, 26), padx=36)

        ctk.CTkButton(frame, text="▶ Iniciar Automação", command=self.executar_automacao, width=320, height=48, fg_color="#1B5E20").pack(padx=36)

    def esperar_imagem(self, nome_imagem, tempo_limite=60):
        """Aguarda uma imagem aparecer na tela para prosseguir o loop."""
        caminho = obter_caminho(nome_imagem)
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"Imagem não encontrada: {nome_imagem}")
        
        inicio = time.time()
        while (time.time() - inicio) < tempo_limite:
            try:
                posicao = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
                if posicao:
                    return posicao
            except:
                pass
            time.sleep(0.5)
        return None

    def executar_automacao(self):
        try:
            self.coluna_sel = self.combo_colunas.get()
            if not self.coluna_sel or self.coluna_sel == "Selecione a coluna...":
                messagebox.showwarning("Aviso", "Selecione a coluna de custo!")
                return

            df_processar = self.df_banco.copy()
            col_item_nome = self.df_banco.columns[0] # Primeira coluna da tabela

            self.master.iconify()
            registrar_log(self.usuario_logado, f"Início — Coluna: {self.coluna_sel}")
            time.sleep(2)

            for index, row in df_processar.iterrows():
                item_codigo = str(row[col_item_nome])
                custo_valor = formatar_numero_limpo(row[self.coluna_sel])

                # 1. Grupo3
                pos = self.esperar_imagem('Grupo3.png')
                if not pos: break
                pyautogui.click(pos)
                time.sleep(1.5)
                pyautogui.write("3")
                pyautogui.press('enter')
                time.sleep(2)

                # 2. Escrever Item
                pyautogui.write(item_codigo)
                pyautogui.press('enter')
                time.sleep(2)

                # 3. Alteracao
                pos = self.esperar_imagem('Alteracao.png')
                if not pos: break
                pyautogui.click(pos)
                time.sleep(2)

                ''' # 4. cue
                pos = self.esperar_imagem('cue.png')
                if not pos: break
                pyautogui.click(pos)
                pyautogui.press('right', presses=2, interval=0.2)
                pyautogui.write(custo_valor)
                pyautogui.press('enter')
                time.sleep(0.5)

                # 5. cmup
                pos = self.esperar_imagem('cmup.png')
                if not pos: break
                pyautogui.click(pos)
                pyautogui.press('right', presses=2, interval=0.2)
                pyautogui.write(custo_valor)
                pyautogui.press('enter')'''

                # 6. Pausa e Verificação de Continuidade
                registrar_log(self.usuario_logado, "Item OK", self.coluna_sel, item_codigo)
                
                self.master.deiconify()
                continuar = messagebox.askyesno(
                    "Automação Pausada", 
                    f"Item {item_codigo} finalizado.\nDeseja continuar para o próximo?\n\n(Clique em 'Não' para fechar o programa)"
                )
                
                if not continuar:
                    registrar_log(self.usuario_logado, "Encerrado pelo usuário no popup")
                    self.master.destroy()
                    return

                self.master.iconify()

            messagebox.showinfo("Sucesso", "Processamento concluído!")

        except Exception as e:
            erro_detalhado = traceback.format_exc()
            messagebox.showerror("Erro Crítico", f"Falha na execução:\n\n{erro_detalhado}")
        finally:
            if self.master.winfo_exists():
                self.master.deiconify()

if __name__ == "__main__":
    root = ctk.CTk()
    app = AutomacaoApp(root)
    root.mainloop()