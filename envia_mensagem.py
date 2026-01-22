import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd
import time
import re
import threading
from datetime import datetime
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de WhatsApp - Excel")
        self.root.geometry("750x550")

        self.file_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        # Frame Arquivo
        frame_file = tk.LabelFrame(self.root, text="Seleção de Arquivo", padx=10, pady=10)
        frame_file.pack(padx=10, pady=5, fill="x")
        tk.Label(frame_file, text="Arquivo Excel:").pack(side="left")
        tk.Entry(frame_file, textvariable=self.file_path, width=50).pack(side="left", padx=5)
        tk.Button(frame_file, text="Procurar", command=self.select_file).pack(side="left")

        # Frame Ação
        frame_action = tk.Frame(self.root, padx=10, pady=5)
        frame_action.pack(fill="x")
        self.btn_run = tk.Button(frame_action, text="INICIAR DISPAROS", bg="#25D366", fg="white", font=("Arial", 10, "bold"), command=self.start_thread)
        self.btn_run.pack(fill="x", pady=5)

        # Frame Log
        frame_log = tk.LabelFrame(self.root, text="Log de Execução", padx=10, pady=10)
        frame_log.pack(padx=10, pady=5, fill="both", expand=True)
        self.log_area = scrolledtext.ScrolledText(frame_log, height=15)
        self.log_area.pack(fill="both", expand=True)

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)

    def select_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.file_path.set(filename)
            self.log(f"Arquivo selecionado: {filename}")

    def clean_phone(self, phone):
        phone = str(phone)
        phone = re.sub(r'\D', '', phone)
        if not phone: return None
        # Smart DDI: Se tiver 10 ou 11 dígitos, adiciona 55.
        if 10 <= len(phone) <= 11:
            phone = f"55{phone}"
        return phone

    def start_thread(self):
        if not self.file_path.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo Excel.")
            return
        self.btn_run.config(state="disabled")
        thread = threading.Thread(target=self.run_automation)
        thread.start()

    def run_automation(self):
        start_time = time.time()
        driver = None
        
        try:
            self.log("Lendo arquivo Excel...")
            try:
                df = pd.read_excel(self.file_path.get(), sheet_name="Planilha1", dtype=str)
                self.log("Aba 'Planilha1' carregada.")
            except ValueError:
                self.log("Aba 'Planilha1' não encontrada. Usando a primeira aba...")
                df = pd.read_excel(self.file_path.get(), sheet_name=0, dtype=str)

            col_tel = "TEL_AJUSTADO"
            col_msg = "TEXTO MENSAGEM"

            # Fallback para índices
            if col_tel not in df.columns:
                if len(df.columns) > 8: col_tel = df.columns[8]
                else: raise Exception("Coluna de telefone não encontrada.")
            if col_msg not in df.columns:
                if len(df.columns) > 10: col_msg = df.columns[10]
                else: raise Exception("Coluna de mensagem não encontrada.")

            self.log("Iniciando Chrome...")
            options = webdriver.ChromeOptions()
            # options.add_argument("--headless") 
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.get("https://web.whatsapp.com")

            self.log("Aguardando login...")
            while True:
                try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "pane-side")))
                    self.log("Login OK!")
                    break
                except:
                    time.sleep(2)

            total_enviados = 0
            total_erros = 0

            # O XPath fornecido pelo usuário
            USER_XPATH = '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[4]/div/span/button/div/div/div[1]/span'

            for index, row in df.iterrows():
                try:
                    raw_phone = row[col_tel]
                    message_text = row[col_msg]

                    if pd.isna(raw_phone) or pd.isna(message_text): continue
                    phone = self.clean_phone(raw_phone)
                    if not phone:
                        total_erros += 1
                        continue

                    self.log(f"Processando: {phone}...")
                    msg_encoded = quote(str(message_text))
                    link = f"https://web.whatsapp.com/send?phone={phone}&text={msg_encoded}"
                    driver.get(link)
                    
                    try:
                        wait = WebDriverWait(driver, 25)
                        
                        # 1. Espera a caixa de texto aparecer (sinal que carregou)
                        wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]')))
                        time.sleep(2) # Pausa técnica para renderização

                        # --- TENTATIVA 1: XPath do Usuário ---
                        try:
                            btn_usuario = driver.find_element(By.XPATH, USER_XPATH)
                            btn_usuario.click()
                            self.log(f"-> Clicado no botão (XPath Usuário)")
                            time.sleep(3)
                            total_enviados += 1
                            continue # Se deu certo, vai para o próximo
                        except Exception as e:
                            self.log("-> XPath do usuário falhou, tentando método reserva...")

                        # --- TENTATIVA 2: Método Genérico (Backup) ---
                        try:
                            # Tenta achar o botão pelo ícone padrão
                            send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                            send_btn.click()
                            self.log(f"-> Clicado no botão (Backup)")
                            time.sleep(3)
                            total_enviados += 1
                        except:
                            # Se tudo falhar, tenta o Enter
                            driver.switch_to.active_element.send_keys(Keys.ENTER)
                            self.log(f"-> Enviado via ENTER")
                            time.sleep(3)
                            total_enviados += 1

                    except Exception as e:
                        # Verificação de erro de número inválido
                        try:
                            body = driver.find_element(By.TAG_NAME, "body").text
                            if "número de telefone não está no WhatsApp" in body or "url is invalid" in body:
                                self.log(f"-> ERRO: Número {phone} não tem WhatsApp.")
                            else:
                                self.log(f"-> Erro no envio: {phone} (Timeout ou Botão não achado)")
                        except:
                             self.log(f"-> Erro crítico no envio: {phone}")
                        
                        total_erros += 1

                except Exception as e:
                    self.log(f"Erro linha {index+2}: {str(e)}")
                    total_erros += 1

            tempo = time.time() - start_time
            resumo = f"Fim!\nTempo: {tempo:.2f}s | Enviados: {total_enviados} | Erros: {total_erros}"
            self.log("-" * 30)
            self.log(resumo)
            messagebox.showinfo("Concluído", resumo)

        except Exception as e:
            self.log(f"ERRO CRÍTICO: {str(e)}")
            messagebox.showerror("Erro", str(e))
        finally:
            self.btn_run.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()