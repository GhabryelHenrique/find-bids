import tkinter as tk
from tkinter import ttk
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pdfkit
from tkinter import filedialog
from selenium.webdriver.support.select import Select
import os
import subprocess

class GuiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Download de Nota Fiscal")
        self.browser = None  # Inicializa o browser como None

        self.label = tk.Label(
            root, text="Preencha com as credencias do UdiDigital:")
        self.label.pack(padx=20, pady=10)

        input_frame = tk.Frame(root)
        input_frame.pack(padx=20, pady=10)

        # Adicione os campos de entrada para login e senha
        self.login_label = tk.Label(input_frame, text="Login (CNPJ):")
        self.login_label.grid(row=0, column=0, padx=5, pady=5)
        self.login_entry = tk.Entry(input_frame)
        self.login_entry.grid(row=1, column=0, padx=5, pady=5)

        # Vincule a função de validação aos campos de entrada
        self.login_entry.bind("<KeyRelease>", self.validate_cnpj_format)
        self.password_label = tk.Label(input_frame, text="Senha:")
        self.password_label.grid(row=0, column=1, padx=5, pady=5)
        self.password_entry = tk.Entry(input_frame, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        self.label2 = tk.Label(root, text="Adicione os filtros nescessários:")
        self.label2.pack(padx=20, pady=10)

        input_frame2 = tk.Frame(root)
        input_frame2.pack(padx=20, pady=10)

        # Adicione os campos de entrada para CNPJ, mês e ano
        self.cnpj_label = tk.Label(input_frame2, text="CNPJ:")
        self.cnpj_label.grid(row=2, column=0, padx=5, pady=5)
        self.cnpj_entry = tk.Entry(input_frame2)
        self.cnpj_entry.grid(row=3, column=0, padx=5, pady=5)

        self.month_label = tk.Label(input_frame2, text="Mês:")
        self.month_label.grid(row=2, column=1, padx=5, pady=5)
        self.month_combobox = ttk.Combobox(input_frame2, values=[
                                           "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"])
        self.month_combobox.grid(row=3, column=1, padx=5, pady=5)

        self.year_label = tk.Label(input_frame2, text="Ano:")
        self.year_label.grid(row=2, column=2, padx=5, pady=5)
        self.year_combobox = ttk.Combobox(input_frame2, values=list(
            map(str, range(2000, 2031))))  # Converta os anos inteiros para strings
        self.year_combobox.grid(row=3, column=2, padx=5, pady=5)

        self.save_dir = ""  # Variável para armazenar o diretório de salvamento

        # Botão para selecionar um diretório
        self.select_dir_button = tk.Button(
            root, text="Selecionar Diretório", command=self.select_save_directory)
        self.select_dir_button.pack(padx=20, pady=5)

        self.label3 = tk.Label(root, text="")
        self.label3.pack(padx=20, pady=10)

        self.start_button = tk.Button(
            root, text="Iniciar Download", command=self.start_download, state="disabled")
        self.start_button.pack(padx=20, pady=5)

        self.login_entry.bind("<KeyRelease>", self.check_input_fields)
        self.password_entry.bind("<KeyRelease>", self.check_input_fields)

        self.progressbar = ttk.Progressbar(
            root, orient="horizontal", length=300, mode="determinate")  # Mude para "determinate"
        self.progressbar.pack(padx=20, pady=10)

        # Label para mostrar a porcentagem
        self.progressbar_label = tk.Label(root, text="Progresso: 0%")
        self.progressbar_label.pack()

    def validate_cnpj_format(self, event=None):
        login = self.login_entry.get()

        login_digits = ''.join(filter(str.isdigit, login))

        if len(login_digits) == 14:
            # CNPJ válido, mude a cor do texto para preto
            self.login_entry.config(foreground="black")
        else:
            # CNPJ inválido, mude a cor do texto para vermelho
            self.login_entry.config(foreground="red")

        self.check_input_fields()

    def check_input_fields(self, event=None):
        login = self.login_entry.get()
        senha = self.password_entry.get()

        if len(login) == 14 and len(senha) > 0:
            self.start_button.config(state="normal")
        else:
            self.start_button.config(state="disabled")

    def start_download(self):
        self.progressbar.start()
        self.progressbar_label.config(text="Progresso: 0%")
        self.download_thread = Thread(target=self.download)
        self.download_thread.start()

    def update_progress(self, value):
        self.progressbar["value"] = value
        self.progressbar_label.config(text=f"Progresso: {value}%")

    def select_save_directory(self):
        self.save_dir = filedialog.askdirectory()
        if self.save_dir:
            self.start_button.config(state="normal")
            print(f"Diretório selecionado: {self.save_dir}")
        else:
            self.start_button.config(state="disabled")

    def download(self):
        try:
            login = self.login_entry.get()
            senha = self.password_entry.get()
            cnpjFilter = self.cnpj_entry.get()
            month = self.month_combobox.get()
            year = self.year_combobox.get()

            options = Options()
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--chrome-user-data-dir=C:/Users/ghabr/AppData/Local/Google/Chrome/User Data")
            self.browser = webdriver.Chrome(options=options)
            self.browser.get(
                'https://udigital.uberlandia.mg.gov.br/NotaFiscal/')
            time.sleep(1)
            # Resto do seu código de download aqui

            # Localize o iframe primeiro
            iframe = self.browser.find_element(
                By.XPATH, '//*[@id="principal"]')

            # Alterne o contexto para o iframe
            self.browser.switch_to.frame(iframe)

            self.browser.find_element(
                By.XPATH, '/html/body/div[2]/div[1]/div/div[1]/div/div[2]/ul/li[2]/a').click()
            time.sleep(1)

            cnpjLogin = self.browser.find_element(
                By.XPATH, '//*[@id="rLogin"]')
            cnpjLogin.clear()
            cnpjLogin.send_keys(login)  # Use o login inserido pelo usuário

            senhaLogin = self.browser.find_element(
                By.XPATH, '//*[@id="rSenha"]')
            senhaLogin.clear()
            senhaLogin.send_keys(senha)  # Use a senha inserida pelo usuário

            self.browser.find_element(By.XPATH, '//*[@id="btnEntrar"]').click()
            time.sleep(1)

            self.browser.find_element(
                By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td[2]/div[5]/ul/li/a').click()
            time.sleep(1)

            self.browser.find_element(
                By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td[2]/div[9]/ul/li/a').click()
            time.sleep(1)

            if (cnpjFilter):
                elemento = self.browser.find_element(
                    By.XPATH, '//*[@id="rCpfCnpjCN"]')
                elemento.clear()
                elemento.send_keys(cnpjFilter)

            if (month):
                mesSelecionado = self.browser.find_element(
                    By.XPATH, '//*[@id="rMesCompetenciaCN"]')
                select_month = Select(mesSelecionado)
                select_month.select_by_value(month)

            if (year):
                anoSelecionado = self.browser.find_element(
                    By.XPATH, '//*[@id="rAnoCompetenciaCN"]')
                select_year = Select(anoSelecionado)
                select_year.select_by_value(year)

            self.browser.find_element(
                By.XPATH, '//*[@id="btnConsultar"]').click()
            time.sleep(1)
            self.browser.find_element(
                By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td[4]/span/form/table[2]/tbody/tr/td[2]/table[3]/tbody/tr[4]/td[1]/a').click()
            time.sleep(1)

            janelas = self.browser.window_handles

            self.browser.switch_to.window(janelas[1])
            time.sleep(1)

            notaFiscal = self.browser.find_element(
                By.XPATH, '//*[@id="visualizacao-nota"]')

            # Captura o conteúdo do elemento como HTML
            element_html = notaFiscal.get_attribute('outerHTML')

            pdf_filename = "nome_do_arquivo.pdf"  # Nome do arquivo PDF a ser salvo
            # Caminho completo do arquivo PDF
            pdf_path = os.path.join(self.save_dir, pdf_filename)
            config = pdfkit.configuration(
                wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')
            result = subprocess.run(
                ["C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe", "--version"], capture_output=True, text=True)
            print("wkhtmltopdf output:", result.stdout)

            pdfkit.from_string(element_html, pdf_path, configuration=config)

            self.update_progress(100)

            self.browser.close()

            self.browser.switch_to.window(janelas[0])

            time.sleep(1)

            self.browser.close()
            
            self.browser.quit()

            self.progressbar.stop()
            self.label.config(text="Download concluído!")

        except Exception as e:
            self.progressbar.stop()
            self.label.config(text="Ocorreu um erro durante o download.")
            print(e)
        finally:
            if self.browser:
                self.browser.quit()

    def add_close_button(self):
        self.close_button = tk.Button(
            root, text="Fechar", command=self.close_app)
        self.close_button.pack(padx=20, pady=5)

    def close_app(self):
        if self.browser:
            self.browser.quit()
        self.root.destroy()


root = tk.Tk()
app = GuiApp(root)
app.add_close_button()
root.protocol("WM_DELETE_WINDOW", app.close_app)
root.mainloop()
