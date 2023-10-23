import re # Para operações com expressões regulares
import time # Para pausar a execução do código
import pandas as pd # Para trabalhar com dados em formato de tabela
import pyautogui # Para automação de interface gráfica
# pyautogui.alert("O CÓDIGO SERÁ EXECUTADO!!!):") # Exibe um alerta informando o início do código
from selenium import webdriver # Importa o módulo para automatizar o navegador
from selenium.webdriver.common.by import By # Importa mecanismos de busca por elementos
from selenium.webdriver.common.keys import Keys # Importa ações de teclado
from selenium.webdriver.support.ui import WebDriverWait # Importa aguarda do Selenium
from selenium.webdriver.support import expected_conditions as EC # Importa condições esperadas
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk



# Esperar o site carregar 
time.sleep(3)

# Configurar opções do Chrome para não fechar automaticamente
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)

# Inicializar o driver do Selenium para o Chrome com as opções configuradas
driver = webdriver.Chrome(options=chrome_options)

# Abrir o site RPA Challenge
driver.get("https://jornadarpa.com.br/alunos/desafios/sistemas/nf/index.html")
# Maximizar a janela
pyautogui.click(x=981, y=23)

# Função para criar e exibir a janela
def exibir_janela(titulo, mensagem, imagem_path):
    # Crie uma janela
    janela = tk.Tk()
    janela.title(titulo)

    # Defina um tamanho mínimo para a janela (ajuste conforme necessário)
    largura_minima = 400
    altura_minima = 300
    janela.minsize(largura_minima, altura_minima)

    # Obtém a largura e a altura da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcule as coordenadas X e Y para centralizar a janela
    pos_x = (largura_tela - largura_minima) // 2
    pos_y = (altura_tela - altura_minima) // 2

    # Defina as coordenadas da janela
    janela.geometry(f"{largura_minima}x{altura_minima}+{pos_x}+{pos_y}")

    # Configure a janela para ser exibida na frente de outras janelas
    janela.attributes('-topmost', True)

    # Carregue a imagem
    imagem = Image.open(imagem_path)
    imagem = ImageTk.PhotoImage(imagem)

    # Crie um rótulo com a imagem
    label_imagem = tk.Label(janela, image=imagem)
    label_imagem.pack(padx=20, pady=20)

    # Crie um rótulo com a mensagem
    label_mensagem = tk.Label(janela, text=mensagem, font=("Helvetica", 12))
    label_mensagem.pack()

    # Crie um botão "OK" para fechar a janela
    botao_ok = tk.Button(janela, text="OK", command=janela.destroy)
    botao_ok.pack()

    # Exiba a janela
    janela.mainloop()

# Chame a função para exibir a janela no início da execução
exibir_janela("Início da Automação", "Bem-vindo ao Mundo da Automação!", "robo.jpeg")

time.sleep(2)
# Digitar o e-mail
pyautogui.click(x=307, y=403)
pyautogui.write("evento@jornadarpa.com.br")
pyautogui.press("tab")

# Digitar a senha
pyautogui.click(x=275, y=470)
pyautogui.write("evento")

# Pressionar o Botão Entrar
pyautogui.click(x=804, y=532)

# Clicar no Botão Lançar Notas
pyautogui.click(x=259, y=253)

pyautogui.PAUSE = 1


# Ler o arquivo Excel
tabela = pd.read_excel("notasfiscais.xlsx", engine="openpyxl")

print(tabela)

# Itera sobre as linhas da tabela     
for index, linha in tabela.iterrows():
    
    # Preencha os campos com os valores da planilha
    tipo_cliente = linha.iloc[0]  # Valor na coluna 'A'

    if tipo_cliente == 'PJ':
        pyautogui.click(x=163, y=278)  # Seleciona "Pessoa Jurídica"
    elif tipo_cliente == 'PF':
        pyautogui.click(x=682, y=280)  # Seleciona "Pessoa Física"

    # Obtém o valor da segunda coluna (índice 1) da linha atual da planilha e converte-o para uma string.
    cpf_cnpj = str(linha.iloc[1])
    # Define padrões regulares (regex) para verificar se a string 'cpf_cnpj' corresponde a um CPF ou CNPJ válido.
    cpf_pattern = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')  # Padrão de CPF: XXX.XXX.XXX-XX
    cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}') # Padrão de CNPJ: XX.XXX.XXX/XXXX-XX

    if cpf_pattern.match(cpf_cnpj):
        pyautogui.click(x=166, y=331)  # Clique no campo de CPF
        pyautogui.hotkey('ctrl', 'a')  # Seleciona todo o texto no campo
        pyautogui.press("backspace")
        pyautogui.write(cpf_cnpj)
        pyautogui.press("tab")
    elif cnpj_pattern.match(cpf_cnpj):
        pyautogui.click(x=166, y=331)  # Clique no campo de CNPJ
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press("backspace")
        pyautogui.write(cpf_cnpj)
        pyautogui.press("tab")

    nome_razao_social = linha['NOME/RAZÃO SOCIAL']
    pyautogui.write(nome_razao_social)
    pyautogui.press("tab")
   
 
    telefone = linha['TELEFONE']
    pyautogui.write(telefone)
    pyautogui.press("tab")

    email = linha['EMAIL']
    pyautogui.write(email)
    pyautogui.press("tab")

    cep = linha['CEP']
    pyautogui.write(cep)
    pyautogui.press("tab")

    logradouro = linha['LOGRADOURO']
    pyautogui.write(logradouro)
    pyautogui.press("tab")
    time.sleep(3)

    # Preencha o campo de complemento apenas se o valor não for "0" ou vazio
    complemento = linha['COMPLEMENTO']
    pyautogui.write(complemento)
    pyautogui.press("tab")

    bairro = linha['BAIRRO']
    pyautogui.click
    pyautogui.write(bairro)
    pyautogui.press("tab")

    estado = linha['ESTADO']
    pyautogui.click
    pyautogui.write(estado)
    pyautogui.press("tab")

    municipio = linha['MUNICIPIO']
    pyautogui.click
    pyautogui.write(municipio)
    pyautogui.press("tab")
    
    if tipo_cliente == 'PF':
        atividade = linha['ATIVIDADE']
        descricao = linha['DESCRIÇÃO']
        pyautogui.write(atividade)
        pyautogui.press("tab")
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press("backspace")
        pyautogui.write(descricao)
        pyautogui.press("tab")
    else:
        natureza_da_operacao = linha['NATUREZA DA OPERAÇÃO']
        pyautogui.write(natureza_da_operacao)
        pyautogui.press("tab")
        atividade = linha['ATIVIDADE']
        descricao = linha['DESCRIÇÃO']
        pyautogui.write(atividade)
        pyautogui.press("tab")
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press("backspace")
        pyautogui.write(descricao)
        pyautogui.press("tab")    
        
    valor_servico = linha['VALOR']
    pyautogui.click
    pyautogui.write(str(valor_servico))  # Converte o valor numérico para uma string e insere
    
    # Executa scripts para clicar no botão Emitir
    driver.execute_script('document.querySelector("#btnEmitir").click();')
    # Executa scripts para clicar no botão Confirmar
    driver.execute_script('document.querySelector("#avisando > button:nth-child(3)").click();')

    time.sleep(0.5)  # Aguarda 2 segundos para evitar cliques repetidos

    # Executa scripts para clicar no botão Nova Nota Fiscal
    driver.execute_script('document.querySelector("#btnNovoNF").click();')  

#rolar para baixo para visualizar conteúdo adicional
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Chame a função para exibir a mesma janela no final da execução
exibir_janela("Fim da Automação", "Obrigado por usar a automação!", "robof.jpeg")



