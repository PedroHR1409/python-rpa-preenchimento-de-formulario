import pyautogui as pg
import webbrowser
import time
import pandas as pd

def lerArquivoExcel(root: str): # Função para leitura do arquivo com os dados
    df = pd.read_excel(f"{root}.xlsx")
    return df

def abrirNavegador(url: str): # Abre o navegador
   webbrowser.open(url)
   time.sleep(3)

def limparFormulario(): # Função para limpar o forms
    # Descendo a página até o final
    pg.press("end")
    time.sleep(1)

    # Selecionando o botão "Limpar Formulário"
    pg.click(x = 1232, y = 819)
    time.sleep(1)

    # Confirmando limpeza do formulário
    pg.press("tab")
    pg.press("enter")

def inserirDados(archive: str): # Função para inserir os dados no forms
    for i, nome in enumerate(archive['nome']):
        email = archive['email'][i]
        telefone = str(archive['telefone'][i])
        utiliza_servicos = int(archive['utiliza_servicos'][i])
        utiliza_cartao_credito = int(archive['utiliza_cartao_credito'][i])

        # Voltando para o começo do formulário
        pg.press("home")
        time.sleep(1)

        # Preenche o nome
        pg.click(x = 672, y = 714)
        pg.write(nome)
        time.sleep(1)

        # Preenche o e-mail
        pg.press("tab")
        pg.write(email)
        time.sleep(1)

        # Preenche o telefone
        pg.press("tab")
        pg.write(telefone)
        time.sleep(1)
        
        # Descendo a tela
        pg.click(x = 518, y = 563)
        pg.press("down", presses = 5)
        time.sleep(1)

        # Preenche utiliza_servicos
        if utiliza_servicos == 1:
            pg.click(x = 628, y = 489)
        else:
            pg.click(x = 628, y = 534)
        
        time.sleep(1)

        # Preenche utiliza_cartao_credito
        if utiliza_cartao_credito == 1:
            pg.click(x = 628, y = 736)
        else:
            pg.click(x = 628, y = 779)

        # Envia formulário
        pg.click(x = 631, y = 895)
        time.sleep(1)

        # Selecionando a barra de URL
        pg.click(x = 1270, y = 78)
        time.sleep(1)

        # Colocando o URL do site novamente
        pg.write(endereco_forms)
        pg.press("enter")
        time.sleep(1)

# Lê o arquivo contendo os dados para preenchimento e atribui em uma variável
base_para_preenchimento = lerArquivoExcel("base-preenchimento-de-leads")

# Define o endereço do forms
endereco_forms = "https://forms.gle/Wjmbwk5fd4MGo2xK8"

# Executa o programa
if __name__ == "__main__":
    try:    
        abrirNavegador(endereco_forms)
        limparFormulario()
        inserirDados(base_para_preenchimento)
    except KeyboardInterrupt:
        print("Você parou a execução do programa, por favor inicie novamente") 