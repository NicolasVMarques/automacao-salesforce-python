import pyautogui
import time
import pandas as pd
from pyautogui import ImageNotFoundException
from tkinter import *
from tkinter import messagebox
import sys

# Ativa o FAILSAFE para garantir a pausa da automação quando necessário
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

# Carrega a planilha de contatos a serem processados
tabela = pd.read_excel('Indica_robo_19-11-2025.xlsx')

# Coordenadas dos campos na interface do Salesforce
campo_primeiro_nome_x, campo_primeiro_nome_y = 385, 356
campo_sobrenome_x, campo_sobrenome_y = 1076, 413
campo_email_x, campo_email_y = 360, 490
campo_telefone_x, campo_telefone_y = 1074, 489
campo_razao_social_x, campo_razao_social_y = 427, 545
campo_cnpj_x, campo_cnpj_y = 771, 546
botao_confirmar_x, botao_confirmar_y = 683, 618
botao_leads_x, botao_leads_y = 588, 173
botao_novo_lead_x, botao_novo_lead_y = 674, 236
botao_cancelar_x, botao_cancelar_y = 1143, 118

# Email padrão aplicado quando um contato não possui email
email_padrao = 'empresario@gmail.com'

# Coluna adicionada para registrar o status final do lead
tabela["STATUS"] = ""

# Função que executa a automação ao clicar no botão de iniciar
def btn_clicked():

    try:

        # Minimiza a interface
        window.iconify()
        print('Interface minimizada')

        print("Iniciando automação em 5 segundos...")
        print("Posicione a janela de indicação e não mexa no mouse!")
        print("Para pausar a automação: Mova o mouse para o canto superior esquerdo.")
        time.sleep(5)

        # Loop que percorre todos os registros da planilha
        for index, row in tabela.iterrows():

            # Extração e tratamento dos dados da linha atual
            contato = str(row['Contato'])
            telefone = str(row['Telefone'])
            cnpj = str(row['CNPJ'])
            email = str(row['EMAIL'])

            # Aplica email padrão quando o campo está vazio
            if pd.isna(email) or email.strip() == '':
                email_final = email_padrao
            else:
                email_final = email.strip() 

            print(f"\n--- Processando Linha {index + 1} de {len(tabela)} ---")
            print(f"Contato: {contato}")
            print(f"Telefone: {telefone}")
            print(f"CNPJ: {cnpj}")
            print(f"Email: {email_final}")

            # Preenchimento automático dos campos
            pyautogui.click(campo_primeiro_nome_x, campo_primeiro_nome_y)
            time.sleep(0.3)
            pyautogui.write(contato, interval=0.05)

            pyautogui.click(campo_sobrenome_x, campo_sobrenome_y)
            time.sleep(0.3)
            pyautogui.write(contato, interval=0.05)
            
            pyautogui.click(campo_razao_social_x, campo_razao_social_y)
            time.sleep(0.3)
            pyautogui.write(contato, interval=0.05)

            pyautogui.click(campo_telefone_x, campo_telefone_y)
            time.sleep(0.3)
            pyautogui.write(telefone, interval=0.05)

            pyautogui.click(campo_cnpj_x, campo_cnpj_y)
            time.sleep(0.3)
            pyautogui.write(cnpj, interval=0.05)

            pyautogui.click(campo_email_x, campo_email_y)
            time.sleep(0.3)
            pyautogui.write(email_final, interval=0.05)

            print("Clicando em 'Confirmar'...")
            pyautogui.click(botao_confirmar_x, botao_confirmar_y)
            time.sleep(10)

            try:
                # Detecta automaticamente se o Salesforce exibiu erro de CNPJ já cadastrado
                cnpj_indicado = pyautogui.locateOnScreen('print_cnpj_indicado.png',confidence=0.55)

                if cnpj_indicado:
                    print(f'Contato {contato} já estava indicado!')
                    tabela.at[index, "STATUS"] = "Já era indicado"

                    pyautogui.click(botao_cancelar_x, botao_cancelar_y)
                    time.sleep(3)

                else:
                    print(f"Lead {contato} indicado com sucesso!")
                    tabela.at[index, "STATUS"] = "Indicado pelo Robô"

            except ImageNotFoundException:
                # Quando o erro não aparece, considera-se um lead novo e indicado pelo robô
                print(f"Lead {contato} indicado (imagem não encontrada)!")
                tabela.at[index, "STATUS"] = "Indicado pelo Robô"

            print("Aguardando 15 segundos para o próximo registro...")
            time.sleep(15)

            # Retorna à tela de leads para inserir o próximo contato
            pyautogui.click(botao_leads_x, botao_leads_y)
            time.sleep(5)
            
            pyautogui.click(botao_novo_lead_x, botao_novo_lead_y)
            time.sleep(5)

        print('Todos os contatos foram indicados com sucesso!')

        # Exporta a planilha final com o status de cada contato
        tabela.to_excel("lemit_indicados (ROBO).xlsx", index=False)

    except pyautogui.FailSafeException:
        messagebox.showinfo("Pausado", "Automação PAUSADA.")
        print("Automação pausada pelo usuário.")
        return
        

# Gera a interface
window = Tk()

window.geometry("616x357")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 357,
    width = 616,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    311.0, 178.5,
    image=background_img)

img0 = PhotoImage(file = f"img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b0.place(
    x = 300, y = 251,
    width = 100,
    height = 32)

window.resizable(False, False)
window.mainloop()