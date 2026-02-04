import pyautogui
import time
import pandas as pd
from pyautogui import ImageNotFoundException
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox

# Ativa o FAILSAFE para garantir a pausa da automação quando necessária
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

# Coordenadas dos campos na interface do Salesforce
campo_primeiro_nome_x, campo_primeiro_nome_y = 365, 375
campo_sobrenome_x, campo_sobrenome_y = 986, 436
campo_email_x, campo_email_y = 257, 508
campo_telefone_x, campo_telefone_y = 1096, 506
campo_razao_social_x, campo_razao_social_y = 614, 570
campo_cnpj_x, campo_cnpj_y = 728, 570
botao_confirmar_x, botao_confirmar_y = 683, 618
botao_leads_x, botao_leads_y = 591, 210
botao_novo_lead_x, botao_novo_lead_y = 674, 270
botao_cancelar_x, botao_cancelar_y = 1143, 151
# botao_conta_aberta_x, botao_conta_aberta_y = 930,465
# botao não funciona, precisa colocar pra automação dar refresh na página
#  
# Email padrão aplicado quando um contato não possui email
email_padrao = 'empresario@gmail.com'

# Função que executa a automação ao clicar no botão de iniciar
def btn_clicked():

    try:

        # Abre o explorador de arquivos e solicita ao usuário para selecionar um arquivo Excel
        tabela_caminho = askopenfilename(title='selecione um arquivo Excel (.xlsx)')
        tabela = pd.read_excel(tabela_caminho)
        tabela["STATUS"] = ""

        # Minimiza a interface para começar a automação
        window.iconify()
        print('Interface minimizada')

        print("Iniciando automação em 5 segundos...")
        print("Posicione a janela de indicação e não mexa no mouse!")
        print("Para pausar a automação: Mova o mouse para o canto superior esquerdo até aparecer uma janela de pausa.")
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
                # Detecta automaticamente se o Salesforce exibiu tela de conta aberta
                conta_aberta = pyautogui.locateOnScreen('print_conta_aberta.png', confidence=0.55)

                if cnpj_indicado:
                    print(f'Contato {contato} já estava indicado!')
                    tabela.at[index, "STATUS"] = "Já era indicado"

                    pyautogui.click(botao_cancelar_x, botao_cancelar_y)
                    time.sleep(3)

                elif conta_aberta:
                    print(f'Contato {contato} estava com conta aberta!')
                    tabela.at[index, "STATUS"] = "Conta aberta"

                    pyautogui.hotkey('ctrl', 'f5')

                else:
                    print(f"Lead {contato} indicado com sucesso!")
                    tabela.at[index, "STATUS"] = "Indicado pelo Robô"

            except ImageNotFoundException:
                # Quando o erro não aparece, considera-se um lead novo e indicado pelo robô
                print(f"Lead {contato} indicado (imagem não encontrada)!")
                tabela.at[index, "STATUS"] = "Indicado pelo Robô"

            # Salva backup caso a automação dê erro ou pause
            tabela.to_excel("backup_lemit_indicados (ROBO).xlsx", index=False)

            # Retorna à tela de leads para inserir o próximo contato
            print("Aguardando 15 segundos para o próximo registro...")
            time.sleep(15)

            pyautogui.click(botao_leads_x, botao_leads_y)
            time.sleep(5)
            
            pyautogui.click(botao_novo_lead_x, botao_novo_lead_y)
            time.sleep(5)

        print('Todos os contatos foram indicados com sucesso!')

        # Exporta a planilha final com o status de cada contato
        tabela.to_excel("lemit_indicados (ROBO).xlsx", index=False)

    # Cria uma janela quando a pessoa pausa a automação
    except pyautogui.FailSafeException:
        messagebox.showinfo("Pausado", "Automação PAUSADA.")
        print("Automação pausada pelo usuário.")
        return
        

# Gera a interface
window = Tk()
window.title('Automação Salesforce')
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