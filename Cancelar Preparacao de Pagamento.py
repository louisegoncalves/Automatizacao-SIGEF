#OLÁ!
#PROCEDIMENTO: CANCELAR PREPARAÇÃO PAGAMENTO;
#POR: LOUISE-SESDEC;
#ALTERAÇÕES NO CÓDIGO PODEM SER ACESSADAS NO MEU GITHUB: <https://github.com/louisegoncalves/Automatizacao-SIGEF>.

#INSTRUÇÕES
#ATENÇÃO: É OBRIGATÓRIO ABRIR O DEPURADOR DO GOOGLE CHROME PARA EXECUTAR ESSE CÓDIGO
#EXECUTE NO CMD: "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"
#E LOGUE NO SIGEF

#BIBLIOTECAS UTILIZADAS:
import openpyxl
import pyautogui
import sys
import keyboard
from playwright.sync_api import sync_playwright
import time
from datetime import datetime

#VARIÁVEIS IMPORTANTES
robo_deve_parar = False
preparando_pagamento = True

#PLANILHA NO EXCEL:
try:
    book = openpyxl.load_workbook('Para cancelar.xlsx')
    pagina = book['Entrada']
    pagina1 = book['Saída']
except: 
    pyautogui.alert(text='Deu algum erro na planilha.', title='Erro', button='OK')
    sys.exit()     

#INFORMAÇÕES PRELIMINARES
#DATA DE HOJE:
data = datetime.today().strftime("%d/%m/%Y")
#HORA:
hora = datetime.now().strftime('%H:%M:%S')
         
#SE QUISER DESATIVAR AQUELA JANELA DO COMEÇO PODE EXCLUIR ELA AQUI:
pyautogui.alert(text='Procedimento: Cancelar Preparação Pagamento.', title='Início', button='OK')

#FUNÇÃO QUE SERÁ CHAMADA PELA TECLA DE PANICO
def parar_execucao():
    global robo_deve_parar
    print("\n!!! TECLA ESC ACIONADA! ENCERRANDO AUTOMACAO !!!")
    robo_deve_parar = True

#FUNÇÃO QUE ENCERRA O CODIGO E FECHA A PLANILHA COM SEGURANÇA
#A PLANILHA DEVE SEMPRE SER FECHADA ANTES DE ENCERRAR, POIS CORRE O RISCO DE CORROMPER
def verificar_panico_e_sair(workbook):
    global robo_deve_parar
    if robo_deve_parar:
        print("Garantindo que a planilha seja fechada...")
        if workbook:
            workbook.close()
        pyautogui.alert('Tecla ESC acionada. Automação encerrada.')
        sys.exit()

#DEFINA SUA TECLA DE PÂNICO
tecla_de_panico = "Esc" 
keyboard.add_hotkey(tecla_de_panico, parar_execucao)
print(f"Robô iniciado. Pressione a tecla '{tecla_de_panico}' a qualquer momento para abortar com seguranca.")

#AQUI ELE VAI PEDIR PARA ABRIR O SIGEF PELO DEPURADOR DO GOOGLE
pyautogui.confirm(text='Aperte OK quando o SIGEF estiver logado no depurador do Google Chrome', title='Depurador do Chrome' , buttons=['OK'])

#PORTA DO DEPURADOR DO GOOGLE CHROME
CHROME_DEBUG_URL = "http://localhost:9222"

if robo_deve_parar:
    verificar_panico_e_sair(book)

#EXECUTANDO O PLAYWRIGHT DE FORMA SÍNCRONA
with sync_playwright() as p:
        if robo_deve_parar:
            verificar_panico_e_sair(book)

        #CONECTAR AO NAVEGADOR JÁ ABERTO:
        print(f"Tentando se conectar ao Chrome na porta de depuração: {CHROME_DEBUG_URL}")
        browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
        print("Conexão estabelecida com sucesso!")

        #OBTER A PÁGINA QUE ESTÁ ABERTA:
        #Quando conectamos, precisamos pegar o contexto e a página existentes;
        #Geralmente, a primeira página do primeiro contexto é a que queremos.
        janela = browser.contexts[0]
        guia = janela.pages[0]

        #VERIFICAR A PÁGINA ABERTA:
        print(f"Assumindo o controle da página com o título: '{guia.title()}'")
            
        #LOCALIZANDO O IFRAME:
        frame = guia.frame_locator('iframe[src="/SIGEF2025/SEG/#/SEGControleAcesso?p=1"]')
        
        if robo_deve_parar:
            verificar_panico_e_sair(book)
        
        #INÍCIO                
        print("Iniciando!")
        pesquisar_funcionalidades_sistema = frame.get_by_placeholder("Pesquisar funcionalidades do sistema...")
        pesquisar_funcionalidades_sistema.press("Control+KeyA+Backspace")
        pesquisar_funcionalidades_sistema.press_sequentially("Cancelar Preparação Pagamento")
        funcionalidade_sistema = frame.get_by_title("Cancelar Preparação Pagamento")

        linha = 61
        coluna = 1

        while preparando_pagamento == True:
        
            with guia.expect_popup() as popup_info:
                funcionalidade_sistema.click()
                Cancelar_Preparação_Pagamento = popup_info.value
                Cancelar_Preparação_Pagamento.wait_for_load_state('networkidle', timeout=50000)
            
            #LENDO A PP
            for row in pagina.iter_rows(min_row=linha,max_col=coluna,max_row=linha):
                for cell in row:
                    pp = cell.value
                    pp = str(pp)
            
            if pp != None:
                preparando_pagamento = True
            else:
                preparando_pagamento = False
            
            if pp != "None":
                preparando_pagamento = True
            else:
                preparando_pagamento = False
                sys.exit()

            if robo_deve_parar:
                verificar_panico_e_sair(book)
            else:
                print("Estou na linha " + str(linha) + " da planilha, referente à Preparação de Pagamento " + str(pp) + ".")
    
            Data_Referência = Cancelar_Preparação_Pagamento.locator("#txtDtReferencia_SIGEFData")
            Data_Referência.wait_for(timeout=50000)
            Unidade_Gestora = Cancelar_Preparação_Pagamento.locator("#txtUnidadeGestora")
            Gestão = Cancelar_Preparação_Pagamento.locator("#txtGestao_SIGEFPesquisa")
            Preparação_Pagamento = Cancelar_Preparação_Pagamento.locator("#txtPrepPag")
            Número_da_Preparação_de_Pagamento = Cancelar_Preparação_Pagamento.locator("#txtPrepPagSeq_SIGEFPesquisa")
            Observação = Cancelar_Preparação_Pagamento.locator("#txtObservacao")
            Botão_Confirmar = Cancelar_Preparação_Pagamento.get_by_role("button", name="Confirmar a Operação")
            Botão_Limpar = Cancelar_Preparação_Pagamento.get_by_role("link", name="Limpar a Tela")
            Botão_Fechar = Cancelar_Preparação_Pagamento.get_by_role("link", name="Sair da Transação")
            Botão_Incluir = Cancelar_Preparação_Pagamento.get_by_role("button", name="Incluir o Documento")
            
            Número = pp.replace('2025PP','')
            
            Data_Referência.press("Backspace")
            Data_Referência.press_sequentially(data)
            Unidade_Gestora.press("Backspace")
            Unidade_Gestora.press_sequentially("150001")

            Gestão.press_sequentially("1")
            Número_da_Preparação_de_Pagamento.press_sequentially(Número)

            Observação.press_sequentially("Cancelamento da Preparação de Pagamento " + str(pp) + " para correção de repasse financeiro. Foi emitida equivocadamente como tipo Regularização, e deveria ter sido emitida Descentralizada.")

            Botão_Confirmar.click()

            try:
                Botão_Incluir = Cancelar_Preparação_Pagamento.get_by_role("button", name="Incluir o Documento")
                Botão_Incluir.wait_for(timeout=50000)
                Botão_Incluir.click()
            except:
                preparando_pagamento = False

            try:
                Mensagem_Sucesso = Cancelar_Preparação_Pagamento.get_by_text("Operação realizada com") 
                Mensagem_Sucesso.wait_for(timeout=50000)
                Mensagem = Mensagem_Sucesso.inner_text()
                print(Mensagem)
                if "Operação realizada com sucesso. O número gerado foi " in Mensagem:
                    Nota_Lançamento = Mensagem.replace("Operação realizada com sucesso. O número gerado foi ","")
                    Nota_Lançamento = Nota_Lançamento.replace(".","")
                else:
                    preparando_pagamento = False
            except:
                preparando_pagamento =  False  

            try:
                pagina1.delete_rows(linha,1)
                pagina1.append([Nota_Lançamento])
                book.save('Para Cancelar.xlsx')
            except:                                           
                print("Deu algum erro ao salvar a planilha, a planilha de backup foi solicitada.")
                sys.exit()

            Cancelar_Preparação_Pagamento.close()
            linha = linha + 1

print("Fim.")
if book:
    book.close()
print("\nScript finalizado. A janela de depuracao permanece aberta.")
keyboard.remove_hotkey(tecla_de_panico) 
pyautogui.alert(text='Encerrei por aqui.', title='Fim', button='OK')

            



            
            
            




        


            