#OLÁ!
#PROCEDIMENTO: REALIZAR PRESTAÇÃO DE CONTAS;
#POR: LOUISE-SESDEC;
#ALTERAÇÕES NO CÓDIGO PODEM SER ACESSADAS NO MEU DRIVE: <https://drive.google.com/drive/folders/1TMJkn2RBNvG7LNMlEWmTFi0uw9w5a1eA?usp=drive_link>.


#INSTRUÇÕES
#ATENÇÃO: É OBRIGATÓRIO ABRIR O DEPURADOR DO GOOGLE CHROME PARA EXECUTAR ESSE CÓDIGO
#EXECUTE NO CMD: "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"
#E LOGUE NO SIGEF

#BIBLIOTECAS UTILIZADAS:
from playwright.sync_api import sync_playwright, Page, TimeoutError
import pyautogui
import openpyxl
import keyboard
from datetime import date
from datetime import datetime
import time

#VARIÁVEIS IMPORTANTES
limite = 500
baixas = 0
robo_deve_parar = False
numero_nl = None
numero_nl2 = None

#PLANILHA NO EXCEL:
try:
    book = openpyxl.load_workbook('Planilha de Baixas.xlsx')
    pagina = book['Baixas']
    pagina.append(["UG","GESTAO","CREDOR","NOTA DE EMPENHO","NOTA DE LIQUIDACAO","PREPARACAO DE PAGAMENTO","VALOR","1 PASSO","3 PASSO","DATA DA BAIXA","HORA DA BAIXA","INSTRUMENTO","NATUREZA","PRESTACAO DE CONTAS"])
except: 
    pyautogui.alert(text='Deu algum erro na planilha.', title='Erro', button='OK')

# --- FUNÇÃO QUE SERÁ CHAMADA PELA HOTKEY ---
def parar_execucao():
    global robo_deve_parar
    print("\n!!! TECLA ESC ACIONADA! ENCERRANDO AUTOMACAO !!!")
    robo_deve_parar = True

# --- DEFINA SUA TECLA DE ATALHO ---
# Vamos usar 'k' como você sugeriu. 'esc' também é uma ótima opção.
tecla_de_panico = "Esc" 
keyboard.add_hotkey(tecla_de_panico, parar_execucao)
print(f"--- Robô iniciado. Pressione a tecla '{tecla_de_panico}' a qualquer momento para abortar com seguranca. ---")

pyautogui.confirm(text='Aperte OK quando o SIGEF estiver logado no depurador do Google Chrome', title='Depurador do Chrome' , buttons=['OK'])

#PORTA DO DEPURADOR DO GOOGLE CHROME
CHROME_DEBUG_URL = "http://localhost:9222"

with sync_playwright() as p:
    try:
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

        #PESQUISANDO FUNCIONALIDADE "REALIZAR PRESTAÇÃO DE CONTAS":
        if robo_deve_parar:
            if book:
                print("Garantindo que a planilha seja fechada...")
                book.close()
            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
            exit()
        
        if baixas <= limite:
            pesquisar_funcionalidades_sistema = frame.get_by_placeholder("Pesquisar funcionalidades do sistema...")
            pesquisar_funcionalidades_sistema.press("Control+KeyA+Backspace")
            pesquisar_funcionalidades_sistema.press_sequentially("Realizar Prestação de Contas")
            funcionalidade_sistema = frame.get_by_title("Realizar Prestação de Contas")
            
        
        #COM A JANELA "REALIZAR PRESTAÇÃO DE CONTAS" ABERTA...
        #ESSE COMANDO É ESSENCIAL
        #AQUI ELE RECONHECE O POPUP "REALIZAR PRESTAÇÃO DE CONTAS" E FOCA SOMENTE NELE:

            
            with guia.expect_popup() as popup_info:
                funcionalidade_sistema.click()
                realizar_prestacao_contas = popup_info.value
                realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
            

            #AQUELAS JANELAS QUE PERGUNTAM A UG, GESTAO E EMPENHO ESTÃO AQUI:


            if robo_deve_parar:
                if book:
                    print("Garantindo que a planilha seja fechada...")
                    book.close()
                pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                exit()

            if baixas <= limite:
                #ug = pyautogui.prompt(text='Digite a UG', title='Unidade Gestora' , default='150001')
                #gestao = pyautogui.prompt(text='Digite a Gestão', title='Gestão' , default='00001')
                #empenho = pyautogui.prompt(text='Digite o número da Nota de Empenho', title='Nota de Empenho' , default='0')
                #exercicio = pyautogui.confirm(text='Escolha o Exercício Financeiro', title='Exercício Financeiro' , buttons=['2024', '2025'])
                #unidade_gestora = str(input('Digite a UG: '))
                #gestao = str(input('Digite a Gestão: '))
                #numero_empenho = str(input('Digite o número da Nota de Empenho: '))
                #exercicio_financeiro = str(input('Digite o Exercício Financeiro: '))
                unidade_gestora = '150001'
                gestao = '00001'
                numero_empenho = '154'
                exercicio_financeiro = '2025'
            else:
                pyautogui.alert(text='Deu algum erro.', title='Erro', button='OK')

            while baixas <= limite:
                # --- PONTO DE VERIFICAÇÃO NO INÍCIO DO LOOP ---
                if robo_deve_parar:
                    if book:
                        print("Garantindo que a planilha seja fechada...")
                        book.close()
                    pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                    exit()

                #INFORMAÇÕES PRELIMINARES
                #DATA DE HOJE:
                data_atual = date.today() 
                data_de_baixa = data_atual.strftime("%d/%m/%Y")
                #HORA:
                agora = datetime.now()
                hora = agora.strftime("%H:%M:%S")
                # --- PONTO DE VERIFICAÇÃO NO INÍCIO DO LOOP ---
                if robo_deve_parar:
                    if book:
                        print("Garantindo que a planilha seja fechada...")
                        book.close()
                    pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                    exit()
                
                preencher_unidade_gestora = realizar_prestacao_contas.locator("#txtUnidadeGestora")
                preencher_unidade_gestora.wait_for()
                preencher_unidade_gestora.dblclick()
                preencher_unidade_gestora.press('Delete')
                preencher_unidade_gestora.press_sequentially(unidade_gestora)
                preencher_gestao = realizar_prestacao_contas.locator("#txtGestao_SIGEFPesquisa")
                preencher_gestao.press_sequentially(gestao)
                ponto_interrogacao = realizar_prestacao_contas.locator("#txtPrestacaoContas2_lnkBtnPesquisa")
                with realizar_prestacao_contas.expect_popup() as popup_info:
                    ponto_interrogacao.click()
                    # --- PONTO DE VERIFICAÇÃO NO INÍCIO DO LOOP ---
                    if robo_deve_parar:
                        if book:
                            print("Garantindo que a planilha seja fechada...")
                            book.close()
                        pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                        exit()
                    selecionar_prestacao_contas = popup_info.value
                    selecionar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                    preencher_numero_empenho = selecionar_prestacao_contas.locator("#txtNotaEmpenho2_SIGEFPesquisa")
                    preencher_numero_empenho.wait_for()
                    time.sleep(0.3)
                    preencher_numero_empenho.dblclick()
                    preencher_numero_empenho.press('Delete')
                    preencher_numero_empenho.press_sequentially(numero_empenho)
                    preencher_exercicio = selecionar_prestacao_contas.locator("#txtNEAno")
                    preencher_exercicio.dblclick()
                    preencher_exercicio.press('Delete')
                    preencher_exercicio.press_sequentially(exercicio_financeiro)
                    situacao_prestacao_contas = selecionar_prestacao_contas.locator("#cboSituacao")
                    situacao_prestacao_contas.select_option(label="Paga")
                    botao_confirmar = selecionar_prestacao_contas.get_by_role("button", name="Confirmar a Consulta")
                    botao_confirmar.click()
                    try:
                        selecionar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        locator_candidatos = selecionar_prestacao_contas.get_by_role("cell", name="2025PC", exact=False)
                        todos_os_textos = locator_candidatos.all_inner_texts()
                        numeros_pc = []
                        for texto in todos_os_textos:
                            if texto.startswith("2025PC"):
                                parte_numerica_str = texto[6:]
                                if parte_numerica_str.isdigit():
                                    numeros_pc.append(int(parte_numerica_str))
                        if not numeros_pc:
                            raise Exception("Nenhum número de PC válido foi encontrado na lista de células.")
                            selecionar_prestacao_contas.close()
                            baixas = 999999999999999
                            break

                        menor_numero = min(numeros_pc)
                        nome_exato_da_celula = f"2025PC{menor_numero:06d}"
                        celula_para_clicar = selecionar_prestacao_contas.get_by_role("cell", name=nome_exato_da_celula, exact=True)
                        celula_para_clicar.click()                 

                    except Exception as e:
                        print(f"\nOcorreu um erro ao tentar selecionar a PC: {e}")
                        selecionar_prestacao_contas.close()
                        baixas = 999999999999999
                        break
                
                realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                
                try:

                    #NOME E CPF
                    credor = realizar_prestacao_contas.locator("#txtCredorNotaEmpenho")
                    credor.wait_for()
                    value_credor = credor.input_value()
                    credor = value_credor

                    #NATUREZA
                    natureza = realizar_prestacao_contas.locator('#txtNaturezaDespesa')
                    value_natureza = natureza.input_value()
                    natureza = value_natureza

                    #INSTRUMENTO
                    instrumento = realizar_prestacao_contas.locator("#txtInstrumento")
                    value_instrumento = instrumento.input_value()
                    instrumento = value_instrumento

                    #NOTA DE EMPENHO
                    nota_empenho = realizar_prestacao_contas.locator("#txtNotaEmpenho")
                    value_nota_empenho = nota_empenho.input_value()
                    nota_empenho = value_nota_empenho

                    #VALOR
                    valor = realizar_prestacao_contas.locator("#txtValorPreparacaoPagamento")
                    value_valor = valor.input_value()
                    valor = value_valor

                    #DESPESA CERTIFICADA
                    despesa_certificada = realizar_prestacao_contas.locator("#txtDespesaCertificada")
                    value_despesa_certificada = despesa_certificada.input_value()
                    despesa_certificada = value_despesa_certificada

                    #NOTA LIQUIDAÇÃO
                    nota_liquidacao = realizar_prestacao_contas.locator("#txtNotaLancamento")
                    value_nota_liquidacao = nota_liquidacao.input_value()
                    nota_liquidacao = value_nota_liquidacao

                    #PREPARACAO DE PAGAMENTO
                    preparacao_pagamento = realizar_prestacao_contas.locator("#txtPreparacaoPagamento")
                    value_preparacao_pagamento = preparacao_pagamento.input_value()
                    preparacao_pagamento = value_preparacao_pagamento

                    #ORDEM BANCÁRIA

                    #PRESTACAO DE CONTAS
                    prestacao_contas = realizar_prestacao_contas.locator("#txtPrestacaoContas2_SIGEFPesquisa")
                    value_prestacao_contas = prestacao_contas.input_value()
                    prestacao_contas = value_prestacao_contas

                except Exception as e:
                    print(f"Ocorreu um erro ao tentar extrair o valor do credor: {e}")
                
                if robo_deve_parar:
                        if book:
                            print("Garantindo que a planilha seja fechada...")
                            book.close()
                        pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                        exit()
                botao_prestacao_contas = realizar_prestacao_contas.locator("#btnPrestacaoContas")
                botao_prestacao_contas.click()
                situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                situacao = situacao.input_value()

                while situacao != "Baixa Regular":

                    if robo_deve_parar:
                        if book:
                            print("Garantindo que a planilha seja fechada...")
                            book.close()
                        pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                        exit()
                    situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                    situacao = situacao.input_value()
                    data = realizar_prestacao_contas.locator("#txtDataPrestacaoContas_SIGEFData")
                    data.press_sequentially('h')
                    processo_spp = realizar_prestacao_contas.locator("#txtProcessoSPP_SIGEFPesquisa")
                    processo_spp.press_sequentially('0')
                    observacao = realizar_prestacao_contas.locator("#txtObservacao")
                    operacao = realizar_prestacao_contas.locator("#cboOperacao")
                    botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                    botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                    botao_limpar = realizar_prestacao_contas.get_by_role("link", name="Limpar a Tela")
                    
                    if situacao == 'Paga':
                        if robo_deve_parar:
                            numero_nl = "Não foi feita!" 
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()
                        print('1 PASSO: Prestacao de contas ' + prestacao_contas + ' do credor ' + credor + '.')
                        operacao.select_option(label="Entregue")
                        time.sleep(0.3)
                        observacao.press_sequentially("1º Passo: Entregue. Prestação de Contas de " + instrumento + ", Natureza da Despesa: " + natureza + ", em nome do Credor " + credor + ", Despesa Certificada: " + despesa_certificada + ", Nota Liquidação: " + nota_liquidacao + ", Preparação de Pagamento: " + preparacao_pagamento + ", Nota de Empenho: " + nota_empenho + ".")
                        botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                        botao_confirmar.wait_for()
                        botao_confirmar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)      
                        try:
                            caixa_mensagem_sucesso = realizar_prestacao_contas.locator(".SIGEFMensagemSucesso")
                            caixa_mensagem_sucesso.wait_for(timeout=10000)
                            texto_completo1 = caixa_mensagem_sucesso.inner_text()
                            if "O número gerado foi" in texto_completo1:
                                numero_nl = texto_completo1.split("foi ")[1]
                                numero_nl = numero_nl.strip('.')
                            botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                            botao_consultar.wait_for()
                            botao_consultar.click()
                            realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        except:
                            try:
                                caixa_mensagem_sucesso = realizar_prestacao_contas.locator(".SIGEFMensagemSucesso")
                                caixa_mensagem_sucesso.wait_for(timeout=10000)
                                texto_completo1 = caixa_mensagem_sucesso.inner_text()
                                if "O número gerado foi" in texto_completo1:
                                    numero_nl = texto_completo1.split("foi ")[1]
                                    numero_nl = numero_nl.strip('.')
                                botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                                botao_consultar.wait_for()
                                botao_consultar.click()
                                realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                            except:
                                botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                                botao_consultar.wait_for()
                                botao_consultar.click()
                        if robo_deve_parar:
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()


                    
                    if situacao == 'Entregue':
                        if robo_deve_parar:
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()
                        print('2 PASSO: Prestacao de contas ' + prestacao_contas + ' do credor ' + credor + '.')
                        operacao = realizar_prestacao_contas.locator("#cboOperacao")
                        operacao.select_option(label="Em Análise")
                        time.sleep(0.3)
                        observacao.press_sequentially("2º Passo: Em Análise. Prestação de Contas de " + instrumento + ", Natureza da Despesa: " + natureza + ", em nome do Credor " + credor + ", Despesa Certificada: " + despesa_certificada + ", Nota Liquidação: " + nota_liquidacao + ", Preparação de Pagamento: " + preparacao_pagamento + ", Nota de Empenho: " + nota_empenho + ".")
                        botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                        botao_confirmar.wait_for()
                        botao_confirmar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                        botao_consultar.wait_for()
                        botao_consultar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        if robo_deve_parar:
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()
                    
                    if situacao == 'Em Análise':
                        if robo_deve_parar:
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            exit()
                        print('3 PASSO: Prestacao de contas ' + prestacao_contas + ' do credor ' + credor + '.')
                        operacao = realizar_prestacao_contas.locator("#cboOperacao")
                        operacao.select_option(label="Baixa Regular")
                        time.sleep(0.3)
                        observacao.press_sequentially("3º Passo: Baixa Regular. Prestação de Contas de " + instrumento + ", Natureza da Despesa: " + natureza + ", em nome do Credor " + credor + ", Despesa Certificada: " + despesa_certificada + ", Nota Liquidação: " + nota_liquidacao + ", Preparação de Pagamento: " + preparacao_pagamento + ", Nota de Empenho: " + nota_empenho + ".")
                        botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                        botao_confirmar.wait_for()
                        if robo_deve_parar:
                            numero_nl2 = "Não foi feita!"
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()
                        botao_confirmar.click()

                        try:
                            time.sleep(0.3)
                            caixa_mensagem_sucesso = realizar_prestacao_contas.locator(".SIGEFMensagemSucesso")
                            caixa_mensagem_sucesso.wait_for(timeout=10000)
                            texto_completo2 = caixa_mensagem_sucesso.inner_text()
                            if "O número gerado foi" in texto_completo2:
                                numero_nl2 = texto_completo2.split("foi ")[1]
                                numero_nl2 = numero_nl2.strip('.')
                            botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                            botao_consultar.wait_for()
                            botao_consultar.click()
                            realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        except:
                            try:
                                caixa_mensagem_sucesso = realizar_prestacao_contas.locator(".SIGEFMensagemSucesso")
                                caixa_mensagem_sucesso.wait_for(timeout=10000)
                                texto_completo2 = caixa_mensagem_sucesso.inner_text()
                                if "O número gerado foi" in texto_completo2:
                                    numero_nl2 = texto_completo2.split("foi ")[1]
                                    numero_nl2 = numero_nl2.strip('.')
                                botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                                botao_consultar.wait_for()
                                botao_consultar.click()
                                realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                            except:
                                botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                                botao_consultar.wait_for()
                                botao_consultar.click()
                        
                        if texto_completo1==texto_completo2:
                            numero_nl2 = 'Erro.'
                        
                        botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                        botao_consultar.wait_for()
                        botao_consultar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                        situacao = situacao.input_value()
                        if robo_deve_parar:
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento, natureza,prestacao_contas])
                                book.save('Planilha de Baixas.xlsx')
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()
                        
                    if situacao == 'Baixa Regular':  
                        baixas = baixas + 1
                        numero_de_baixa = str(baixas)
                        print('Número de Notas Lançamento geradas: ' + numero_nl +' e ' + numero_nl2 + '.')
                        pagina.append([unidade_gestora,gestao,credor,nota_empenho,nota_liquidacao,preparacao_pagamento,valor,numero_nl,numero_nl2,data_de_baixa,hora,instrumento,natureza,prestacao_contas])
                        book.save('Planilha de Baixas.xlsx')
                        botao_limpar = realizar_prestacao_contas.get_by_role("link", name="Limpar a Tela")
                        botao_limpar.click()
                        print('Fiz ' + numero_de_baixa +  ' baixas até aqui. Prosseguindo para a próxima...')
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        if robo_deve_parar:
                            if book:
                                print("Garantindo que a planilha seja fechada...")
                                book.close()
                            pyautogui.alert(text='Tecla ESC acionada. Automacao encerrada', title='Tecla de Panico Acionada', button='OK')
                            exit()                       
                        
    except TimeoutError:
        print("\nERRO: Timeout! Não foi possível encontrar um elemento a tempo.")
        if book:
            print("Garantindo que a planilha seja fechada...")
            book.close()
    except Exception as e:
        print(f"\nOcorreu um erro: {e}")
        print("Causas possíveis:")
        print("1. O Chrome não foi iniciado com o comando de depuração (--remote-debugging-port).")
        print("2. A URL de depuração está incorreta.")
        print("3. Não há nenhuma aba (página) aberta no navegador.")
        if book:
            print("Garantindo que a planilha seja fechada...")
            book.close()

if book:
    print("Garantindo que a planilha seja fechada...")
    book.close()
print("\nScript finalizado. A janela de depuracao permanece aberta.")
keyboard.remove_hotkey(tecla_de_panico) 
pyautogui.alert(text='Encerrei por aqui.', title='Fim', button='OK')