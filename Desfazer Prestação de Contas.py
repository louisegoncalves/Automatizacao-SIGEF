from playwright.sync_api import sync_playwright, Page, TimeoutError
import time
import pyautogui
import openpyxl
from datetime import date
from datetime import datetime

limite = 1000

#PLANILHA NO EXCEL:
try:
    book = openpyxl.load_workbook('Planilha de Baixas.xlsx')
    pagina = book['Baixas']
    pagina.append(["UG","GESTAO","CREDOR","NOTA DE EMPENHO","NOTA DE LIQUIDACAO","PREPARACAO DE PAGAMENTO","VALOR","1 PASSO","3 PASSO","DATA DA BAIXA","HORA DA BAIXA","INSTRUMENTO","NATUREZA","PRESTACAO DE CONTAS"])
except: 
    pyautogui.alert(text='Deu algum erro na planilha.', title='Erro', button='OK')

#PORTA DO DEPURADOR DO GOOGLE CHROME
CHROME_DEBUG_URL = "http://localhost:9222"

unidade_gestora = '150001'
gestao = '00001'
exercicio_financeiro = '2025'
numero_empenho = '000154'

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
        pesquisar_funcionalidades_sistema = frame.get_by_placeholder("Pesquisar funcionalidades do sistema...")
        pesquisar_funcionalidades_sistema.press("Control+KeyA+Backspace")
        pesquisar_funcionalidades_sistema.press_sequentially("Realizar Prestação de Contas")
        funcionalidade_sistema = frame.get_by_title("Realizar Prestação de Contas")
        funcionalidade_sistema.click()
        

        guia.pause()
        #COM A JANELA "REALIZAR PRESTAÇÃO DE CONTAS" ABERTA...
        #ESSE COMANDO É ESSENCIAL
        #AQUI ELE RECONHECE O POPUP "REALIZAR PRESTAÇÃO DE CONTAS" E FOCA SOMENTE NELE:
        with guia.expect_popup() as popup_info:
            realizar_prestacao_contas = popup_info.value
            realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
            baixa = 0
            while baixa <= limite:
                preencher_unidade_gestora = realizar_prestacao_contas.locator("#txtUnidadeGestora")
                preencher_unidade_gestora.wait_for()
                preencher_unidade_gestora.dblclick()
                preencher_unidade_gestora.press('Delete')
                preencher_unidade_gestora.press_sequentially(unidade_gestora)
                preencher_gestao = realizar_prestacao_contas.locator("#txtGestao_SIGEFPesquisa")
                preencher_gestao.press_sequentially(gestao)
                ponto_interrogacao = realizar_prestacao_contas.locator("#txtPrestacaoContas2_lnkBtnPesquisa")
                ponto_interrogacao.click()
                with realizar_prestacao_contas.expect_popup() as popup_info:
                    selecionar_prestacao_contas = popup_info.value
                    selecionar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                    preencher_numero_empenho = selecionar_prestacao_contas.locator("#txtNotaEmpenho2_SIGEFPesquisa")
                    preencher_numero_empenho.wait_for()
                    preencher_numero_empenho.dblclick()
                    preencher_numero_empenho.press('Delete')
                    preencher_numero_empenho.press_sequentially(numero_empenho)
                    preencher_exercicio = selecionar_prestacao_contas.locator("#txtNEAno")
                    preencher_exercicio.dblclick()
                    preencher_exercicio.press('Delete')
                    preencher_exercicio.press_sequentially(exercicio_financeiro)
                    situacao_prestacao_contas = selecionar_prestacao_contas.locator("#cboSituacao")
                    situacao_prestacao_contas.select_option(label="Em Análise")
                    botao_confirmar = selecionar_prestacao_contas.get_by_role("button", name="Confirmar a Consulta")
                    botao_confirmar.click()
                    try:
                        selecionar_prestacao_contas.wait_for_load_state()
                        #paginacao = selecionar_prestacao_contas.locator('#pagFormulario_LblPaginacao')
                        #paginacao.wait_for()
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
                            baixa = 999999999999999
                            break
                        menor_numero = min(numeros_pc)
                        nome_exato_da_celula = f"2025PC{menor_numero:06d}"
                        celula_para_clicar = selecionar_prestacao_contas.get_by_role("cell", name=nome_exato_da_celula, exact=True)
                        celula_para_clicar.click()                 

                    except Exception as e:
                        print(f"\nOcorreu um erro ao tentar selecionar a PC: {e}")
                        baixa = 999999999999999
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

                botao_prestacao_contas = realizar_prestacao_contas.locator("#btnPrestacaoContas")
                botao_prestacao_contas.click()
                situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                situacao = situacao.input_value()

                while situacao != "Paga":
                    situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                    situacao = situacao.input_value()
                    print(situacao)
                    data = realizar_prestacao_contas.locator("#txtDataPrestacaoContas_SIGEFData")
                    data.press_sequentially('h')
                    processo_spp = realizar_prestacao_contas.locator("#txtProcessoSPP_SIGEFPesquisa")
                    processo_spp.press_sequentially('0')
                    observacao = realizar_prestacao_contas.locator("#txtObservacao")
                    operacao = realizar_prestacao_contas.locator("#cboOperacao")
                    botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                    botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                    botao_limpar = realizar_prestacao_contas.get_by_role("link", name="Limpar a Tela")
                    
                    if situacao == 'Em Análise':
                        operacao.select_option(label="Entregue")
                        observacao.press_sequentially('Estorno')
                        botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                        botao_confirmar.wait_for()
                        botao_confirmar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                        botao_consultar.wait_for()
                        botao_consultar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                    if situacao == 'Entregue':
                        operacao = realizar_prestacao_contas.locator("#cboOperacao")
                        operacao.select_option(label="Paga")
                        observacao.press_sequentially('Estorno')
                        botao_confirmar = realizar_prestacao_contas.get_by_role("button", name="Confirmar a Operação")
                        botao_confirmar.wait_for()
                        botao_confirmar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000)
                        botao_consultar = realizar_prestacao_contas.get_by_role("button", name="Consultar o Registro")
                        botao_consultar.wait_for()
                        botao_consultar.click()
                        realizar_prestacao_contas.wait_for_load_state('networkidle', timeout=30000) 
                        situacao = realizar_prestacao_contas.locator('#txtSituacaoPrestacaoContas')
                        situacao = situacao.input_value()
                        botao_limpar = realizar_prestacao_contas.get_by_role("link", name="Limpar a Tela")
                        botao_limpar.click()
                
                baixa = baixa + 1
                print("Número de baixas realizadas: " + str(baixa) + ". Prosseguinto para a próxima...")
   
    except TimeoutError:
        print("\nERRO: Timeout! Não foi possível encontrar um elemento a tempo.")
    except Exception as e:
        print(f"\nOcorreu um erro: {e}")
        print("Causas possíveis:")
        print("1. O Chrome não foi iniciado com o comando de depuração (--remote-debugging-port).")
        print("2. A URL de depuração está incorreta.")
        print("3. Não há nenhuma aba (página) aberta no navegador.")

# O script terminará sem fechar o navegador, pois não usamos browser.close()
print("\nScript finalizado. O navegador permanece aberto.")