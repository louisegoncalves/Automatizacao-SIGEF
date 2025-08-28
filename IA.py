import google.generativeai as genai
import pyautogui
import keyboard
import sys

#Primeiro: Você define a chave corretamente em uma variável e configura o genai para usá-la.
#AVISO: Não suba este código com a chave diretamente para o GitHub.
GOOGLE_API_KEY = "AIzaSyDqtryzby5ctWwqUPLnYmo7NHKqtdgcxGc" 

try:
    genai.configure(api_key=GOOGLE_API_KEY)
    #Escolha o modelo que deseja usar.
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    print(f"Erro ao configurar a API: {e}")
    pyautogui.alert(text=f'Nao foi possivel conectar a API do Gemini. Verifique sua chave.\n\nErro: {e}', title='Erro de API', button='OK')
    sys.exit()

try:
    with open('decreto28874.txt', 'r', encoding='utf-8') as f:
        decreto_28874 = f.read()
except FileNotFoundError:
    print("Erro: Arquivo 'decreto28874.txt' não encontrado.")
    exit()
try:
    with open('lei14133.txt', 'r', encoding='utf-8') as f:
            lei_14133 = f.read()
except FileNotFoundError:
    print("Erro: Arquivo 'lei14133.txt' não encontrado.")
    exit()

#VARIÁVEIS IMPORTANTES
robo_deve_parar = False

#ABORTAR OPERAÇÃO
def parar_execucao():
    global robo_deve_parar
    print("\n!!! TECLA ESC ACIONADA! ENCERRANDO!!!")
    robo_deve_parar = True

tecla_de_panico = "Esc" 
keyboard.add_hotkey(tecla_de_panico, parar_execucao)
print(f"--- Robô iniciado. Pressione a tecla '{tecla_de_panico}' a qualquer momento para abortar com seguranca. ---")

prompt_inicial = f"""
Você é um assistente especialista em compras e licitações da Secretaria de Estado da Segurança, Defesa e Cidadania do Estado de Rondônia, sua única fonte de conhecimento é o documento fornecido abaixo.
Responda a TODAS as perguntas do usuário baseando-se APENAS e EXCLUSIVAMENTE no conteúdo dos documentos abaixo.
Se a resposta não estiver nos documentos, diga "Essa informação não está disponível nos documentos que me foram fornecidos".
As respostas devem estar em LÍNGUA PORTUGUESA, seguindo as normas ortográficas, gramaticais e de sintaxe.
Não use nenhum conhecimento externo.
--- INÍCIO DA LEGISLAÇÃO PRIMÁRIA: DECRETO N° 28.874, DE 25 DE JANEIRO DE 2024 ---
{decreto_28874}
--- FIM DA LEGISLAÇÃO PRIMÁRIA ---
--- INÍCIO DA LEGISLAÇÃO SECUNDÁRIA: LEI Nº 14.133, DE 1º DE ABRIL DE 2021 ---
{lei_14133}
--- FIM DA LEGISLAÇÃO SECUNDÁRIA ---
"""

try:
    chat = model.start_chat(history=[
        # Inserimos o prompt inicial como a primeira "fala" do usuário
        {'role': 'user', 'parts': [prompt_inicial]},
        # Adicionamos uma resposta modelo para que a conversa possa começar
        {'role': 'model', 'parts': ["Ok, documento lido e compreendido. Estou pronto para responder suas perguntas com base nele."]}
    ])
except Exception as e:
    print(f"Erro ao iniciar o chat: {e}")
    sys.exit()

print(f"--- Chatbot especialista no documento iniciado. ---")
print(f"--- Pressione '{tecla_de_panico}' a qualquer momento para abortar. ---")

while not robo_deve_parar:
    print("-" * 30)

    #Envie seu prompt (pergunta) para o modelo.
    print("Digite sua pergunta para a IA:")
    prompt = input("Você: ")

    if robo_deve_parar == True:
        break

    #Caso a pessoa solicite para sair:
    if prompt.lower() == 'sair':
        print("Encerrando a conversa...")
        break
    
    try:
        #Envie o prompt para o modelo.
        #response = model.generate_content(prompt)
        response = chat.send_message(prompt)
        #Imprima a resposta.
        print("\nIA:")
        print(response.text)
    except Exception as e:
        print(f"Ocorreu um erro ao chamar a API: {e}")

print("--- Robô finalizado. ---")
sys.exit()