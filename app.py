"Descrever os passos manuais e depois transformar isso em código"
"Ler planilha e guardar informações sobre o nome, telefone e data de vencimento"
"Criar links personalizados do WhatsApp e enviar mensagens para cada cliente com base nos dados da planilha"
#Pressiona CTRL e J e digite "pip install openpyxl" Instalar a biblioteca, openpyxl para automatizar leitura de dados de uma planilha
# importar a biblioteca para ler a planilha

import openpyxl
#para abrir o navegador
import webbrowser
#para instalar o recurso de ler imagem
import pyautogui
#importara o recurso para aguardar o usuario logar na conta
from time import sleep
#para já abrir o navegador com o link do WhatsApp
webbrowser.open('https://web.whatsapp.com/')
#vai aguardar 30 segundos para logar na conta
sleep(30)
#quote permite formatar links para envios de mensagens
from urllib.parse import quote
#digitar o nome da planilha para ler as informações
workbook = openpyxl.load_workbook('alunos.xlsx')
#digite o nome da planilha
pagina_clientes = workbook['Planilha1']
#agora precisamos iterar, "ler as linhas da planilha para pegar as informações"
#passando em cada linha da planilha para pegar as informações, especificar para ler a partir da linha dois porque a primeira linha é o cabeçalho
for linha in pagina_clientes.iter_rows(min_row=2):
    #pausar o loop, quando a linha estiver vazia
    nome = linha [0].value
    if nome is None:
        break
    
#  criar as váriaveis para receber os dados, o indice no começa em zero "sendo zero a primeira célula, 1 a segunda celula da linha e 2 a terceira celula da linha"
    telefone = linha [1].value
    faltas = linha [2].value
    escola = linha[3].value
    nome1 = linha[4].value
    #usando o print para testar se esta conseguindo ler as informações
    print(nome)
    print(telefone)
    print(faltas)
    #clica no play para ver se as informações aparecem no terminal
    #criar os links personalizados no whatsapp
    #criar mensagem
    #{vencimento.strftime("%d/%m/%Y, %H:%M:%S")} para formatar a data no padrão brasileiro
    mensagem = (
        f'💡Olá, Somos da(o) *{escola}*'
        f'                                           '
        f'-------------------------------------------'
f'🔎Verificamos na *Lista de Chamada* que o(a) Estudante *{nome}* possui um *número preocupante* de faltas.'
        f'                                           '
        f'-------------------------------------------'
f'🚨Até a presente data, o(a) Estudante acumula *{faltas} Faltas*'
        f'                                            '
        f'--------------------------------------------'
f'💡Caso as faltas ocorram por *motivos de saúde*, orientamos a família a apresentar o *atestado* na instituição.'
        f'                                            '
        f'--------------------------------------------'
f'🔎Com o amparo do atestado, as faltas serão *abonadas* e assim evitará que o(a) aluno(a) atinja *a reprova*'
        f'                                            '
        f'--------------------------------------------'
f'🚨Conscientizamos que as faltas *prejudicam* o aprendizado e ocasionarão *a reprova*, caso a(o) *Estudante* acumule *50 Faltas*.'
        f'                                            '
        f'--------------------------------------------'
f'🔎Considerando que faltam menos de *2 meses* para finalizar o ano letivo.'
        f'                                             '
        f'------------------------------------------------------'
f'🚨Vamos orientar a família a incentivar a(o) aluna(o) *{nome1}* a frequentar as aulas e *evite* que o(a) Estudante *falte*'
        f'                                           '
        f'----------------------------------------------------'
f'💡Qualquer dúvida, pode retornar este contato. Estamos à disposição.'
)
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    #para abrir o navegador com link da mensagem
    #precisa estar com o Whatsapp logado
    webbrowser.open(link_mensagem_whatsapp)
    #dar uma pausa de 10 segundos para aguardar carregar o link e carregar a pagina
    sleep(10)
    #try serve para verificar se o seguinte codigo deu erro, e depois armazenar na planilha quais mensagens deram erro
    try:
        #no terminal digite: pip install pyautogui para instalar a automação de ler a imagem do botão
        #no terminal digita: pip install pillow
        #pausar a automação
        #pyautogui.locateCenterOnScreen('seta.png') para localizar a seta no centro da tela
        #encontra a seta e a varíavel seta ira armanezar as coordenadas
        seta = pyautogui.locateCenterOnScreen('seta.png')
        #apos encontrar a seta, dar uma pause 3 segundos
        sleep(5)
        #para clicar no botão da seta para enviar a mensagem
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        #após clicar na seta, dar uma pausa de 10 segundos
        #cada vez que executar o codigo, ele abrirá uma nova aba. dessa forma é necessário fechá-la apos enviar a mensagem
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
        UsarNessa = pyautogui.locateCenterOnScreen('UsarNessa.png')
        sleep(5)
        pyautogui.click(UsarNessa[0], UsarNessa[1])
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')

print("Concluído, pode fechar esta janela!")
