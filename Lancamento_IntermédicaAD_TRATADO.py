import time
import PyPDF2 as pyf
import shutil
import dateutil.parser
import pyautogui
import keyboard
from pathlib import Path
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def entrar_lancamento():
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txfUsuario"]'))).clear()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txfUsuario"]'))).send_keys('****')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txfSenha"]'))).send_keys('****')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen19"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen20"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="SCP"]/div'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen901"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen941"]'))).click()

def carregando():
    wait_menor.until(EC.visibility_of_element_located((By.XPATH, '//*[text()="Carregando..."]')))
    wait.until(EC.invisibility_of_element_located((By.XPATH, '//*[text()="Carregando..."]')))

def reaproveitar_titulo():
    loc_titulo.clear()
    loc_titulo.send_keys(n_reaproveitar, Keys.TAB)
    carregando()

def novo_titulo():
    loc_titulo.clear()
    loc_titulo.send_keys(n_titulo, Keys.TAB)
    carregando()

def qual_xpath_sim():
    try:
        wait_menor.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-comp-1297"]/tbody/tr[2]/td[2]')))
        xpath_sim = '//*[@id="ext-comp-1297"]/tbody/tr[2]/td[2]'
    except:
        xpath_sim = '//*[@id="ext-comp-1305"]/tbody/tr[2]/td[2]'
    return xpath_sim

def preencher_apropriacao():
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen330"]/div/table/tbody/tr/td[1]'))).click()
    wait.until(EC.presence_of_element_located((By.ID, 'ext-gen873'))).click()

info_boletos = (input('Informe o caminho dos Boletos: '))
info_boletos = info_boletos.replace('\\', '/')

navegador = Chrome()
navegador.get('https://****')
navegador.maximize_window()

wait = WebDriverWait(navegador, 60, poll_frequency=0.5)
wait_menor = WebDriverWait(navegador, 1, poll_frequency=0.5)

entrar_lancamento()

iframe = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="2294_IFrame"]')))
navegador.switch_to.frame(iframe)

caminho_boletos = Path('{}'.format(info_boletos))

nome_pasta = 'BOLETOS LANCADOS'
Path('{}/{}'.format(caminho_boletos, nome_pasta)).mkdir()
caminhoBol_lancados = Path('{}/{}'.format(caminho_boletos, nome_pasta))

# Criar lista com todos os boletos:
boletos = caminho_boletos.iterdir()

cnpj = '****'
n_reaproveitar = '****'

for boleto in boletos:
    nome_arquivo = boleto.name
    path_caminhoBol = Path('{}/{}'.format(caminho_boletos, nome_arquivo))

    # Transforma o caminho em string para a biblioteca 'keyboard' reconhecer o que será escrito:
    caminho_bol = "{}".format(path_caminhoBol)

    # Se o arquivo for um pdf, irá executar o comando:
    if nome_arquivo[-3:] == 'pdf':
        # Abrir o pdf:
        boleto_abrir = pyf.PdfFileReader('{}/{}'.format(caminho_boletos, nome_arquivo))
        # Acessa a primeira página do pdf onde está o boleto:
        pagina = boleto_abrir.pages[0]
        # Extrai o texto da pagina:
        info_boleto = pagina.extractText()
        # Armazena o texto na variável:
        texto_final = info_boleto

        # Valor do boleto:
        pi_valor = texto_final.find('(+) Mora / Multa')
        pf_valor = texto_final.find('Data do DocumentoData')
        valor = texto_final[pi_valor + 16 + 2:pf_valor]

        # Código de barras do boleto:
        pi_codbarr = texto_final.find('(+) Outros Acréscimos')
        cod_barras = texto_final[pi_codbarr + 21:pi_codbarr + 75]
        # Apenas números no código de barras:
        cod_barras = cod_barras.replace('.', '').replace(' ', '')

        # Número do título:
        pi_tit = texto_final.find('DocumentoBeneficiário')
        pf_tit = texto_final.find('/', pi_tit)
        n_titulo = texto_final[pi_tit + 21:pf_tit - 2]

        # Vencimento:
        pi_venc = texto_final.find('DocumentoBeneficiário')
        pf_venc = texto_final.find('N(=)', pi_venc)
        vencimento_bol = texto_final[pi_venc + 29:pf_venc]

        # Fazer o parse obter um objeto datetime
        datetime_object = dateutil.parser.parse(vencimento_bol, dayfirst=True)

        # Formatar o datetime para o formato desejado (dia/mês/ano)
        vencimento = datetime_object.strftime("%d%m%y")  # "y": ano abreviado / "Y": ano completo.

        # Emissão:
        pf_emissao = texto_final.find('Espécie do Documento')
        pi_emissao = pf_emissao - 10

        dt_emissao = texto_final[pi_emissao:pf_emissao]

        emissao_object = dateutil.parser.parse(dt_emissao, dayfirst=True)
        emissao = emissao_object.strftime("%d%m%y")

        # Elementos do formulário:
        loc_fornecedor = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="hpcClienteFornecedor"]')))
        loc_titulo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="hpfTitulo"]')))
        loc_docfiscal = wait.until(EC.element_to_be_clickable((By.ID, 'txfDocumentoFiscal')))
        loc_dataemissao = wait.until(EC.element_to_be_clickable((By.ID, 'mdfDataEmissao')))
        loc_dataentrada = wait.until(EC.element_to_be_clickable((By.ID, 'mdfEntrada')))
        loc_vencimento = wait.until(EC.element_to_be_clickable((By.ID, 'mdfVencimento')))
        loc_campovalor = wait.until(EC.element_to_be_clickable((By.ID, 'mnfValor')))
        loc_cgdbarras = wait.until(EC.element_to_be_clickable((By.ID, 'txfIdentificacao')))
        loc_anexo = wait.until(EC.element_to_be_clickable((By.ID, 'ext-gen33')))

        # Inserir Fornecedor:
        loc_fornecedor.send_keys(cnpj, Keys.TAB)

        a = False
        while not a:
            reaproveitar_titulo()
            novo_titulo()
            xpath_sim = qual_xpath_sim()
            try:
                wait_menor.until(EC.element_to_be_clickable((By.XPATH, '{}'.format(xpath_sim)))).click()
                carregando()
                a = True
            except:
                continue

        # PREENCHER LANÇAMENTO:
        loc_docfiscal.clear()
        time.sleep(1)
        if n_reaproveitar == '20865420':
            loc_dataemissao.clear()
            loc_dataemissao.send_keys(emissao)
            loc_dataentrada.clear()
            loc_dataentrada.send_keys(emissao, Keys.TAB)
            carregando()
            time.sleep(1)
            loc_vencimento.clear()
            loc_vencimento.send_keys(vencimento, Keys.TAB)
            carregando()
        loc_campovalor.clear()
        loc_campovalor.send_keys(valor, Keys.TAB)
        carregando()
        time.sleep(1)
        try:
            preencher_apropriacao()
        except:
            time.sleep(2)
            preencher_apropriacao()
        time.sleep(1)
        loc_cgdbarras.clear()
        loc_cgdbarras.send_keys(cod_barras)

        # UPLOAD DO ARQUIVO NO SISTEMA:
        # Abrir tela de anexo:
        loc_anexo.click()
        try:
            carregando()
        except:
            pass
        iframe_anexo = wait.until(EC.element_to_be_clickable((By.ID, 'PanelAnexosDaTela_IFrame')))
        navegador.switch_to.frame(iframe_anexo)
        # Clicar em "Escolher arquivos":
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cntArquivoUpload_Content"]'))).click()
        # Aguarda tela para selecionar arquivo ficar disponível:
        while not pyautogui.locateOnScreen('Google_Abrir.PNG'):
            time.sleep(0.5)
        # Informar qual o arquivo no diretório para upload:
        keyboard.write(caminho_bol)
        keyboard.press('enter')
        time.sleep(1.5)
        # Clicar em "Inserir":
        wait.until(EC.element_to_be_clickable((By.ID, 'ext-gen111'))).click()
        navegador.switch_to.default_content()
        navegador._switch_to.frame(iframe)
        try:
            carregando()
        except:
            pass
        # Fechar a tela de anexo:
        wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ButtonFecharAnexosTela"]/tbody/tr[2]/td[2]'))).click()
        time.sleep(1.5)

        # Salvar o lançamento:
        wait.until(EC.element_to_be_clickable((By.ID, 'ext-gen29'))).click()  # salvar
        carregando()

        # Caso apareça alguma crítica, irá interromper o lançamento:
        try:
            wait_menor.until(EC.visibility_of_element_located((By.XPATH, '//*[text()="Lista de mensagens"]')))
            print('Sistema apresentou crítica ao salvar o lançamento.\n'
                  'Favor verificar.')
            break
        except:
            pass

        # Atualizar para o mais recente o modelo de reaproveitamento das informações:
        n_reaproveitar = n_titulo

        onde_colar = Path('{}/{}'.format(caminhoBol_lancados, nome_arquivo))
        shutil.move(caminho_boletos / nome_arquivo, onde_colar)