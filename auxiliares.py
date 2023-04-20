"""
Todas as funções de suporte
"""

import os
import sys
import time


def caminhoprojeto(subpasta=''):
    """
    :param subpasta: adiciona o caminho da subpasta dada como entrada (caso preenchido).
    :return: o caminho do projeto ao qual a rotina está inserida.
    """
    import errno

    try:
        if getattr(sys, 'frozen', False):
            caminho = os.path.dirname(sys.executable)
        else:
            caminho = os.path.dirname(os.path.abspath(__file__))

        if len(subpasta) > 0:
            if os.path.isdir(caminho + '\\' + subpasta):
                return caminho + '\\' + subpasta
            else:
                os.mkdir(caminho + '\\' + subpasta)

            if os.path.isdir(caminho + '\\' + subpasta):
                return caminho + '\\' + subpasta
            else:
                return ''
        else:
            if os.path.isdir(caminho):
                return caminho
            else:
                return ''

    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass


def ultimoarquivo(caminho, extensao):
    """
    :param caminho: diretório onde pesquisa será realizada.
    :param extensao: extensão do arquivo que está sendo buscado o último alterado.
    :return: retorna o caminho completo do último arquivo atualizado da extensão e caminho dado como entrada.
    """
    lista_arquivos = os.listdir(caminho)
    ultimadata = 0
    ultimoatualizado = ''
    for arquivo in lista_arquivos:
        if extensao.upper() in arquivo.upper():
            if os.path.getmtime(caminho + '/' + arquivo) > ultimadata:
                ultimoatualizado = caminho + '/' + arquivo
                ultimadata = os.path.getmtime(caminho + '/' + arquivo)

    return ultimoatualizado


def renomeararquivo(nomeantigo, novonome):
    """
    :param nomeantigo: nome do arquivo a ser renomeado (endereço completo).
    :param novonome: nome novo do arquivo (endereço completo).
    """
    if os.path.isfile(to_raw(novonome)):
        os.remove(to_raw(novonome))
    time.sleep(0.5)
    os.rename(to_raw(nomeantigo), to_raw(novonome))


def to_raw(string):
    """
    :param string: texto a ser tratado.
    :return: string com o prefixo f para criar strings literais formatadas e r usado para tornar a string numa
    string literal bruta e ignorar caracteres especiais 'dentro' dela como o '\', por exemplo.
    """
    return fr"{string}"


def retornorvalorlista(lista, indice):
    """

    :param lista: lista a ser realizada a busca.
    :param indice: item a ser retornado.
    :return: item do índice dado como entrada.
    """
    import numpy as np

    lst = np.array(lista)
    return lst[indice]


def escreverlistaexcelog(caminho, lista):
    """

    :param caminho: caminho do arquivo a ser escrito.
    :param lista: lista a ser adicionada no arquivo do caminho dado.
    """
    import pandas as pd

    df = pd.DataFrame(lista)

    writer = pd.ExcelWriter(caminho, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Lista', index=False)
    writer.save()


def acertardataatual():
    """

    :return: retorna data e hora do sistema formatado.
    """
    from datetime import datetime

    textodata = datetime.now()
    return str(textodata.strftime('%Y_%m_%d_%H_%M_%S'))


def caminhospadroes(caminho):
    """

    :param caminho: código dos caminhos padrões (dúvidas, confira lista abaixo).
    :return: retorna o caminho padrão selecionado (str)
    """
    import ctypes.wintypes

    # CSIDL	                        Decimal	Hex	    Shell	Description
    # CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools for an individual user.
    # CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group.
    # CSIDL_APPDATA	                26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data.
    # CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder containing the objects in the user's Recycle Bin.
    # CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to CD.
    # CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer.
    # CSIDL_COMMON_ALTSTARTUP	    30	    0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users.
    # CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The file system directory containing application data for all users.
    # CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and folders that appear on the desktop for all users.
    # CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users.
    # CSIDL_COMMON_MUSIC	        53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users.
    # CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system directory that serves as a repository for image files common to all users.
    # CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that contains the programs and folders that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains the programs that appear in the Startup folder for all users.
    # CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are available to all users.
    # CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup.
    # CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder representing Network Connections, containing network and dial-up connections.
    # CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel applications.
    # CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies.
    # CSIDL_DESKTOP	                0	    0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace.
    # CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to physically store file objects on the desktop.
    # CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    # CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a common repository for the user's favorite items.
    # CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts.
    # CSIDL_HISTORY	                34	    0x22	 	    The file system directory that serves as a common repository for Internet history items.
    # CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files.
    # CSIDL_LOCAL_APPDATA	        28	    0x1C	5.0	    The file system directory that serves as a data repository for local (nonroaming) applications.
    # CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder representing the My Documents desktop item.
    # CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files.
    # CSIDL_MYVIDEO	                14	    0x0E	6.0	    The file system directory that serves as a common repository for video files.
    # CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may exist in the My Network Places virtual folder.
    # CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
    # CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards, CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.)
    # CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums.
    # CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing installed printers.
    # CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder.
    # CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder.
    # CSIDL_PROGRAM_FILESX86	    42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems.
    # CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across applications.
    # CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems.
    # CSIDL_PROGRAMS	            2	    0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories).
    # CSIDL_RECENT	                8	    0x08	 	    The file system directory that contains shortcuts to the user's most recently used documents.
    # CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data.
    # CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system directory that contains sample music.
    # CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists.
    # CSIDL_SAMPLE_PICTURES	        66	    0x42	Vista	The file system directory that contains sample pictures.
    # CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items.
    # CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory containing Start menu items.
    # CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group.
    # CSIDL_SYSTEM	                37	    0x25	5.0	    The Windows System folder.
    # CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems.
    # CSIDL_TEMPLATES	            21	    0x15	 	    The file system directory that serves as a common repository for document templates.
    # CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

    if caminho != 80:
        csidl_personal = caminho  # Caminho padrão
        shgfp_type_current = 0  # Para não pegar a pasta padrão e sim a definida como documentos

        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.Shell32.SHGetFolderPathW(None, csidl_personal, None, shgfp_type_current, buf)

        return buf.value
    else:
        if os.name == 'nt':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
            return location
        else:
            return os.path.join(os.path.expanduser('~'), 'Downloads')


def caminhoselecionado(tipojanela=1, titulojanela='Selecione o caminho/arquivo:',
                       tipoarquivos=('Todos os Arquivos', '*.*'), caminhoini=caminhospadroes(5), arquivoinicial=''):
    """

    :param tipojanela: 1 — Seleciona Arquivo (Padrão); 2 — Seleciona caminho para salvar arquivo; 3 — Seleciona diretório.
    :param titulojanela: cabeçalho exibido na caixa de diálogo.
    :param tipoarquivos: extensão dos arquivos permitidos da seleção.
    :param caminhoini: caminho inicial.
    :param arquivoinicial: arquivo inicial.
    :return:
    """
    import tkinter as tk
    from tkinter import filedialog

    'Cria a janela raiz'
    root = tk.Tk()
    root.withdraw()

    if tipojanela == 1:
        retorno = filedialog.askopenfilename(title=titulojanela,
                                             initialdir=caminhoini, filetypes=tipoarquivos, initialfile=arquivoinicial)
        if retorno is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None

    elif tipojanela == 2:
        name = filedialog.asksaveasfile(mode='w', defaultextension='.txt',
                                        filetypes=tipoarquivos,
                                        initialdir=caminhoini,
                                        title=titulojanela, initialfile=arquivoinicial)
        if name is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None
        text2save = str(name.name)
        name.write('')
        retorno = text2save

    elif tipojanela == 3:
        name = filedialog.askdirectory(initialdir=caminhoini, title=titulojanela)
        if name is None:  # askdirectory return `None` if dialog closed with "cancel".
            return None
        text2save = name
        retorno = text2save

    else:
        return

    return retorno


def criarinputbox(titulo, mensagem, substituircaracter='', valorinicial=''):
    """
    :param valorinicial: valor pré-preenchido na caixa de texto.
    :param titulo: cabeçalho da caixa de recebimento de dados do usuário (inputbox).
    :param mensagem: mensagem (normalmente descritiva ao 'input') para orientar o usuário.
    :param substituircaracter: caso seja um campo de senha informar o parâmetro para que a digitação não fique visível.
    :return: janela com as opções escolhidas.
    """
    import tkinter as tk
    from tkinter import simpledialog

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    # the input dialog
    user_inp = simpledialog.askstring(title=titulo, prompt=mensagem, initialvalue=valorinicial, show=substituircaracter)
    if user_inp is None:
        user_inp = 0

    root.attributes("-topmost", True)

    return user_inp


def left(s, amount):
    """

    :param s: string.
    :param amount: quantidade de caracteres.
    :return: retorna à esquerda de uma string
    """
    return s[:amount]


def right(s, amount):
    """

    :param s: string.
    :param amount: quantidade de caracteres.
    :return: retorna à direita de uma string
    """
    return s[-amount:]


def mid(s, offset, amount):
    """

    :param s: string.
    :param offset: início do "corte" da string.
    :param amount: quantidade de caracteres.
    :return:
    """
    return s[(offset - 1):(offset - 1) + amount]


def adicionarcabecalhopdf(arquivo, arquivodestino, cabecalho, reportlab=None):
    """
    :param arquivo: arquivo PDF de "entrada".
    :param arquivodestino: arquivo PDF de saída (já com o cabeçalho).
    :param cabecalho: texto do cabeçalho a ser adicionado.
    :return:
    """
    import PyPDF2
    import io
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics

    packet = io.BytesIO()
    pdfmetrics.registerFont(TTFont('Arial', 'arial-bold.ttf'))
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Arial', 10)
    can.drawString(500, 820, cabecalho)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    new_pdf = PyPDF2.PdfFileReader(packet)
    # read your existing PDF
    with open(arquivo, 'rb') as p:
        existing_pdf = (p.readlines())
    arquivopdf = reset_eof_of_pdf_return_stream(existing_pdf)
    arquivoacertado = mid(arquivo, 1, len(arquivo) - 4) + "_acertado" + right(arquivo, 4)
    if os.path.isfile(arquivoacertado):
        os.remove(arquivoacertado)
    # write to new pdf
    with open(arquivoacertado, 'wb') as f:
        f.writelines(arquivopdf)

    if os.path.isfile(arquivoacertado) and os.path.isfile(arquivo):
        os.remove(arquivo)

    renomeararquivo(arquivoacertado, arquivo)

    if os.path.isfile(arquivoacertado):
        os.remove(arquivoacertado)

    existing_pdf = PyPDF2.PdfFileReader(arquivo)

    output = PyPDF2.PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    paginas = existing_pdf.getNumPages()
    for pagina in range(paginas):
        page = existing_pdf.getPage(pagina)
        if pagina == 0:
            page.mergePage(new_pdf.getPage(pagina))
        output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open(arquivodestino, "wb")
    output.write(outputStream)
    outputStream.close()
    if os.path.isfile(arquivo):
        os.remove(arquivo)


def reset_eof_of_pdf_return_stream(pdf_stream_in: list):
    # find the line position of the EOF
    for i, x in enumerate(pdf_stream_in[::-1]):
        if b'%%EOF' in x:
            actual_line = len(pdf_stream_in) - i
            # print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
            break

    # return the list up to that point
    return pdf_stream_in[:actual_line]


def extrairtextopdf(caminho):
    import PyPDF2
    import re
    import pandas as pd

    texto = ''
    reader = PyPDF2.PdfFileReader(caminho)
    for pagina in range(reader.getNumPages()):
        p = reader.getPage(pagina)
        texto += p.extractText()

    lista = re.findall(r'([\d].[\d]{3}.[\d]{3}-[\d])[\d]{2}', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            linha = linha.replace('.', '')
            linha = linha.replace('-', '')
            linha = linha.replace('\n', '')
            listalimpa.append(linha)
    df = pd.DataFrame(listalimpa, columns=['Inscricao'])

    lista = re.findall(r'COMPETÊNCIA([\d]{4})[\d]{2}', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(linha.replace('\n', ''))
    df['Competencia'] = listalimpa

    lista = re.findall(r'([\d]{2}/[\d]{2}/[\d]{4})[\d]{2}', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(linha.replace('\n', ''))
    df['Vencimentos'] = listalimpa

    lista = re.findall(r'CONTRIBUINTE([\D]*)[\d]{2}.', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(linha.replace('\n', ''))
    df['Contribuinte'] = listalimpa

    lista = re.findall(r'AUTENTICAÇÃO AUTOMÁTICA[\D]PARA USO DO BANCO[\D]([\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d])[\D]', texto)
    df['Codigo de Barras'] = lista

    lista = re.findall(r'GUIA/COTA([\d]{2}/[\d]{2})', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(linha.replace('\n', ''))
    df['Guia'] = listalimpa

    lista = re.findall(r'VALOR TOTAL([\d]*,[\d]{2})[\d]{2}.', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(linha)
    df['Valor Total'] = listalimpa

    lista = re.findall(r'GUIA/COTA[\d]{2}/([\d]{2})', texto)
    listalimpa = []
    for indice, linha in enumerate(lista):
        if indice % 2:
            listalimpa.append(str(int(linha)))
    df['Parcela'] = listalimpa

    return df


def hora(timezone, pedaco=''):
    import requests
    import datetime

    url = 'http://worldtimeapi.org/api/timezone/' + timezone
    resposta = requests.get(url)
    try:
        match pedaco.upper():
            case 'DATA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).date()

            case 'HORA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).time()

            case _:
                return datetime.datetime.fromisoformat(resposta.json()['datetime'])

    except Exception:
        match pedaco.upper():
            case 'DATA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).date()

            case 'HORA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).time()

            case _:
                return datetime.datetime.fromisoformat(resposta.json()['datetime'])


def timezones_disponiveis():
    import requests

    url = 'http://worldtimeapi.org/api/timezone/'
    resposta = requests.get(url)
    timezones = resposta.json()
    return timezones


def criarpasta(caminho, cliente):
    caminhofinal = os.path.join(caminho, cliente)
    if not os.path.isdir(caminhofinal):
        os.mkdir(caminhofinal)


def escreverlistaexcelog(caminho, lista):
    """
    :param caminho: caminho do arquivo a ser escrito.
    :param lista: lista a ser adicionada no arquivo do caminho dado.
    """
    import pandas as pd

    df = pd.DataFrame(lista)

    writer = pd.ExcelWriter(caminho, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Lista', index=False)
    writer._save()


def acertardataatual():
    """
    :return: retorna data e hora do sistema formatado.
    """
    from datetime import datetime

    textodata = datetime.now()
    return textodata.strftime('%Y_%m_%d_%H_%M_%S')