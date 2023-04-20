import sensiveis as senha
import os
import auxiliares as aux
import messagebox
import sys
import banco as bd
import shutil
import re


if os.path.isfile(aux.caminhoprojeto() + '/' + 'Scai.WMB'):
    caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                          tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                          caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
else:
    if os.path.isdir(aux.caminhoprojeto()):
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                              caminhoini=aux.caminhoprojeto())
    else:
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])


caminhoacriar = aux.caminhoselecionado(3)

if len(caminhobanco) == 0:
    messagebox.msgbox('Selecione o caminhob do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

# diretório de origem
caminhoorigem = aux.caminhoselecionado(3, titulojanela='Pasta de Origem')

if len(caminhoorigem) == 0:
    messagebox.msgbox('Selecione a pasta de origem!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

if len(caminhoacriar) == 0:
    messagebox.msgbox('Selecione a pasta de destino!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

data = bd.Banco(caminhobanco)

resultado = data.consultar(senha.sqlclientes)

data.fecharbanco()

for cliente in resultado:
    aux.criarpasta(caminhoacriar, cliente[0])

# dicionário com pares de pastas para cada tipo de arquivo
pastas = {
    "ADITIVOS ADMINISTRAÇÃO": "Aditivo",
    "ADITIVOS LOCAÇÃO": "Aditivo",
    "CONTRATOS ADMINISTRAÇÃO": "Contrato",
    "CONTRATOS LOCAÇÃO": "ContratoLoc"
}

# loop nos arquivos da pasta de origem
for pasta, subpastas, arquivos in os.walk(caminhoorigem):
    for arquivo in arquivos:
        # extrair o código a partir do nome do arquivo usando expressão regular
        codigo = re.findall(r'\d+', arquivo)[0].zfill(4)

        # determinar o tipo de arquivo a partir do nome da pasta atual
        tipo_arquivo = None
        for chave in pastas.keys():
            if chave in pasta:
                tipo_arquivo = chave
                break

        # se não encontrou o tipo de arquivo, pular para o próximo arquivo
        if not tipo_arquivo:
            continue

        # determinar o diretório de destino com base no código e no tipo de arquivo
        pasta_destino = os.path.join(caminhoacriar, codigo[:4])
        if not os.path.exists(pasta_destino):
            os.mkdir(pasta_destino)

        # pasta_destino = os.path.join(pasta_destino, pastas[tipo_arquivo])

        # copiar o arquivo para o diretório de destino com o nome padronizado
        nome_padronizado = f"{pastas[tipo_arquivo]}{codigo}.pdf"
        shutil.copy2(os.path.join(pasta, arquivo), os.path.join(pasta_destino, nome_padronizado))
