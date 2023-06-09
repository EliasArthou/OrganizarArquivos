import sensiveis as senha
import os
import auxiliares as aux
import messagebox
import sys
import banco as bd
import shutil
import re

listaexcel = []
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

if len(caminhobanco) == 0:
    messagebox.msgbox('Selecione o caminhob do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

# diretório de origem
caminhoorigem = aux.caminhoselecionado(3, titulojanela='Pasta de Origem')

if len(caminhoorigem) == 0:
    messagebox.msgbox('Selecione a pasta de origem!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

caminhoacriar = aux.caminhoselecionado(3, 'Caminho onde organizar')

if len(caminhoacriar) == 0:
    messagebox.msgbox('Selecione a pasta de destino!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

if not os.path.isdir(os.path.join(caminhoacriar, 'Ativo')):
    os.mkdir(os.path.join(caminhoacriar, 'Ativo'))

if not os.path.isdir(os.path.join(caminhoacriar, 'Deletados')):
    os.mkdir(os.path.join(caminhoacriar, 'Deletados'))

data = bd.Banco(caminhobanco)

resultado = data.consultar(senha.sqlclientes)

data.fecharbanco()

# dicionário com pares de pastas para cada tipo de arquivo
pastas = {
    "ADITIVOS ADMINISTRAÇÃO": "AditivoAdm",
    "ADITIVOS LOCAÇÃO": "AditivoLoc",
    "CONTRATOS ADMINISTRAÇÃO": "ContratoAdm",
    "CONTRATOS LOCAÇÃO": "ContratoLoc",
    "APOLICES": "Apolices",
    "CONTRATOS DELETADOS - ADMINISTRAÇÃO": "ContratoAdm",
    "CONTRATOS DELETADOS - LOCAÇÃO": "ContratoLoc"
}
# Cabeçalho do arquivo de Saída
listachaves = ['Origem', 'Destino']

# loop nos arquivos da pasta de origem
for pasta, subpastas, arquivos in os.walk(caminhoorigem):
    for arquivo in arquivos:
        # extrair o código a partir do nome do arquivo usando expressão regular
        # codigo = re.findall(r'\d+', arquivo)
        codigo = re.findall(r'\d{4,}', arquivo)
        origem = os.path.join(pasta, arquivo)
        destino = ''
        if isinstance(codigo, list) and len(codigo) > 0:
            codigo = codigo[0].zfill(4)

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
            if codigo[:4] in resultado:
                pasta_destino = os.path.join(caminhoacriar, 'Ativo', codigo[:4])
            else:
                pasta_destino = os.path.join(caminhoacriar, 'Deletados', codigo[:4])
            if not os.path.exists(pasta_destino):
                os.mkdir(pasta_destino)

            # copiar o arquivo para o diretório de destino com o nome padronizado
            nome_padronizado = f"{pastas[tipo_arquivo]}{codigo}.pdf"

            # verifique se já existe um arquivo com o mesmo nome na pasta de destino
            if os.path.exists(os.path.join(pasta_destino, nome_padronizado)):
                # se o arquivo já existir, adicione um número de sequência até encontrar um nome de arquivo que não existe
                i = 2
                while True:
                    novo_nome = f"{pastas[tipo_arquivo]}{codigo}_{i}.pdf"
                    if not os.path.exists(os.path.join(pasta_destino, novo_nome)):
                        nome_padronizado = novo_nome
                        break
                    i += 1

            destino = os.path.join(pasta_destino, nome_padronizado)
            shutil.move(os.path.join(pasta, arquivo), os.path.join(pasta_destino, nome_padronizado))
        dadosarquivos = [origem, destino]

        listaexcel.append(dict(zip(listachaves, dadosarquivos)))

aux.escreverlistaexcelog(os.path.join(caminhoacriar, 'Log_' + aux.acertardataatual() + '.xlsx'), listaexcel)
