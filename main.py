import sensiveis as senha
import os
import auxiliares as aux
import messagebox
import sys
import banco as bd

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

if len(caminhoacriar) == 0:
    messagebox.msgbox('Selecione a pasta de destino!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

data = bd.Banco(caminhobanco)

resultado = data.consultar(senha.sqlclientes)

data.fecharbanco()

for cliente in resultado:
    aux.criarpasta(caminhoacriar, cliente[0])
