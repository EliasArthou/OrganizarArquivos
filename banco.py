import os
import pypyodbc as pyodbc
import sensiveis as senha


class Banco:
    """
    Criado para se conectar e realizar operações no banco de dados
    """

    def __init__(self, caminho):
        self.conxn = None
        self.cursor = None
        self.constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + caminho + ';Pwd=' + senha.senhabanco
        self.abrirconexao()

    def abrirconexao(self):
        if len(self.constr) > 0:
            self.conxn = pyodbc.connect(self.constr)
            self.cursor = self.conxn.cursor()

    def consultar(self, sql):
        """
        :param sql: Código sql a ser executado (uma consulta SQL).
        :return: O resultado da consulta em uma lista.
        """
        self.cursor.execute(sql)
        resultado = [item[0] for item in self.cursor.fetchall()]
        return resultado

    def adicionardf(self, tabela, df, indicelimpeza=-1):
        for linha in df.values:
            my_list = [str(x) for x in linha]
            if len(my_list) > 0:
                if indicelimpeza != -1:
                    self.abrirconexao()
                    strSQL = "DELETE * FROM [%s] WHERE Barras = '%s'" % (tabela, my_list[indicelimpeza])
                    self.cursor.execute(strSQL)
                    self.conxn.commit()

                strSQL = "INSERT INTO %s VALUES [" % tabela + "] ('" + "', '".join(my_list) + "')"
                self.cursor.execute(strSQL)
                self.conxn.commit()

            print("('" + "', '".join(my_list) + "')")
        # df.to_csv('df.csv', sep=';', encoding='utf-8', index=False)

        # RUN QUERY
        strSQL = "INSERT INTO %s SELECT * FROM [text;HDR=Yes;FMT=Delimited(;);Database=D:\Projetos\Extrair Imposto\].df.csv" % tabela

        self.cursor.execute(strSQL)
        self.conxn.commit()

        self.conxn.close()  # CLOSE CONNECTION
        os.remove('df.csv')

    def fecharbanco(self):
        """
        Fecha a conexão com o banco de dados
        """
        self.cursor.close()
