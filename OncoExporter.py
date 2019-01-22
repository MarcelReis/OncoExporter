def printm(msg, *args):
    if "b" in args:
        print("#"*64)
        print(addSpace(msg).center(64, "#"))
        if "e" in args:
            print("#"*64, end="")
        else:
            print("#"*64)
    if "l" in args:
        print("#"+addSpace(msg).ljust(62, " ")+"#")
    if "w" in args:
        input("")
    return

def addSpace(sentence):
    return " "+sentence+" "

printm("BEM VINDO AO ONCOEXPORTER V0.9", "b")
printm("IMPORTANDO BIBLIOTECAS", "l")
import pyodbc
import pandas as pd

printm("ESTABELECENDO CONEXÃO", "l")
conn_str = (
    r'DRIVER=Microsoft Access Driver (*.mdb);'
    r'DBQ=O:\Banco\BancoNovo.mdb;' #Especifica caminho onde se encontra banco de dados do Oncosis, precisa estar mapeado em alguma unidade
    r'PWD=arthur@01'
    )
cnxn = pyodbc.connect(conn_str)

sql_query = '''
SELECT Prontuario, Nome, DataCadastro, Procedencia, Sexo, Idade,
DataNascimento, Convenio, Cor, Profissao, Endereco, EnderecoNum, EnderecoComp,
Bairro, Telefone, Cidade, UF, Cep, Responsavel, Parentesco, EnderecoResp,
TelefoneResp, MedicoRecepcao, TelefoneMedico, EnderecoMedico, CidadeMedico,
EstadoCivil, Natural, Nacionalidade, Tratamento, CPF, Mae, NaturalUF,
TipoLogradouro, CodigoIbge, CodigoIbgeNatural, Identidade,
cartaosus, DataExpedicao FROM Cliente
'''

printm("RECUPERANDO CADASTROS", "l")
df = pd.read_sql_query(sql_query, cnxn)

printm("SALVANDO O ARQUIVO", "l")
df.to_csv('OncoExporter_Register.csv') #Salva o arquivo como csv simples

printm("", "l")
printm("CONCLUIDO COM SUCESSO! APERTE ENTER PARA FECHAR", "b", "w", "e")
