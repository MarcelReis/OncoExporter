#Beta Version

print("# IMPORTANDO BIBLIOTECAS")
import pyodbc
import pandas as pd

print("# ESTABELECENDO CONEXÃO")
conn_str = (
    r'DRIVER=Microsoft Access Driver (*.mdb);'
    r'DBQ=O:\Banco\BancoNovo.mdb;' #Especifica caminho onde se encontra banco de dados do Oncosis, precisa estar mapeado em alguma unidade
    r'PWD=arthur@01'
    )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()

sql_query = '''
SELECT Prontuario, Nome, DataCadastro, Documento, Procedencia, Sexo, Idade, DataNascimento, Convenio, Cor, Profissao, Endereco, EnderecoNum, EnderecoComp, Bairro, Telefone, Cidade, UF, Cep, Responsavel, Parentesco, EnderecoResp, TelefoneResp, MedicoRecepcao, TelefoneMedico, EnderecoMedico, CidadeMedico, EstadoCivil, Natural, Nacionalidade, Tratamento, CPF, Mae, NaturalUF, DataDiagnostico, TipoLogradouro, CodigoIbge, CodigoIbgeNatural, Identidade, cartaosus, DataExpedicao FROM Cliente 
'''

print("# RECUPERANDO CADASTROS")
df = pd.read_sql_query(sql_query, cnxn)

print("# SALVANDO O ARQUIVO")
df.to_csv('OncoExporter_Register.csv') #Salva o arquivo como csv simples

input("\n\n# CONCLUIDO COM SUCESSO! APERTE ENTER PARA FECHAR")