import psycopg2


# Conexão com o banco de dados PostgreSQL
conn = psycopg2.connect(
    host="localhost",
    port=5433,
    database="Multiverso",
    user="postgres",
    password="Multiverso@Educa"
)

# Cursor
cur = conn.cursor()

################# Traking ###############################


# Se a tabela não existe, cria
cur.execute('''CREATE TABLE Traking_Integracao (
               ATIVIDADE INTEGER,
               PROCESSO VARCHAR(5000),
               CHECKLIST_DE_INTEGRAÇÃO VARCHAR(5000),
               NÍVEL_DE_CRITICIDADE VARCHAR(5000),
               OBJETIVOS_ESPERADOS_COM_ESSA_ATIVIDADE VARCHAR(5000),
               STATUS_DE_REALIZAÇÃO VARCHAR(5000),
               ÁREA_RESPONSÁVEL VARCHAR(5000),
               DIA_D CHARACTER(10),
               TIMING_DIAS VARCHAR(5000),
               FINANCEIRO VARCHAR(5000),
               PEDAGÓGICO VARCHAR(5000),
               RH VARCHAR(5000),
               JURÍDICO VARCHAR(5000),
               MARKETING VARCHAR(5000),
               ATENDIMENTO VARCHAR(5000),
               DADOS VARCHAR(5000),
               OPERACIONAL VARCHAR(5000),
               TI VARCHAR(5000),
               EXPANSÃO VARCHAR(5000),
               MA VARCHAR(5000),
               FORNECEDOR VARCHAR(5000),
               OBSERVAÇÕES_INFORMAÇÕES_EXTRAS VARCHAR(5000),
               Data DATE,
               Setor VARCHAR(5000),
               Area VARCHAR(5000),
               Macro VARCHAR(5000),
               DIA VARCHAR(5000),
               FRENTE VARCHAR(5000)
               );''')

# Commit e fechamento da conexão
conn.commit()
cur.close()
conn.close()