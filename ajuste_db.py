import sqlite3

conn = sqlite3.connect("producao_calcados.db")
c = conn.cursor()
try:
    c.execute("ALTER TABLE ops ADD COLUMN tipo TEXT NOT NULL DEFAULT 'Masculino'")
    print("Coluna 'tipo' adicionada com sucesso!")
except Exception as e:
    print("Erro ou coluna já existe:", e)
conn.commit()
conn.close()