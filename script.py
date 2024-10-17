import pandas as pd

# Carica il file Excel
df = pd.read_excel(r"C:\Users\lbertellini\Downloads\matricole.xlsx")

# Crea una colonna "normalizzata" rimuovendo spazi, punti e altri caratteri non numerici
df['Matricola_normalizzata'] = df['MATRICOLA'].astype(str).str.replace(r'\W+', '', regex=True)

# Trova i duplicati sulla base della colonna normalizzata
duplicati = df[df.duplicated('Matricola_normalizzata', keep=False)]

# Visualizza i risultati
print(duplicati)

# Puoi anche salvare i duplicati in un file Excel per ulteriori analisi
duplicati.to_excel('duplicati_matricola.xlsx', index=False)