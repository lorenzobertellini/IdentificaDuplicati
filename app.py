import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Funzione per selezionare il file e identificare i duplicati
def identifica_duplicati():
    # Chiede all'utente di selezionare un file Excel
    file_path = filedialog.askopenfilename(title="Seleziona il file Excel", filetypes=[("Excel files", "*.xlsx")])
    
    if file_path:
        try:
            # Carica il file Excel
            df = pd.read_excel(file_path)
            
            # Verifica che la colonna 'Matricola' esista
            if 'MATRICOLA' not in df.columns:
                messagebox.showerror("Errore", "La colonna 'MATRICOLA' non è presente nel file.")
                return
            
            # Normalizza i numeri di matricola rimuovendo spazi, punti e altri caratteri
            df['Matricola_normalizzata'] = df['MATRICOLA'].astype(str).str.replace(r'\W+', '', regex=True)
            
            # Trova i duplicati sulla base della colonna normalizzata
            duplicati = df[df.duplicated('Matricola_normalizzata', keep=False)]
            
            if duplicati.empty:
                messagebox.showinfo("Risultato", "Nessun duplicato trovato.")
            else:
                # Salva i duplicati in un nuovo file Excel
                output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Salva il file Excel con i duplicati")
                
                if output_path:
                    duplicati.to_excel(output_path, index=False)
                    messagebox.showinfo("Successo", f"File con duplicati salvato con successo in {output_path}")
        
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore: {str(e)}")

# Configurazione dell'interfaccia grafica
root = tk.Tk()
root.title("Identifica Duplicati Matricola")

# Bottone per eseguire il processo di identificazione duplicati
btn = tk.Button(root, text="Carica File Excel e Identifica Duplicati", command=identifica_duplicati)
btn.pack(pady=20)

# Esegue l'interfaccia
root.mainloop()