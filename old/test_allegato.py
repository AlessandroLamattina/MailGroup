# Apre il primo file in modalità lettura
with open('dizionario.txt', 'r',encoding='utf-8') as file1:
    # Legge i dati dal primo file e li carica in una lista
    data = file1.readlines()

# Modifica tutte le prime lettere di ciascun elemento nella lista
modified_data = [line.title() for line in data]
lowercase_data = [line.lower() for line in data]
upper_case = [line.upper() for line in data]
# Modifica i dati con la prima lettera di ogni parola in maiuscolo
titlecase_data = [line.title() for line in data]

# Apre il file in modalità "append"
with open('output.txt','a',encoding='utf-8') as file2:
    # Scrive i dati in minuscolo nel secondo file
    file2.writelines(lowercase_data)
    file2.writelines(titlecase_data)
    file2.writelines(upper_case)
