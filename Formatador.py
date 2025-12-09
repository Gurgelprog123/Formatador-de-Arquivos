
import os
import win32com.client

# Diretório atual 
diretorio_atual = os.getcwd()

# Inicializa o Word em modo invisível
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# Percorre todos os arquivos da pasta atual
for arquivo in os.listdir(diretorio_atual):
    if arquivo.lower().endswith(".docx"):  # Ignora tudo que não for .docx
        caminho_docx = os.path.join(diretorio_atual, arquivo)
        caminho_pdf = os.path.join(diretorio_atual, arquivo.replace(".docx", ".pdf"))

        # Abre o documento no Word
        doc = word.Documents.Open(caminho_docx)

        # Salva como PDF (FileFormat=17)
        doc.SaveAs(caminho_pdf, FileFormat=17)
        doc.Close()
    if arquivo.lower().endswith(".doc"):  # Ignora tudo que não for .docx
        caminho_docx = os.path.join(diretorio_atual, arquivo)
        caminho_pdf = os.path.join(diretorio_atual, arquivo.replace(".doc", ".pdf"))

        # Abre o documento no Word
        doc = word.Documents.Open(caminho_docx)

        # Salva como PDF (FileFormat=17)
        doc.SaveAs(caminho_pdf, FileFormat=17)
        doc.Close()

# Fecha o Word
word.Quit()

