import os
import shutil
import tempfile
from flask import Flask, request, render_template, send_file, make_response
from PyPDF2 import PdfWriter
from PIL import Image
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Define o caminho do arquivo de fonte TrueType para a fonte Arial
font_path = "arial.ttf"
# Registra a fonte Arial no sistema de métricas de PDF do reportlab
pdfmetrics.registerFont(TTFont('Arial', font_path))

# Cria uma instância do objeto Flask para iniciar o aplicativo
app = Flask(__name__, template_folder=os.path.abspath('.'))

# Define uma rota para a página inicial que renderiza o template 'upload.html'
@app.route('/')
def index():
    return render_template('upload.html')

# Define uma rota para lidar com a requisição de upload de arquivos
@app.route('/upload', methods=['POST'])
def upload():
    # Verifica se a chave 'files[]' existe no dicionário de arquivos da requisição
    if 'files[]' not in request.files:
        return "Nenhum arquivo enviado!"

    # Obtém uma lista de arquivos enviados pelo usuário
    files = request.files.getlist('files[]')
    # Define uma lista com as extensões de arquivo suportadas
    supported_extensions = {'.png', '.jpg', '.jpeg', '.txt', '.docx'}
    # Cria uma lista para armazenar os caminhos dos arquivos PDF gerados
    pdf_files = []
    # Cria um diretório temporário usando o módulo tempfile
    temp_folder = tempfile.mkdtemp()

    # Itera por cada arquivo enviado pelo usuário
    for file in files:
        # Obtém o nome do arquivo e sua extensão
        filename, ext = os.path.splitext(file.filename)
        ext = ext.lower()

        # Verifica se a extensão do arquivo está na lista de extensões suportadas
        if ext not in supported_extensions:
            return f"Formato de arquivo não suportado: {ext}"

        # Trata diferentes extensões de arquivo
        if ext == '.pdf':
            pdf_files.append(file)
        elif ext in {'.png', '.jpg', '.jpeg'}:
            pdf_path = convert_image_to_pdf(file, temp_folder)
            pdf_files.append(pdf_path)
        elif ext == '.txt':
            pdf_path = convert_text_to_pdf(file, temp_folder)
            pdf_files.append(pdf_path)
        elif ext == '.docx':
            pdf_path = convert_docx_to_pdf(file, temp_folder)
            pdf_files.append(pdf_path)

    # Define o caminho para o arquivo ZIP comprimido
    compressed_file_path = os.path.join(temp_folder, 'compressed.rar')
    # Compacta os arquivos PDF em um arquivo ZIP usando a função compress_files_to_rar
    compress_files_to_rar(pdf_files, compressed_file_path)

    # Cria uma resposta HTTP com o arquivo ZIP como anexo
    response = make_response(send_file(compressed_file_path, as_attachment=True))
    # Define o nome do arquivo ZIP baixado pelo usuário
    response.headers["Content-Disposition"] = "attachment; filename=compressed.rar"

    return response

# Função para converter uma imagem para um arquivo PDF usando a biblioteca PIL
def convert_image_to_pdf(image_file, output_folder):
    # Salva a imagem temporariamente em uma pasta temporária
    image_path = os.path.join(output_folder, image_file.filename)
    image_file.save(image_path)

    # Define o caminho do arquivo PDF
    pdf_path = os.path.join(output_folder, f"{os.path.splitext(image_file.filename)[0]}.pdf")

    # Abre a imagem usando a biblioteca PIL e a converte para o formato PDF
    img = Image.open(image_path)
    img.convert('RGB').save(pdf_path, "PDF")

    return pdf_path

# Função para converter um arquivo de texto (TXT) para um arquivo PDF usando a biblioteca reportlab
def convert_text_to_pdf(txt_file, output_folder):
    # Salva o arquivo TXT temporariamente em uma pasta temporária
    txt_path = os.path.join(output_folder, txt_file.filename)
    txt_file.save(txt_path)

    # Define o caminho do arquivo PDF
    pdf_path = os.path.join(output_folder, f"{os.path.splitext(txt_file.filename)[0]}.pdf")

    # Define o tamanho da página como A4 (21 cm x 29.7 cm)
    page_width, page_height = A4

    # Lê o conteúdo do arquivo de texto
    with open(txt_path, 'r', encoding='utf-8') as txt:
        content = txt.read()

    # Cria um canvas (página) para o arquivo PDF
    pdf_canvas = canvas.Canvas(pdf_path, pagesize=A4)

    # Define as margens (72 pontos = 1 polegada)
    left_margin = 72
    right_margin = page_width - 72
    top_margin = page_height - 72
    bottom_margin = 72

    # Calcula a altura disponível para o texto em cada página
    available_height = top_margin - bottom_margin

    # Separa o conteúdo em linhas
    lines = content.splitlines()

    current_y = top_margin
    for line in lines:
        pdf_canvas.setFont("Helvetica", 12)
        text_width = pdf_canvas.stringWidth(line, "Helvetica", 12)

        # Verifica se a linha cabe na página
        if current_y - pdf_canvas._leading >= bottom_margin:
            pdf_canvas.drawString(left_margin, current_y, line)
            current_y -= pdf_canvas._leading
        else:
            # Cria uma nova página e recomeça a escrever do topo
            pdf_canvas.showPage()
            current_y = top_margin
            pdf_canvas.drawString(left_margin, current_y, line)
            current_y -= pdf_canvas._leading

    # Salva o arquivo PDF
    pdf_canvas.save()

    return pdf_path

# Função para converter um arquivo DOCX para um arquivo PDF usando a biblioteca reportlab
def convert_docx_to_pdf(docx_file, output_folder):
    # Salva o arquivo DOCX temporariamente em uma pasta temporária
    docx_path = os.path.join(output_folder, docx_file.filename)
    docx_file.save(docx_path)

    # Define o caminho do arquivo PDF
    pdf_path = os.path.join(output_folder, f"{os.path.splitext(docx_file.filename)[0]}.pdf")

    # Define o tamanho da página como A4 (21 cm x 29.7 cm)
    page_width, page_height = A4

    # Cria um objeto Document para o arquivo DOCX
    doc = Document(docx_path)
    # Cria um canvas (página) para o arquivo PDF
    pdf_canvas = canvas.Canvas(pdf_path, pagesize=A4)

    # Define as margens (72 pontos = 1 polegada)
    left_margin = 72
    right_margin = page_width - 72
    top_margin = page_height - 72
    bottom_margin = 72

    current_y = top_margin
    for para in doc.paragraphs:
        lines = para.text.splitlines()
        for line in lines:
            pdf_canvas.setFont("Helvetica", 12)
            text_width = pdf_canvas.stringWidth(line, "Helvetica", 12)

            # Verifica se a linha cabe na página
            if current_y - pdf_canvas._leading >= bottom_margin:
                pdf_canvas.drawString(left_margin, current_y, line)
                current_y -= pdf_canvas._leading
            else:
                # Cria uma nova página e recomeça a escrever do topo
                pdf_canvas.showPage()
                current_y = top_margin
                pdf_canvas.drawString(left_margin, current_y, line)
                current_y -= pdf_canvas._leading

    # Salva o arquivo PDF
    pdf_canvas.save()

    return pdf_path

# Função para compactar os arquivos PDF em um arquivo RAR
def compress_files_to_rar(input_paths, output_path):
    # Cria um diretório temporário usando o módulo tempfile
    with tempfile.TemporaryDirectory() as temp_folder:
        # Move os arquivos de entrada para o diretório temporário
        for input_path in input_paths:
            shutil.move(input_path, os.path.join(temp_folder, os.path.basename(input_path)))

        # Cria o arquivo ZIP usando a função make_archive do módulo shutil
        shutil.make_archive(os.path.splitext(output_path)[0], 'zip', temp_folder)

        # Move o arquivo ZIP criado para o caminho de saída
        shutil.move(os.path.splitext(output_path)[0] + '.zip', output_path)

    return output_path

# Executa o aplicativo Flask quando o arquivo é executado como um programa independente
if __name__ == '__main__':
    app.run(debug=True)
