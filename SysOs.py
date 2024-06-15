import ffmpeg
import moviepy.editor as mp
import os.path
from docx2pdf import convert
from fpdf import FPDF
import textwrap
from PIL import Image
import sys
import PyPDF2
from pdf2docx import Converter
import docx
from txt2docx import txt2docx
import docx2txt
from striprtf.striprtf import rtf_to_text


text_extensions = ['pdf', 'txt', 'docx']
audio_extensions = ['mp3', 'wav', 'ogg', 'avi', '3ga', '4xm', 'aac',
                    'htk', 'g722', 'weba', 'opus']
video_extensions = ['mov', 'vmv', 'gif', 'm4a', 'flac', 'mp4',
                    'mkv', 'h261', 'h262', 'h263', 'h265',
                    'dct', 'cak', 'av1', 'xvv', 'webm', 'paf',
                    'p64', 'pmp']
image_extensions = ['jpeg', 'png', 'img', 'bmp', 'gif', 'pfm',
                    'spider', 'tiff', 'blp', 'icns', 'ico', 'jpg',
                    'tga', 'webp', 'dds', 'dib', 'eps', 'im', 'msp',
                    'pcx', 'xbm', 'jpg']


file_path = sys.argv[1]
#  From colleagues to people
base_file_name = os.path.basename(file_path)
file_name_without_extension, file_extension_with_dot = os.path.splitext(base_file_name)
file_extension = file_extension_with_dot[1:]
absolute_path = file_path
directory = os.path.dirname(file_path)
expected_extension = input('expected_extension: ').lower()
if file_name_without_extension.endswith('(by SysOs)'):
    new_file_path = os.path.join(directory, f"{file_name_without_extension}.{expected_extension}")
else:
    new_file_path = os.path.join(directory, f"{file_name_without_extension} (by SysOs).{expected_extension}")


def pdf_to_txt():
    pdf_file_obj = open(absolute_path, 'rb')
    reader = PyPDF2.PdfReader(pdf_file_obj)
    pages = len(reader.pages)
    file1 = open(new_file_path, 'w')
    for i in range(pages):
        page_obj = reader.pages[i]
        text = page_obj.extract_text()
        file1.writelines(text)


def rtf_to_txt():
    rtf_file = open(absolute_path, 'r')
    rtf_content = rtf_file.read()
    rtf_text = rtf_to_text(rtf_content)
    txt_file = open(new_file_path, "w")
    txt_file.write(rtf_text)


def docx_and_doc_to_txt():  # good
    text = docx2txt.process(absolute_path)
    with open(new_file_path, "w") as file:
        file.write(text)


def doc_to_docx():
    doc_file = absolute_path
    doc = docx.Document(doc_file)
    doc.save(new_file_path)


def audio_and_video_converter():
    if expected_extension == 'mp4' and file_extension == 'gif':
        gif_to_mp4()
        return
    stream = ffmpeg.input(base_file_name)
    stream = ffmpeg.output(stream, new_file_path)
    ffmpeg.run(stream)


def txt_to_docx():
    txt2docx.txt2docx(absolute_path, new_file_path)


def pdf_to_docx():
    cv = Converter(absolute_path)
    cv.convert(new_file_path, start=0, end=None)


def text_to_pdf():
    file = open(absolute_path)
    text = file.read()
    file.close()
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(True)
    pdf.add_page()
    pdf.set_font(family='Helvetica', size=11)
    splitted = text.split('\n')
    for line in splitted:
        lines = textwrap.wrap(line)
        if len(lines) == 0:
            pdf.ln()
        for wrap in lines:
            pdf.cell(0, 16, wrap, ln=1)
    pdf.output(new_file_path, 'F')


def txt_extensions_converter():
    if expected_extension == 'txt':
        if file_extension == 'pdf':
            pdf_to_txt()
        elif file_extension == 'docx' or file_extension == 'doc':
            docx_and_doc_to_txt()
        elif file_extension == 'rtf':
            rtf_to_txt()
    elif expected_extension == 'docx':
        if file_extension == 'pdf':
            pdf_to_docx()
        elif file_extension == 'txt':
            txt_to_docx()
    elif expected_extension == 'pdf':
        if file_extension == 'docx':
            convert(absolute_path, new_file_path)
        elif file_extension == 'txt':
            text_to_pdf()


def convert_image():
    with Image.open(absolute_path) as img:
        try:
            img.save(new_file_path, format=expected_extension.upper())
        except:
            rgb_image = img.convert('RGB')
            rgb_image.save(new_file_path)


def gif_to_mp4():
    clip = mp.VideoFileClip(absolute_path)
    clip.write_videofile(new_file_path)


def main():
    if expected_extension == 'gif':
        if file_extension in image_extensions:
            convert_image()
        else:
            audio_and_video_converter()
    elif expected_extension == 'pdf':
        if file_extension in image_extensions:
            convert_image()
        else:
            txt_extensions_converter()
    elif expected_extension in video_extensions or expected_extension in audio_extensions:
        audio_and_video_converter()
    elif expected_extension in text_extensions:
        txt_extensions_converter()
    elif expected_extension in image_extensions:
        convert_image()


if __name__ == '__main__':
    main()
