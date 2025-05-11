import asyncio

from langchain_community.document_loaders import PyPDFLoader
from context_pptx import process_presentation
import os

def read_file_pdf(file_path):
    loader = PyPDFLoader(file_path, extract_images=True)
    return loader.load()

def read_file_pptx(file_path):
    slide_contexts_list = asyncio.run(process_presentation(file_path))
    return slide_contexts_list

def read_file(file_path):
    extension = os.path.splitext(file_path)[-1].lower()

    if extension == ".pdf":
        slide_contexts = read_file_pdf(file_path)
    elif extension == ".pptx":
        slide_contexts = read_file_pptx(file_path)
    else:
        print("Данный тип файла не поддерживается.")
        return None

    return slide_contexts
  
