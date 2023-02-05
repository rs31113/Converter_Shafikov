from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ParseMode, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from aiogram.utils.markdown import hlink
from PIL import Image
from docx import Document
from fpdf import FPDF
import os
from pytube import YouTube
import pandas as pd
from moviepy.editor import VideoFileClip
from config import TOKEN
from pdf2docx import parse
from typing import Tuple


def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
    if pages:
        pages = [int(i) for i in list(pages) if i.isnumeric()]
    result = parse(pdf_file=input_file,
                   docx_with_path=output_file, pages=pages)
    return result


def download(youtube_link):
    video = YouTube(youtube_link)
    video = video.streams.get_highest_resolution()
    video.download('storage', 'test.mp4')


async def reply_to_user(extension, kb, message, valid=True):
    if valid:
        await message.answer(f"Тип файла - {extension.upper()}.\n"
                             f"Выберите тип для конвертации:", reply_markup=kb)
    else:
        await message.answer("Извините, данный тип файла не поддерживается.")


bot = Bot(token=TOKEN)
dp = Dispatcher(bot)
bot_link = hlink('@shafikov_bot', 'https://t.me/shafikov_bot')
mode = ParseMode.HTML


im_btn_1 = InlineKeyboardButton('PDF', callback_data='pdf photo')
im_btn_2 = InlineKeyboardButton('PNG', callback_data='png photo')
im_btn_3 = InlineKeyboardButton('BMP', callback_data='bmp photo')
im_btn_4 = InlineKeyboardButton('JPEG', callback_data='jpeg photo')
im_kb = InlineKeyboardMarkup(row_width=2).add(im_btn_1, im_btn_2, im_btn_3, im_btn_4)

txt_btn_1 = InlineKeyboardButton('PDF', callback_data='pdf text')
txt_btn_2 = InlineKeyboardButton('DOCX', callback_data='docx text')
txt_kb = InlineKeyboardMarkup(row_width=2).add(txt_btn_1, txt_btn_2)

vid_btn = InlineKeyboardButton('MP3', callback_data='mp3 audio')
vid_kb = InlineKeyboardMarkup().add(vid_btn)


@dp.callback_query_handler(text=['docx pdf file', 'docx txt doc', 'docx text file', 'txt text file'])
async def convert_docx(callback_query: types.CallbackQuery):
    source_format = callback_query.data.split()[0]
    if source_format == 'txt':
        f = open("storage/test.txt", "r")
        f = f.readlines()
        text = [elem for elem in f]
        text = '\n'.join(text)
        await callback_query.message.answer(text)
    else:
        convert_to = callback_query.data.split()[1]
        path = f'storage/test.{convert_to}'
        if convert_to == 'text' or convert_to == 'txt':
            doc = Document(f"storage/test.docx")
            result = [p.text for p in doc.paragraphs]
            result = '\n'.join(result)
            if convert_to == 'text':
                await callback_query.message.answer(result)
            else:
                with open(path, "w") as f:
                    f.write(result)
                f.close()
                await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
                os.remove(path)
            os.remove("storage/test.docx")


@dp.callback_query_handler(text=['png jpeg file', 'png jpg file', 'png pdf file'])
async def convert_png(callback_query: types.CallbackQuery):
    convert_to = callback_query.data.split()[1]
    path = f"storage/test.{convert_to}"
    await bot.answer_callback_query(callback_query.id)
    image_1 = Image.open('storage/test.png')
    im_1 = image_1.convert('RGB')
    im_1.save(path)
    await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
    os.remove(path)
    os.remove('storage/test.png')


@dp.callback_query_handler(text=['pdf docx file'])
async def convert_pdf(callback_query: types.CallbackQuery):
    path = 'storage/test.docx'
    convert_pdf2docx('storage/test.pdf', path)
    await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
    os.remove('storage/test.pdf')
    os.remove(path)


@dp.callback_query_handler(text=['xlsx csv file'])
async def convert_xlsx(callback_query: types.CallbackQuery):
    path = 'storage/test.csv'
    read_file = pd.read_excel("storage/test.xlsx")
    read_file.to_csv(path)
    await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
    os.remove("storage/test.xlsx")
    os.remove(path)


@dp.callback_query_handler(text='mp3 audio')
async def convert_video(callback_query: types.CallbackQuery):
    video = VideoFileClip('storage/test.mp4')
    audio = video.audio
    path = 'storage/test.mp3'
    audio.write_audiofile(path)
    audio.close()
    video.close()
    await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
    os.remove('storage/test.mp4')
    os.remove(path)


@dp.callback_query_handler(text=['pdf text', 'docx text'])
async def convert_text(callback_query: types.CallbackQuery):
    message_format = callback_query.data.split()[0]
    path = f'storage/test.{message_format}'
    if message_format == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=15)
        for line in open("storage/test.txt", "r+", encoding='utf-8'):
            pdf.cell(200, 10, txt=line, ln=1, align='L')
        pdf.output(path)
        await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
        os.remove(path)
    elif message_format == "docx":
        document = Document()
        with open('storage/test.txt') as f:
            for line in f:
                document.add_paragraph(line)
        document.save(path)
        await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
        os.remove(path)
    os.remove('storage/test.txt')


@dp.callback_query_handler(text=['png photo', 'pdf photo', 'bmp photo', 'jpeg photo'])
async def convert_photo(callback_query: types.CallbackQuery):
    to_convert = callback_query.data.split()[0]
    path = f"storage/test.{to_convert}"
    await bot.answer_callback_query(callback_query.id)
    image_1 = Image.open('storage/test.jpg')
    im_1 = image_1.convert('RGB')
    im_1.save(path)
    await callback_query.message.answer_document(InputFile(path), caption=bot_link, parse_mode=mode)
    os.remove(path)
    os.remove('storage/test.jpg')


@dp.message_handler(commands=["start"])
async def send_welcome(message: types.Message):
    await message.answer("Привет! Конвертер файлов к "
                         "вашим услугам!\n\nЯ могу изменить "
                         "формат вашего файла.\nПросто отправь мне любой файл, а я "
                         "предложу варианты для конвертации.")


@dp.message_handler(content_types=['photo'])
async def photo_processing(message: types):
    extension = "jpg"
    kb = im_kb
    await message.photo[-1].download('storage/test.jpg')
    await reply_to_user(extension, kb, message)


@dp.message_handler(content_types=['text'])
async def text_processing(message: types):
    try:
        download(message.text)
        await message.answer_document(InputFile('storage/test.mp4'))
        os.remove('storage/test.mp4')
    except:
        extension = "txt"
        kb = txt_kb
        file = open('storage/test.txt', 'w+')
        file.write(message.text)
        file.close()
        await reply_to_user(extension, kb, message)


@dp.message_handler(content_types=['video'])
async def video_processing(message: types):
    extension = "mp4"
    kb = vid_kb
    await message.video.download('storage/test.mp4')
    await reply_to_user(extension, kb, message)


@dp.message_handler(content_types=['document'])
async def file_processing(message: types):
    valid = True
    extension = message.document.file_name.split(".")[1].lower()
    kb = InlineKeyboardMarkup(row_width=2)
    if extension == "xlsx":
        doc_btn_1 = InlineKeyboardButton('CSV', callback_data='xlsx csv file')
        kb.add(doc_btn_1)
    elif extension == "docx":
        doc_btn_1 = InlineKeyboardButton('PDF', callback_data='docx pdf file')
        doc_btn_2 = InlineKeyboardButton('TXT', callback_data='docx txt file')
        doc_btn_3 = InlineKeyboardButton('text message', callback_data='docx text file')
        kb.add(doc_btn_1, doc_btn_2, doc_btn_3)
    elif extension == "txt":
        kb.add(InlineKeyboardButton('text message', callback_data='txt text file'))
    elif extension == "pdf":
        doc_btn_1 = InlineKeyboardButton('DOCX', callback_data='pdf docx file')
        kb.add(doc_btn_1)
    elif extension == "png":
        doc_btn_1 = InlineKeyboardButton('JPEG', callback_data='png jpeg file')
        doc_btn_2 = InlineKeyboardButton('PDF', callback_data='png pdf file')
        doc_btn_3 = InlineKeyboardButton('JPG', callback_data='png jpg file')
        kb.add(doc_btn_1, doc_btn_2, doc_btn_3)
    elif extension == "jpg":
        kb = im_kb
    else:
        valid = False
    if valid:
        await message.document.download(f"storage/test.{extension}")
    await reply_to_user(extension, kb, message, valid)


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
