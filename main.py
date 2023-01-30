from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import KeyboardButton, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from PIL import Image
from docx import Document
from fpdf import FPDF
import os
from moviepy.editor import VideoFileClip


TOKEN = '5958019510:AAFUzneTvzc1043uLDNscqH632BMmfCoyKc'


bot = Bot(token=TOKEN)
dp = Dispatcher(bot)


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


@dp.callback_query_handler(text=['pdf jpg document', 'pdf docx document', 'pdf png document', 'pdf jpeg document'])
async def testing(callback_query: types.CallbackQuery):
    pass



@dp.callback_query_handler(text='mp3 audio')
async def convert_video(callback_query: types.CallbackQuery):
    video = VideoFileClip('storage/test.mp4')
    audio = video.audio
    audio.write_audiofile('storage/test.mp3')
    audio.close()
    video.close()
    await callback_query.message.answer_document(InputFile("storage/test.mp3"))
    os.remove('storage/test.mp4')
    os.remove('storage/test.mp3')


@dp.callback_query_handler(text=['pdf text', 'docx text'])
async def convert_text(callback_query: types.CallbackQuery):
    message_format = callback_query.data.split()[0]
    if message_format == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=15)
        for line in open("storage/test.txt", "r+", encoding='CP-866'):
            pdf.cell(200, 10, txt=line, ln=1, align='L')
        pdf.output("storage/test.pdf")
        await callback_query.message.answer_document(InputFile("storage/test.pdf"))
        os.remove('storage/test.pdf')
    elif message_format == "docx":
        document = Document()
        with open('storage/test.txt') as f:
            for line in f:
                document.add_paragraph(line)
        document.save('storage/test.docx')
        await callback_query.message.answer_document(InputFile("storage/test.docx"))
        os.remove('storage/test.docx')
    os.remove('storage/test.txt')


@dp.callback_query_handler(text=['png photo', 'pdf photo', 'bmp photo', 'jpeg photo'])
async def convert_photo(callback_query: types.CallbackQuery):
    message_format = callback_query.data.split()[0]
    print(message_format)
    path = f"storage/test.{message_format}"
    await bot.answer_callback_query(callback_query.id)
    image_1 = Image.open('storage/test.jpg')
    im_1 = image_1.convert('RGB')
    im_1.save(path)
    await callback_query.message.answer_document(InputFile(path))
    os.remove(path)
    os.remove('storage/test.jpg')


@dp.message_handler(commands=["start"])
async def send_welcome(message: types.Message):
    await message.answer("Привет! Конвертер файлов к вашим услугам!\n\nЯ могу изменить "
                         "формат вашего файла.\nПросто отправь мне любой файл, а я "
                         "предложу варианты для конвертации.")


@dp.message_handler(content_types=['any'])
async def controller(message: types):
    valid = True
    try:
        extension = message.document.file_name.split(".")[1]
        kb = InlineKeyboardMarkup(row_width=2)
        if extension == "xlsx":
            doc_btn_1 = InlineKeyboardButton('PDF', callback_data='xlsx pdf document')
            doc_btn_2 = InlineKeyboardButton('CSV', callback_data='xslx csv document')
            kb.add(doc_btn_1, doc_btn_2)
        elif extension == "docx":
            doc_btn_1 = InlineKeyboardButton('PDF', callback_data='docx pdf document')
            doc_btn_2 = InlineKeyboardButton('TXT', callback_data='docx txt document')
            kb.add(doc_btn_1, doc_btn_2)
        elif extension == "txt":
            kb = text
        elif extension == "pdf":
            doc_btn_1 = InlineKeyboardButton('PNG', callback_data='pdf png document')
            doc_btn_2 = InlineKeyboardButton('TXT', callback_data='pdf txt document')
            doc_btn_3 = InlineKeyboardButton('DOCX', callback_data='pdf docx document')
            doc_btn_4 = InlineKeyboardButton('JPG', callback_data='pdf jpg document')
            kb.add(doc_btn_1, doc_btn_2, doc_btn_3, doc_btn_4)
        elif extension == "png":
            doc_btn_1 = InlineKeyboardButton('JPEG', callback_data='png jpeg document')
            doc_btn_2 = InlineKeyboardButton('PDF', callback_data='png pdf document')
            doc_btn_3 = InlineKeyboardButton('JPG', callback_data='png jpg document')
            kb.add(doc_btn_1, doc_btn_2, doc_btn_3)
        elif extension == "jpg":
            kb = im_kb
        else:
            valid = False
        if valid:
            await message.document.download(f"storage/test.{extension}")
    except AttributeError:
        if message.content_type == "text":
            extension = "txt"
            kb = txt_kb
            file = open('storage/test.txt', 'w+')
            file.write(message.text)
            file.close()
        elif message.content_type == "photo":
            extension = "jpg"
            kb = im_kb
            await message.photo[-1].download('storage/test.jpg')
        else:
            extension = "mp4"
            kb = vid_kb
            await message.video.download('storage/test.mp4')
    if valid:
        await message.answer(f"Тип файла - {extension.upper()}.\n"
                             f"Выберите тип для конвертации:", reply_markup=kb)
    else:
        await message.answer("Извините, данный тип файла не поддерживается.")


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
