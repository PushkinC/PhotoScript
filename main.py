from PIL import Image, ImageDraw, ImageFont, ImageTk
from tkinter import Label, Tk, Entry
from openpyxl import load_workbook
import os

file_xlsx = load_workbook('book.xlsx')

def save_to_excel(data):
    global file_xlsx
    sheets = file_xlsx.sheetnames
    sheet_name = sheets[0]
    sheet = file_xlsx[sheet_name]
    max_row = sheet.max_row + 1
    sheet[f'A{max_row}'].value = data[0]
    sheet[f'B{max_row}'].value = data[1]
    sheet[f'C{max_row}'].value = data[2]
    sheet[f'D{max_row}'].value = data[3]
    sheet[f'E{max_row}'].value = f'=B{max_row} - C{max_row}'
    sheet[f'F{max_row}'].hyperlink = data[4]
    sheet[f'F{max_row}'].value = data[4]
    sheet[f'F{max_row}'].style = "Hyperlink"




amount = len(os.listdir('input'))
cost = '200'
size = '12'
nam = '20'

l = []
for i in os.listdir('output'):
    l.append(int(i.rstrip('.jpg')))

if len(l) == 0:
    l = [0]

data = []
cur_num = max(l)
for _, filename in enumerate(os.listdir('input')):
    i = _ + cur_num
    print(f'Обрабатываю {filename} ({i + 1 - cur_num}/{amount})')
    im = Image.open(f'input/{filename}')
    im_size = im.size


    # cost = input('Введите цену: ')
    # size = input('Введите размер: ')
    # nam = input('Введите ст-ть: ')

    data = [i + 1, cost, nam, size, f'output/{i + 1}.jpg']

    new_im = Image.new('RGB', (im_size[0], im_size[1] + 400), color=(255,255,255))
    new_im.paste(im)
    new_size = new_im.size



    font = ImageFont.truetype('arial.ttf', size=250)
    text = f'{cost} руб.,  Размер {size},  Арт. {i + 1}'

    draw_im = ImageDraw.Draw(new_im)
    _,_,w,h = draw_im.textbbox((0, 0), text, font=font)

    draw_im.text(((new_size[0] - w) // 2, new_size[1] - h - 50), text, (0, 0, 0), font=font)

    new_im.save(f'output/{i + 1}.jpg')
    save_to_excel(data)

file_xlsx.save('book.xlsx')



