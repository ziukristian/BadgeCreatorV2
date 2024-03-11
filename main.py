import os, sys
import pip
from PIL import Image,ImageDraw,ImageFont
import pandas as pd
import math
import datetime
from pypdf import PdfWriter

FILE_NAME = 'Lista.xlsx'
FILE_SHEET = 'Lista Totale'
HEADER_INDEX = 3
COL_RANGE = 'A:BY'

colorePacci="#808080"
coloreAngel="#FF0000"
font = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 180)
fontBold = ImageFont.truetype("Fonts/Montserrat-Bold.ttf", 180)
fontAzienda = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 150)
fontTrainer = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 100)
fontPacci = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 100)
fontSmallerBold = ImageFont.truetype("Fonts/Montserrat-Bold.ttf", 150)
fontSmaller = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 150)


class Partecipante:
    def __init__(self, nome='',cognome='',azienda='',trainer='',rifrequentante='',account='',colore='',testo=''):
        self.nome = '' if pd.isnull(nome) else nome
        self.cognome = '' if pd.isnull(cognome) else cognome
        self.azienda = '' if pd.isnull(azienda) else azienda
        self.trainer = '' if pd.isnull(trainer) else trainer
        self.rifrequentante = '' if pd.isnull(rifrequentante) else rifrequentante
        self.account = '' if pd.isnull(account) else account
        self.colore= '#FFFFFF' if pd.isnull(colore) else colore
        self.testo = '#000000' if pd.isnull(testo) else testo
def scrivi(testo,w,h,font,draw,hex="#000000"):
    W, H = (w, h)
    _, _, w, h = draw.textbbox((0, 0), testo, font=font)
    draw.text(((W - w) / 2, (H - h) / 2), testo, font=font, fill=hex)
def getTextWidth(font,text):
    left, top, right, bottom = font.getbbox(text)
    return right - left
def genera(partecipante,i,writer,numeroPartecipanti):
    IMAGE = Image.open("badge.jpg")
    DRAW = ImageDraw.Draw(IMAGE)

    if (partecipante.account=='Pacci'):
        DRAW.polygon([(0, 0), (IMAGE.width * (33 / 100), 0), (0, IMAGE.width * (33 / 100))], fill=colorePacci)

    if (partecipante.account=='Client Angel'):
        DRAW.polygon([(0, 0), (IMAGE.width * (33 / 100), 0), (0, IMAGE.width * (33 / 100))], fill=coloreAngel)

    #if (partecipante.rifrequentante == 'X'):
        #if(partecipante.account=='Client Angel' or partecipante.account=='Pacci'):
            #draw.polygon([(size[0], 0), (size[0], size[0] * (33 / 100)), (size[0] - (size[0] * (33 / 100)), 0)],fill=coloreRifrequentante)
        #else:
            #draw.polygon([(size[0], 0), (size[0], size[0] * (27 / 100)), (size[0] - (size[0] * (27 / 100)), 0)],fill=coloreRifrequentante)

    DRAW.polygon([(0, 0), (IMAGE.width * (27 / 100), 0), (0, IMAGE.width * (27 / 100))], fill=partecipante.colore)

    scrivi(partecipante.trainer, IMAGE.width * (0.2 if partecipante.trainer=="Pacci" else 0.14), IMAGE.width * 0.14, fontPacci if partecipante.trainer=="Pacci" else fontTrainer,DRAW,partecipante.testo)

    #AZIENDA
    azienda = partecipante.azienda.strip().upper()
    if getTextWidth(fontAzienda,azienda) > (IMAGE.width * 0.85):
        aziendaArray = azienda.split(' ', 2)
    else:
        aziendaArray = [azienda]

    if (len(aziendaArray) == 1):
        scrivi(aziendaArray[0], IMAGE.width, IMAGE.width * 1.76, fontAzienda, DRAW)
    if (len(aziendaArray) == 2):
        scrivi(aziendaArray[0], IMAGE.width, IMAGE.width * 1.63, fontAzienda, DRAW)
        scrivi(aziendaArray[1], IMAGE.width, IMAGE.width * 1.76, fontAzienda, DRAW)
    if (len(aziendaArray) == 3):
        fontDaUsare = fontAzienda
        left, top, right, bottom = fontAzienda.getbbox(aziendaArray[0])
        w = right - left

        if w > (IMAGE.width * 0.85):
            fontDaUsare = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 90)
        scrivi(aziendaArray[0], IMAGE.width, IMAGE.width * 1.63, fontDaUsare, DRAW)
        left, top, right, bottom = fontAzienda.getbbox(aziendaArray[1])
        w = right - left
        if w > (IMAGE.width * 0.85):
            fontDaUsare = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 90)
        scrivi(aziendaArray[1], IMAGE.width, IMAGE.width * 1.76, fontDaUsare, DRAW)
        left, top, right, bottom = fontAzienda.getbbox(aziendaArray[2])
        w = right - left
        if w > (IMAGE.width * 0.85):
            fontDaUsare = ImageFont.truetype("Fonts/Montserrat-Regular.ttf", 90)
        scrivi(aziendaArray[2], IMAGE.width, IMAGE.width * 1.89, fontDaUsare, DRAW)

    #NOME
    nome = partecipante.nome.strip().upper()
    if getTextWidth(fontAzienda,nome) > (IMAGE.width * 0.85):
        scrivi(nome, IMAGE.width, IMAGE.width * 1.03, fontSmallerBold, DRAW)
    else:
        scrivi(nome, IMAGE.width, IMAGE.width * 1.03, fontBold, DRAW)

    #COGNOME
    cognome = partecipante.cognome.strip().upper()
    if getTextWidth(fontAzienda,cognome) > (IMAGE.width * 0.85):
        scrivi(cognome, IMAGE.width, IMAGE.width * 1.22, fontSmaller, DRAW)
    else:
        scrivi(cognome, IMAGE.width, IMAGE.width * 1.22, font, DRAW)

    IMAGE.save('result'+str(i)+'.pdf')
    IMAGE.close()
    writer.append('result' + str(i) + '.pdf')
    os.remove('result'+str(i)+'.pdf')
    print(f"[{int((i/numeroPartecipanti)*100)}%] - {partecipante.nome} {partecipante.cognome}")

def main():
    print("Ricerca partecipanti in corso...")
    partecipanti=[]
    df = pd.read_excel(FILE_NAME,sheet_name=FILE_SHEET,header=HEADER_INDEX,usecols=COL_RANGE)
    df = df.reset_index()
    for index, row in df.iterrows():
        if (row['ESPORTA BADGE'] == 'SI' and row['ESPORTATO'] != 'SI'):
            partecipanti.append(
                Partecipante(row['Nome K-Chain'], row['Cognome K-Chain'], row['AZIENDA'], row["Account durante evento"],
                             row["Altre Edizioni Confermato"], row['Account blackship'], row['COLORE'], row['TESTO'])
            )
    partecipanti.sort(key=lambda x: x.nome.strip() + x.cognome.strip())
    print(f"Trovati {len(partecipanti)} partecipanti")
    print("Generazione badge in corso...")
    writer = PdfWriter()
    for i, x in enumerate(partecipanti):
        genera(x, i, writer,len(partecipanti))
    writer.write(f"{datetime.datetime.now().strftime("%y%m%d-%H%M")}-Badges.pdf")
    print("GENERAZIONE TERMINATA")


main()
