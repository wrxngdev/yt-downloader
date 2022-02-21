import os
import xlrd2 as xlrd
import pytube
from os.path import expanduser

home = expanduser("~")
musicFolder = os.path.join(home, r"Documents\MusicDownloads") # Route zu MusicOrdner(Universal)
excelFile = os.path.join(home, r"Documents\MusicDownloads\Excel\youtubelinks.xlsx") # Route zur Excel datei(Universal)


try:
    f = open(excelFile, "r")
    f.close()

    excel_workbook = xlrd.open_workbook(excelFile)
    sheet = excel_workbook.sheet_by_index(0)
    

    for row in range(sheet.nrows):
        youtube_link = sheet.cell_value(row, 0)
        youtube_mp3 = pytube.YouTube(youtube_link)
        stream = youtube_mp3.streams.get_audio_only()
        stream.download(musicFolder)
        print(youtube_link + " ist Fertig installiert")
    print("")
    print("Der Vorgang wurde beendet und alle Links wurden installiert!")
    print("")
except FileNotFoundError:
    print("")
    print("Die Datei existiert nicht bitte erstellen sie diese!")
    print("1. Erstelle bei den Documents Ordner einen Ordner namens 'MusicDownloads'.")
    print("2. Erstelle in diesem Ordner noch ein Ordner namens 'Excel'.")
    print("3. In dem Excel ordner erstellst du eine Excel Datei namens 'youtubelinks.xlsx'")
    print("4. Nun kannst du die Links in die erste Spalte der Excel datei einfügen und die yt-downloader.exe starten.")
    print("Die Dateien solltest du nun in dem 'MusicDownloads' Ordner finden'.")
    print("")