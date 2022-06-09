# 
# Łączy pliki *.pdf|*.PDF rozpoczynające się od liczby 1-9 i znaku kropki
# 

import fnmatch, os, io, re, time, fitz

from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont

registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

# Funkcja zwracająca stronę załącznika
def get_Attachment(packet, app_number, file_name, title):
	can = canvas.Canvas(packet, pagesize=(210 * mm, 20 * mm))
	can.setFont('Arial-Bold', 10)
	can.drawString(16 * mm, 10 * mm, 'Załącznik ' + str(app_number) + '. ' + str(title))
	can.setFillColor(HexColor(0xE5402C))
	can.rect(16 * mm, 5 * mm, 178 * mm, 0.8 * mm, stroke=0, fill=1)
	can.save()
	return can

# Funkcje sortujące
def atoi(text):
    return int(text) if text.isdigit() else text

def natural_keys(text):
    return [ atoi(c) for c in re.split(r'(\d+)', text) ]

# Wyrażenie regularne do wybrania plików do połączenia
pattern = re.compile("[0-9]+\\..*\\.(pdf|PDF)")

pdfs = [a for a in os.listdir() if pattern.match(a)]

pdfs.sort(key=natural_keys)

fileName = str(pdfs[0][3:15]) + ' Ocena ryzyka z załącznikami.pdf'

# Usuwa istniejący plik
if os.path.exists(fileName):
  os.remove(fileName)

merger = PdfFileMerger(strict=False)

# Tworzy tablicę tytułów załączników
pdf = fitz.open(pdfs[0])
page = pdf.load_page(pdf.page_count - 1)
if re.search('Załączniki', page.get_text()):
	attachmentCase = 'Załączniki'
else:
	attachmentCase = 'ZAŁĄCZNIKI'
lst = page.get_text().split(attachmentCase);
lst = lst[1].split()
pattern = re.compile("[0-9]+\\..*")
title = []
titles = []
for l in lst:
	if not l == 'Załącznik':
		if not pattern.match(l):
			title.append(l)
	else:
		if len(title) > 0:
			t = ' '.join(title)
			titles.append(t)
		title = []
titles.append(' '.join(title))
titles.append([])

# Łącz jeśli znaleziono więcej niż jeden plik
if len(pdfs) > 1:
	print('Znalezione pliki do połączenia:')
	for p in pdfs:
		print('\t'+ str(p))
	for index, pdf in enumerate(pdfs, start=1):
		packet = io.BytesIO()
		# Pobierz obiekt załącznika
		can = get_Attachment(packet, index, pdf, titles[index - 1])
		packet.seek(0)
		attachment = PdfFileReader(packet)
		# Dolącz pdf do całości
		merger.append(PdfFileReader(open(pdf, 'rb')))
		# Jeśli pdf nie jest ostatnim dołącz stronę załącznika do całości
		if index < len(pdfs):
			merger.append(attachment, bookmark='Załącznik ' + str(index) + '. '+ str(titles[index - 1]))
	# Zapisz plik całości na dysk
	with open(fileName, 'wb') as fout:
		merger.write(fout)
	merger.close()
	# Sprawdź czy plik istnieje na dysku
	if os.path.exists(fileName):
		print('Pliki połączone w:\n\t' + str(fileName))
	else:
		print('Błąd zapisu!')
else:
	print('Brak wystarcząjącej liczby plików do połączenia!')

input('\nNaciśnij Enter...')
