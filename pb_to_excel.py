import tkinter as tk
from tkinter.messagebox import showerror
import requests
import string
try:
	from BeautifulSoup import BeautifulSoup
except ImportError:
	from bs4 import BeautifulSoup
import os
import shutil
import base64
import xlwt
from datetime import datetime
title_list = ['Nominativo', 'Indirizzo', 'numero/i', 'Assente', 'Appuntamento']
field_list = ['nome', 'indirizzo', 'numero']

Error = "Impossibile trovare nominativi a questo indirizzo. {0} \n" \
        "Controlla che sia stato inserito correttamente. \n" \
        "Esempio indirizzo : Via Rossi, Roma (RO) "


datas = []

output = ''
page=1

class App(tk.Tk):

	def __init__(self):
		super().__init__()
		self.geometry("900x550")
		self.title("CERCA INDIRIZZO E CREA EXCEL")
		self.grid_columnconfigure(0, weight=1)
		self.welcome_label = tk.Label(self,
			text="Inserisci l'indirizzo da cercare",
			font=("Helvetica", 15))
		self.welcome_label.grid(row=0, column=0, sticky="N", padx=20, pady=10)

		self.text_input = tk.Entry()
		self.text_input.grid(row=1, column=0, sticky="WE", padx=10)

		self.download_button = tk.Button(text="CERCA INDIRIZZO E CREA EXCEL", command=self.search_nominativi)
		self.download_button.grid(row=2, column=0, sticky="WE", pady=10, padx=10)



	def generate_xls_attachment(self, datas, lista_campi=None, lista_titoli=None,
	                            ):
		if not lista_campi or type(lista_campi) is dict:
			lista_campi = field_list
		if not lista_titoli or type(lista_titoli) is dict:
			lista_titoli = title_list

		name = f'File {self.text_input.get()} {datetime.today().strftime("%d-%m-%Y")}.xls'

		path_xlsx = str('C:/Users/polik/Downloads')

		PATH = path_xlsx + "/" + name

		# Nel caso esiste già il file lo cancello
		try:
			os.remove(PATH)
		except:
			pass
		wb = xlwt.Workbook()
		try:
			# Aggiungo una pagina al file Excel
			sheet = wb.add_sheet("Nominativi", cell_overwrite_ok=True)

		except:
			# Nel caso esiste già prendo la prima
			sheet = wb.get_sheet(0)

		bold = xlwt.easyxf('font: bold on')
		# style_tot = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
		# money = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
		#     num_format_str='€#,####.##')
		style_t = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;',
			'font: bold on; align: vert centre, horiz center')
		# Scrivo i titoli delle colonne
		for titolo in lista_titoli:
			sheet.write(0, lista_titoli.index(titolo), titolo, style_t)

		riga = 1
		for line in datas:
			for field in lista_campi:
				keys = line.keys()
				# Se il campo è presente come campo nella fattura lo scrivo
				if field in keys:
					new_field = line[field]

					sheet.write(riga, lista_campi.index(field), new_field)

			riga += 1

		# Salvo il file

		wb.save(PATH)




	def search_nominativi(self):
		global datas
		global page
		global output

		def get_indirizzo(indirizzo, params=False):
			global page
			global output
			def find_nominativi(html_body):
				global datas
				try:
					nominativi = html_body.findAll('section')
				except:
					#showerror("Error", message=str(Error))
					return Error

				for nom in nominativi:
					data = {'nome': '', 'indirizzo': '', 'numero': ''}
					if nom.contents:
						nom_1 = nom.contents
						if nom_1[0].contents:
							con = nom_1[0].contents
							nome_indirizzo = [item.text for item in con[1].contents[0].contents]
							data['nome'] = nome_indirizzo[0]
							data['indirizzo'] = nome_indirizzo[1]
							numeri = [numero.rstrip(string.ascii_uppercase) for numero in con[2].contents[0].text]
							data['numero'] = ''.join(numeri)
					datas.append(data)
				if not nominativi:
					return Error
				return False
			payload = {"dv": indirizzo}
			if params:
				payload = payload | params

			response = requests.get("http://www.paginebianche.it/cerca-da-indirizzo",
				params=payload)

			text_response = response.text
			html = text_response  # the HTML code you've written above
			parsed_html = BeautifulSoup(html, features="html.parser")
			html_body = parsed_html.body.find('div', attrs={'class': 'search-listing'})

			risultato = find_nominativi(html_body)
			if risultato:#se torna c'è un errore
				errore = risultato.replace('{0}',indirizzo)
				output +=f'{errore}\n'
			others = parsed_html.body.find('a', attrs={'class': 'click-load-others'})
			if others:
				c = others
				if c.attrs:
					if c.attrs['data-pageurl'] and c.attrs['data-nextpage']:
						page = page + 1
						params = {'p': page}

						get_indirizzo(indirizzo, params)

		if not self.text_input:
			output += "Aggiungi una parola o una frase al campo input!"
			return


		user_input = self.text_input.get()
		indirizzi = user_input.split(';')
		for indirizzo in indirizzi:

			globals()['page']=1
			get_indirizzo(indirizzo.strip())



		output += f"Sono stati trovati in totale {len(datas)} nominativi"
		datas =sorted(datas, key=lambda d:(d['indirizzo'] , d['nome']))
		self.generate_xls_attachment(datas)


		textwidget = tk.Text()
		textwidget.insert(tk.END, output)
		textwidget.grid(row=3, column=0, sticky="WE", padx=10, pady=10)

		credits_label = tk.Label(self, text="Scarica Excel")
		credits_label.grid(row=4, column=0, sticky="S", pady=10)


if __name__ == "__main__":
	app = App()
	app.mainloop()
