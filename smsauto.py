from ipaddress import IPv4Address
from pyairmore.request import AirmoreSession 
from pyairmore.services.messaging import MessagingService
from openpyxl import load_workbook
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
from progress.bar import IncrementalBar
import time


filepath = filedialog.askopenfilename(filetypes=[("Ficher Excel", ".xlsx .xls")])
columns = "AZERTYUIOPQSDFGHJKLMWXCVBN"
length = 1000

workbook = load_workbook(filename=filepath, read_only=True)
worksheet = workbook.active

phone_numbers = []

bar = IncrementalBar("Analyse du fichier Excel en cours...", max=26000)

for column in columns:
	for i in range(length):
		bar.next()
		cell = "{}{}".format(column, i+1)
		number = worksheet[cell].value
		number = str(number)
		delchar=[" ","/"]
		
		for char in delchar:
			number = number.replace(char, "")
		try:
			number = int(number)
		except:
			a=1	
		finally:
			if str(type(number)) == "<class 'int'>":
				number = str(number)
				if len(number) == 9:
					number = "0" + number
					phone_numbers.append(number)
				elif len(number) == 10:
					phone_numbers.append(number)
				elif number == "None":
					t = 1
				else:
					print (f"{number} n'est pas un numéro valide")
		
bar.finish()
time.sleep(2)
phone_numbers = list( dict.fromkeys(phone_numbers) )
print(f"\n{phone_numbers}")


ip = simpledialog.askstring(title="IP de connexion", prompt="Merci de renseigner l'adresse IP du téléphone (voir application)")

if ip == None:
	messagebox.showwarning(title="Annulation", message="Envois annulés")
	exit()

ip = IPv4Address(ip)
session = AirmoreSession(ip)


if session.is_server_running:
	messagebox.showwarning(title="Autorisation", message="Merci de confirmer la connexion sur votre téléphone (Cliquez OK pour continuer)")
	was_accepted = session.request_authorization()
	if was_accepted:
		print("Connexion autorisée")
		service = MessagingService(session)

		valid = 0
		while valid == 0:
				mess = simpledialog.askstring(title="Message à envoyer", prompt="Merci de rentrer le message à envoyer")
				if mess == None:
					messagebox.showwarning(title="Annulation", message="Envois annulés")
					exit()
				v = messagebox.askquestion(title="Confimer le message", message=f"Voulez-vous vraiment envoyer ce message à {len(phone_numbers)} numéros de téléphone? \n\n{mess}")
				if v == "yes":
					valid = 1
		bar = IncrementalBar("Analyse du fichier Excel en cours...", max=len(phone_numbers))
		for num in phone_numbers:
			try:
				service.send_message(num, mess)
			except MessageRequestGSMError:
				print(f"\nEchec de l'envoi du message à {num}, le numéro peut être valide mais le problème vient plutôt du telephone utilisé")
			else:
				print(f"Message envoyé à {num}")
			bar.next()
			bar.finish()
		messagebox.showwarning(title="Fin", message="Tout les messages ont été envoyés")
	else:
		print("\nLa connexion n'a pas été autorisée merci de réessayer")
else:
	print("La connection à échoué, merci de relancer l'application avant de réessayer")
