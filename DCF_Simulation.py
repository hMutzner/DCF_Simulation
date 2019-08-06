import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt

# Füegen Sie hier den Namen der EXCEL Datei ein
file_name =r"C:\Username\Projekte\weiterbauen\DCF saniert unsaniert.xlsx"

Mappe = xw.Book(file_name)
Blatt = Mappe.sheets['Modell']

# Parameter für die Dreiecksverteilung aus EXCEL Arbeitsblatt lesen
minimum = Blatt.range('Minimum').value
maximum = Blatt.range('Maximum').value
modus = Blatt.range('Modus').value

n = 1000 # Anzahl Simulationen

# Leere Listen für die Ergebnisse erzeugen 
KW_saniert = []
KW_unsaniert = []

# Simulieren
for i in range(n):
    # Dreiecksverteilte Zufallszahlen erzeugen
    liste_zufallszahlen = np.random.triangular(minimum,modus,maximum,len(Blatt.range('Teuerung')))
    # Zufallszahlen ins Arbeitsblatt schreiben
    Blatt.range('Teuerung').options(transpose=True).value = liste_zufallszahlen 
    # Ergebnisse (Kapitalwerte) einsammeln
    KW_unsaniert.append(Blatt.range('Kapitalwert_unsaniert').value)
    KW_saniert.append(Blatt.range('Kapitalwert_saniert').value)

# Ergebnisse darstellen
Klassenbreite = 10000
tiefster =  np.floor(min(KW_unsaniert+KW_saniert)/Klassenbreite)*Klassenbreite 
hoechster = np.ceil(max(KW_unsaniert+KW_saniert)/Klassenbreite)*Klassenbreite+Klassenbreite
Klassen = np.arange(tiefster,hoechster,Klassenbreite)

plt.figure(dpi=1200)
plt.hist(KW_unsaniert,Klassen,alpha = 0.5, label = 'unsaniert', color = 'r')
plt.hist(KW_saniert,Klassen,alpha = 0.5, label = 'saniert',color = 'b')
plt.xlabel("Kapitalwerte")
plt.ylabel("Häufigkeit")
plt.legend(loc='upper left')

plt.savefig(r"C:\Username\Projekte\weiterbauen\histog.png")
