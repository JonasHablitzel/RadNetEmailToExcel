# RadNetEmailToExcel
In diesem kleinen Projekt wurde eine automatisierte umwandlung von den Standard Meldemails zu einer Excel, welche alle Informationen beinhaltet umgesetzt

 http://www.rad-net.de/
 
## Abhängigkeiten
folgende python pakete werden benötigt
```bash
pip install openpyxl
pip install python-imap
pip install emails
pip install yaml
```
## Benutzung
anpassen der Werte im config file und dann folgenden befehl ausführen
```bash
python main.py -c config.yml
```