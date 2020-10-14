# RadNetEmailToExcel
In diesem kleinen Projekt wurde eine automatisierte umwandlung von den Standard Meldeemails zu einer Excel, welche alle Informationen beinhaltet umgesetzt
## Abhängigkeiten
folgende python pakete werden benötigt
```bash
pip install openpyxl
pip install python-imap
pip install emails
pip install yaml
```
## Benutzung
anpassen der Werte im config file und dann folgenden befehl
```bash
python main.py -c config.yml
```