# my_VbaRegistry
Auslesen von Registry_Keys und Prüfung per WindowsScripting_Host

# Verwendung/Setup

# Beispiele 
Schlüssel auslesen - TMP-Verzeichnis<br>
Debug.Print registry_Key_Read("HKEY_CURRENT_USER\Environment\TMP")

Schlüssel prüfen - Netzlaufwerk H vorhanden<br>
Debug.Print registry_Key_Exists("HKEY_CURRENT_USER\Network\H\RemotePath")

