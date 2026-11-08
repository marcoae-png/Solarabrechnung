# Solarabrechnung

Lokales Windows-Programm zum Erstellen von Quartalsabrechnungen aus Smart-me-Excel-Exporten und Solar-Manager-Kennzahlen. Rechnungen werden als HTML gespeichert und können im Browser als PDF ausgegeben werden.

## Einrichtung

1. Python 3.11 oder neuer installieren und beim Setup zum PATH hinzufügen.
2. Im Programmordner eine Eingabeaufforderung öffnen und `py -3 -m pip install -r requirements.txt` ausführen.
3. `Solarabrechnung.vbs` starten und im Programm unter «Absender und Zahlung» IBAN sowie Absenderadresse eintragen und speichern.
4. Im Bereich «Bezüger und Zähler» Namen, Adressen und Smart-me-Zählernummern eintragen und speichern.
5. Die Tarife im Programm einstellen und speichern.
6. Den eigenen Solar-Manager-API-Schlüssel speichern und Smart-me-Quartalsexporte importieren.

Die Programmversion speichert Rechnungen und Register im Unterordner `Ausgabe`. API- und Mailbox-Passwörter werden im Windows-Anmeldeinformationsspeicher des jeweiligen Windows-Benutzers abgelegt.

## Konfiguration

IBAN und Absenderadresse werden im Programm unter «Absender und Zahlung» eingetragen. Bezüger und Smart-me-Zählernummern werden im Bereich «Bezüger und Zähler» gepflegt. Beim ersten Speichern legt das Programm `abrechnungskonfiguration.json` lokal an. Die Zählernummer muss mit der Nummer in der Smart-me-Datei übereinstimmen. Monatsproduktionen werden für jeden Monat des importierten Quartals einzeln aus Solar Manager geladen.

Für den optionalen E-Mail-Versand `email`, `smtp_host` und `smtp_port` eintragen. Unterstützt wird SMTP über SSL; das Mailbox-Passwort wird im Programm gespeichert.

## Smart-me-Import

Alle Dateien eines Imports müssen denselben vollen Kalenderquartalszeitraum enthalten. Erwartet werden Bezüger und Zeitraum in B2/B3 sowie die Gesamtbezugszeile `Import Total` oder `Bezug Total` in Spalte E.

## Voraussetzungen und Hinweise

- Windows, Python 3.11+, Microsoft Edge oder Google Chrome für PDF-Ausgabe und Internetzugang zu Solar Manager.
- Benötigte Python-Pakete: `openpyxl` und `reportlab`.
- Der Betreiber muss seine Solar-Manager-Berechtigung, Bankdaten, Tarife und Kundenangaben selbst konfigurieren und prüfen.
- Vor produktivem Versand QR-Rechnungen und Tarifberechnungen mit den eigenen Bank- und Abrechnungsanforderungen abgleichen.

Nach Änderungen an der Konfiguration das Programm neu starten. Die Beispielkonfiguration enthält absichtlich keine persönlichen Daten; vor der ersten Rechnung Absender, IBAN, Tarife und mindestens einen Bezüger eintragen.
