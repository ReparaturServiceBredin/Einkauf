# Täglicher Einkauf über ITscope

`einkauf.py` liest eine Bedarfsliste (Excel/CSV), fragt jede Herstellernummer bei der
ITscope Produkt-API ab und wählt pro Artikel den günstigsten **eigenen** Lieferanten
mit ausreichendem Lagerbestand. Das Ergebnis ist eine Excel-Datei in `ausgabe\`.

## Ergebnis-Excel

Blatt **Einkauf**: Herstellernummer, Bezeichnung, Menge, bester Lieferant, Preis netto,
Lagerbestand, zweitbester Lieferant + Preis, Gesamtsumme (Formel Menge × Preis),
ITscope-Link, Status, Hinweis. Darunter die Gesamtsumme und die Liste aller Artikel,
die nicht gefunden wurden oder nicht bestellbar sind.

| Farbe   | Status                            | Bedeutung                                                    |
|---------|-----------------------------------|--------------------------------------------------------------|
| –       | `OK`                              | Günstigster Lieferant mit Bestand ≥ Menge                     |
| gelb    | `TEILBESTAND`                     | Niemand hat genug – günstigster mit Teilbestand gewählt       |
| gelb    | `KEIN BESTAND`                    | Keiner meiner Lieferanten hat Bestand – günstigster Preis    |
| orange  | `KEIN ANGEBOT (meine Lieferanten)`| Produkt gefunden, aber kein Angebot von meinen Lieferanten   |
| **rot** | `NICHT GEFUNDEN`                  | ITscope kennt die Herstellernummer nicht                     |
| **rot** | `FEHLER`                          | API-/Datenfehler bei diesem Artikel (Details im Log)         |

Blatt **Alle Angebote**: alle berücksichtigten Angebote je Artikel zum Nachvollziehen.

Fehler bei einzelnen Artikeln brechen den Lauf nicht ab. Nur bei falschen Zugangsdaten
(HTTP 401/403) wird abgebrochen, weil dann alle Artikel betroffen wären.
Gleiche Herstellernummern in der Bedarfsliste werden zusammengefasst (Mengen addiert).

**Rate Limit:** Zwischen zwei Anfragen liegt mindestens `min_abstand_sekunden`
(Standard 1 s). Bei HTTP 429 wird der `Retry-After`-Header beachtet, bei 5xx/Netzwerk-
fehlern mit wachsender Wartezeit bis zu `max_versuche`-mal wiederholt.

---

## 1. Installation (Windows, einmalig)

1. **Python installieren** (3.11 oder neuer): <https://www.python.org/downloads/windows/>
   Beim Setup unbedingt **„Add python.exe to PATH“** anhaken.
2. Diesen Ordner z. B. nach `C:\Einkauf` kopieren.
3. **Eingabeaufforderung** (`cmd`) öffnen und eingeben:
   ```bat
   cd /d C:\Einkauf
   python -m venv .venv
   .venv\Scripts\python.exe -m pip install -r requirements.txt
   ```
4. **Zugangsdaten:** `.env.example` nach `.env` kopieren und ausfüllen:
   ```bat
   copy .env.example .env
   notepad .env
   ```
   ```
   ITSCOPE_ACCOUNT_ID=deine-account-id
   ITSCOPE_API_KEY=dein-api-key
   ```
5. **`config.toml` anpassen** (mit Notepad öffnen):
   - `[eingabe]` `spalte_herstellernummer` / `spalte_menge`: exakte Spaltenüberschriften
     eurer Bedarfsliste, `datei`: wo die tägliche Liste liegt.
   - `[lieferanten]` `meine`: Namen (wie in ITscope angezeigt) oder Lieferanten-IDs der
     Lieferanten, bei denen ihr ein Kundenkonto habt. **Ohne Eintrag werden alle
     Angebote berücksichtigt.**
   - `[api]` `such_pfad` und `[felder]`: siehe Abschnitt „Abgleich mit der API-Doku“.

## 2. Starten

```bat
cd /d C:\Einkauf
.venv\Scripts\python.exe einkauf.py                         :: Datei aus config.toml
.venv\Scripts\python.exe einkauf.py "D:\Bedarf\heute.xlsx"  :: andere Datei
```

Probelauf ohne API mit den Beispieldaten:

```bat
.venv\Scripts\python.exe einkauf.py beispiel\Bedarfsliste.csv --mock beispiel\api_antworten
```

Ergebnis: `ausgabe\Einkauf_JJJJ-MM-TT_HHMM.xlsx`, Protokoll: `logs\einkauf_JJJJ-MM-TT.log`.
Ist die Ergebnisdatei gerade in Excel geöffnet, wird automatisch unter neuem Namen gespeichert.

## 3. Täglich um 14:45 per Aufgabenplanung

**Variante A – ein Befehl** (`cmd` öffnen, Pfad ggf. anpassen):

```bat
schtasks /Create /TN "Einkauf ITscope" /SC DAILY /ST 14:45 /TR "\"C:\Einkauf\start_einkauf.bat\""
```

Nur werktags (Mo–Fr):

```bat
schtasks /Create /TN "Einkauf ITscope" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 14:45 /TR "\"C:\Einkauf\start_einkauf.bat\""
```

Sofort testen: `schtasks /Run /TN "Einkauf ITscope"` – danach in
`logs\aufgabenplanung.log` und `ausgabe\` nachsehen.

**Variante B – per Oberfläche:**

1. Startmenü → **Aufgabenplanung** → rechts **„Aufgabe erstellen…“**.
2. Reiter **Allgemein**: Name „Einkauf ITscope“. Soll es auch ohne Anmeldung laufen:
   „Unabhängig von der Benutzeranmeldung ausführen“ wählen (Passwort wird abgefragt).
3. Reiter **Trigger** → Neu → **Täglich**, Start **14:45:00** (oder Wöchentlich Mo–Fr).
4. Reiter **Aktionen** → Neu → Programm starten:
   - Programm/Skript: `C:\Einkauf\start_einkauf.bat`
   - **Starten in:** `C:\Einkauf`
5. Reiter **Einstellungen**: „Aufgabe so schnell wie möglich nach einem verpassten Start
   ausführen“ anhaken. OK.

Hinweise: Liegt die Bedarfsliste auf einem Netzlaufwerk, in `config.toml` den UNC-Pfad
verwenden – in einfachen Anführungszeichen: `datei = '\\server\freigabe\Bedarfsliste.xlsx'` –, da Laufwerksbuchstaben
ohne Anmeldung nicht verbunden sind. Rückgabewert ≠ 0 bedeutet: Lauf komplett
gescheitert (Konfiguration, Bedarfsliste oder Zugangsdaten) – sichtbar in der
Aufgabenplanung unter „Letztes Ausführungsergebnis“.

## 4. Abgleich mit der API-Doku (wichtig vor dem ersten echten Lauf)

Die Feldnamen der API-Antwort sind in `config.toml` unter `[felder]` hinterlegt – jeweils
als Liste möglicher Pfade (Punkt-getrennt für verschachtelte Felder, z. B. `supplier.name`).
Vorgehen:

1. Eine echte Antwort für eine bekannte Herstellernummer abrufen, z. B. im Browser
   `https://api.itscope.com/2.1/products/search/hstpid=CF259A/standard.json`
   (Login = Account-ID / API-Key) – oder die Beispielantwort aus der Doku nehmen.
2. `[api] such_pfad` an die in der Doku beschriebene Suche nach Herstellernummer anpassen.
3. Unter `[felder]` die tatsächlichen Feldnamen für Produktliste, Bezeichnung, Angebote,
   Lieferant, Preis und Lagerbestand eintragen.
4. Falls ITscope Angebote mit eigenen Konditionen per Feld kennzeichnet, dieses unter
   `[lieferanten] eigene_konditionen_feld` eintragen.
5. Die echte Antwort als `beispiel\api_antworten\<Herstellernummer>.json` speichern und mit
   `--mock` testen.
