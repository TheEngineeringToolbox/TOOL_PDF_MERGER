# Rapportage Merger Tool

De Rapportage Merger Tool voegt een hoofdrapport en de bijbehorende bijlagen samen tot één complete PDF. De applicatie maakt automatisch voorbladen voor de bijlagen, zet Word-, Excel- en PowerPoint-bestanden om naar PDF en voegt alles in de juiste volgorde samen.

## Benodigdheden

- Windows
- Microsoft Word voor het hoofdrapport en voorbladen
- Microsoft Excel en/of PowerPoint wanneer deze bestandstypen als bijlage worden gebruikt
- Een correct opgebouwde projectmap

De tool wordt gebruikt via de Windows-EXE. Python, VS Code en Git zijn niet nodig voor normaal gebruik.

## Projectmap voorbereiden

Maak de projectmap als volgt op:

```text
Projectmap
├── Rapportage.docx
└── Bijlage
    ├── Bijlage A - Locatie overzicht
    │   ├── 01 Overzicht.pdf
    │   └── 02 Tekening.pdf
    ├── Bijlage B - Tekeningen
    │   ├── 01 Plattegrond.pdf
    │   └── 02 Doorsnede.xlsx
    └── Bijlage 1 - Berekening
        └── 01 Berekening.docx
```

Het hoofdrapport mag een andere bestandsnaam hebben, maar moet een `.docx`-bestand zijn. De map met bijlagen moet `Bijlage` heten. Bijlagemappen moeten beginnen met `Bijlage`, gevolgd door een nummer of letter, een koppelteken en een omschrijving. Bijvoorbeeld:

```text
Bijlage A - Locatie overzicht
Bijlage 1 - Berekening
```

De tool ondersteunt PDF, Word (`.doc` en `.docx`), Excel (`.xls`, `.xlsx` en `.xlsm`) en PowerPoint (`.ppt` en `.pptx`). Bestanden worden binnen elke bijlage op natuurlijke bestandsnaamvolgorde verwerkt. Tijdelijke bestanden (`~$...`), `desktop.ini`, `Thumbs.db`, `.tmp`- en `.bak`-bestanden worden overgeslagen.

## Een complete PDF maken

1. Start bijvoorbeeld `Rapportage Merger Tool.exe`.
2. Klik op **Kies DOCX...**.
3. Selecteer het hoofdrapport (`.docx`). De projectmap wordt automatisch bepaald.
4. Controleer in **Preview PDF-opbouw** of het rapport, alle bijlagen, de nummering, de volgorde en de bestanden per bijlage kloppen.
5. Klik op **GENEREER COMPLETE PDF**.
6. Wacht tot de voortgang 100% bereikt en de status **Gereed** toont.

De PDF wordt opgeslagen in de map `PDF rapportage` binnen de projectmap. De bestandsnaam is `<naam van hoofdrapport>_compleet.pdf`, bijvoorbeeld `Rapportage_compleet.pdf`.

## Volgorde van de PDF

De uiteindelijke PDF bestaat uit het hoofdrapport, gevolgd door het voorblad van elke bijlage en daarna de bestanden uit die bijlage. Bijlagen met nummers komen vóór bijlagen met letters. Nummers worden numeriek gesorteerd (`Bijlage 2` vóór `Bijlage 10`), gevolgd door letterbijlagen (`Bijlage A`, `Bijlage B`). Pas namen aan wanneer de volgorde niet klopt en genereer opnieuw.

## Bijlagen wijzigen

Pas het betreffende bestand aan in de juiste bijlagemap. Selecteer daarna het hoofdrapport opnieuw en genereer de complete PDF opnieuw. De voorbladen en de samengestelde PDF worden opnieuw gemaakt.

## Belangrijke aandachtspunten

- Sluit de uitvoer-PDF in Adobe Reader, Edge of een andere PDF-viewer voordat je opnieuw genereert; een geopende PDF kan niet worden overschreven.
- Wijzig of verwijder het centrale voorbladtemplate en het TOOL-logo niet.
- Gebruik geen tijdelijk Word-bestand dat met `~$` begint.
- Controleer vóór het genereren altijd de preview.

## Problemen oplossen

### De map `Bijlage` wordt niet gevonden

Controleer of de map direct in de projectmap staat en exact `Bijlage` heet. Controleer ook of de geselecteerde DOCX in de juiste projectstructuur staat.

### Geen bijlagen gevonden

Controleer of de submappen volgens dit patroon zijn benoemd:

```text
Bijlage A - Omschrijving
Bijlage 1 - Omschrijving
```

Elke herkende bijlagemap moet ten minste één verwerkbaar bestand bevatten.

### Een bestand kan niet worden geconverteerd

Controleer of het bestandstype wordt ondersteund en of de bijbehorende Microsoft Office-toepassing beschikbaar is. Sluit het bestand wanneer het door een andere toepassing is vergrendeld.

### Het genereren mislukt bij het opslaan

Sluit de bestaande uitvoer-PDF en probeer opnieuw. Controleer daarnaast of je schrijfrechten hebt op de projectmap.

Gebruik bij een probleem de exacte foutmelding uit het logvenster en vermeld de applicatieversie.

## Versie controleren

De versie staat in de titelbalk van de applicatie. Je kunt deze ook controleren via de Windows-eigenschappen van de `.exe`: rechtermuisknop op het bestand, **Eigenschappen**, tabblad **Details**.

