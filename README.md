# TOOL PDF Merger

Windows-app voor het samenvoegen van rapportages en bijlagen. Zie [README.txt](README.txt) voor de gebruikershandleiding.

## Ontwikkelen in VS Code

1. Installeer Python 3.10 of nieuwer (met tkinter) en Git. Voor de volledige conversie is Microsoft Office nodig.
2. Open deze **map** in VS Code en installeer de aanbevolen extensies.
3. Kies via `Ctrl+Shift+P` **Python: Create Environment**, kies **Venv** en selecteer `requirements-dev.txt`. Kies daarna zo nodig **Python: Select Interpreter** en selecteer `.venv`.
4. Met **Tasks: Run Task** kun je de taken hieronder starten. De setup-taak kan dependencies later opnieuw installeren.

| Taak | Functie |
| --- | --- |
| Setup: dependencies installeren | Installeert runtime-dependencies en Ruff in de lokale `.venv`-omgeving |
| Code: formatteren | Maakt de Python-opmaak consistent |
| Code: opmaak controleren | Controleert opmaak zonder wijzigingen |
| Code: lint controleren | Controleert imports en veelvoorkomende codefouten |
| Tests: uitvoeren | Voert de unittest-tests uit zonder Office te starten |
| Controle: alles | Voert opmaakcontrole, lint en tests achtereenvolgens uit; ook via `Ctrl+Shift+B` |
| App: starten | Start de desktopapp voor een handmatige test |
| GitHub: push huidige branch (na controles) | Pusht commits van de huidige branch naar `origin`, uitsluitend als alle controles slagen |

De taken gebruiken rechtstreeks `.venv/Scripts/python.exe` in de projectmap. Python wordt bij opslaan automatisch geformatteerd. Met `F5` kun je debuggen; tests zijn ook beschikbaar in het Testing-paneel.

## Naar GitHub pushen

De remote `origin` hoort te verwijzen naar https://github.com/TheEngineeringToolbox/TOOL_PDF_MERGER.git.
Controleer je wijzigingen in Source Control en maak daar eerst een commit. Voer daarna de pushtaak uit. Deze maakt geen commit en gebruikt geen force-push. GitHub-aanmelding en schrijfrechten op de repository zijn vereist.

## Testbereik

De automatische tests controleren bijlage- en bestandsvolgorde, genegeerde tijdelijke bestanden, rapportvalidatie en paginavolgorde bij PDF-samenvoeging. Test Office-conversie en de GUI handmatig met een voorbeeldrapport en bijlagen. De centrale template- en logopaden staan bovenin `app.py` en moeten voor die test bereikbaar zijn.

## EXE bouwen en versiebeheer

Voer eerst **Setup: dependencies installeren** uit. Start daarna **Build: EXE maken** via **Tasks: Run Task**. Deze taak controleert opmaak, lint en tests en bouwt vervolgens met PyInstaller. `build/build_exe.bat` gebruikt dezelfde bouwprocedure en de lokale `.venv`.

`APP_VERSION` in `app.py` is de centrale versiebron voor de app en de Windows EXE-eigenschappen. Gebruik **Versie: patch verhogen** voor reparaties, **Versie: minor verhogen** voor functionaliteit of **Versie: major verhogen** voor incompatibele wijzigingen. Controleer en commit de versiewijziging; een build verhoogt de versie niet.

Elke build komt in een eigen map onder `dist/`, met versie, Git-commit, eventuele `dirty`-markering en UTC-buildtijd. `build-info.json` bevat herkomst en SHA-256; `dependencies.txt` legt de gebruikte pakketten vast. Bestaande builds worden niet verwijderd. Nieuw gegenereerde bestanden onder `temp/` en `dist/` worden genegeerd door Git; historische, al gevolgde bestanden blijven voorlopig gevolgd.

### Branches en releases

Ontwikkel op `codex/development`; `master` blijft de stabiele branch. De bestaande pushtaak pusht de huidige branch. Maak voor een release een pull request van `codex/development` naar `master`. Bouw na het samenvoegen vanuit een schone checkout van `master`, test de EXE met Office en maak vervolgens een Git-tag `vX.Y.Z` bij die commit. Publiceer de geteste EXE bij die release. Builds met `dirty` bevatten lokale wijzigingen en zijn ontwikkelbuilds. De taken maken geen tags, commits of GitHub-releases automatisch.

## Licentie

TOOL PDF Merger valt onder de aangepaste **TOOL Engineers B.V. Free Use and Redistribution License 1.0** in `LICENSE.txt`.

Samengevat mag de software gratis worden gebruikt, ook intern binnen commerciële organisaties, en gratis in ongewijzigde vorm worden doorgegeven met behoud van de licentie- en copyrightvermeldingen. Zonder voorafgaande schriftelijke toestemming van **TOOL Engineers B.V.** mag de software niet worden verkocht, verhuurd, tegen betaling toegankelijk worden gemaakt, in een betaald softwarepakket worden opgenomen of in gewijzigde vorm worden verspreid. De volledige tekst in `LICENSE.txt` is leidend.

Derde-partijcomponenten behouden hun eigen licenties. Zie `THIRD_PARTY_NOTICES.txt`. Voer na installatie van `requirements-dev.txt` handmatig `python build/check_licenses.py` uit om de actuele runtime-licenties te controleren. Dezelfde controle draait automatisch in GitHub Actions en als onderdeel van de releasebuild.

Een geslaagde releasebuild plaatst naast de EXE ook `LICENSE.txt`, `THIRD_PARTY_NOTICES.txt` en een uit de actieve buildomgeving gegenereerd `THIRD_PARTY_LICENSES.txt`. Nieuwe dependencylicenties die niet in de gecontroleerde allow-list vallen blokkeren de compliancecheck totdat ze expliciet zijn beoordeeld.
