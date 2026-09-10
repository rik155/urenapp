# UrenApp MVP

Mobiele urenregistratie voor medewerkers met centrale Excel-uitvoer.

## Wat werkt al
- Naam + weeknummer
- Uren per dag en meerdere klussen per dag
- Optioneel 0-urencontract met veld `Vrije uren`
- Werknemer ziet geen periode
- Excel rekent periode automatisch uit in blokken van 4 weken
- Elke medewerker krijgt automatisch een eigen tabblad
- Tabblad `Overzicht`
- Tabblad `Periode-overzicht`
- Opnieuw indienen van dezelfde medewerker/week vervangt de vorige inzending
- Kantooroverzicht via `/admin`
- Excel downloaden via `/api/excel`
- PWA-basis zodat de website als app-icoon op telefoon kan worden gezet

## Starten
```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```
Of via Docker:
```bash
docker build -t urenapp .
docker run -d --name urenapp -p 8000:8000 -v $(pwd)/data:/app/data urenapp
```

## Belangrijk voor de definitieve live Excel-koppeling
Deze MVP schrijft direct naar `/data/urenregistratie.xlsx` op de server. Voor *hetzelfde Excel-bestand dat kantoor al in Microsoft 365 gebruikt* voegen we in de volgende stap Microsoft Graph / OneDrive of SharePoint toe. De app-architectuur is hiervoor voorbereid; het serverbestand kan dan vervangen worden door de centrale Microsoft 365-workbook-opslag.
