# Aktindsigtsrobot – Procesbeskrivelse

Denne robot er bygget til at automatisere afgørelsesprocessen for aktindsigtsanmodninger i OpenOrchestrator-frameworket. Den benytter kø-baseret afvikling og interagerer med både SharePoint og KMD Nova.

## Formål

Robotten håndterer DeskPro-baserede sager ved at:
- Indsamle metadata fra køelementer og eksterne API'er
- Udtrække dokumentoversigter fra SharePoint
- Generere og udfylde afgørelsesdokumenter baseret på gældende lovgivning
- Uploade dokumenter til korrekte mapper i SharePoint
- Opdatere sagen og tilhørende opgaver i KMD Nova

---

## Procesoverblik

### 1. Initialisering
Ved opstart hentes nødvendige konstanter og credentials:
- SharePoint URL og login
- API-nøgler til KMD Nova og Aktbob
- Tidsstempel for KMD-token valideres og fornyes ved behov

### 2. Købehandling
For hvert køelement indlæses:
- DeskPro ID og titel
- Navn og e-mail på ansøger
- Afdeling og lovgrundlag
- Modtagelsesdato for anmodningen

### 3. Dataindsamling
- Dokumentlisten (Excel) lokaliseres i SharePoint under `Dokumentlister/{DeskProTitel}`
- Dokumentlisten analyseres for titler, aktindsigtsvurderinger og begrundelser
- Eventuel beskrivelse hentes fra ekstern Aktbob API

### 4. Afgørelsesdokument
- En hovedskabelon vælges ud fra afdeling og lovgrundlag
- Pladsholdere som `[Ansøgernavn]`, `[Deskprotitel]` mv. udfyldes
- Relevante begrundelser mappes til minifrases og flettes ind ved `[RELEVANTE_TEKSTER]`
- Interne dokumentbegrundelser samles og indsættes som punktopstilling

### 5. Upload og link
- Det færdige dokument gemmes som `Afgørelse.docx` og uploades til SharePoint:
  `Aktindsigter/{DeskProTitel}`
- Der sendes et link til mappen tilbage til DeskPro via API

### 6. KMD Nova-opdatering
Via `AfslutSag.py`:
- Der hentes relevante CaseUuid'er
- Sagen opdateres til status "Afsluttet", og attributter som titel, dato og kategori udfyldes
- Tre specifikke opgaver lukkes
  - “05. Klar til sagsbehandling”
  - “25. Afslut/henlæg sagen”
  - “11. Tidsreg: Sagsbehandling”

## Bemærkninger

- Dokument- og begrundelsesskabeloner kan tilpasses i `doc_map_by_lovgivning` i `process.py`
- Midlertidige filer slettes automatisk medmindre andet angives
  

