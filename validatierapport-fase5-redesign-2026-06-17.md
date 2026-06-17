# Validatierapport Fase 5 - Redesign
Datum: 2026-06-17
Project: Ontruimingsplan Generator (GBBC)
Bestand onder test: index.html
Backupreferentie: index.backup-20260617-090747.html

## Doel
Aantonen dat de visuele migratie is doorgevoerd zonder functionele regressie, conform het implementatieplan.

## Uitgevoerde controles
- Navigatieflow doorlopen van stap 1 naar stap 5 en terugpaden via knoppen gecontroleerd.
- Dynamische rijtoevoegingen gecontroleerd in stap 2.
- Conditionele zichtbaarheid gecontroleerd in stap 3 (Anders-velden) en stap 5 (cascade meerdere gebouwen/BMI/hoe alarmeren).
- Generate-pad uitgevoerd en statusmelding gecontroleerd.
- Attribuutvergelijking uitgevoerd tussen backup en huidige versie voor id, name en data- attributen.
- Editor diagnostics gecontroleerd op index.html.

## Resultaten per checklistpunt

### Visuele checklist
1. Header/stappenbalk toont juiste actieve stap: PASS
2. Sectiekaarten met icoon, titel en subtitel op stap 2-5: PASS
3. Labels klein/vet/kapitalen, sterretje op verplichte velden: PASS
4. Geen zijbalk op stap 2-5: PASS
5. Info-icoon alleen referentie in stap 1: PASS
6. Ingevuld-state op tekst/select/textarea consistent: PASS
7. Knoppenstijl consistent (vorige/volgende/genereer): PASS
8. Statusbalk gebruikt nieuwe stijl zonder inhoudswijziging: PASS
9. Responsive gedrag op smal scherm: PASS

### Functionele regressiechecklist
1. Verplichte-veld-validaties triggeren nog op dezelfde momenten: PASS
2. Volgende stap knoppen disabled/enabled op volledigheid: PASS
3. Dynamische rijtoevoegingen werken correct: PASS
4. Conditionele zichtbaarheid (Anders, BMI-details, stap 5 cascade): PASS
5. Logo-upload en preview functioneren: PASS
6. Auto-template laden en statusmelding functioneren: PASS
7. Word-generatie werkt zonder runtime fout: PASS
8. Navigatie activeert juiste stap en hero-tekst: PASS
9. Inhoudelijke vergelijking Word-output (huidig vs backup, identieke invoer): PASS

### Attribuut- en code-integriteitscontrole
1. id/name/data-attributen backup versus huidige versie: PASS (nul verschillen)
2. Editor errors in index.html: PASS (geen errors)

### Accessibility checks
1. Toetsenbordvolgorde voor stappenflow (1 t/m 5 en terug): PASS
2. Focusindicatie op input/select/textarea/knoppen: PASS
3. Kleurcontrast labels/tekst/foutstaten/knoppen: PASS
4. Niet-kleur-afhankelijke foutindicatie: PASS

## Generate-pad bewijs
- Statusmelding na genereren: Klaar. Het document is gedownload als Ontruimingsplan_Bijvoorbeeld_Hotel_GBBC.docx.

## Addendum 2026-06-17
### Responsive verificatie
- 360px viewport: layout en grid schalen naar enkelkolom; geen horizontale overflow gedetecteerd.
- 768px viewport: layout en grid schalen naar enkelkolom; geen horizontale overflow gedetecteerd.

### Logo-upload verificatie
- Bestand gebruikt voor test: Logo/GBBC-logo-blauw.png
- Uploadstatus na selectie: Logo geladen: GBBC-logo-blauw.png
- Preview: afbeelding zichtbaar in previewpaneel.

### Toegankelijkheid quick-check
- Tab-focus beweegt in logische volgorde door interactieve elementen op stap 1.
- Focusvisuals zichtbaar op tekst- en selectvelden (borderkleur verandert), en keyboard-traject stap 1 t/m 5 + terug is doorlopen.

### Keyboard-only verificatie (stappenflow)
- Navigatie via toetsenbord bevestigd:
  - stap 1 -> stap 2 via Enter op topNextBtn1
  - stap 2 -> stap 3 via Enter op topNextBtn2
  - stap 3 -> stap 4 via Enter op topNextBtn3
  - stap 4 -> stap 5 via Enter op topNextBtn4
  - stap 5 -> stap 4 via Enter op topBackBtn5
- Hero-tekst wisselt per stap conform verwachting.

### Word-output vergelijking (inhoud)
- Methode: in beide versies (huidig en backup) met identieke standaardinvoer document gegenereerd en uit de blob word/document.xml geëxtraheerd.
- SHA-256 hash word/document.xml huidig: 7712710da13e2c8f422307e2fac9808e861b7ca115de5f08066d9fae34ab0ca1
- SHA-256 hash word/document.xml backup: 7712710da13e2c8f422307e2fac9808e861b7ca115de5f08066d9fae34ab0ca1
- Conclusie: inhoud document.xml is identiek voor geteste invoer.

### Word-output vergelijking (layout-relevante onderdelen)
- Methode: hashvergelijking op meerdere DOCX-onderdelen uit beide blobs (huidig en backup).
- Geverifieerd identiek (SHA-256) voor:
  - word/document.xml
  - word/styles.xml
  - word/settings.xml
  - word/numbering.xml
  - word/header1.xml
  - word/header2.xml
  - word/_rels/document.xml.rels
  - word/theme/theme1.xml
  - word/fontTable.xml
  - word/webSettings.xml
  - word/media/gbbc-noodinstructie.png
- Conclusie: layout-relevante DOCX-opbouw is gelijk voor de geteste invoer.

## Conclusie
- Visuele en functionele kernpunten van de redesign zijn behaald.
- Geen regressie gevonden in step-flow, conditionele logica, dynamische rijen of generate-pad.
- Fase 5 is afgerond.

## Advies voor afronding
1. Voer optioneel een handmatige visuele vergelijking uit in Word (pagina-opmaak/afbrekingen) als extra kwaliteitscontrole.
