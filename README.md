# Doslista

En webbsida för att söka och kontrollera läkemedelsrader från `Doslista.xlsx`.

## Starta

Sidan är statisk och behöver inte Vite, Node.js eller npm för att köras. Öppna projektmappen i VS Code och starta `index.html` med tillägget **Live Server**. Öppna inte filen direkt som `file://`, eftersom webbläsaren då blockerar läsningen av Excel-filen. HyperFormula och SheetJS laddas från jsDelivr, så internetanslutning behövs när sidan öppnas.

Arbetsboken innehåller läkemedelsdata. Håll därför GitHub-repot privat om informationen inte ska vara offentlig; en publicerad webbsida gör också arbetsboken tillgänglig för besökare.

## Uppdatera läkemedelslistan

Lägg till eller ändra läkemedel i `Doslista.xlsx` och ersätt filen i projektmappen. Gränssnittet har ingen uppladdningsfunktion. Alla flikar med kolumnerna **Benämning**, **Substans** och **Kategori ATC** blir tillgängliga i djurväljaren; nya djurflikar läggs därmed till automatiskt.

Sökningen omfattar Benämning, Substans och Kategori ATC. När en rad väljs kan vikt anges i fältet för kolumn M. Resultatkolumnerna N–T hämtas från radens Excel-formler via HyperFormula; sidan skapar inte egna dosformler eller skriver tillbaka ändringar till originalfilen.
