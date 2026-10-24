# Doslista

En webbsida för att söka och kontrollera läkemedelsrader från `Doslista.xlsx`.

## Starta

```sh
npm install
npm run dev
```

Öppna adressen som Vite visar. Sidan läser `Doslista.xlsx` från projektmappen när den startar.

## Uppdatera läkemedelslistan

Lägg till eller ändra läkemedel i `Doslista.xlsx`. För att publicera ändringarna byter du ut arbetsboken i projektmappen och startar om eller distribuerar om webbsidan. Gränssnittet har ingen uppladdningsfunktion. Alla flikar med kolumnerna **Benämning**, **Substans** och **Kategori ATC** blir tillgängliga i djurväljaren; nya djurflikar läggs därmed till automatiskt.

Sökningen omfattar Benämning, Substans och Kategori ATC. När en rad väljs kan vikt anges i fältet för kolumn M. Resultatkolumnerna N–T hämtas från radens Excel-formler via HyperFormula; sidan skapar inte egna dosformler eller skriver tillbaka ändringar till originalfilen.
