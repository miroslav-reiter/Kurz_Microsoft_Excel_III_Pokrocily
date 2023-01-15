# 📊 Kurz Microsoft Excel III. Pokročilý

[Kurz Microsoft Excel III. Pokročilý](https://www.it-academy.sk/kurz/microsoft-excel-iii-pokrocily/) je pre teba vhodný, ak máš skúsenosti s Excelom alebo si absolvoval kurz Microsoft Excel II. Mierne pokročilý. Naučíme ťa modifikovať používateľské rozhranie a markantne urýchliť svoju prácu. Osvojíš si zabezpečenie dát v tabuľkách, tvorbu formulárov či vyhľadávacie funkcie tzv. lookupy. Tvorba reportov bude pre teba samozrejmosťou. Ako absolvent kurzu Microsoft Excel III. Pokročilý zvládneš pokročilé analýzy dát, prácu s maticami či kontingenčnými tabuľkami.

## ❓ Čo je to Microsoft Excel
Microsoft Excel je **tabuľkový procesor od** spoločnosti Microsoft navrhnutý pre operačný systém **Microsoft Windows**, **Mac Os**, **Android** a **iO**S. Je súčasťou **kancelárskeho balíka Microsoft Office** spolu s aplikáciami Microsoft Word, Microsoft PowerPoint, Microsoft Outlook, Microsoft Access atď.

## 🙋 Verzie a edície Microsoft Excel
Najaktuálnešia/najnovšia verzia je verzia **Microsoft Excel 365 (Office 365)**. Na trhu sú aj standalone verzie : 2000, 2002, 2003, 2007, 2010, 2013, 2016, 2019

**TIP:** Verzie zistíme na Karte Domov (Home) > Konto (Account) > Čo je Excel
![verzia](https://user-images.githubusercontent.com/24510943/212565132-3a9892b7-d660-4e8e-b883-45794a06fc50.png)


## ⚓ Odkazy na kurzy
[Prezenčné Kurzy Microsoft Excel](https://www.it-academy.sk/kategoria/kancelarske-baliky/kurzy-excel/)  
[Online Kurzy Microsoft Excel](https://www.vita.sk/?s=excel)  

## 📁 Súbory a Materiály
Dostupné na GitHube alebo na kurze od lektora

## 🧰 Stránky a nástroje na precvičovanie Microsoft Excel
1. [Microsoft 365](https://www.microsoft.com/sk-sk/microsoft-365/excel)
2. [ASAP Utilities](http://www.asap-utilities.com/excel-tips-shortcuts.php)
3. [Microsoft Excel Alza Návod](https://www.alza.sk/microsoft-excel-navod)
4. [FinStat Firmy s najväčšími tržbami](https://finstat.sk/databaza-financnych-udajov?sort=sales-desc&years=2020)
5. [FinStat Najziskovejšie Firmy](https://finstat.sk/databaza-financnych-udajov?sort=profit-desc&years=2020)
6. [FinStat Najväčší zamestnávatelia](https://finstat.sk/databaza-firiem-organizacii?sort=empl-desc)
7. [FinStat Najväčšie univerzity](https://finstat.sk/databaza-neziskoviek?sort=revenue-desc&tab=revenue&legalform=382)

## 📔 Dokumentácia Microsoft Excel a Guidelines
1. [Microsoft Excel help & learning](https://support.microsoft.com/en-us/excel)
2. [Premium templates](https://templates.office.com/en-us/premium-templates)
3. [Analyze Data in Microsoft Excel](https://support.microsoft.com/en-us/office/analyze-data-in-excel-3223aab8-f543-4fda-85ed-76bb0295ffc4)
4. [Microsoft Excel functions (alphabetical)](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
5. [The Ultimate Guide to Using Microsoft Excel](https://blog.hubspot.com/marketing/microsoft-excel)

## 📎Obsah Kurzu
### 📑 Microsoft Excel má 3 typy Hárkov (Sheets)
1. **Pracovný Hárok (Worksheet) (Shift + F11)**
2. Makro Hárok (Macro Sheet) (Ctrl + F11)
3. Grafový Hárok (Graph Sheet) (F11)

![harky](https://user-images.githubusercontent.com/24510943/212564384-aa4f4b9a-1b41-419b-b67a-6b5dfa0053cc.png)

### 🔥 Duplikácia a Kopírovanie Formátu
* Hromadné Vkladanie, Generovanie Hodnôt (Ctrl + Enter)
* Kopírovanie Formátu Metlička (2-klik na metlu)

## 💡 Snippety
### 🗔 Zobraz Prehľadové Okno s Hárkami 
```vb
Sub WbTab()
' Zobraz Prehľadové Okno s Hárkami (Taby)
    Application.CommandBars("Workbook tabs").ShowPopup
End Sub
```
