# Každodenní obsluha a řešení chyb

## Význam stavů

| Stav | Význam |
| --- | --- |
| `K ODESLÁNÍ` | Připraveno pro příští běh. Hodnota musí být přesná. |
| `ZPRACOVÁVÁ SE` | Automat řádky rezervoval. Nevracet automaticky. |
| `ODESLÁNO` | Označená zpráva byla potvrzena ve složce Odeslaná pošta Classic Outlooku. |
| `ČÁSTEČNĚ ODESLÁNO` | Jeden příjemce uspěl a druhý selhal. |
| `CHYBA VALIDACE` | Chybí údaj nebo jsou duplicitní řádky v rozporu. |
| `CHYBA` | Outlook zprávu nepotvrdil v Odeslané poště nebo nastala jiná provozní chyba. |

`K ODESLÁNÍ` je současně schválení obsahu. Dokud je tato hodnota nastavena,
neupravujte příjemce, oslovení, žadatele, číslo RM, projekt, výzvu ani dotaci.
Před opravou nejprve odstraňte stav, proveďte změnu a kontrolu a teprve potom vraťte
přesné `K ODESLÁNÍ`. Automat před každým žadatelem kontroluje, že se schválená data
od načtení nezměnila.

Podrobný stav a datum se zapisují zvlášť pro starostu a tajemníka. Hodnota za
svislítkem je identifikátor konkrétního pokusu.

## Denní kontrola

1. Ověřte poslední výsledek úlohy **SIOLA Email Automation** v Plánovači úloh.
2. V tabulce vyfiltrujte `CHYBA`, `CHYBA VALIDACE`, `ČÁSTEČNĚ ODESLÁNO` a
   `ZPRACOVÁVÁ SE`.
3. Provozní logy jsou v
   `%LOCALAPPDATA%\SIOLA Email Automation\data\logs`.

Při selhání se přihlášenému uživateli zobrazí upozornění a v datové složce vznikne
soubor `ACTION_REQUIRED.txt`. Logy starší než 90 dní se automaticky mažou.

## Bezpečné opakování

Před změnou chybového stavu vždy vyhledejte adresáta a předmět v Outlooku ve složce
**Odeslaná pošta** a zkontrolujte také **Pošta k odeslání**.

- Pokud zpráva existuje, příjemci ji znovu neposílejte.
- Pokud prokazatelně neexistuje, opravte příčinu a změňte obecný `Stav` na
  `K ODESLÁNÍ` u **všech řádků stejného žadatele, které patří do společného
  e-mailu**. Nikdy nevracejte jen jeden projekt z víceprojektového e-mailu.
- Dílčí stav `ODESLÁNO` pro již úspěšného příjemce ponechte. Automat jej přeskočí.
- Pokud si nejste jistí, nic neresetujte a úlohu vypněte.

Stará hodnota `ZPRACOVÁVÁ SE` znamená, že běh mohl skončit mezi Outlookem a zápisem
do Google tabulky. Automat takový řádek záměrně sám neopakuje.

## Bezpečnostní zámek

Po tvrdém vypnutí počítače může zůstat soubor `automation.lock`. Před jeho
odstraněním je nutná kontrola Odeslané pošty, Pošty k odeslání a všech řádků
`ZPRACOVÁVÁ SE`. Nejprve v Plánovači úloh ověřte, že úloha neběží. Skript odmítne
odstranit zámek, pokud jej stále drží aktivní proces.
Odstranění provede správce příkazem:

```powershell
pwsh -File .\Invoke-SiolaAutomation.ps1 -ForceUnlock
```

Nikdy nenastavujte automatické mazání zámku ani automatické vracení chybových stavů
na `K ODESLÁNÍ`.

## Přesun na jiný počítač

1. Na starém počítači zakažte úlohu **SIOLA Email Automation**.
2. Na novém počítači proveďte instalaci, VALIDATE a TEST včetně kontroly náhledu.
3. Spusťte `TAKE_OVER.cmd` a napište přesně `PREVZIT`.

Tabulka se tím přiřadí nové instalaci. I kdyby starý počítač zůstal omylem zapnutý,
jeho další LIVE běh bude odmítnut dříve, než rezervuje řádky nebo odešle zprávu.
