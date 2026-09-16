# Každodenní obsluha a řešení chyb

## Význam stavů

| Stav | Význam |
| --- | --- |
| `K ODESLÁNÍ` | Připraveno pro příští běh. Hodnota musí být přesná. |
| `ZPRACOVÁVÁ SE` | Automat řádky rezervoval. Nevracet automaticky. |
| `ODESLÁNO` | Označená zpráva byla potvrzena ve složce Odeslaná pošta Classic Outlooku. |
| `KONTAKTOVÁNO JINÝM PROJEKTEM` | Řádek patří k jiné výzvě stejného žadatele; alespoň jeden e-mail pro jeho první výzvu byl potvrzen jako odeslaný. |
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

Má-li žadatel více výzev ve stavu `K ODESLÁNÍ`, rozhoduje nejnižší číslo řádku.
E-mail obsahuje pouze projekty z výzvy na tomto řádku. Řádky ostatních výzev se
rezervují společně, ale jejich dílčí stavy starosty a tajemníka se nemění. Stav
`KONTAKTOVÁNO JINÝM PROJEKTEM` se jim zapíše až po potvrzeném kontaktu alespoň
jednoho příjemce. Při úplném neúspěchu dostanou stejně jako vybraná výzva stav
`CHYBA`, aby se další výzva další den neodeslala místo ní.

## Denní kontrola

1. Ověřte poslední výsledek úlohy **SIOLA Email Automation** v Plánovači úloh.
2. V tabulce vyfiltrujte `CHYBA`, `CHYBA VALIDACE`, `ČÁSTEČNĚ ODESLÁNO` a
   `ZPRACOVÁVÁ SE`.
3. Provozní logy jsou v
   `%LOCALAPPDATA%\SIOLA Email Automation\data\logs`.

Při selhání se přihlášenému uživateli zobrazí upozornění a v datové složce vznikne
soubor `ACTION_REQUIRED.txt`. Logy starší než 90 dní se automaticky mažou.
Chyby validace vracejí Plánovači úloh nenulový výsledek, i když se platné zprávy
předtím bezpečně zpracovaly. Zbývající zprávy mohou být odloženy před dosažením
pětihodinového limitu běhu; zůstanou `K ODESLÁNÍ` pro další den.

## Dočasné vypnutí automatu

Automat lze pozastavit bez odinstalace a bez změny tabulky:

1. V nabídce Start vyhledejte a otevřete **Plánovač úloh**.
2. Vlevo otevřete **Knihovna Plánovače úloh**.
3. Vyberte úlohu **SIOLA Email Automation**.
4. Vpravo klikněte na **Zakázat**.
5. Ověřte, že je úloha zakázaná. Řádky `K ODESLÁNÍ` zůstanou připravené a při
   zakázané úloze se automaticky neodešlou.

Zakázání zabrání budoucím spuštěním, ale **nezastaví právě probíhající běh**. Pokud
Plánovač ukazuje stav **Spuštěno**, běžně nechte automat doběhnout a sledujte jeho
výsledek. Tlačítko **Ukončit** použijte jen v nouzi: tvrdé ukončení může nastat mezi
odesláním zprávy a zápisem výsledku do tabulky. Potom automat znovu nezapínejte,
neměňte řádky `ZPRACOVÁVÁ SE` a postupujte podle částí **Bezpečné opakování** a
**Bezpečnostní zámek** níže.

Pro opětovné zapnutí spusťte v instalační složce postupně `VALIDATE.cmd`,
`TEST.cmd`, zkontrolujte testovací zprávy a úplný HTML náhled a nakonec spusťte
`ENABLE_LIVE.cmd`. Nezapínejte úlohu přímo v Plánovači; tento postup znovu ověří
obsah, program i výchozí podpis Classic Outlooku.

## Když počítač není v naplánovaný čas připravený

Úloha je nastavená na jeden denní čas, na dodatečné spuštění po zmeškaném čase a
jen pro přihlášeného uživatele, který provedl instalaci. Sama počítač neprobudí.

| Situace v naplánovaný čas | Co se stane |
| --- | --- |
| Počítač je vypnutý | Nic se neodešle. Po zapnutí a přihlášení správného uživatele Windows zařadí zmeškaný běh; obvykle začne přibližně po 10 minutách. |
| Počítač spí | Automat počítač neprobudí. Zmeškaný běh se zařadí po probuzení, jakmile je k dispozici přihlášená relace správného uživatele. |
| Počítač běží, ale uživatel je odhlášený | Běh čeká na přihlášení tohoto uživatele. |
| Obrazovka je zamknutá, ale uživatel zůstal přihlášený | Zamknutí není odhlášení; Plánovač může úlohu spustit. Classic Outlook musí být v této relaci správně nakonfigurovaný. |
| Není internet nebo Outlook/Google nejsou dostupné | Běh bezpečně skončí chybou a vytvoří log a `ACTION_REQUIRED.txt`. Automat nemá nastavené automatické opakování po chybě; další běžný pokus je následující den. |
| Předchozí běh ještě pokračuje | Nový běh se souběžně nespustí; nový požadavek Plánovač ignoruje. |

Dodatečný běh nezačne nutně ihned po zapnutí nebo probuzení. Windows takové běhy
řadí do fronty; jeho standardní zpoždění je přibližně 10 minut. Počítač proto po
přihlášení hned nevypínejte. Výsledek vždy ověřte podle části **Denní kontrola**.

Pokud se počítač restartuje nebo vypne **během odesílání**, nejde jen o zmeškaný
čas. Přerušený běh se automaticky neopakuje a může po něm zůstat stav
`ZPRACOVÁVÁ SE` nebo soubor `automation.lock`. Před jakýmkoli ručním opakováním
zkontrolujte Odeslanou poštu, Poštu k odeslání, logy a dotčené řádky podle pokynů
níže.

## Bezpečné opakování

Před změnou chybového stavu vždy vyhledejte adresáta a předmět v Outlooku ve složce
**Odeslaná pošta** a zkontrolujte také **Pošta k odeslání**.

- Pokud zpráva existuje, příjemci ji znovu neposílejte.
- Pokud prokazatelně neexistuje, opravte příčinu a změňte obecný `Stav` na
  `K ODESLÁNÍ` u **všech řádků stejného žadatele, které patří do společného
  e-mailu**. Nikdy nevracejte jen jeden projekt z víceprojektového e-mailu.
- Pokud úplně selhal kontakt pro žadatele s více výzvami a stav `CHYBA` proto
  dostaly i jeho další výzvy, vraťte na `K ODESLÁNÍ` všechny tyto dotčené řádky
  společně. Jinak by se po opravě správně nezapsal stav
  `KONTAKTOVÁNO JINÝM PROJEKTEM`.
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
