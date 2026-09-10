# Instalace na počítači s Windows

Instalaci proveďte nejprve v bezpečném režimu. Naplánovaná úloha se během instalace
vytvoří jako vypnutá a žádný ostrý e-mail se neodešle.

## 1. Přepnutí na Classic Outlook

V New Outlook vypněte přepínač **New Outlook** v pravém horním rohu a potvrďte
návrat. Pokud přepínač není dostupný, vyhledejte v nabídce Start aplikaci
**Outlook (classic)**. Classic Outlook musí být nainstalovaný a zvolený odesílající
účet v něm musí umět ručně odeslat zprávu. Ručně odeslaná
zpráva se musí uložit do složky **Odeslaná pošta**; bez ukládání odeslaných kopií
automat z bezpečnostních důvodů nepotvrdí odeslání.

Microsoft uvádí, že New Outlook nepodporuje Outlook Object Model ani COM. Tato
automatizace proto s New Outlook nefunguje:
https://support.microsoft.com/en-us/outlook/getstarted/feature-comparison-between-new-outlook-and-classic-outlook

## 2. Vytvoření přístupu ke Google tabulce

Tento krok provede správce Google účtu:

1. Otevřete https://console.cloud.google.com/ a vytvořte samostatný projekt,
   například `SIOLA-email-automation`.
2. V **APIs & Services > Library** zapněte **Google Sheets API**.
3. V **IAM & Admin > Service Accounts** vytvořte účet bez dalších rolí projektu.
4. U účtu otevřete **Keys > Add key > Create new key > JSON** a soubor bezpečně
   stáhněte. Neposílejte jej e-mailem a neukládejte jej do sdílené složky.
5. Z JSON souboru zjistěte hodnotu `client_email`.
6. V Google tabulce použijte **Sdílet** a tomuto `client_email` udělte roli
   **Editor**. Nikomu dalšímu se oprávnění nemění.

Klíč dovoluje měnit sdílenou tabulku. Pokud unikne, v Google Cloud jej okamžitě
smažte a vytvořte nový.

## 3. Instalace PowerShellu 7

Soubor `windows/INSTALL.cmd` se pokusí PowerShell 7 nainstalovat pomocí Windows
Package Manageru. Pokud instalace není povolena, požádejte správce počítače o
instalaci **Microsoft PowerShell 7**. Windows může během automatické instalace
zobrazit žádost o oprávnění správce; tu je nutné potvrdit.

## 4. Instalace automatu

1. Zkopírujte celou složku projektu na cílový počítač.
2. Dvakrát klikněte na `windows/INSTALL.cmd`.
3. Vyberte stažený Google service-account JSON klíč.
4. Zadejte Google Sheet URL, odesílající účet, vlastní testovací e-mail, kontaktní
   údaje podpisu a denní čas spuštění. Instalátor nemá žádné předvyplněné provozní
   identifikátory ani osobní údaje.
5. Instalátor zkopíruje provozní soubory do
   `%LOCALAPPDATA%\SIOLA Email Automation` a omezí přístup ke konfiguraci, klíči,
   logům a náhledům na aktuálního uživatele a systémový účet.
6. Až po úspěšném vytvoření vypnuté úlohy nabídne trvalé smazání původního JSON
   klíče. Smazání vyžaduje napsat přesně `SMAZAT`; soubor se nepřesouvá do Koše.
   Pokud jej ponecháte, bezpečně jej odstraňte později.
7. Naplánovaná úloha **SIOLA Email Automation** zůstane vypnutá.

Úloha běží pouze tehdy, když je tento uživatel ve Windows přihlášený. Classic
Outlook nemusí zůstat otevřený; musí však být správně nakonfigurovaný pod stejným
uživatelem.

## 5. Povinná kontrola VALIDATE

V instalační složce dvakrát klikněte na `VALIDATE.cmd`. Režim pouze čte tabulku.
Nic nemění a nic neodesílá.

Opravte všechny hlášené chyby. Kontrola vyžaduje tyto přesné hlavičky v prvním
řádku listu `Obce a města`:

`Výzva`, `Číslo RM`, `Žadatel`, `Název akce`, `Dotace (Kč)`,
`Oslovení - TAJEMNÍK`, `Email - TAJEMNÍK`, `Oslovení - STAROSTA`,
`Email - STAROSTA`, `Stav`, `Stav STAROSTA`, `Datum e-mailu STAROSTA`,
`Stav TAJEMNÍK`, `Datum e-mailu TAJEMNÍK`.

Sloupce `Datum e-mailu STAROSTA` a `Datum e-mailu TAJEMNÍK` jednou nastavte přes
**Formát > Číslo > Datum a čas**. Automat zapisuje nativní číselnou hodnotu data;
bez formátu by Google tabulka mohla zobrazit pouze pořadové číslo.

## 6. Povinný TEST

Spusťte `TEST.cmd`. Automat zkontroluje všechny připravené žadatele a vytvoří jeden
HTML soubor s náhledem všech zpráv. Skutečně odešle pouze zprávy pro nejvýše první
tři žadatele a všechny přesměruje na testovací adresu zadanou při instalaci.
Tabulku nezmění.

U každé zprávy zkontrolujte:

- žlutý TEST proužek a původního příjemce;
- správný předmět, obec, projekty, výzvu a součet dotací;
- starosta má `Oslovení - STAROSTA`;
- tajemník má `Oslovení - TAJEMNÍK`;
- tajemník je samostatná zpráva, nikoli CC;
- Calibri 12, tučné pasáže a celý podpis.

V automaticky otevřeném HTML náhledu navíc projděte nebo vyhledejte všechny obce,
příjemce a oslovení. Úspěšný TEST vytvoří potvrzení platné 24 hodin.
Potvrzení obsahuje otisk všech připravených zpráv a provozních souborů. Jakákoli
změna připravených řádků, podpisu nebo programu proto vyžaduje nový TEST. Náhled se
po zapnutí LIVE odstraní; bez zapnutí se smaže při nejbližším dalším spuštění,
jakmile je starší než dva dny.

Pokud Outlook zobrazí bezpečnostní dotaz na programové odesílání, nepovolujte ostrý
režim, dokud správce neověří zabezpečení Outlooku a aktuální antivirus.

## 7. Zapnutí LIVE

Až po schválení testů spusťte `ENABLE_LIVE.cmd`. Skript znovu provede
VALIDATE. Potom požádá o napsání přesného slova `LIVE` a teprve následně zapne
naplánovanou úlohu. Při prvním zapnutí zároveň tabulku interně přiřadí této jediné
instalaci. Druhá instalace se stejnou tabulkou bude bezpečně odmítnuta.
Pokud právě běží stará instalace, převzetí se odmítne až do vypršení jejího
běhového zámku. Starou úlohu přesto vždy nejprve ručně vypněte.

Výchozí dávka je 50 žadatelů za den. Starosta a tajemník mohou znamenat až dvě
samostatné zprávy na jednoho žadatele.
Automat označí vybrané řádky dočasnými interními značkami, které se při řazení
přesunou spolu s řádkem. Před každým e-mailem znovu ověří úplný obsah, členství
řádků, svou rezervaci a běhový zámek. Bezpečné seřazení proto nezamění cílový řádek;
změna schváleného obsahu automat zastaví.
