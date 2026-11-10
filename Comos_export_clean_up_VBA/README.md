# COMOS Excel-exportok formázása

A makró a COMOS adatbázisból frissen exportált cső- és szeleplisták `Query` munkalapját tisztítja és formázza. Az eredmény az eredeti fájl mellett, új néven készül el. A makró egyszer telepíthető az Excel személyes makró-munkafüzetébe, így nem kell minden exportba bemásolni, és az export `.xlsx` formátumban maradhat.

## Gyors használat

Az egyszeri telepítés után:

1. Nyisd meg az eredeti exportot az asztali Excelben.
2. Kattints az export munkafüzetére, például a `Query` lapjára.
3. Nyomj **Alt+F8**-at, válaszd a `PERSONAL.XLSB!FormatComosExport` makrót, majd kattints a **Run / Futtatás** gombra. A megjelenő név a modul nevét is tartalmazhatja.
4. Várd meg a sikeres mentésről szóló üzenetet. Az üzenet a kimeneti útvonalat és a számmá alakított cellák számát is tartalmazza.
5. Ellenőrizd a látható fejlécet, a szűrőt és a számértékeket. OneDrive használatakor várd meg a szinkronizálás befejezését is.

A futtatás után már az új, formázott munkafüzet van megnyitva.

## Feltételek

- Windowsos asztali Microsoft Excel szükséges. A megoldás Excel webes verziójában és Macen nem futtatható.
- A makrók futtatását a vállalati biztonsági szabályoknak megfelelően engedélyezni kell. Az összes makró feltétel nélküli engedélyezése nem szükséges és nem javasolt.
- A forrásfájl legyen már elmentve, és tartalmazzon pontosan `Query` nevű munkalapot.
- A fejlécben pontosan `Line No.` szerepeljen. A makró az első sorban keresi; ha az első sor rejtett, akkor a második sorban.
- A munkalapok védelme legyen jelszó nélkül feloldható. A makró nem kerül meg jelszavas védelmet.
- Támogatott kiterjesztések: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`. A makró megtartja a munkafüzet fájlformátumát.
- OneDrive-/SharePoint-webcímmel megnyitott fájlnál az adott mappának szinkronizálva kell lennie a Windows OneDrive-kliensével, és a szinkronmappának elérhetőnek kell lennie.

A makró egy futtatással egy munkafüzetet dolgoz fel, nem vizsgál át rekurzívan egy teljes könyvtárat.

## Egyszeri telepítés a PERSONAL.XLSB-be

A `PERSONAL.XLSB` az Excel személyes, jellemzően rejtett makró-munkafüzete. Az Excel induláskor betölti, ezért a benne tárolt makrók más megnyitott munkafüzeteken is futtathatók.

### Ha még nincs személyes makró-munkafüzet

1. Excelben válaszd a **Developer / Fejlesztőeszközök → Record Macro / Makró rögzítése** parancsot.
2. A tárolási helynél válaszd a **Personal Macro Workbook / Személyes makró-munkafüzet** lehetőséget.
3. Indítsd el, majd azonnal állítsd le a rögzítést.
4. Az Excel bezárásakor engedélyezd a személyes makró-munkafüzet mentését.

Ha a Developer lap nem látható, a **File → Options → Customize Ribbon** beállításban kapcsolható be.

### A formázómodul importálása

1. Nyomj **Alt+F11**-et a VBA-szerkesztő megnyitásához.
2. Ha a bal oldali projektlista nem látható, nyomj **Ctrl+R**-t.
3. Jelöld ki a **VBAProject (PERSONAL.XLSB)** projektet. Ne az exportfájl projektjét válaszd.
4. Válaszd a **File → Import File…** parancsot, és importáld a [ComosExportFormatter.bas](ComosExportFormatter.bas) fájlt.
5. Ellenőrizd, hogy a **Modules** alatt megjelent a `ComosExportFormatter` modul.
6. Válaszd a **Debug → Compile VBAProject** parancsot, majd mentsd a személyes makró-munkafüzetet **Ctrl+S**-sel.
7. **Alt+F11**-gyel térj vissza Excelbe. Bezáráskor is engedélyezd a `PERSONAL.XLSB` mentését, ha az Excel rákérdez.

Importáláskor a `.bas` fájl első, `Attribute VB_Name = "ComosExportFormatter"` sora maradjon a fájlban: ezt az importáló kezeli.

### Meglévő verzió frissítése

1. A `PERSONAL.XLSB` projekt **Modules** részében jobb kattintás a régi `ComosExportFormatter` modulra, majd **Remove ComosExportFormatter…**.
2. Az exportálási kérdésre válaszolj **No**-val, ha nincs szükséged a régi kód mentésére. Más modulokat ne törölj.
3. Importáld az új `.bas` fájlt a fenti lépésekkel, fordítsd le a projektet, és mentsd a személyes munkafüzetet.

Alternatívaként a meglévő modul teljes tartalma kézzel is lecserélhető. Ilyenkor **az első `Attribute VB_Name…` sort ne másold be**: a kódszerkesztőben az első sor `Option Explicit` legyen. A további deklarációkat, köztük a `#If VBA7 Then` blokkot is másold át.

### Saját gomb a gyorselérési eszköztárra

1. Nyisd meg a **File → Options → Quick Access Toolbar** beállítást.
2. A parancsforrásnál válaszd a **Macros** lehetőséget.
3. Add hozzá a `PERSONAL.XLSB`-ben található `FormatComosExport` makrót.
4. A **Modify** gombbal válassz ikont és például `COMOS formázás` nevet.

Ezután az export aktiválása után egy gombnyomással indítható a feldolgozás. A személyes makró-munkafüzetnek ugyanabban az Excel-példányban kell megnyitva lennie, mint az exportnak.

## Mi történik futtatáskor?

1. Ellenőrzi az aktív munkafüzetet, a mentési útvonalat, a `Query` lapot és a `Line No.` fejlécet. Nem engedi a makrót tartalmazó munkafüzet vagy egy bővítmény formázását.
2. **Még a munkalapok módosítása előtt** új néven menti a megnyitott munkafüzetet. A további műveletek már ezen a másolaton történnek.
3. Megpróbálja feloldani az összes munkalap védelmét, üres jelszóval.
4. Aktiválja a `Query` lapot, és eltávolítja annak korábbi automatikus szűrőjét.
5. Törli a `Line No.` oszlop előtti összes oszlopot. A szeleplista mintájában például így tűnik el az első `SystemUID` oszlop.
6. A tényleges tartalom jobb szélétől balra haladva törli az összefüggő, rejtett záróoszlopokat. Az első látható oszlopnál megáll; közbenső rejtett oszlopokat nem töröl.
7. Ha a fejléc a második sorban volt, törli az előtte lévő rejtett technikai sort. Ellenőrzi, hogy az új `A1` értéke `Line No.`, és láthatóvá teszi az első sort.
8. A `Query` összes cellájáról eltávolítja a közvetlen kitöltést és a régi szegélyeket. Bekapcsolja a sortörést (**Wrap Text**) és a függőleges középre igazítást.
9. A lent leírt szabályok szerint valódi Excel-számmá alakítja a numerikus szövegkonstansokat.
10. Félkövérre állítja az első sort, és szűrőt kapcsol be a teljes adattartományon.
11. Az `A1`-től az utolsó tartalmas sorig és oszlopig vékony fekete külső és belső cellarácsot rajzol, átlós szegélyek nélkül. A tartományhatárt értékek és képletek alapján keresi, nem pusztán a korábbi formázások alapján.
12. Automatikus szélességet állít a látható adatoszlopokra, majd automatikus magasságot a látható adatsorokra.
13. Normál nézetre vált, megszünteti a korábbi rögzítést és felosztást, ismét láthatóvá teszi és méretezi az első sort, majd a bal felső sarokra görget. **Csak ezután rögzíti az első sort**, oszloprögzítés nélkül.
14. Menti a formázott fájlt, visszaállítja a számítási, eseménykezelési és képernyőfrissítési alkalmazásbeállításokat, majd összesítést mutat.

A `Query_CSI` fedőlap és az egyéb lapok tartalmát, formázását nem alakítja át; a védelem feloldása azonban minden munkalapot érint. A többi rejtett adatsort nem fedi fel, és a feltételes formázási szabályokat nem törli.

## Fájlnév és az eredeti megőrzése

Példa:

```text
Forrás:    Valve_List_MED.xlsx
Eredmény:  Valve_List_MED_formatted_261110_1435.xlsx
```

A név vége `_formatted_YYMMDD_HHMM`, ahol az időpont a futás elején, a kimeneti név képzésekor érvényes számítógépes helyi idő. A példában ez 2026. november 10., 14:35. A VBA-formátum `yymmdd_hhnn`, mert VBA-ban az `nn` jelenti a percet.

- Az eredeti fájlt a makró nem formázza át és nem írja felül: először `SaveAs` művelettel létrehozza az új fájlt, és csak utána változtat a lapokon.
- A másolat a forrás mappájában készül, az eredeti kiterjesztéssel.
- Ha ugyanaz a kimeneti fájlnév már létezik, a makró leáll, nem írja felül. Ugyanazon forráson ugyanabban a percben ismételt futtatáskor ez előfordulhat; a következő percben új név keletkezik.
- Új exporthoz mindig az eredeti exportot indítsd. Már formázott fájl újrafuttatásakor újabb `_formatted_…` utótag kerül a névhez.
- A megőrzés a makró által végzett módosításokra vonatkozik. A korábban kézzel végzett vagy AutoSave által már elmentett változásokat nem vonja vissza.

## OneDrive és SharePoint

Az Excel egy szinkronizált fájl elérési útját helyi Windows-útvonal helyett `https://…sharepoint.com/…` címként is visszaadhatja. Ezen a címen a hagyományos VBA `Dir` fájlellenőrzés nem működik.

A makró ezért a Windows aktuális felhasználójának OneDrive-bejegyzéseit olvassa:

```text
HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive
```

Az `UrlNamespace` és `MountPoint` értékek alapján a webcímet a hozzá tartozó szinkronmappára képezi le. Több találat esetén a leghosszabb illeszkedő webes gyökeret választja, és csak létező mappát fogad el. Az URL-kódolt szóközöket és UTF-8 ékezeteket a Windows `UrlUnescapeW` függvényével dekódolja. A registryt csak olvassa, nem módosítja.

**Nincs szükség a fájl kézi átmásolására vagy OneDrive-on kívüli mentésre.** Az eredmény ugyanabba a szinkronizált mappába kerül; a feltöltést a OneDrive-kliens végzi.

Ez nem közvetlen SharePoint-feltöltő vagy Microsoft Graph-kliens. Ha a mappa nincs szinkronizálva, nincs hozzá megtalálható megfeleltetés, vagy a vállalati szabályok tiltják a szükséges Windows-szolgáltatásokat, a makró nem tudja automatikusan feloldani az útvonalat. Ilyenkor a mappát kell szinkronizálni, illetve az elérési jogosultságokat ellenőrizni.

## Szövegből szám: pontos szabályok

A felismerés a [Python-szkript](convert_excel_cells_to_numbers.py) szigorú szabályát követi: opcionális előjel, legalább egy számjegy, majd opcionálisan pont és további számjegyek. A szöveg eleji és végi támogatott szóközjellegű karaktereket levágja; belső elválasztókat nem távolít el.

| Eredeti szöveg | Eredmény |
| --- | --- |
| `123` | Valódi szám: `123` |
| `00123` | Valódi szám: `123`; a vezető nullák eltűnnek |
| `000` | Valódi szám: `0` |
| `12.345` | Valódi szám, három megjelenített tizedesjeggyel |
| `12.340` | Valódi szám, három megjelenített tizedesjeggyel, a záró nulla megtartásával |
| ` -0.50 ` | Valódi szám, két megjelenített tizedesjeggyel |
| `90222243` | Valódi szám; ez az egyetlen dokumentumazonosítót tartalmazó PID-cellákra is érvényes |
| `90222243 / 90222272` | Változatlan szöveg |
| `DN 25`, `PN 16`, `M0` | Változatlan szöveg |
| `1,23`, `1 234`, `12%`, `1E3` | Változatlan szöveg |
| `1234567890123456` | Szöveg marad a pontossági korlát miatt |

A kimenet tizedeselválasztójának megjelenése az Excel területi beállításától függhet, de a ponttal érkező exportadatból valódi szám készül. A cella számalakja a bemenet tizedesjegyeinek számát követi.

Az Excel pontossági és tartománykorlátai miatt a makró kihagyja azokat a számszövegeket, amelyek a vezető nullák elhagyása után 15-nél több számjegyet tartalmaznak, túl sok tizedesjegyet igényelnek, vagy nem ábrázolhatók biztonságosan. A kihagyott numerikus szövegek számát külön jelzi. A nem numerikus szöveg nem ebbe a számlálóba tartozik.

A konverzió a `Query` teljes adattartományának szövegkonstansaira vonatkozik, nem csak kijelölt oszlopokra. Képleteket és már számtípusú cellákat nem módosít.

## Hibakeresés

A hibaüzenet `Step:` sora megmutatja, melyik lépésnél állt le a feldolgozás; az `Error` a hibaszámot és leírást, az `Output:` pedig az ismert kimeneti útvonalat tartalmazza. Hibabejelentéshez a teljes üzenetet érdemes elküldeni.

| Üzenet vagy jelenség | Teendő |
| --- | --- |
| `Compile error: Syntax error`, kiemelt `Attribute` sor | Kézi bemásolásnál töröld az `Attribute VB_Name…` sort. Importálásnál hagyd benne a `.bas` fájlban. |
| `Select the export, not the macro workbook` | A kód a `PERSONAL.XLSB`-ben legyen, az aktív munkafüzet pedig az export. Ne az export saját VBA-projektjéből indítsd ugyanazt a kódot. |
| `No matching OneDrive sync folder was found` | Ellenőrizd az érintett mappa szinkronizálását és a OneDrive működését. Nem szükséges OneDrive-on kívülre másolni a fájlt. |
| `Bad file name or number` | Ellenőrizd, hogy a friss modul fut-e, és nézd meg a `Step:` és `Output:` sorokat. A régi verzió webcímen végzett `Dir` ellenőrzése okozott ilyen hibát; más útvonalprobléma is kiválthatja. |
| `Output already exists` | Várj a következő percig, vagy nevezd át/helyezd át a korábbi kimenetet. |
| Nincs `Query` lap vagy nincs megfelelő `Line No.` fejléc | Ellenőrizd, hogy valóban a megfelelő export aktív-e, és a fájl megfelel-e az elvárt szerkezetnek. A makró nem találja ki a hiányzó fejléc helyét. |
| Jelszóval védett munkalap | Oldd fel a védelmet az ismert jelszóval, megfelelő jogosultsággal, majd indítsd újra az eredeti exportból. |
| Nem látszik az első sor | Ellenőrizd, hogy a friss verzió futott-e. Kézi ellenőrzéshez a névmezőbe írd `1:1`, majd **Home → Format → Hide & Unhide → Unhide Rows**. |
| `FreezePanes` / 1004 hiba | A friss verzió normál nézetre vált és a lap tetejére görget a rögzítés előtt. Ha így is hibázik, küldd el a teljes üzenetet; az ablakállapotot külön kell vizsgálni. |

Ha a hiba **a másolat létrehozása előtt** történt, a makró még nem kezdett formázni. Ha **utána** történt, az eredeti export nem lett formázva, de a `_formatted_…` másolat részlegesen elkészülhetett, és az AutoSave is menthetett rá. Zárd be a hibás másolatot, töröld azt, majd az eredeti exportból próbáld újra. A puszta „bezárás mentés nélkül” felhős AutoSave mellett nem garantálja a részleges változások eldobását.

## A mappa fájljai és ellenőrzés

- [ComosExportFormatter.bas](ComosExportFormatter.bas): az Excelbe telepítendő teljes formázómakró.
- [convert_excel_cells_to_numbers.py](convert_excel_cells_to_numbers.py): a számfelismerés alapjául szolgáló, önálló Python-segédprogram. Rekurzívan dolgoz fel `.xlsx` és `.xlsm` fájlokat, alapértelmezésben helyben ment, és támogatja a `--dry-run` kapcsolót. Nem a teljes VBA-formázási folyamatot valósítja meg; a makróhoz Python nem szükséges.
- [convert_to_number.vba](convert_to_number.vba): korábbi, bővebb számformátum-felismerést használó konverter. Nem ez a teljes COMOS-formázó, és nem szükséges mellé telepíteni.
- [Pipe_list_MED.xlsx](Pipe_list_MED.xlsx) és [Valve_List_MED.xlsx](Valve_List_MED.xlsx): referencia-exportok.
- [Test-ComosExportFormatter.ps1](Test-ComosExportFormatter.ps1): Excel COM-automatizálásra épülő integrációs tesztszkript. Külön Excel-példányt és ideiglenes munkafüzeteket nyit; nem része a napi használatnak. A VBA-projekt programozott eléréséhez külön Excel-biztonsági engedély szükséges lehet, ez nem követelménye a kézi makrófuttatásnak.

A legutóbbi OneDrive-útvonal-feloldást és URL-dekódolást Excel indítása nélkül ellenőriztük. A teljes integrációs teszt sikeres lefutása a jelenlegi verzióra nincs igazolva; a korábbi automatizált próbák RPC-hibákba ütköztek. A kódsorrend vagy a szerkesztői diagnosztika ellenőrzése nem helyettesíti a tényleges Exceles futtatást.

Új verzió kézi ellenőrzéséhez használj referencia-exportot, és nézd meg: az eredeti fájl megmaradt-e, a fejléc látható-e az első sorban, görgetéskor rögzítve marad-e, a szűrő a teljes táblázatot lefedi-e, a vezető nullás számok átalakultak-e, a `/` jellel elválasztott PID-k változatlanok-e, valamint az eredmény újranyitás után is megfelelő-e.