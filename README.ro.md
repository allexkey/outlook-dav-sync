# outlook-dav-sync

Sincronizare în ambele sensuri între un **server CalDAV / CardDAV** și **Legacy Outlook pentru Mac**.

Outlook pentru Mac nu se poate conecta la servere CalDAV sau CardDAV, deși Calendar și Contacts de la Apple pot. Acest program rulează în fundal și ține sincronizate cu serverul un calendar și o agendă de contacte din secțiunea **On My Computer** a Outlook. Servere tipice: cPanel, Nextcloud, Radicale, Baïkal.

*English: [README.md](README.md)*

> Proiectul nu este afiliat cu Microsoft și nu este aprobat de Microsoft.

## Cerințe

| Cerință | Detalii |
|---|---|
| macOS | Testat pe macOS 27. |
| Outlook | **Legacy Outlook**, testat cu 16.115. „New Outlook” nu poate fi controlat din scripturi; comuți din **Outlook → Legacy Outlook**. |
| Python | 3.9 sau mai nou. Se folosește doar biblioteca standard, deci nu trebuie instalat nimic cu `pip`. |
| Server | Orice server CalDAV/CardDAV accesibil prin HTTPS, cu utilizator și parolă. Testat cu cPanel (`cpdavd`, portul 2080). |

## Instalare

```bash
git clone <adresa-repo> outlook-dav-sync
cd outlook-dav-sync
./install.sh
```

Ce face instalarea:

1. Verifică cerințele.
2. Copiază programul în `~/Library/Application Support/outlook-dav-sync/bin`. După instalare poți șterge folderul clonat.
3. Cere adresa serverului și utilizatorul.
   - Merge orice adresă de pe server, de exemplu `https://mail.domeniu.ro:2080/`.
   - Parola intră direct în **Keychain-ul macOS**. Nu e scrisă niciodată într-un fișier.
4. Găsește calendarele și agendele de pe server și te lasă să alegi ce sincronizezi.
5. Face o **rulare de probă**, care arată ce s-ar schimba fără să modifice nimic. Apoi te întreabă înainte de prima sincronizare reală.
6. Instalează un job în fundal (launchd) care sincronizează la fiecare 3 minute, cât timp Outlook e deschis.

macOS va întreba o singură dată dacă Python poate **controla Microsoft Outlook**. Apasă **OK**. Dacă ai apăsat „Don't Allow”, reactivezi permisiunea din **System Settings → Privacy & Security → Automation**.

Opțiuni utile:

```bash
./install.sh --reconfigure     # schimbi serverul, contul sau colecțiile
./install.sh --interval 300    # sincronizare la 5 minute
```

Pentru actualizare, rulezi din nou `./install.sh` dintr-o versiune nouă a repo-ului. Setările și starea sincronizării se păstrează.

## Utilizare

Lucrezi în cele două foldere create în Outlook. Implicit se numesc **CalDAV Calendar** și **CardDAV Contacts**; la instalare le poți da alt nume. Modificările circulă în ambele sensuri în câteva minute.

```bash
APP="$HOME/Library/Application Support/outlook-dav-sync/bin/davsync.py"
python3 "$APP" run              # sincronizează acum
python3 "$APP" run --dry-run    # arată ce s-ar schimba
python3 "$APP" discover         # colecțiile de pe server
tail -f ~/Library/Logs/outlook-dav-sync.log
```

Dacă se schimbă parola de e-mail, o înlocuiești în Keychain:

```bash
security add-generic-password -s outlook-dav-sync -a tu@domeniu.ro -U -w
```

## Reguli de sincronizare

**Detectarea modificărilor**
- Pe server, prin ETag.
- În Outlook, prin **conținut**, nu doar prin data modificării. Outlook atinge singur unele elemente, de exemplu șterge memento-urile evenimentelor trecute, iar asta nu contează ca editare.

**Ce se trimite din Outlook spre server**
- La evenimente: titlul, locul, descrierea, orele, „toată ziua”, recurența și excepțiile ei.
- La contacte: numele, firma, funcția, telefoanele, e-mailurile, adresele, paginile web, nota și ziua de naștere.
- Restul datelor de pe server rămân neatinse: participanți, organizator, memento-uri, poza contactului.

**Conflicte.** Dacă un element e modificat în ambele locuri între două rulări, **câștigă serverul**. Versiunea din Outlook se salvează în `backups/`.

**Ștergeri**
- Orice element șters se salvează întâi ca backup.
- O rulare care ar șterge mai mult de 20% din elemente se oprește. Dacă ștergerea e intenționată, rulezi o dată cu `--allow-mass-delete`.

**Prima rulare.** Înainte de orice altceva, toate elementele de pe server se salvează în `backups/initial-…`.

**Memento-urile nu se sincronizează.** Fiecare parte își păstrează memento-urile ei.

## Dezinstalare

```bash
./uninstall.sh
```

Dezinstalarea oprește jobul din fundal. Apoi te întreabă separat dacă vrei să ștergi folderele din Outlook, parola din Keychain, și setările împreună cu backup-urile. Pe server nu se modifică nimic.

## Depanare

| Problemă | Soluție |
|---|---|
| `Outlook bridge failed … not allowed` | Permite Python în **System Settings → Privacy & Security → Automation → Microsoft Outlook**. |
| `No password in Keychain` | Rulează comanda `security add-generic-password` de mai sus. |
| `HTTP 401` | Utilizator sau parolă greșite. Rulează `./install.sh --reconfigure`. |
| `refusing to delete N items` | Verifică dacă ștergerea e intenționată, apoi rulează o dată cu `--allow-mass-delete`. |

Licență: [MIT](LICENSE).
