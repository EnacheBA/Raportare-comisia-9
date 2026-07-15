# Ghid complet de instalare și utilizare

Acest ghid descrie fluxul recomandat pe Windows. Interfața serviciilor externe
se poate modifica; comparați pașii cu documentația oficială indicată la final.

## 0. Înainte de a începe

Aveți nevoie de:

- Windows 10/11 pe 64 de biți sau un sistem compatibil cu Anaconda/Jupyter;
- spațiu liber pentru Anaconda, cache și fișierele Excel;
- un profil Google Scholar public;
- un cont SerpApi și o cheie API;
- acces instituțional autorizat la Web of Science;
- Microsoft Excel sau LibreOffice pentru verificarea rezultatelor;
- șablonul de raportare corespunzător Comisiei 9.

Faceți copii de siguranță. Nu utilizați singura copie a șablonului sau a
exporturilor.

## 1. Instalarea Anaconda și configurarea Jupyter Notebook

### 1.1 Descărcarea

1. Deschideți pagina oficială:
   https://www.anaconda.com/download
2. Descărcați **Anaconda Distribution – Windows 64-Bit Graphical Installer**.
3. Opțional, verificați hash-ul SHA-256 al kitului conform instrucțiunilor
   Anaconda.

### 1.2 Instalarea pe Windows

1. Deschideți kitul descărcat din directorul Downloads.
2. Selectați **Next**, citiți termenii și selectați **I Agree** dacă îi
   acceptați.
3. Alegeți **Just Me**, opțiunea recomandată pentru un utilizator obișnuit.
4. Alegeți un director fără caractere speciale; evitați căi foarte lungi.
5. Păstrați **Create shortcuts**.
6. Nu selectați adăugarea permanentă în PATH dacă nu știți că aveți nevoie;
   Anaconda recomandă utilizarea Anaconda Prompt/Navigator.
7. Finalizați instalarea.

Consultați ghidul oficial:
https://www.anaconda.com/docs/getting-started/anaconda/install/windows-gui-install

> Pentru utilizare într-o organizație, verificați termenii Anaconda și politica
> instituției. Proiectul nu acordă o licență pentru Anaconda Distribution.

### 1.3 Verificarea

Din meniul Start deschideți **Anaconda Prompt** și rulați:

```bash
conda --version
conda list
```

Dacă sunt afișate versiunea și lista pachetelor, instalarea este funcțională.

### 1.4 Descărcarea proiectului

Varianta Git:

```bash
git clone ADRESA_REPOSITORY-ULUI
cd raportare-omec-3019-2025-comisia-9
```

Varianta fără Git:

1. în GitHub selectați **Code → Download ZIP**;
2. extrageți arhiva într-un director propriu, de exemplu
   `D:\Raportare_OMEC_3019`;
3. nu lucrați direct din arhiva ZIP.

### 1.5 Crearea mediului izolat

În Anaconda Prompt, din rădăcina proiectului:

```bash
conda env create -f environment.yml
conda activate raportare-omec-3019
```

Pentru actualizarea unui mediu deja creat:

```bash
conda env update -f environment.yml --prune
```

Verificați kernelul:

```bash
python --version
python -c "import openpyxl, pandas, requests, xlrd; print('Mediu OK')"
```

### 1.6 Pornirea Jupyter

Din același director:

```bash
jupyter lab
```

sau:

```bash
jupyter notebook
```

În browser deschideți `Raportare_OMEC_3019_2025.ipynb`. Kernelul trebuie să
fie mediul `raportare-omec-3019`. Dacă nu apare:

```bash
python -m ipykernel install --user --name raportare-omec-3019 --display-name "Python (raportare OMEC 3019)"
```

## 2. Obținerea unei chei API SerpApi

### 2.1 Crearea contului

1. Deschideți https://serpapi.com/users/sign_up.
2. Creați contul și confirmați adresa de e-mail, dacă vi se solicită.
3. Deschideți https://serpapi.com/pricing.
4. Alegeți planul **Free**.
5. La 15 iulie 2026, planul Free este listat cu 250 de căutări/lună.
6. Deschideți Dashboard/Your Account și copiați cheia API privată.

Planurile și limitele se pot schimba. Planul Starter era listat la 25 USD/lună
pentru 1.000 de căutări, nu 25 EUR. Verificați prețul, moneda, taxele,
reînnoirea și politica de rambursare înainte de achiziție.

### 2.2 Protejarea cheii

Metoda recomandată este fișierul local `.env`:

1. copiați `.env.example` și redenumiți copia `.env`;
2. completați:

```text
SERPAPI_API_KEY=cheia_dumneavoastra
CROSSREF_MAILTO=adresa_email
```

3. nu puneți ghilimele dacă cheia nu le conține;
4. nu publicați fișierul; `.gitignore` îl exclude automat.

Notebookul poate solicita cheia printr-un câmp mascat dacă variabila nu este
setată. Nu introduceți cheia direct într-o celulă salvată.

### 2.3 Estimarea consumului

Indicele Hirsch nu este o unitate de facturare. Fiecare profil implică:

- cel puțin o căutare pentru lista publicațiilor;
- interogări pentru detaliile fiecărui articol;
- una sau mai multe pagini pentru citările fiecărui articol;
- reluări dacă forțați reîmprospătarea.

Pentru profiluri cu h-index sub 20, planul Free poate fi suficient în multe
cazuri, dar nu există garanție. Pentru profiluri mai mari sau cu multe citări,
urmăriți soldul și luați în calcul planul plătit.

Cache-ul reduce repetarea cererilor. Nu setați `FORCE_REFRESH=True` fără
motiv.

## 3. Pregătirea profilului Google Scholar

Profilul trebuie să fie public. Articolele unui profil privat nu sunt vizibile
public și nu pot fi colectate complet.

### 3.1 Crearea/completarea profilului

1. autentificați-vă în contul Google;
2. deschideți https://scholar.google.com;
3. selectați **My profile / Profilul meu**;
4. completați numele, afilierea, domeniile de interes și e-mailul instituțional;
5. adăugați publicațiile corecte și eliminați asocierile greșite;
6. combinați versiunile duplicate numai după verificare.

### 3.2 Activarea vizibilității publice

În pagina profilului:

1. utilizați opțiunea de editare a profilului;
2. identificați setarea **Profile visibility / Vizibilitatea profilului**;
3. activați **Make my profile public / Faceți profilul public**;
4. salvați modificarea;
5. deschideți URL-ul profilului într-o fereastră Incognito/Private în care nu
   sunteți autentificat;
6. verificați că numele și lista publicațiilor sunt vizibile.

Etichetele pot varia. Ajutor oficial:
https://scholar.google.com/intl/us/scholar/help.html

### 3.3 Copierea URL-ului

URL-ul trebuie să conțină parametrul `user`, de exemplu:

```text
https://scholar.google.com/citations?hl=en&user=IDENTIFICATOR&view_op=list_works
```

Notebookul extrage automat identificatorul. Nu folosiți URL-ul unei căutări
generale sau al unei publicații individuale.

## 4. Exportarea datelor Web of Science

Instrucțiunile complete sunt în
[EXPORT_WEB_OF_SCIENCE.md](EXPORT_WEB_OF_SCIENCE.md). Rezumat:

1. autentificați-vă în contul autorizat E-nformation;
2. accesați Web of Science – Core Collection;
3. căutați cercetătorul și validați setul;
4. exportați articolele în Excel cu **Full Record**;
5. creați **Citation Report**;
6. pentru fiecare lucrare deschideți lista de citări și adăugați toate paginile
   în **Marked List**;
7. exportați **Marked List → Documents** în Excel cu **Full Record**;
8. salvați fișierele ca `Articole_WOS.xls(x)` și `Citari_WOS.xls(x)`.

## 5. Organizarea directorului

Copiați fișierele în `data/input/`:

```text
data/input/
├── Articole_WOS.xls
├── Citari_WOS.xls
└── Criterii_CNATDCU_2025.xlsx
```

Sunt acceptate și extensiile .xlsx. Dacă numele diferă, modificați căile în
celula de configurare a notebookului.

## 6. Rularea notebookului

### Celula A – instalarea/verificarea bibliotecilor

Celula verifică pachetele. Într-un mediu creat cu `environment.yml`, nu ar
trebui să instaleze nimic.

### Celula B – configurarea

Completați:

- `PROFILE_URL`;
- căile exporturilor WoS și ale șablonului;
- `FORCE_REFRESH=False`;
- opțional limite mici pentru un test.

Pentru primul test:

```python
MAX_ARTICLES = 2
MAX_CITATIONS_PER_ARTICLE = 5
```

După validare, utilizați `None` pentru colectarea completă.

### Celula C – testul configurației

Verifică directoarele, profilul, existența intrărilor și prezența cheii fără a
afișa secretul.

### Celula D – Google Scholar

Creează imediat un fișier Excel gol valid, apoi scrie checkpointuri. În timpul
rulării, urmăriți:

- mesajele din notebook;
- fișierul .log;
- fișierul .status.json;
- Excelul parțial.

Dacă rularea se oprește, păstrați cache-ul și rulați din nou aceeași celulă.
Articolele deja salvate sunt încărcate din cache. Dacă SerpApi nu returnează
rezultate pentru un articol, acesta este păstrat cu informațiile disponibile,
iar fluxul continuă.

### Celula E – clasificarea J/C

Regula clasifică drept C o denumire care conține, fără diferență între litere
mari și mici:

- conference;
- symposium;
- proceedings;
- un acronim fără spații între paranteze rotunde.

Orice altă denumire nenulă devine J. Regula poate clasifica greșit și trebuie
verificată manual în `Articole_Google.xlsx`.

### Celula F – separarea WoS/Google

Corelează:

- `Articole_Google.xlsx`;
- `Articole_WOS.xls(x)`;
- `Citari_WOS.xls(x)`.

WoS are prioritate. Rândurile mutate în fișierul WoS sunt eliminate din
fișierul Google, astfel încât ieșirile să aibă intrări unice.

### Celula G – completarea raportului

Completează numai categoriile implementate. Articolele-părinte sunt bold;
pentru categoriile de citări sunt marcate cu fond verde și urmate de citările
lor. Rândurile Google se termină cu „Google Scholar”.

Șablonul-sursă nu este suprascris.

## 7. Checklist final

- [ ] profilul Google este public și corect;
- [ ] numărul articolelor exportate este plauzibil;
- [ ] toate rândurile galbene/incomplete au fost verificate;
- [ ] toate clasificările J/C au fost verificate;
- [ ] fiecare citare este sub articolul corect;
- [ ] autocitările au fost tratate conform cerințelor aplicabile;
- [ ] fișierele WoS și Google nu conțin duplicate;
- [ ] DOI-urile și ISSN-urile au fost confirmate în surse;
- [ ] raportul se deschide fără reparații în Excel;
- [ ] formulele nu conțin #REF!, #VALUE!, #DIV/0! sau #NAME?;
- [ ] totalurile au fost recalculate în Excel;
- [ ] punctajele corespund formei oficiale actualizate;
- [ ] a fost arhivată o copie a tuturor intrărilor și ieșirilor.

## 8. Probleme frecvente

### „Lipsește cheia SERPAPI_API_KEY”

Completați .env, reporniți kernelul sau introduceți cheia în câmpul mascat.

### „Google hasn't returned any results”

Fluxul actual trece la următorul articol și păstrează datele disponibile.
Consultați logul. Nu ștergeți cache-ul înainte de diagnostic.

### Rularea pare blocată

Verificați data ultimei modificări a fișierelor .log, .status.json și .xlsx.
Dacă nu există progres peste limita configurată, opriți kernelul și reluați
celula; checkpointurile rămân disponibile.

### Fișierul .xls nu poate fi citit

Confirmați că `xlrd` este instalat și fișierul este un Excel 97–2003 real,
nu HTML redenumit. Încercați deschiderea și salvarea ca .xlsx în Excel.

### Clasificări greșite

Corectați manual coloana Type din `Articole_Google.xlsx`, apoi reluați doar
etapele de separare și raportare.

### Formule nerecalculate

Deschideți rezultatul în Microsoft Excel, utilizați **Formulas → Calculate
Now**, salvați și verificați totalurile.

