# Exportarea articolelor și citărilor din Web of Science

Instrucțiunile descriu interfața cunoscută la 15 iulie 2026. Denumirile și
poziția butoanelor pot fi modificate de Clarivate.

## 1. Accesul

1. deschideți https://www.e-nformation.ro/;
2. autentificați-vă cu un cont asociat unei instituții care are acces;
3. din profilul de acces selectați resursa Clarivate/Web of Science;
4. accesați pachetul care include **Web of Science – Core Collection, InCites
   Journal Citation Reports, Derwent Innovations Index, Clarivate Analytics**;
5. în Web of Science selectați explicit **Web of Science Core Collection**, nu
   „All Databases”, dacă metodologia dumneavoastră cere Core Collection.

Accesul depinde de abonamentul instituției. Dacă resursa nu apare, contactați
biblioteca sau responsabilul instituțional.

## 2. Obținerea fișierului Articole_WOS

### 2.1 Căutarea

1. în câmpul de căutare alegeți **Author**;
2. introduceți numele cercetătorului în forma cerută de interfață;
3. repetați sau rafinați după variante de nume, inițiale, afiliere, ORCID,
   ResearcherID și interval de ani;
4. eliminați lucrările altor persoane cu nume similar;
5. includeți toate lucrările relevante și documentați filtrele aplicate.

O căutare doar după nume poate produce rezultate false pozitive sau poate
omite variante ale numelui.

### 2.2 Exportul articolelor

1. în lista de rezultate selectați **Export**;
2. alegeți **Excel**;
3. în fereastra de export selectați **Records from 1 to n**, unde n este
   ultimul rezultat relevant;
4. la **Record Content** selectați **Full Record**;
5. executați **Export**;
6. salvați fișierul în `data/input/` ca:
   - `Articole_WOS.xls`, sau
   - `Articole_WOS.xlsx`.

Dacă interfața limitează numărul de înregistrări pe export, exportați intervale
consecutive și reuniți-le într-un singur registru, cu un singur rând de antet.

## 3. Obținerea fișierului Citari_WOS

### 3.1 Citation Report

1. reveniți la setul validat de publicații;
2. în partea superioară selectați **Citation Report**;
3. notați:
   - numărul publicațiilor;
   - Times Cited;
   - Citing Articles;
   - eventualele valori fără autocitări, dacă sunt afișate;
4. păstrați o captură sau un jurnal al datei și setului de filtre.

**Times Cited** poate fi mai mare decât numărul de documente citante unice,
deoarece același document poate cita mai multe articole ale cercetătorului.
Pentru exportul unei liste fără duplicate trebuie urmărit numărul de
**Citing Articles** distincte, nu simpla sumă a tuturor celulelor.

### 3.2 Adăugarea citărilor în Marked List

Pentru fiecare articol din tabelul Citation Report:

1. identificați numărul de citări din indexul ales;
2. faceți clic pe număr pentru a deschide **Citing Articles**;
3. pe pagina rezultatelor selectați **Add to Marked List**;
4. alegeți **All records on page**;
5. dacă există mai multe pagini, navigați pe fiecare pagină și repetați;
6. reveniți la Citation Report și continuați cu articolul următor;
7. nu ștergeți lista marcată între articole.

Dacă un document citează două articole, Marked List îl păstrează o singură
dată. Acesta este comportamentul dorit pentru un fișier cu citări unice.

### 3.3 Verificarea Marked List

1. selectați **Marked List** din partea stângă;
2. deschideți fila/secțiunea **Documents**;
3. verificați numărul documentelor;
4. comparați-l cu valoarea **Citing Articles** distincte din Citation Report;
5. dacă există diferențe, verificați:
   - pagini neadăugate;
   - filtre diferite;
   - baze/indexuri diferite;
   - documente duplicate;
   - acces limitat la anumite înregistrări.

Nu presupuneți automat că suma citărilor pe articole trebuie să fie egală cu
numărul documentelor din Marked List.

### 3.4 Exportul citărilor

1. din **Marked List → Documents** selectați **Export**;
2. alegeți **Excel**;
3. selectați **Records from 1 to n**;
4. la **Record Content** selectați **Full Record**;
5. executați exportul;
6. salvați în `data/input/` ca:
   - `Citari_WOS.xls`, sau
   - `Citari_WOS.xlsx`.

Clarivate documentează exportul listelor marcate și menționează limite care
pot ajunge la 1.000 de înregistrări per operațiune, în funcție de opțiune.
Pentru volume mai mari, exportați intervale și reuniți-le controlat.

## 4. Controlul calității

- deschideți ambele fișiere și verificați anteturile;
- nu eliminați coloanele din Full Record înainte de rularea proiectului;
- nu schimbați denumirile câmpurilor;
- păstrați exporturile originale nemodificate;
- notați data exportului și baza/indexul selectat;
- verificați DOI-urile și titlurile pentru un eșantion;
- arhivați Citation Report ca dovadă a reconcilierii.

## 5. Surse oficiale

- Export Records:
  https://webofscience.help.clarivate.com/en-us/Content/export-records.htm
- Creating a Citation Report:
  https://webofscience.zendesk.com/hc/en-us/articles/20016866431121-Creating-a-Citation-Report
- Citation Reports:
  https://webofscience.zendesk.com/hc/en-us/articles/20016829622033-Citation-Reports
- Saving and Exporting Marked Lists:
  https://webofscience.zendesk.com/hc/en-us/articles/20135824927505-Saving-and-Exporting-Marked-Lists
- Întrebări frecvente E-nformation:
  https://www.e-nformation.ro/intrebari-frecvente

