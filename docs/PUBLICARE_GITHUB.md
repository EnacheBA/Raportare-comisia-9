# Publicarea proiectului pe GitHub

## 1. Verificări înainte de publicare

1. lucrați numai în directorul acestui proiect;
2. confirmați că .env nu este prezent în lista Git;
3. confirmați că directoarele data nu conțin fișiere personale care urmează
   să fie publicate;
4. căutați chei API și eliminați-le din toate fișierele și din istoricul Git;
5. modificați ADRESA_REPOSITORY-ULUI în README.md și CITATION.cff;
6. completați datele de contact pentru raportarea privată a vulnerabilităților;
7. rulați:

       python -m unittest discover -s tests -v
       python tools/check_notebook.py

## 2. Crearea repository-ului

În GitHub:

1. selectați New repository;
2. introduceți un nume, de exemplu raportare-omec-3019-2025-comisia-9;
3. alegeți Public;
4. nu adăugați automat README, .gitignore sau license, deoarece există deja;
5. creați repository-ul.

## 3. Inițializarea locală

Din Anaconda Prompt, în rădăcina proiectului:

    git init
    git branch -M main
    git add .
    git status

Examinați lista. Nu continuați dacă apar .env, fișiere .xls/.xlsx reale,
cache, loguri sau date personale.

    git commit -m "Versiunea inițială a proiectului"
    git remote add origin ADRESA_REPOSITORY-ULUI.git
    git push -u origin main

## 4. După publicare

- deschideți Actions și verificați că testele trec;
- verificați randarea notebookului direct în GitHub;
- verificați licențele și CITATION.cff;
- creați o versiune Release și arhivați checksum-ul;
- nu atașați exporturi personale la Release;
- publicați schimbările normative numai după verificarea sursei oficiale.

## 5. Dacă a fost publicată o cheie

1. revocați/regenerați cheia la furnizor imediat;
2. eliminați cheia din fișiere;
3. curățați istoricul Git cu un instrument adecvat, de exemplu git-filter-repo;
4. forțați actualizarea repository-ului numai după coordonarea contribuitorilor;
5. presupuneți că cheia veche a fost copiată și nu o reutilizați.

