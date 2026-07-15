# Confidențialitate și protecția datelor

Proiectul rulează local, dar transmite către servicii externe interogări
necesare colectării și îmbogățirii metadatelor:

- SerpApi primește identificatorul profilului și interogările Google Scholar;
- Crossref poate primi titluri, autori și adresa de contact configurată;
- Google Scholar este sursa rezultatelor accesate prin serviciul selectat.

## Date care pot apărea în fișiere

Exporturile și rapoartele pot conține nume, afilieri, adrese de e-mail,
identificatori, titluri, linkuri și alte metadate despre autori și persoane
care citează. Unele date pot constitui date cu caracter personal.

## Măsuri incluse

- cheia API nu este scrisă în notebook sau în fișierele de stare;
- fișierul .env este exclus din Git;
- directoarele data/input, data/intermediate, data/output și data/cache sunt
  excluse din Git;
- proiectul păstrează datele local, în directoarele indicate;
- utilizatorul poate șterge cache-ul și rezultatele după realizarea copiilor
  necesare.

## Obligațiile utilizatorului

Înainte de publicare sau partajare:

- eliminați cheile, tokenurile și adresele private;
- verificați dacă aveți dreptul să distribuiți exporturile WoS;
- minimizați sau anonimizați datele cu caracter personal când este necesar;
- stabiliți temeiul și perioada de păstrare conform regulilor instituției și
  legislației aplicabile, inclusiv GDPR unde este cazul;
- nu publicați rapoarte personale fără acord și fără verificare.

Acest document este informativ și nu constituie consultanță juridică.

