# Politica de securitate

## Secrete

Nu transmiteți și nu publicați:

- cheia SerpApi;
- parole sau cookie-uri;
- fișierul .env;
- exporturi personale care nu trebuie distribuite.

Dacă o cheie a fost introdusă într-un commit, considerați-o compromisă:
revocați/regenerați cheia în contul furnizorului și eliminați secretul din
istoricul Git. Ștergerea doar din ultimul commit nu este suficientă.

## Raportarea vulnerabilităților

Nu includeți secrete sau date personale într-un issue public. Contactați
menținătorul printr-un canal privat indicat în profilul repository-ului.
Includeți pașii de reproducere, versiunea și impactul estimat.

## Fișiere Excel

Deschideți numai exporturi provenite din surse de încredere. Proiectul nu
execută macrocomenzi, însă fișierele Office pot conține legături externe sau
conținut activ. Scanați fișierele și păstrați copii de siguranță.

