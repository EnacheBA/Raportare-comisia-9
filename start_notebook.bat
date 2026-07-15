@echo off
call conda activate raportare-omec-3019
if errorlevel 1 (
  echo Mediul raportare-omec-3019 nu a putut fi activat.
  echo Rulati: conda env create -f environment.yml
  pause
  exit /b 1
)
jupyter lab Raportare_OMEC_3019_2025.ipynb

