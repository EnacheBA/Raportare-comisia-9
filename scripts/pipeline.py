"""Orchestrarea etapelor proiectului de raportare OMEC 3019/2025."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from .classifica_articole import clasifica_fisier
from .completeaza_criterii_cnatdcu import complete_workbook
from .construieste_articole_citari_wos import (
    construieste_fisiere,
    verifica_fisiere_generate,
)
from .google_scholar_export import ExportConfig, run_export


@dataclass(frozen=True)
class ProjectPaths:
    root: Path

    @property
    def input_dir(self) -> Path:
        return self.root / "data" / "input"

    @property
    def intermediate_dir(self) -> Path:
        return self.root / "data" / "intermediate"

    @property
    def output_dir(self) -> Path:
        return self.root / "data" / "output"

    @property
    def cache_dir(self) -> Path:
        return self.root / "data" / "cache" / "google_scholar"

    @property
    def google_raw(self) -> Path:
        return self.intermediate_dir / "Articole_Google_neclasificat.xlsx"

    @property
    def google_classified(self) -> Path:
        return self.intermediate_dir / "Articole_Google.xlsx"

    @property
    def wos_unique(self) -> Path:
        return self.output_dir / "Articole_Citari_WoS.xlsx"

    @property
    def google_unique(self) -> Path:
        return self.output_dir / "Articole_Citari_Google.xlsx"

    @property
    def report(self) -> Path:
        return self.output_dir / "Criterii_CNATDCU_2025_completat.xlsx"

    def prepare(self) -> None:
        for directory in (
            self.input_dir,
            self.intermediate_dir,
            self.output_dir,
            self.cache_dir,
        ):
            directory.mkdir(parents=True, exist_ok=True)


def author_id_from_profile_url(profile_url: str) -> str:
    """Extrage parametrul user din URL-ul public Google Scholar."""
    parsed = urlparse(profile_url.strip())
    if parsed.netloc.casefold() not in {"scholar.google.com", "scholar.google.ro"}:
        raise ValueError("URL-ul trebuie să fie un profil Google Scholar.")
    author_id = (parse_qs(parsed.query).get("user") or [""])[0].strip()
    if not author_id:
        raise ValueError("URL-ul nu conține parametrul user al profilului.")
    return author_id


def run_google_scholar(
    paths: ProjectPaths,
    *,
    profile_url: str,
    serpapi_api_key: str | None = None,
    crossref_mailto: str = "",
    backend: str = "serpapi",
    force_refresh: bool = False,
    max_articles: int | None = None,
    max_citations_per_article: int | None = None,
) -> dict[str, Any]:
    """Colectează profilul și citările; salvează checkpointuri după fiecare articol."""
    paths.prepare()
    key = (serpapi_api_key or os.getenv("SERPAPI_API_KEY", "")).strip()
    if backend == "serpapi" and not key:
        raise ValueError("Lipsește cheia SERPAPI_API_KEY.")
    config = ExportConfig(
        author_id=author_id_from_profile_url(profile_url),
        profile_url=profile_url,
        output_path=paths.google_raw,
        cache_dir=paths.cache_dir,
        backend=backend,
        serpapi_api_key=key,
        crossref_mailto=crossref_mailto,
        force_refresh=force_refresh,
        max_articles=max_articles,
        max_citations_per_article=max_citations_per_article,
        progress_interval=15,
        scholar_stall_timeout=180,
        checkpoint_every_citations=10,
    )
    return run_export(config)


def run_classification(paths: ProjectPaths) -> dict[str, Any]:
    paths.prepare()
    return clasifica_fisier(paths.google_raw, paths.google_classified)


def run_unique_split(
    paths: ProjectPaths,
    *,
    articole_wos: str | Path,
    citari_wos: str | Path,
    fuzzy_threshold: float = 0.90,
) -> dict[str, Any]:
    paths.prepare()
    result = construieste_fisiere(
        articole_google=paths.google_classified,
        articole_wos=articole_wos,
        citari_wos=citari_wos,
        output_wos=paths.wos_unique,
        output_google=paths.google_unique,
        prag_fuzzy_citari=fuzzy_threshold,
        verbose=True,
    )
    result["verificare"] = verifica_fisiere_generate(paths.wos_unique, paths.google_unique)
    return result


def run_report(
    paths: ProjectPaths,
    *,
    template_path: str | Path,
) -> dict[str, Any]:
    paths.prepare()
    return complete_workbook(
        template_path=template_path,
        wos_path=paths.wos_unique,
        google_path=paths.google_unique,
        output_path=paths.report,
    )


def check_required_inputs(*paths: str | Path) -> None:
    missing = [str(Path(path)) for path in paths if not Path(path).expanduser().exists()]
    if missing:
        raise FileNotFoundError("Lipsesc fișierele:\n- " + "\n- ".join(missing))

