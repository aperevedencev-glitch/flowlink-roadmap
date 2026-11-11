"""Загрузка настроек из config.toml."""
from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


@dataclass
class Feed:
    law: str              # "44", "223", "615"
    endpoint_law: str     # "44" или "223": какой сервер ЕИС
    subsystem: str        # RGK, RD223, RD615
    document_types: list[str]
    kind: str = "contract"            # "protocol" — протоколы закупок 223-ФЗ
    regions: list[str] | None = None  # свой список регионов (["all"] — все)
    lookup_protocols: bool = False    # договор без поставщика → протоколы по номеру закупки


@dataclass
class Config:
    root: Path
    raw: dict
    feeds: list[Feed] = field(default_factory=list)

    # --- пути -------------------------------------------------------------
    def path(self, key: str) -> Path | None:
        value = self.raw.get("paths", {}).get(key, "")
        if not value:
            return None
        p = Path(value)
        return p if p.is_absolute() else self.root / p

    # --- ЕИС --------------------------------------------------------------
    @property
    def eis(self) -> dict:
        return self.raw.get("eis", {})

    @property
    def token(self) -> str:
        return os.environ.get("EIS_TOKEN") or self.eis.get("token", "")

    def endpoint(self, endpoint_law: str) -> str:
        key = f"{self.eis.get('access', 'LE')}_{endpoint_law}"
        return self.eis["endpoints"][key]

    @property
    def regions(self) -> list[str]:
        return expand_regions(self.eis.get("regions", []))

    def feed_regions(self, feed: "Feed") -> list[str]:
        return expand_regions(feed.regions) if feed.regions else self.regions

    # --- период -----------------------------------------------------------
    @property
    def competitive_from(self) -> date:
        return date.fromisoformat(self.raw.get("period", {}).get("competitive_from", "2017-07-01"))

    @property
    def noncompetitive_from(self) -> date:
        return date.fromisoformat(self.raw.get("period", {}).get("noncompetitive_from", "2026-03-01"))

    @property
    def matching(self) -> dict:
        return {"date_tolerance_days": 3, "price_tolerance_pct": 1.0, **self.raw.get("matching", {})}

    @property
    def alerts(self) -> dict:
        return self.raw.get("alerts", {})

    @property
    def nostroy(self) -> dict:
        return self.raw.get("nostroy", {"mode": "file"})


# Коды субъектов РФ, как в ЕИС (orgRegion). Сверьте со справочником ЕИС.
ALL_REGIONS = [f"{i:02d}" for i in range(1, 80)] + ["80", "81", "83", "84", "85", "86", "87",
                                                     "89", "90", "91", "92", "93", "94", "95", "99"]


def expand_regions(values) -> list[str]:
    values = [str(v) for v in (values or [])]
    if any(v.lower() == "all" for v in values):
        return list(ALL_REGIONS)
    return [v for v in values if v.lower() != "wins"]


def load_config(path: str | Path) -> Config:
    path = Path(path)
    with open(path, "rb") as fh:
        raw = tomllib.load(fh)
    feeds = [
        Feed(
            law=str(f["law"]),
            endpoint_law=str(f.get("endpoint_law", "44")),
            subsystem=f["subsystem"],
            document_types=list(f.get("document_types", [])),
            kind=f.get("kind", "contract"),
            regions=[str(r) for r in f["regions"]] if f.get("regions") else None,
            lookup_protocols=bool(f.get("lookup_protocols", False)),
        )
        for f in raw.get("eis", {}).get("feeds", [])
    ]
    return Config(root=path.parent.resolve(), raw=raw, feeds=feeds)
