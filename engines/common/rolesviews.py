"""Acesso ao RolesViews.xlsx (assets/). Devolve sempre um BytesIO novo, com cursor em 0."""
import io
from functools import lru_cache
from pathlib import Path

_ASSET = Path(__file__).resolve().parents[2] / "assets" / "RolesViews.xlsx"


@lru_cache(maxsize=1)
def _rolesviews_bytes() -> bytes:
    return _ASSET.read_bytes()


def get_rolesviews_bytesio() -> io.BytesIO:
    return io.BytesIO(_rolesviews_bytes())
