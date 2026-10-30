"""Offline tests for the reusable watchlist exporter."""

import json

import pandas as pd
import pytest

from tools.export_watchlist import (
    export_files,
    fetch_watchlist,
    load_symbols,
    make_tables,
    select_quote,
    validate_history,
)


def history_frame():
    return pd.DataFrame(
        {
            "date": ["2026-10-02", "2026-10-01"],
            "open": [105.0, 100.0],
            "high": [112.0, 105.0],
            "low": [99.0, 98.0],
            "close": [110.0, 100.0],
            "volume": [2000, 1000],
            "is_anomaly": [False, False],
        }
    )


def test_watchlist_is_normalized_and_unique(tmp_path):
    path = tmp_path / "watchlist.txt"
    path.write_text(" ffc \nHUBC\n\n")
    assert load_symbols(path) == ["FFC", "HUBC"]
    path.write_text("FFC\nFFC\n")
    with pytest.raises(ValueError, match="duplicate"):
        load_symbols(path)
    path.write_text("../invalid\n")
    with pytest.raises(ValueError, match="alphanumeric"):
        load_symbols(path)


def test_exact_quote_wins_over_temporary_suffix():
    screener = pd.DataFrame({"symbol": ["HUBCXD", "HUBC"], "price": [190.0, 194.0]})
    assert select_quote(screener, "HUBC")["price"] == 194.0
    assert select_quote(screener.iloc[:1], "HUBC")["symbol"] == "HUBCXD"
    assert select_quote(screener, "SYS") == {}
    assert select_quote(pd.DataFrame(), "SYS") == {}


def test_history_validation():
    validate_history(history_frame(), "TEST", "2026-10-01", "2026-10-02")
    with pytest.raises(ValueError, match="duplicate"):
        validate_history(pd.concat([history_frame()] * 2), "TEST", "2026-10-01", "2026-10-02")
    with pytest.raises(ValueError, match="outside"):
        validate_history(history_frame(), "TEST", "2026-10-02", "2026-10-02")
    bad = history_frame()
    bad.loc[0, "high"] = 90
    with pytest.raises(ValueError, match="constraints"):
        validate_history(bad, "TEST", "2026-10-01", "2026-10-02")
    bad = history_frame()
    bad.loc[0, "volume"] = -1
    with pytest.raises(ValueError, match="volume"):
        validate_history(bad, "TEST", "2026-10-01", "2026-10-02")


def test_partial_failure_is_recorded(monkeypatch):
    class FakeClient:
        def symbols(self, **kwargs):
            return pd.DataFrame(
                {"symbol": ["FFC"], "name": ["Test"], "sector_name": ["Test sector"]}
            )

        def screener(self, **kwargs):
            return pd.DataFrame({"symbol": ["FFC"], "price": [110.0]})

        def stocks(self, symbol, **kwargs):
            if symbol == "MISSING":
                raise RuntimeError("upstream unavailable")
            return history_frame()

    monkeypatch.setattr("tools.export_watchlist.PSXClient", FakeClient)
    quotes, histories, coverage, metadata = fetch_watchlist(
        ["FFC", "MISSING"], "2026-10-01", "2026-10-02"
    )
    assert len(quotes) == len(histories) == 1
    assert coverage.set_index("symbol").loc["MISSING", "status"] == "error"
    assert metadata["FFC"]["name"] == "Test"


def test_export_formats_and_provenance(tmp_path):
    quotes = [
        {
            "requested_symbol": "FFC",
            "data": {
                "symbol": "FFC",
                "price": 110.0,
                "market_cap": None,
                "pe_ratio": 10.0,
                "dividend_yield": 5.0,
                "change_pct": 10.0,
                "change_1y_pct": 20.0,
                "free_float": None,
                "volume_avg_30d": 1500,
            },
            "meta": {"timestamp": "2026-10-02T16:00:00+05:00", "cached": False},
        }
    ]
    histories = [
        {
            "requested_symbol": "FFC",
            "data": history_frame().to_dict("records"),
            "meta": {"timestamp": "2026-10-02T16:00:00+05:00", "cached": True},
        }
    ]
    coverage = pd.DataFrame(
        {"symbol": ["FFC"], "status": ["ok"], "rows": [2], "quote_found": [True], "error": [""]}
    )
    snapshot, history = make_tables(quotes, histories, {})
    assert snapshot.iloc[0]["period_return_pct"] == pytest.approx(10.0)
    assert snapshot.iloc[0]["open_pkr"] == 105.0
    assert pd.isna(snapshot.iloc[0]["market_cap_pkr"])
    assert history.iloc[0]["date"] == "2026-10-01"
    assert history.iloc[0]["source_cached"]
    export_files(tmp_path, quotes, histories, coverage, {}, {"source": "offline test"})
    assert len(list((tmp_path / "history").glob("*.csv"))) == 1
    assert json.loads((tmp_path / "manifest.json").read_text())["history_rows"] == 2
    assert pd.read_excel(tmp_path / "psx_watchlist.xlsx", sheet_name=None).keys() == {
        "Read me",
        "Snapshot",
        "Daily history",
        "Coverage",
    }
