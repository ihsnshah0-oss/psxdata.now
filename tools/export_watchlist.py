#!/usr/bin/env python3
"""Export a PSX watchlist to CSV, JSON and a formatted Excel workbook.

Run from the repository root, with direct access to dps.psx.com.pk:
    pip install -e . openpyxl
    python tools/export_watchlist.py --start 2026-10-01 --end 2026-10-29 \
        --output exports/psx-25-2026-10-29

The default watchlist is examples/watchlist.txt. No observations are fabricated:
failed symbols are reported in coverage.csv and cause a nonzero exit status.
Excel export requires openpyxl; CSV/JSON export does not.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from psxdata import PSXClient

ROOT = Path(__file__).resolve().parents[1]
HISTORY_COLUMNS = ["date", "open", "high", "low", "close", "volume", "is_anomaly"]
PKT = ZoneInfo("Asia/Karachi")


def load_symbols(path: Path) -> list[str]:
    symbols = [s.strip().upper() for s in path.read_text().splitlines() if s.strip()]
    if not symbols or any(not re.fullmatch(r"[A-Z0-9]+", s) for s in symbols):
        raise ValueError("Watchlist must contain nonempty alphanumeric PSX symbols")
    if len(symbols) != len(set(symbols)):
        raise ValueError("Watchlist contains duplicate symbols")
    return symbols


def select_quote(screener: pd.DataFrame, symbol: str) -> dict:
    """Prefer an exact match; otherwise allow the observed temporary XD suffix."""
    if screener.empty or "symbol" not in screener:
        return {}
    codes = screener["symbol"].astype(str)
    matches = screener[codes.eq(symbol)]
    if matches.empty:
        matches = screener[codes.eq(symbol + "XD")]
    if len(matches) > 1:
        raise ValueError(f"Ambiguous quote for {symbol}: {matches['symbol'].tolist()}")
    return matches.iloc[0].to_dict() if len(matches) else {}


def validate_history(frame: pd.DataFrame, symbol: str, start: str, end: str) -> None:
    missing = set(HISTORY_COLUMNS) - set(frame.columns)
    if missing:
        raise ValueError(f"{symbol}: missing history columns {sorted(missing)}")
    dates = pd.to_datetime(frame["date"], errors="raise")
    if dates.duplicated().any():
        raise ValueError(f"{symbol}: duplicate historical dates")
    if not dates.between(start, end).all():
        raise ValueError(f"{symbol}: history outside requested range")
    prices = frame[["open", "high", "low", "close"]].apply(pd.to_numeric, errors="raise")
    if not prices.notna().all().all() or (prices <= 0).any().any():
        raise ValueError(f"{symbol}: missing or nonpositive OHLC price")
    if (prices.max(axis=1) != prices["high"]).any() or (prices.min(axis=1) != prices["low"]).any():
        raise ValueError(f"{symbol}: OHLC high/low constraints failed")
    volume = pd.to_numeric(frame["volume"], errors="raise")
    if volume.isna().any() or (volume < 0).any():
        raise ValueError(f"{symbol}: invalid volume")


def fetch_watchlist(symbols: list[str], start: str, end: str, cache: bool = True) -> tuple:
    """Use the repository's token-aware, rate-limited PSX client."""
    client = PSXClient()
    metadata = client.symbols(cache=cache).set_index("symbol").to_dict("index")
    screener = client.screener(cache=cache)
    quotes, histories, coverage = [], [], []
    for symbol in symbols:
        retrieved = datetime.now(PKT).isoformat()
        try:
            quote = select_quote(screener, symbol)
            history = client.stocks(symbol, start=start, end=end, cache=cache)
            if history.empty:
                raise ValueError("No historical rows returned")
            validate_history(history, symbol, start, end)
            quotes.append(
                {
                    "requested_symbol": symbol,
                    "data": quote,
                    "meta": {"timestamp": retrieved, "source": "PSX portal"},
                }
            )
            histories.append(
                {
                    "requested_symbol": symbol,
                    "data": history.to_dict("records"),
                    "meta": {"timestamp": retrieved, "source": "PSX portal"},
                }
            )
            coverage.append(
                {
                    "symbol": symbol,
                    "status": "ok",
                    "rows": len(history),
                    "quote_found": bool(quote),
                    "error": "",
                }
            )
        except Exception as exc:
            # Keep partial results, but never turn a failed request into zero prices.
            logging.exception("Export failed for %s", symbol)
            coverage.append(
                {
                    "symbol": symbol,
                    "status": "error",
                    "rows": 0,
                    "quote_found": False,
                    "error": str(exc),
                }
            )
    return quotes, histories, pd.DataFrame(coverage), metadata


def make_tables(quotes: list[dict], histories: list[dict], metadata: dict) -> tuple:
    quote_rows, history_rows = [], []
    for response in quotes:
        symbol = response["requested_symbol"]
        quote = response["data"]
        info = metadata.get(symbol, {})
        quote_rows.append(
            {
                "symbol": symbol,
                "portal_symbol": quote.get("symbol", symbol),
                "company_name": info.get("name", ""),
                "sector": info.get("sector_name", ""),
                "price_pkr": quote.get("price"),
                "change_pct": quote.get("change_pct"),
                "pe_ratio_ttm": quote.get("pe_ratio"),
                "dividend_yield_pct": quote.get("dividend_yield"),
                "change_1y_pct": quote.get("change_1y_pct"),
                "volume_avg_30d_shares": quote.get("volume_avg_30d"),
                "market_cap_pkr": quote.get("market_cap"),
                "free_float_shares": quote.get("free_float"),
                "quote_retrieved_at": response["meta"]["timestamp"],
                "portal_url": f"https://dps.psx.com.pk/company/{symbol}",
                "source_url": response["meta"].get("url", "https://dps.psx.com.pk/screener"),
            }
        )
    for response in histories:
        symbol = response["requested_symbol"]
        for row in response["data"]:
            history_rows.append(
                {
                    "symbol": symbol,
                    **{c: row[c] for c in HISTORY_COLUMNS},
                    "source_retrieved_at": response["meta"]["timestamp"],
                    "source_cached": response["meta"].get("cached", False),
                }
            )
    snapshot = pd.DataFrame(quote_rows)
    history = pd.DataFrame(history_rows)
    if not history.empty:
        history["date"] = pd.to_datetime(history["date"]).dt.strftime("%Y-%m-%d")
        history = history.sort_values(["symbol", "date"]).reset_index(drop=True)
        stats = []
        for symbol, group in history.groupby("symbol", sort=False):
            group = group.sort_values("date")
            stats.append(
                {
                    "symbol": symbol,
                    "trading_date": group.iloc[-1]["date"],
                    "open_pkr": group.iloc[-1]["open"],
                    "high_pkr": group.iloc[-1]["high"],
                    "low_pkr": group.iloc[-1]["low"],
                    "volume_shares": group.iloc[-1]["volume"],
                    "history_rows": len(group),
                    "period_return_pct": (group.iloc[-1]["close"] / group.iloc[0]["close"] - 1)
                    * 100,
                    "period_high_pkr": group["high"].max(),
                    "period_low_pkr": group["low"].min(),
                    "history_retrieved_at": group.iloc[-1]["source_retrieved_at"],
                    "history_cached": group.iloc[-1]["source_cached"],
                }
            )
        snapshot = snapshot.merge(pd.DataFrame(stats), on="symbol", how="left")
        preferred = [
            "symbol", "portal_symbol", "price_pkr", "change_pct", "volume_shares",
            "company_name", "sector", "period_return_pct", "pe_ratio_ttm",
            "dividend_yield_pct", "change_1y_pct",
        ]
        columns = [c for c in preferred if c in snapshot.columns]
        snapshot = snapshot[columns + [c for c in snapshot.columns if c not in columns]]
    return snapshot, history


def write_workbook(
    path: Path,
    snapshot: pd.DataFrame,
    history: pd.DataFrame,
    coverage: pd.DataFrame,
    notes: list[tuple[str, str]],
) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        snapshot.to_excel(writer, sheet_name="Snapshot", index=False)
        history.to_excel(writer, sheet_name="Daily history", index=False)
        coverage.to_excel(writer, sheet_name="Coverage", index=False)
        pd.DataFrame(notes, columns=["Topic", "Details"]).to_excel(
            writer, sheet_name="Read me", index=False
        )
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "C2" if sheet.title in {"Snapshot", "Daily history"} else "A2"
            sheet.sheet_view.showGridLines = False
            sheet.row_dimensions[1].height = 30
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for column in sheet.columns:
                letter = get_column_letter(column[0].column)
                sheet.column_dimensions[letter].width = min(
                    48, max(14, max(len(str(c.value or "")) for c in column[:100]) + 2)
                )
                label = str(sheet.cell(1, column[0].column).value)
                for cell in column[1:]:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0" if "shares" in label else "#,##0.00"
            if sheet.max_row > 1:
                table = Table(
                    displayName=re.sub(r"[^A-Za-z]", "", sheet.title) + "Table",
                    ref=sheet.dimensions,
                )
                table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                sheet.add_table(table)
            if sheet.title == "Snapshot":
                from openpyxl.formatting.rule import CellIsRule

                for label in ("change_pct", "change_1y_pct", "period_return_pct"):
                    if label in snapshot.columns:
                        letter = get_column_letter(snapshot.columns.get_loc(label) + 1)
                        cells = f"{letter}2:{letter}{sheet.max_row}"
                        sheet.conditional_formatting.add(
                            cells,
                            CellIsRule(
                                operator="lessThan", formula=["0"], font=Font(color="B91C1C")
                            ),
                        )
                        sheet.conditional_formatting.add(
                            cells,
                            CellIsRule(
                                operator="greaterThan", formula=["0"], font=Font(color="166534")
                            ),
                        )
            if sheet.title == "Read me":
                sheet.column_dimensions["A"].width = 28
                sheet.column_dimensions["B"].width = 110
                for row in sheet.iter_rows(min_row=2):
                    row[1].alignment = Alignment(wrap_text=True, vertical="top")
                    sheet.row_dimensions[row[0].row].height = 45


def export_files(
    output: Path,
    quotes: list,
    histories: list,
    coverage: pd.DataFrame,
    metadata: dict,
    manifest: dict,
) -> None:
    output.mkdir(parents=True, exist_ok=True)
    snapshot, history = make_tables(quotes, histories, metadata)
    manifest = {
        **manifest,
        "requested_symbols": len(coverage),
        "successful_symbols": int(coverage["status"].eq("ok").sum()),
        "history_rows": len(history),
    }
    for name, frame in [
        ("psx_snapshot", snapshot),
        ("psx_daily_history", history),
        ("coverage", coverage),
    ]:
        frame.to_csv(output / f"{name}.csv", index=False)
    for name, data in [
        ("psx_snapshot", quotes),
        ("psx_daily_history", histories),
        ("manifest", manifest),
    ]:
        (output / f"{name}.json").write_text(json.dumps(data, indent=2, default=str) + "\n")
    (output / "history").mkdir(exist_ok=True)
    for symbol, group in history.groupby("symbol"):
        group.to_csv(output / "history" / f"{symbol}.csv", index=False)
    notes = [
        (k.replace("_", " ").title(), str(v))
        for k, v in manifest.items()
        if k not in {"symbols", "source_urls"}
    ]
    notes += [
        (
            "Currency / volume",
            "Prices in PKR per share; volume and free float in shares; market cap in PKR.",
        ),
        (
            "Period return",
            "Last close / first close in the exported range minus one. "
            "Not a total return and not necessarily month-to-date.",
        ),
        (
            "Corporate actions",
            "Temporary XD suffixes are preserved in portal_symbol. "
            "Returns are not adjusted for payouts.",
        ),
        (
            "Missing values",
            "Blank means unavailable, not zero. Source-reported zero P/E or yield "
            "is retained, not interpreted as a valuation signal.",
        ),
        (
            "Usage",
            "Information only, not investment advice. Check PSX market-data licensing "
            "before redistribution or commercial use.",
        ),
    ]
    write_workbook(output / "psx_watchlist.xlsx", snapshot, history, coverage, notes)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--watchlist", type=Path, default=ROOT / "examples/watchlist.txt")
    parser.add_argument("--start", default=str(date.today() - timedelta(days=365)))
    parser.add_argument("--end", default=str(date.today()))
    parser.add_argument("--output", type=Path, default=ROOT / "exports/psx-watchlist")
    parser.add_argument("--no-cache", action="store_true")
    args = parser.parse_args()
    start, end = date.fromisoformat(args.start), date.fromisoformat(args.end)
    if start > end:
        parser.error("--start must not be after --end")
    symbols = load_symbols(args.watchlist)
    quotes, histories, coverage, metadata = fetch_watchlist(
        symbols, str(start), str(end), not args.no_cache
    )
    manifest = {
        "source": "https://dps.psx.com.pk",
        "retrieved_at": datetime.now(PKT).isoformat(),
        "start": str(start),
        "end": str(end),
        "symbols": symbols,
        "metadata_source": "Current PSX /symbols endpoint",
    }
    export_files(args.output, quotes, histories, coverage, metadata, manifest)
    print(f"Exported {coverage.status.eq('ok').sum()}/{len(symbols)} symbols to {args.output}")
    return 0 if coverage.status.eq("ok").all() else 1


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    raise SystemExit(main())
