"""Collect Opinet daily average price data for two seven-day periods."""
from __future__ import annotations

import csv
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

BASE_URL = "https://www.opinet.co.kr/api"
OUT_DIR = Path(__file__).resolve().parent / "opinet_data"


def load_api_key() -> str:
    env_path = Path(__file__).resolve().parent / ".env"
    for raw in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if line.startswith("opinet_api_key="):
            value = line.split("=", 1)[1].strip()
            if value:
                return value
    raise RuntimeError(".env에 opinet_api_key가 없습니다.")


def request_json(endpoint: str, api_key: str, params: dict[str, str]) -> dict:
    query = {"out": "json", "code": api_key, **params}
    req = Request(f"{BASE_URL}/{endpoint}.do?{urlencode(query)}", headers={"User-Agent": "opinet-collector/1.0"})
    with urlopen(req, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8-sig"))
    return payload


def extract_rows(payload: dict) -> list[dict]:
    candidates = [payload]
    if isinstance(payload, dict):
        for key in ("RESULT", "result", "RESULTS", "results", "data", "DATA"):
            if isinstance(payload.get(key), dict):
                candidates.append(payload[key])
    for obj in candidates:
        if not isinstance(obj, dict):
            continue
        for key, value in obj.items():
            if isinstance(value, list) and all(isinstance(x, dict) for x in value):
                return value
            if isinstance(value, dict) and any(k in value for k in ("DATE", "AREA_CD", "PRODCD")):
                return [value]
    return []


def response_error(payload: dict) -> str | None:
    text = json.dumps(payload, ensure_ascii=False)
    if any(token in text.lower() for token in ("error", "fail", "invalid", "오류", "실패")):
        return text[:500]
    return None


def write_csv(path: Path, rows: list[dict]) -> None:
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=keys or ["_empty"], extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def safe_name(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣_-]+", "_", value)


def main() -> int:
    api_key = load_api_key()
    today = datetime.now(ZoneInfo("Asia/Seoul")).date()
    recent_end = today - timedelta(days=1)
    recent_start = recent_end - timedelta(days=6)
    previous_end = recent_start - timedelta(days=1)
    previous_start = previous_end - timedelta(days=6)
    periods = [("recent_7d", recent_start, recent_end), ("previous_7d", previous_start, previous_end)]
    OUT_DIR.mkdir(exist_ok=True)
    calls = 0
    errors: list[dict] = []
    files: list[dict] = []

    try:
        area_payload = request_json("areaCode", api_key, {})
        calls += 1
    except Exception as exc:
        raise RuntimeError(f"지역코드 API 호출 실패: {exc}") from exc
    area_rows = extract_rows(area_payload)
    sido_rows = [r for r in area_rows if str(r.get("AREA_CD", "")).isdigit() and len(str(r.get("AREA_CD"))) == 2]
    if not sido_rows:
        raise RuntimeError("지역코드 API 응답에서 2자리 시도 코드가 없습니다.")
    sido_codes = [str(r["AREA_CD"]).zfill(2) for r in sido_rows]
    write_csv(OUT_DIR / "sido_codes.csv", sido_rows)

    specs = [
        ("national_product", "dateAvgRecentPrice", {}),
        ("national_brand", "datePollAvgRecentPrice", {}),
    ]
    for period_name, start, end in periods:
        anchor = end.strftime("%Y%m%d")
        for label, endpoint, base_params in specs:
            try:
                payload = request_json(endpoint, api_key, {**base_params, "date": anchor})
                calls += 1
                err = response_error(payload)
                if err:
                    errors.append({"endpoint": endpoint, "period": period_name, "area": "00", "error": err})
                rows = extract_rows(payload)
                path = OUT_DIR / f"{period_name}_{label}_{anchor}.csv"
                write_csv(path, rows)
                files.append({"file": path.name, "endpoint": endpoint, "period": period_name, "area": "00", "rows": len(rows)})
            except Exception as exc:
                calls += 1
                errors.append({"endpoint": endpoint, "period": period_name, "area": "00", "error": str(exc)})
        for sido in sido_codes:
            try:
                payload = request_json("dateAreaAvgRecentPrice", api_key, {"area": sido, "date": anchor})
                calls += 1
                err = response_error(payload)
                if err:
                    errors.append({"endpoint": "dateAreaAvgRecentPrice", "period": period_name, "area": sido, "error": err})
                rows = extract_rows(payload)
                path = OUT_DIR / f"{period_name}_area_product_{sido}_{anchor}.csv"
                write_csv(path, rows)
                files.append({"file": path.name, "endpoint": "dateAreaAvgRecentPrice", "period": period_name, "area": sido, "rows": len(rows)})
            except Exception as exc:
                calls += 1
                errors.append({"endpoint": "dateAreaAvgRecentPrice", "period": period_name, "area": sido, "error": str(exc)})
            time.sleep(0.1)

    metadata = {
        "collected_at_kst": datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds"),
        "today_used_by_runtime": today.isoformat(),
        "today_excluded": True,
        "periods": [{"name": n, "start": s.isoformat(), "end": e.isoformat()} for n, s, e in periods],
        "sido_codes": sido_codes,
        "sido_count": len(sido_codes),
        "sigungu_codes_queried": False,
        "products_and_brands": "API defaults (parameters omitted)",
        "response_files": files,
        "total_api_calls": calls,
        "errors": errors,
        "missing": [],
    }
    (OUT_DIR / "collection_metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output_dir": str(OUT_DIR), "sido_count": len(sido_codes), "files": len(files), "total_api_calls": calls, "errors": len(errors)}, ensure_ascii=False))
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
