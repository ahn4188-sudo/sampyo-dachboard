from __future__ import annotations

import csv
import json
import os
import subprocess
import sys
import threading
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent.parent
APP = Path(__file__).resolve().parent
CACHE = APP / "cache"
STATIC = APP / "static"
RETRY_SECONDS = 600
LOCK = threading.Lock()
STATUS = {"state": "대기", "message": "", "last_attempt": None, "next_retry": None}

PRODUCT_NAMES = {"B034": "고급휘발유", "B027": "보통휘발유", "D047": "자동차용경유", "C004": "실내등유", "K015": "자동차용부탄"}


def kst_today() -> date:
    return datetime.now(ZoneInfo("Asia/Seoul")).date()


def read_csv(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def legacy_cache() -> dict | None:
    """Convert the previously collected CSV set into the API payload shape."""
    source = ROOT / "opinet_data"
    metadata = source / "collection_metadata.json"
    if not metadata.exists():
        return None
    meta = json.loads(metadata.read_text(encoding="utf-8"))
    if meta.get("today_used_by_runtime") != kst_today().isoformat():
        return None
    national, brands, areas = [], [], []
    for p in source.glob("*.csv"):
        rows = read_csv(p)
        period = "최근 7일" if p.name.startswith("recent_7d") else "직전 7일"
        if "national_product" in p.name:
            national += [{"기간": period, "날짜": r.get("DATE", ""), "제품코드": r.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(r.get("PRODCD", ""), r.get("PRODCD", "")), "평균가격(원/리터)": num(r.get("PRICE"))} for r in rows]
        elif "national_brand" in p.name:
            brands += [{"기간": period, "날짜": r.get("DATE", ""), "제품코드": r.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(r.get("PRODCD", ""), r.get("PRODCD", "")), "상표코드": r.get("POLL_DIV_CD", ""), "상표명": r.get("POLL_DIV_CD", ""), "평균가격(원/리터)": num(r.get("PRICE"))} for r in rows]
        elif "area_product" in p.name:
            areas += [{"기간": period, "날짜": r.get("DATE", ""), "지역코드": str(r.get("AREA_CD", "")).zfill(2), "지역명": r.get("AREA_NM", r.get("AREA_CD", "")), "제품코드": r.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(r.get("PRODCD", ""), r.get("PRODCD", "")), "평균가격(원/리터)": num(r.get("PRICE"))} for r in rows]
    return {"asof": kst_today().isoformat(), "national": national, "brands": brands, "areas": areas, "products": [[k, PRODUCT_NAMES.get(k, k)] for k in sorted({r["제품코드"] for r in national if r.get("제품코드")})], "regions": sorted({(r["지역코드"], r["지역명"]) for r in areas if r.get("지역코드")}), "source": "기존 CSV 캐시"}


def collect_fresh() -> dict:
    """Reuse the existing collector's API helpers without running Excel/email code."""
    sys.path.insert(0, str(ROOT))
    import collect_opinet as collector
    key = collector.load_api_key()
    today = kst_today(); recent_end = today - timedelta(days=1); recent_start = recent_end - timedelta(days=6); previous_end = recent_start - timedelta(days=1); previous_start = previous_end - timedelta(days=6)
    area_rows = collector.extract_rows(collector.request_json("areaCode", key, {}))
    sido = [str(r["AREA_CD"]).zfill(2) for r in area_rows if str(r.get("AREA_CD", "")).isdigit() and len(str(r.get("AREA_CD"))) == 2]
    national, brands, areas = [], [], []
    for period, end in (("최근 7일", recent_end), ("직전 7일", previous_end)):
        anchor = end.strftime("%Y%m%d")
        for row in collector.extract_rows(collector.request_json("dateAvgRecentPrice", key, {"date": anchor})):
            national.append({"기간": period, "날짜": row.get("DATE", ""), "제품코드": row.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(row.get("PRODCD", ""), row.get("PRODCD", "")), "평균가격(원/리터)": num(row.get("PRICE"))})
        for row in collector.extract_rows(collector.request_json("datePollAvgRecentPrice", key, {"date": anchor})):
            brands.append({"기간": period, "날짜": row.get("DATE", ""), "제품코드": row.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(row.get("PRODCD", ""), row.get("PRODCD", "")), "상표코드": row.get("POLL_DIV_CD", ""), "상표명": row.get("POLL_DIV_CD", ""), "평균가격(원/리터)": num(row.get("PRICE"))})
        for code in sido:
            for row in collector.extract_rows(collector.request_json("dateAreaAvgRecentPrice", key, {"area": code, "date": anchor})):
                areas.append({"기간": period, "날짜": row.get("DATE", ""), "지역코드": code, "지역명": row.get("AREA_NM", code), "제품코드": row.get("PRODCD", ""), "제품명": PRODUCT_NAMES.get(row.get("PRODCD", ""), row.get("PRODCD", "")), "평균가격(원/리터)": num(row.get("PRICE"))})
    return {"asof": today.isoformat(), "national": national, "brands": brands, "areas": areas, "products": [[k, PRODUCT_NAMES.get(k, k)] for k in sorted({r["제품코드"] for r in national if r.get("제품코드")})], "regions": sorted({(r["지역코드"], r["지역명"]) for r in areas if r.get("지역코드")}), "source": "오피넷 API"}


def cache_path(day: date) -> Path:
    return CACHE / f"data_{day.isoformat()}.json"


def ensure_cache() -> dict | None:
    today = kst_today(); current = cache_path(today); stale = sorted(CACHE.glob("data_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if current.exists():
        data = json.loads(current.read_text(encoding="utf-8")); STATUS.update(state="정상", message="오늘 자료 사용 중", next_retry=None); return data
    with LOCK:
        if current.exists():
            return json.loads(current.read_text(encoding="utf-8"))
        legacy = legacy_cache()
        if legacy:
            CACHE.mkdir(exist_ok=True); current.write_text(json.dumps(legacy, ensure_ascii=False), encoding="utf-8"); STATUS.update(state="정상", message="기존 자료를 오늘 캐시로 재사용", next_retry=None); return legacy
        last = STATUS.get("last_attempt")
        if last and (datetime.now().timestamp() - last) < RETRY_SECONDS and stale:
            STATUS.update(state="이전 자료", message="갱신 실패 후 재시도 대기 중", next_retry=int(last + RETRY_SECONDS)); return json.loads(stale[0].read_text(encoding="utf-8"))
        STATUS.update(state="갱신 중", message="오피넷 API에서 자료 수집 중", last_attempt=datetime.now().timestamp())
        try:
            data = collect_fresh(); CACHE.mkdir(exist_ok=True); current.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8"); STATUS.update(state="정상", message="최신 자료 갱신 완료", next_retry=None); return data
        except Exception as exc:
            STATUS.update(state="이전 자료", message=f"갱신 실패: {type(exc).__name__}", next_retry=int(datetime.now().timestamp() + RETRY_SECONDS))
            return json.loads(stale[0].read_text(encoding="utf-8")) if stale else None


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/data":
            data = ensure_cache()
            if data is None: self.send_error(503, "자료가 없습니다."); return
            data["status"] = {k: v for k, v in STATUS.items() if v is not None}; self.send_json(data); return
        if path == "/api/status": self.send_json(STATUS); return
        target = STATIC / ("index.html" if path == "/" else path.lstrip("/"))
        if target.exists() and target.is_file():
            body = target.read_bytes(); self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body); return
        self.send_error(404)
    def send_json(self, payload):
        body=json.dumps(payload,ensure_ascii=False).encode("utf-8"); self.send_response(200); self.send_header("Content-Type","application/json; charset=utf-8"); self.send_header("Cache-Control","no-store"); self.send_header("Content-Length",str(len(body))); self.end_headers(); self.wfile.write(body)
    def log_message(self, *_): pass

if __name__ == "__main__":
    CACHE.mkdir(exist_ok=True); port=int(os.environ.get("OPINET_PORT","8000")); print(f"http://127.0.0.1:{port}"); ThreadingHTTPServer(("127.0.0.1",port),Handler).serve_forever()
