"""Клиент сервиса отдачи информации ЕИС (getDocsIP / getDocsLE).

Сервис не ищет по ИНН поставщика. Он отдаёт архивы документов за дату
по региону заказчика и подсистеме (RGK, RD223, RD615). Модуль скачивает
архивы и сам отбирает договоры членов СРО (см. parser.py).

Доступ с 01.01.2025 — только с токеном пользователя машиночитаемых данных,
полученным после регистрации в ЕИС через ЕСИА.
"""
from __future__ import annotations

import logging
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
import uuid
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

import requests

from .config import Config, Feed

log = logging.getLogger(__name__)

# Единый сервис getDocsIP использует пространство имён 44-ФЗ и для 223-ФЗ
NS = {
    "44": "http://zakupki.gov.ru/fz44/get-docs-ip/ws",
    "223": "http://zakupki.gov.ru/fz44/get-docs-ip/ws",
}


class EisError(RuntimeError):
    pass


class EisClient:
    def __init__(self, cfg: Config, session: requests.Session | None = None):
        self.cfg = cfg
        self.http = session or requests.Session()
        # папки, которые пакетная загрузка уже пыталась получить и не смогла:
        # разбор дня не запрашивает их повторно по одной (повтор — в конце загрузки)
        self.failed_folders: set[Path] = set()
        # use_system_proxy = false — ходить в ЕИС напрямую, мимо прокси/VPN из настроек Windows
        if not cfg.eis.get("use_system_proxy", True) and hasattr(self.http, "trust_env"):
            self.http.trust_env = False
        self._last_call = 0.0
        self._last_dl = 0.0
        self._lock = threading.Lock()
        if not cfg.token:
            log.warning("Токен ЕИС не задан: сервис отдачи информации ответит отказом")

    # ------------------------------------------------------------------ SOAP
    def build_region_request(self, feed: Feed, region: str, day: date, doc_type: str) -> str:
        ns = NS[feed.endpoint_law]
        doc_tag = "documentType223" if feed.endpoint_law == "223" else "documentType44"
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="{ns}">
  <soapenv:Header>
    <{token_tag}>{escape(self.cfg.token)}</{token_tag}>
  </soapenv:Header>
  <soapenv:Body>
    <ws:getDocsByOrgRegionRequest>
      <index>
        <id>{uuid.uuid4()}</id>
        <createDateTime>{datetime.now().strftime('%Y-%m-%dT%H:%M:%S')}</createDateTime>
        <mode>{self.cfg.eis.get('mode', 'PROD')}</mode>
      </index>
      <selectionParams>
        <orgRegion>{escape(region)}</orgRegion>
        <subsystemType>{escape(feed.subsystem)}</subsystemType>
        <{doc_tag}>{escape(doc_type)}</{doc_tag}>
        <periodInfo>
          <exactDate>{day.isoformat()}</exactDate>
        </periodInfo>
      </selectionParams>
    </ws:getDocsByOrgRegionRequest>
  </soapenv:Body>
</soapenv:Envelope>"""

    def build_nsi_request(self, code: str, law: str = "223") -> str:
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        tag = "nsiCode223" if law == "223" else "nsiCode44"
        return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="{NS[law]}">
  <soapenv:Header>
    <{token_tag}>{escape(self.cfg.token)}</{token_tag}>
  </soapenv:Header>
  <soapenv:Body>
    <ws:getNsiRequest>
      <index>
        <id>{uuid.uuid4()}</id>
        <createDateTime>{datetime.now().strftime('%Y-%m-%dT%H:%M:%S')}</createDateTime>
        <mode>{self.cfg.eis.get('mode', 'PROD')}</mode>
      </index>
      <selectionParams>
        <{tag}>{escape(code)}</{tag}>
        <nsiKind>all</nsiKind>
      </selectionParams>
    </ws:getNsiRequest>
  </soapenv:Body>
</soapenv:Envelope>"""

    def build_reestr_request(self, subsystem: str, number: str, law: str = "223") -> str:
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="{NS[law]}">
  <soapenv:Header>
    <{token_tag}>{escape(self.cfg.token)}</{token_tag}>
  </soapenv:Header>
  <soapenv:Body>
    <ws:getDocsByReestrNumberRequest>
      <index>
        <id>{uuid.uuid4()}</id>
        <createDateTime>{datetime.now().strftime('%Y-%m-%dT%H:%M:%S')}</createDateTime>
        <mode>{self.cfg.eis.get('mode', 'PROD')}</mode>
      </index>
      <selectionParams>
        <subsystemType>{escape(subsystem)}</subsystemType>
        <reestrNumber>{escape(number)}</reestrNumber>
      </selectionParams>
    </ws:getDocsByReestrNumberRequest>
  </soapenv:Body>
</soapenv:Envelope>"""

    def fetch_by_reestr(self, subsystem: str, number: str, law: str = "223") -> list[Path]:
        """Все документы по реестровому номеру (для RI223 — закупка и её протоколы), с кэшем."""
        cache = self.cfg.path("cache_dir") or Path("data/cache")
        folder = cache / "byreestr" / subsystem / number
        done = folder / ".done"
        if done.exists():
            return sorted(folder.glob("*.zip"))
        urls = self._ask(law, self.build_reestr_request(subsystem, number, law))
        paths = [self.download_archive(u, folder / f"part{i:03d}.zip") for i, u in enumerate(urls)]
        folder.mkdir(parents=True, exist_ok=True)
        done.write_text(datetime.now().isoformat())
        return paths

    def _post(self, endpoint_law: str, body: str) -> bytes:
        self._throttle()
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        # при повторе после обрыва связи нужен новый id запроса: повтор с прежним ЕИС отклоняет
        # («указан неуникальный гуид»)
        resp = self._net(lambda: self.http.post(
            self.cfg.endpoint(endpoint_law),
            data=_fresh_id(body).encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8", token_tag: self.cfg.token},
            timeout=self.cfg.eis.get("timeout_sec", 60),
        ))
        if resp.status_code != 200:
            raise EisError(f"ЕИС вернул HTTP {resp.status_code}: {resp.text[:300]}")
        return resp.content

    # Сбои на стороне ЕИС, которые проходят сами: по журналам — чаще всего вечером
    # (около 19:15–19:45 и 23:00 по Москве). Раньше такой ответ сразу считался ошибкой
    # и день или закупка пропускались; теперь робот ждёт и спрашивает снова.
    TRANSIENT = ("Непредвиденная ошибка", "HTTP 500", "HTTP 502", "HTTP 503", "HTTP 504")

    def _ask(self, endpoint_law: str, body: str) -> list[str]:
        """Запрос ссылок на архивы с ожиданием при временном сбое ЕИС."""
        limit = float(self.cfg.eis.get("outage_wait_max_min", 45)) * 60
        waited, pause = 0.0, 60.0
        while True:
            try:
                return parse_archive_urls(self._post(endpoint_law, body))
            except EisError as exc:
                if not any(t in str(exc) for t in self.TRANSIENT) or waited >= limit:
                    raise
                with self._lock:
                    now = time.time()
                    if now - getattr(self, "_outage_logged", 0) > 300:
                        self._outage_logged = now
                        log.warning("Временный сбой ЕИС (%s). Жду %.0f с и повторяю запросы.",
                                    str(exc)[:80], pause)
                time.sleep(pause)
                waited += pause
                pause = min(pause * 2, 600)

    def fetch_nsi(self, code: str, dest_dir: Path, law: str = "223") -> list[Path]:
        urls = self._ask(law, self.build_nsi_request(code, law))
        return [self.download_archive(u, dest_dir / f"{code}_{i:03d}.zip") for i, u in enumerate(urls)]

    def _net(self, call):
        """Запрос с ожиданием при обрыве связи (выключился VPN, пропал интернет).

        Без этого при обрыве модуль быстро «пролистывал» дни с ошибками.
        Ждём до net_wait_max_min минут (по умолчанию 60), паузы растут до 5 минут.
        """
        limit = float(self.cfg.eis.get("net_wait_max_min", 60)) * 60
        waited, pause = 0.0, 15.0
        while True:
            try:
                return call()
            except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError,
                    requests.exceptions.ContentDecodingError) as exc:   # обрыв посреди скачивания архива
                if waited >= limit:
                    raise
                log.warning("Нет связи с ЕИС (%s). Жду %.0f с и пробую снова.", type(exc).__name__, pause)
                time.sleep(pause)
                waited += pause
                pause = min(pause * 2, 300)

    def _throttle(self):
        """Не чаще min_interval_sec между запросами — и из нескольких потоков тоже."""
        interval = float(self.cfg.eis.get("min_interval_sec", 1.5))
        with self._lock:
            now = time.monotonic()
            slot = max(now, self._last_call + interval)
            self._last_call = slot
        if slot > now:
            time.sleep(slot - now)

    def _parallel(self, func, items):
        """Запросы к ЕИС в несколько потоков (post_workers), в пределах общего лимита частоты."""
        workers = max(1, int(self.cfg.eis.get("post_workers", 3)))
        if workers == 1 or len(items) < 2:
            return [func(x) for x in items]
        with ThreadPoolExecutor(max_workers=workers) as pool:
            return list(pool.map(func, items))

    def request_archives(self, feed: Feed, region: str, day: date, doc_type: str) -> list[str]:
        """Возвращает ссылки на архивы (их может быть несколько)."""
        body = self.build_region_request(feed, region, day, doc_type)
        return self._ask(feed.endpoint_law, body)

    def download_archive(self, url: str, dest: Path) -> Path:
        """Скачивает архив. Архив готовится на стороне ЕИС, поэтому повторяем."""
        if dest.exists() and dest.stat().st_size > 0:
            return dest
        dest.parent.mkdir(parents=True, exist_ok=True)
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        retries = int(self.cfg.eis.get("archive_retries", 6))
        pause = float(self.cfg.eis.get("archive_retry_wait_sec", 20))
        last = None
        for attempt in range(1, retries + 1):
            self._throttle()
            r = self._net(lambda: self.http.get(url, headers={token_tag: self.cfg.token}, timeout=120))
            if r.status_code == 200 and r.content[:2] == b"PK":
                dest.write_bytes(r.content)
                return dest
            last = f"HTTP {r.status_code}, {len(r.content)} байт"
            log.info("Архив ещё не готов (%s), попытка %s из %s", last, attempt, retries)
            time.sleep(pause)
        raise EisError(f"Архив не скачан: {url} ({last})")

    # ------------------------------------------------------ пакетная загрузка
    # ЕИС готовит архив не сразу: по ссылке какое-то время отвечает 404.
    # Поэтому сначала заказываем все архивы за день, а потом забираем готовые
    # по кругу — ожидание идёт параллельно, а не по очереди.

    def _try_download(self, url: str, dest: Path) -> bool:
        # общий для всех потоков интервал между скачиваниями
        interval = float(self.cfg.eis.get("download_interval_sec", 0.3))
        with self._lock:
            now = time.monotonic()
            slot = max(now, self._last_dl + interval)
            self._last_dl = slot
        if slot > now:
            time.sleep(slot - now)
        token_tag = self.cfg.eis.get("token_header", "individualPerson_token")
        r = self._net(lambda: self.http.get(url, headers={token_tag: self.cfg.token}, timeout=120))
        if r.status_code == 200 and r.content[:2] == b"PK":
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(r.content)
            return True
        if r.status_code == 429:
            log.warning("ЕИС просит снизить частоту (HTTP 429), пауза 60 с")
            time.sleep(60)
        return False

    def download_many(self, groups: dict[Path, list[str]]) -> dict[Path, list[Path]]:
        """groups: папка → ссылки. Возвращает скачанные папки (в каждой ставится .done)."""
        jobs = [(folder, url, folder / f"part{i:03d}.zip")
                for folder, urls in groups.items() for i, url in enumerate(urls)]
        pending = [j for j in jobs if not (j[2].exists() and j[2].stat().st_size > 0)]
        rounds = int(self.cfg.eis.get("archive_rounds", 12))
        pause = float(self.cfg.eis.get("archive_retry_wait_sec", 20))
        for rnd in range(1, rounds + 1):
            # скачиваем в несколько потоков: архивы большие, по одному — долго
            ok = self._parallel(lambda j: self._try_download(j[1], j[2]), pending)
            pending = [j for j, got in zip(pending, ok) if not got]
            if not pending:
                break
            if rnd < rounds:
                log.info("Архивов ещё не готово: %s из %s, жду %.0f с", len(pending), len(jobs), pause)
                time.sleep(pause)
        failed = {j[0] for j in pending}
        self.failed_folders |= failed
        done: dict[Path, list[Path]] = {}
        for folder, urls in groups.items():
            if folder in failed:
                continue
            folder.mkdir(parents=True, exist_ok=True)
            (folder / ".done").write_text(datetime.now().isoformat())
            done[folder] = sorted(folder.glob("*.zip"))
        if failed:
            log.warning("Не дождались архивов в %s папках — они будут запрошены заново", len(failed))
        return done

    def prefetch_day(self, feed: Feed, regions: list[str], day: date, errors: list[str] | None = None) -> int:
        """Заказать архивы за день по всем регионам и скачать их пакетом. Возвращает число новых папок."""
        return self.prefetch_days(feed, regions, [day], errors).get(day, 0)

    def prefetch_days(self, feed: Feed, regions: list[str], days: list[date],
                      errors: list[str] | None = None) -> dict[date, int]:
        """То же сразу за несколько дней: ЕИС готовит архивы всех дней одновременно,
        и ожидание готовности не повторяется для каждого дня отдельно.
        Возвращает {день: число новых папок}."""
        cache = self.cfg.path("cache_dir") or Path("data/cache")
        todo = []
        for day in days:
            for region in regions:
                for doc_type in feed.document_types:
                    folder = cache / feed.subsystem / region / day.isoformat() / doc_type
                    if not (folder / ".done").exists():
                        todo.append((day, region, doc_type, folder))
        if not todo:
            return {}
        t0 = time.monotonic()
        if len(days) > 1:
            log.info("%s: заказываю архивы за %s дн. (%s – %s), запросов %s", feed.subsystem, len(days),
                     f"{min(days):%d.%m.%Y}", f"{max(days):%d.%m.%Y}", len(todo))

        def ask(item):
            day, region, doc_type, folder = item
            try:
                return day, folder, self.request_archives(feed, region, day, doc_type)
            except (EisError, OSError, ET.ParseError) as exc:
                msg = f"{day} {feed.subsystem} регион {region}: {exc}"
                log.error(msg)
                if errors is not None:
                    errors.append(msg)
                self.failed_folders.add(folder)
                return day, folder, None
        # 07.11: заказ (644 запроса на неделю) и скачивание (~1,5 тыс. архивов) шли друг за другом.
        # Теперь архивы дня начинают скачиваться, пока заказываются следующие дни.
        from concurrent.futures import ThreadPoolExecutor
        answers = []
        with ThreadPoolExecutor(max_workers=1) as downloader:
            pending = []
            for day in sorted({t[0] for t in todo}, reverse=True):
                part = [x for x in self._parallel(ask, [t for t in todo if t[0] == day]) if x[2] is not None]
                answers += part
                groups = {f: u for _, f, u in part}
                if groups:
                    pending.append(downloader.submit(self.download_many, groups))
            for fut in pending:
                fut.result()
        if len(days) > 1:
            log.info("%s: архивы за %s – %s скачаны за %.0f с", feed.subsystem,
                     f"{min(days):%d.%m.%Y}", f"{max(days):%d.%m.%Y}", time.monotonic() - t0)
        counts: dict[date, int] = {}
        for day, _, _ in answers:
            counts[day] = counts.get(day, 0) + 1
        return counts

    def prefetch_reestr(self, subsystem: str, numbers: list[str], law: str = "223") -> None:
        """Заказать документы по многим реестровым номерам и скачать их пакетом."""
        cache = self.cfg.path("cache_dir") or Path("data/cache")
        todo = [n for n in numbers if not (cache / "byreestr" / subsystem / n / ".done").exists()]
        if todo:
            log.info("Запрашиваю документы по %s реестровым номерам", len(todo))

        def ask(number):
            folder = cache / "byreestr" / subsystem / number
            try:
                return folder, self._ask(law, self.build_reestr_request(subsystem, number, law))
            except (EisError, OSError, ET.ParseError) as exc:
                log.error("закупка %s: %s", number, exc)
                return folder, None
        groups = {f: u for f, u in self._parallel(ask, todo) if u is not None}
        if groups:
            self.download_many(groups)

    def unpurge_days(self, feed: Feed, regions: list[str], days: list[date], mark: str | None = None) -> int:
        """Очищенные папки этих дней снова считать нескачанными (см. purge_folder)."""
        cache = self.cfg.path("cache_dir") or Path("data/cache")
        n = 0
        for day in days:
            for region in regions:
                for doc_type in feed.document_types:
                    n += unpurge_folder(cache / feed.subsystem / region / day.isoformat() / doc_type, mark)
        if n:
            log.info("%s: папок с удалёнными архивами, которые нужно скачать заново: %s", feed.subsystem, n)
        return n

    def fetch_day(self, feed: Feed, region: str, day: date) -> list[Path]:
        """Все архивы подсистемы за день по региону, с кэшем на диске."""
        cache = self.cfg.path("cache_dir") or Path("data/cache")
        paths: list[Path] = []
        for doc_type in feed.document_types:
            folder = cache / feed.subsystem / region / day.isoformat() / doc_type
            done = folder / ".done"
            if done.exists():
                paths.extend(sorted(folder.glob("*.zip")))
                continue
            urls = self.request_archives(feed, region, day, doc_type)
            for i, url in enumerate(urls):
                paths.append(self.download_archive(url, folder / f"part{i:03d}.zip"))
            folder.mkdir(parents=True, exist_ok=True)
            done.write_text(datetime.now().isoformat())
        return paths


# 08.11: архивы ЕИС заняли весь диск (930 ГБ). Разобранный день больше не нужен целиком:
# найденные договоры уже в data/work. Архивы такой папки удаляются, отметка .done остаётся
# (день не скачивается снова), а в .purged записывается отметка разбора (_scan_mark).
# Если список членов СРО или правила разбора изменятся, отметка не совпадёт, и день
# скачается заново, а не будет молча пропущен.
PURGED = ".purged"
PURGE_FLAG = "purged.flag"   # в корне кэша: архивы части дней удалены (для пересборки)


def purge_folder(folder: Path, mark: str) -> int:
    """Удалить архивы разобранной папки дня, оставить .done. Возвращает освобождённые байты."""
    zips = list(folder.glob("*.zip"))
    if not zips:
        return 0
    marker = folder / PURGED
    try:
        marker.write_text(mark, encoding="utf-8")
        written = True
    except OSError:          # диск полон: сначала освобождаем место, потом пишем отметку
        written = False
    freed = 0
    for z in zips:
        try:
            size = z.stat().st_size
            z.unlink()
            freed += size
        except OSError:
            pass
    if not written:
        marker.write_text(mark, encoding="utf-8")
    return freed


def unpurge_folder(folder: Path, mark: str | None = None) -> bool:
    """Снять отметку скачанности с очищенной папки, чтобы день скачался заново.

    mark=None — всегда (точечная загрузка); иначе только если день разбирали по другим
    правилам или другому списку членов."""
    marker = folder / PURGED
    if not marker.exists():
        return False
    if mark is not None:
        try:
            if marker.read_text(encoding="utf-8").strip() == mark:
                return False
        except OSError:
            pass
    (folder / ".done").unlink(missing_ok=True)
    marker.unlink(missing_ok=True)
    return True


_ID_RE = re.compile(r"<id>[^<]*</id>")


def _fresh_id(body: str) -> str:
    """Новый глобальный идентификатор запроса (index/id) для каждой отправки."""
    return _ID_RE.sub(f"<id>{uuid.uuid4()}</id>", body, count=1)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_archive_urls(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    fault = [e for e in root.iter() if _local(e.tag) in ("faultstring", "errorInfo", "message")]
    urls = [e.text.strip() for e in root.iter() if _local(e.tag) == "archiveUrl" and e.text]
    if not urls and fault:
        raise EisError("ЕИС вернул ошибку: " + "; ".join((f.text or "").strip() for f in fault))
    return urls
