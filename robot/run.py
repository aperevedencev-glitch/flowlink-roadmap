"""Запуск модуля на компьютере ассоциации.

    python run.py probe        — проверка токена на одной закупке (1–2 минуты)
    python run.py load 2026    — загрузка за год: сначала 223-ФЗ, затем 44-ФЗ и 615-ПП
    python run.py load 2026 223 — загрузка за год только по 223-ФЗ
    python run.py reconcile    — сверка и Excel для оператора
    python run.py rebuild      — пересобрать договоры из скачанного (отбор по ГрК) и сверить
    python run.py probe615     — какие типы документов ПП-615 отдаёт ЕИС (диагностика)
    python run.py ranges 223 2024-02-22:2024-02-24 2024-10-13:2024-10-13
                               — догрузить отдельные дни (после сбоев ЕИС) и сверить
    python run.py dashboards   — обновить HTML-дашборды руководителя и сотрудника (без токена);
                                 раз в сутки подтягивает сведения о членах из реестра НОСТРОЙ
    python run.py reestr       — сразу обновить сведения из реестра НОСТРОЙ и дашборды
    python run.py history223   — 223-ФЗ по всем регионам с 01.07.2017 по сегодня, по годам
                                 от свежих к старым; уже скачанные дни не повторяются
    python run.py history 615 44 223 rebuild
                               — вся история по очереди с одним вводом токена, в конце пересборка
    python run.py daily        — ежедневный прогон: последние 3 дня по всем законам, сверка, оповещение
    python run.py mark-sent    — отметить последний пакет nostroy_*.csv как загруженный в НОСТРОЙ
    python run.py etalon       — проверить правила на решениях оператора (data\\decisions.xlsx)
    python run.py targeted223 [файл.xlsx]
                               — точечная загрузка 223-ФЗ по списку реестровых номеров (часы, а не сутки)
    python run.py subjects223  — дописать предмет договорам 223-ФЗ из извещений и сверить
    python run.py audit223     — отчёт по каждому договору списка, без запросов в ЕИС
    python run.py targeted44   — точечная загрузка 44-ФЗ по списку в папке data
    python run.py roles44      — роль члена СРО в договорах 44-ФЗ, выпавших при пересборке
    python run.py versions     — договоры, чья новая версия не прошла отбор; перед загрузкой в НОСТРОЙ
    python run.py missing223   — поискать ненайденные договоры с окном 60 дней и сверить
    python run.py purge [--dry-run]
                               — удалить архивы уже разобранных дней (освободить диск), без токена
    python run.py tests        — автотесты робота
    python run.py status       — что сейчас делает робот
    python run.py menu         — пульт: все действия в одном окне

Токен ЕИС спрашивается при запуске и никуда не записывается.
Можно заранее задать переменную окружения EIS_TOKEN.
"""
import getpass
import logging
import os
import sys
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
os.chdir(HERE)

# Сертификаты из хранилища Windows: сервис ЕИС подписан российским корневым
# сертификатом (Минцифры), которого нет во встроенном списке Python.
# Проверка подлинности сервера при этом остаётся включённой.
try:
    import truststore  # noqa: E402
    truststore.inject_into_ssl()
except ImportError:
    print("Не установлен пакет truststore: запустите «1 Проверка токена.bat», он доставит его сам.")

from eis2nostroy import cli  # noqa: E402


class Tee:
    """Всё, что видно в окне, пишется и в файл лога."""

    def __init__(self, stream, fh):
        self.stream, self.fh = stream, fh

    def write(self, data):
        self.stream.write(data)
        self.fh.write(data)
        self.fh.flush()
        return len(data)

    def flush(self):
        self.stream.flush()
        self.fh.flush()


def need_token():
    if not os.environ.get("EIS_TOKEN"):
        t = getpass.getpass("Вставьте токен ЕИС (символы не видны) и нажмите Enter: ").strip()
        if not t:
            sys.exit("Токен не введён.")
        os.environ["EIS_TOKEN"] = t


def notify_done(cmd: str, rc: int, log_file: Path) -> None:
    """Итог долгой операции — в Telegram или на почту, если оповещения настроены в config.toml."""
    try:
        from eis2nostroy.alerts import notify
        from eis2nostroy.config import load_config
        cfg = load_config("config.toml")
        if not (cfg.alerts.get("telegram_bot_token") or cfg.alerts.get("smtp_host")):
            return
        lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()[-400:]
        keep = [l[24:] if l[:2] == "20" else l for l in lines
                if any(k in l for k in ("Новых договоров", "ошибок", "Записей после сверки", "К передаче",
                                         "Решения оператора", "Эталон", "Пересобрано", "Уже идёт"))]
        notify(cfg.alerts, f"Flowlink: {cmd} — " + ("готово" if rc == 0 else "есть ошибки"), "\n".join(keep[-8:]) or "без подробностей")
    except Exception as exc:  # noqa: BLE001 — оповещение не должно ронять робота
        print(f"Оповещение не отправлено: {exc}")


def status() -> int:
    from eis2nostroy.ops import lock_info, version_line
    print(version_line())
    work = HERE / "data" / "work"
    busy = lock_info(work / "robot.lock")
    print("Сейчас идёт:", f"«{busy['what']}» с {busy['since']}" if busy else "ничего (робот свободен)")
    prog = work / "rebuild_progress.json"
    if prog.exists():
        import json
        p = json.loads(prog.read_text(encoding="utf-8"))
        done = sum(f["done"] for f in p["feeds"].values())
        total = sum(len(f["archives"]) for f in p["feeds"].values())
        print(f"Незаконченная пересборка от {p.get('started')}: разобрано {done} из {total}. "
              f"Запустите пересборку ещё раз — она продолжит.")
    logs = sorted((x for x in (HERE / "data" / "logs").glob("*.log") if not x.name.startswith(("status_", "menu_"))),
                  key=lambda x: x.stat().st_mtime)[-3:]
    for lf in logs:
        tail = [l for l in lf.read_text(encoding="utf-8", errors="replace").splitlines() if l.strip()][-1:]
        print(f"{lf.name}: {tail[0][:150] if tail else ''}")
    return 0


MENU = [
    ("1", "Ежедневный прогон: последние 3 дня, сверка, дашборды", ["daily"], True),
    ("2", "Вся история по очереди: ПП-615 → 44-ФЗ → 223-ФЗ → пересборка", ["history", "615", "44", "223", "rebuild"], True),
    ("3", "Только сверка (минута)", ["reconcile"], False),
    ("4", "Пересборка по ГрК (часы; можно прервать — продолжит)", ["rebuild"], False),
    ("5", "Отметить последний пакет как загруженный в НОСТРОЙ", ["mark-sent"], False),
    ("6", "Проверка правил: автотесты и эталон по решениям оператора", ["check"], False),
    ("7", "Обновить реестр НОСТРОЙ и дашборды", ["reestr"], False),
    ("8", "Состояние робота", ["status"], False),
    ("9", "Точечная загрузка 223-ФЗ по списку договоров (часы)", ["targeted223"], True),
]


def menu() -> int:
    import subprocess
    while True:
        print("\n" + "=" * 64 + "\n  FLOWLINK — пульт робота\n" + "=" * 64)
        for key, title, _, _ in MENU:
            print(f"  {key}. {title}")
        print("  Д. Открыть дашборд руководителя    Р. Открыть файл решений\n  0. Выход")
        choice = input("\nВыберите действие: ").strip().upper()
        if choice in ("0", ""):
            return 0
        if choice in ("Д", "L", "Р", "H"):
            f = (HERE / "data" / "out" / "Flowlink — дашборд руководителя.html") if choice in ("Д", "L") \
                else (HERE / "data" / "decisions.xlsx")
            try:
                os.startfile(f)  # noqa: S606
            except OSError:
                print(f"Файла {f.name} пока нет — сначала запустите сверку (3).")
            continue
        item = next((m for m in MENU if m[0] == choice), None)
        if not item:
            continue
        if item[3]:
            need_token()  # токен спрашивается один раз за сеанс пульта и передаётся запускам
        commands = [["tests"], ["etalon"]] if item[2] == ["check"] else [item[2]]
        for c in commands:
            subprocess.call([sys.executable, str(HERE / "run.py"), *c], env=os.environ.copy())


def update_dashboards() -> int:
    """Пересобрать HTML-дашборды из последней сверки. Ошибка здесь не должна ронять загрузку."""
    try:
        import dashboards
        return dashboards.build()
    except Exception:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        print("Дашборды не обновились — пришлите лог из папки data\\logs")
        return 1


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "probe"
    (HERE / "data" / "logs").mkdir(parents=True, exist_ok=True)
    log_file = HERE / "data" / "logs" / f"{cmd}_{date.today():%Y%m%d}.log"
    fh = open(log_file, "a", encoding="utf-8")
    sys.stdout = Tee(sys.__stdout__, fh)
    sys.stderr = Tee(sys.__stderr__, fh)
    print(f"=== {cmd} {date.today():%d.%m.%Y}, Python {sys.version.split()[0]} ===")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(sys.stdout)])
    if cmd == "probe":
        need_token()
        rc = cli.main(["-c", "config.toml", "probe"])
        print("\nПРОВЕРКА ПРОЙДЕНА" if rc == 0 else "\nПРОВЕРКА НЕ ПРОЙДЕНА — пришлите файл лога из папки data\\logs")
        return rc
    if cmd == "load":
        need_token()
        year = int(sys.argv[2]) if len(sys.argv) > 2 else date.today().year
        start = date(year, 1, 1)
        end = min(date(year, 12, 31), date.today() - timedelta(days=1))
        args = ["-c", "config.toml", "fetch", "--from", start.isoformat(), "--to", end.isoformat()]
        for law in sys.argv[3:]:
            args += ["--law", law]
        rc = cli.main(args)
        cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "reconcile":
        return cli.main(["-c", "config.toml", "reconcile"])
    if cmd == "daily":
        need_token()
        return cli.main(["-c", "config.toml", "run", "--days", "3"])
    if cmd == "history":
        # вся история с одним вводом токена: run.py history 615 44 223 rebuild
        need_token()
        rc = 0
        first, last = "2017-07-01", (date.today() - timedelta(days=1)).isoformat()
        for step in sys.argv[2:] or ["615", "44", "223", "rebuild"]:
            if step == "rebuild":
                print("\n=== Пересборка по ГрК ===")
                rc |= cli.main(["-c", "config.toml", "rebuild"])
                continue
            print(f"\n=== {step}, {first} – {last}, все регионы ===")
            r = cli.main(["-c", "config.toml", "fetch", "--from", first, "--to", last, "--law", step])
            if r == 3:  # другая загрузка уже идёт
                return r
            rc |= r
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "mark-sent":
        return cli.main(["-c", "config.toml", "mark-sent", *sys.argv[2:]])
    if cmd == "subjects223":
        need_token()
        rc = cli.main(["-c", "config.toml", "subjects223"])
        if rc != 3:
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "versions":
        return cli.main(["-c", "config.toml", "versions"])
    if cmd == "roles44":
        return cli.main(["-c", "config.toml", "roles44", *sys.argv[2:]])
    if cmd == "audit223":
        found = sorted((HERE / "data").glob("Договоры 223-ФЗ*.xlsx"), key=lambda p: p.stat().st_mtime)
        if not found:
            print("Нет списка: положите в папку data файл «Договоры 223-ФЗ … .xlsx».")
            return 1
        print(f"Список: {found[-1].name}")
        return cli.main(["-c", "config.toml", "audit223", str(found[-1])])
    if cmd == "missing223":
        found = sorted([p for p in (HERE / "data" / "out").glob("targeted223_не_найдены_*.xlsx") if "окно" not in p.name],
                       key=lambda p: p.stat().st_mtime)
        if not found:
            print("Нет файла targeted223_не_найдены_ДАТА.xlsx в папке data\\out.")
            return 1
        print(f"Список: {found[-1].name}, окно 60 дней")
        need_token()
        rc = cli.main(["-c", "config.toml", "targeted223", str(found[-1]), "--window", "60"])
        if rc != 3:
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "targeted44":
        found = sorted((HERE / "data").glob("Договоры 44-ФЗ*.xlsx"), key=lambda p: p.stat().st_mtime)
        if not found:
            print("Нет списка: положите в папку data файл «Договоры 44-ФЗ … .xlsx».")
            return 1
        print(f"Список: {found[-1].name}")
        need_token()
        rc = cli.main(["-c", "config.toml", "targeted44", str(found[-1])])
        if rc != 3:
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "targeted223":
        if len(sys.argv) > 2:
            lst = Path(sys.argv[2])
        else:
            found = sorted([*(HERE / "data").glob("Договоры 223-ФЗ*.xlsx"), *(HERE / "data" / "out").glob("Договоры 223-ФЗ*.xlsx")],
                           key=lambda p: p.stat().st_mtime)
            if not found:
                print("Нет списка: положите в папку data файл «Договоры 223-ФЗ … .xlsx» с колонками "
                      "«Реестровый номер» и «Дата заключения».")
                return 1
            lst = found[-1]
        print(f"Список: {lst.name}")
        need_token()
        rc = cli.main(["-c", "config.toml", "targeted223", str(lst)])
        if rc != 3:
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "etalon":
        return cli.main(["-c", "config.toml", "etalon"])
    if cmd == "tests":
        import unittest
        suite = unittest.defaultTestLoader.discover(str(HERE / "tests"))
        ok = unittest.TextTestRunner(stream=sys.stdout, verbosity=2).run(suite).wasSuccessful()
        print("\nАВТОТЕСТЫ ПРОЙДЕНЫ" if ok else "\nАВТОТЕСТЫ НЕ ПРОЙДЕНЫ — правки не ставить, пришлите лог")
        return 0 if ok else 1
    if cmd == "status":
        return status()
    if cmd == "menu":
        return menu()
    if cmd == "history223":
        need_token()
        first = date(2017, 7, 1)
        last = date.today() - timedelta(days=1)
        rc = 0
        for year in range(last.year, first.year - 1, -1):
            start = max(date(year, 1, 1), first)
            end = min(date(year, 12, 31), last)
            print(f"\n=== 223-ФЗ, {start:%d.%m.%Y} – {end:%d.%m.%Y} ===")
            rc |= cli.main(["-c", "config.toml", "fetch", "--from", start.isoformat(),
                            "--to", end.isoformat(), "--law", "223"])
            # сверка после каждого года: свежие результаты видны, не дожидаясь конца
            cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "ranges":
        # догрузка отдельных периодов: run.py ranges 223 2024-02-22:2024-02-24 ...
        need_token()
        law = sys.argv[2] if len(sys.argv) > 2 else "223"
        rc = 0
        for span in sys.argv[3:]:
            a, _, b = span.partition(":")
            b = b or a
            print(f"\n=== {law}-ФЗ, догрузка {a} – {b} ===")
            rc |= cli.main(["-c", "config.toml", "fetch", "--from", a, "--to", b, "--law", law])
        # сверка по действующим правилам (текущий ремонт, «Требует проверки»); полная
        # пересборка всего кэша (rebuild) здесь не нужна — она идёт много часов
        cli.main(["-c", "config.toml", "reconcile"])
        return rc
    if cmd == "dashboards":
        # раз в сутки — свежие сведения о членах из открытого реестра НОСТРОЙ (взносы в КФ, уровни, страхование)
        try:
            import nostroy_reestr
            nostroy_reestr.refresh(max_age_hours=20)
        except Exception as exc:  # noqa: BLE001 — без реестра дашборды строятся по прежним сведениям
            print(f"Реестр НОСТРОЙ не обновился: {exc}")
        return update_dashboards()
    if cmd == "reestr":
        import nostroy_reestr
        ok = nostroy_reestr.refresh()
        update_dashboards()
        return 0 if ok else 1
    if cmd == "probe615":
        need_token()
        return cli.main(["-c", "config.toml", "probe615"])
    if cmd == "rebuild":
        # без запросов в ЕИС: пересобрать договоры из скачанного с отбором по ГрК
        return cli.main(["-c", "config.toml", "rebuild"])
    if cmd == "purge":
        # без запросов в ЕИС: удалить архивы уже разобранных дней, отметки скачанности остаются
        extra = ["--dry-run"] if "--dry-run" in sys.argv[2:] else []
        return cli.main(["-c", "config.toml", "purge-cache"] + extra)
    print(__doc__)
    return 2


LONG = ("load", "history223", "ranges", "rebuild", "history", "daily", "targeted223", "targeted44")
NOTIFY = ("load", "history223", "ranges", "rebuild", "history", "targeted223", "targeted44")  # daily оповещает сам

if __name__ == "__main__":
    try:
        code = main()
        cmd = sys.argv[1] if len(sys.argv) > 1 else ""
        if cmd in LONG + ("reconcile", "mark-sent"):
            update_dashboards()
        if cmd in NOTIFY:
            notify_done(cmd, code, HERE / "data" / "logs" / f"{cmd}_{date.today():%Y%m%d}.log")
    except SystemExit:
        raise
    except BaseException:  # noqa: BLE001 — любой сбой должен попасть в лог
        import traceback
        traceback.print_exc()
        code = 1
    sys.exit(code)
