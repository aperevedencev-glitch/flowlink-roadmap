"""Проверки робота Flowlink. Запуск: python -m unittest discover -s tests (или «run.py tests»).

Каждая проверка — случай, на котором робот уже ошибался или может ошибиться.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from eis2nostroy import decisions as D  # noqa: E402
from eis2nostroy.models import CONFIRMED, EXCLUDED, IN_NOSTROY, NEEDS_CHECK, Contract, Member, ReconciledRecord  # noqa: E402
from eis2nostroy.ops import RobotBusy, run_lock  # noqa: E402
from eis2nostroy.parser import parse_contract_xml  # noqa: E402
from eis2nostroy.reconcile import reconcile  # noqa: E402

MEMBER, SUB, CUST = "5001058387", "5018199756", "7701000001"


def contract_xml(supplier: str, sub: str | None = None, subject: str = "Капитальный ремонт здания школы",
                 reg: str = "3500000000026000001") -> bytes:
    sub_block = f"<subContractorsInfo><subContractor><inn>{sub}</inn></subContractor></subContractorsInfo>" if sub else ""
    return f"""<export><contract><regNum>{reg}</regNum><signDate>2026-03-01</signDate>
<customer><inn>{CUST}</inn><fullName>ГБУ Тест</fullName></customer>
<priceInfo><price>15000000.00</price></priceInfo>
<contractSubject>{subject}</contractSubject>
<suppliersInfo><supplier><inn>{supplier}</inn>{sub_block}</supplier></suppliersInfo>
</contract></export>""".encode("utf-8")


def rec(status=CONFIRMED, reg="3500000000026000001", price=15e6) -> ReconciledRecord:
    c = Contract(source="eis_44", law="44", supplier_inn=MEMBER, reg_number=reg, sign_date=date(2026, 3, 1),
                 price=price, subject="Капитальный ремонт", customer_name="ГБУ Тест")
    return ReconciledRecord(member=Member(MEMBER, "ООО Тест"), status=status, primary=c, sources=["eis_44"])


class ParserTests(unittest.TestCase):
    def test_supplier_is_found(self):
        got = parse_contract_xml(contract_xml(MEMBER), "44", {MEMBER})
        self.assertEqual([c.supplier_inn for c in got], [MEMBER])

    def test_subcontractor_is_not_a_party(self):
        # 02.11: член СРО-субподрядчик считался подрядчиком (88,6 млрд ₽ у одного члена)
        got = parse_contract_xml(contract_xml("7700000000", sub=SUB), "44", {SUB})
        self.assertEqual(got, [])

    def test_supplier_kept_when_it_has_subcontractors(self):
        got = parse_contract_xml(contract_xml(MEMBER, sub=SUB), "44", {MEMBER, SUB})
        self.assertEqual({c.supplier_inn for c in got}, {MEMBER})

    def test_pp615_is_always_grk(self):
        got = parse_contract_xml(contract_xml(MEMBER, reg="350000000002600001"), "615", {MEMBER})
        self.assertEqual(got[0].extra.get("grk"), "ГрК")


class ReconcileTests(unittest.TestCase):
    def test_cheap_contract_is_not_excluded(self):
        # заключение юриста: цена до 10 млн ₽ — не основание исключать
        c = parse_contract_xml(contract_xml(MEMBER), "44", {MEMBER})[0]
        c.price = 2_000_000
        c.extra["grk"] = "ГрК"   # без кода ОКПД2 договор с 10.11 уходит на проверку
        recs = reconcile({MEMBER: Member(MEMBER, "ООО Тест")}, [c], date(2017, 7, 1), date(2026, 3, 1))
        self.assertEqual(recs[0].status, CONFIRMED)


class DecisionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_queue_and_decisions_roundtrip(self):
        path = self.tmp / "decisions.xlsx"
        records = [rec(NEEDS_CHECK, "1"), rec(CONFIRMED, "2")]
        D.apply_decisions(records, {})
        self.assertEqual(D.sync_queue(path, records, {}), 1)        # в очередь попал только спорный
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb["Решения"]
        ws.cell(4, 11).value = "Исключить: текущий ремонт"
        wb.save(path)
        dec = D.load_decisions(path)
        records = [rec(NEEDS_CHECK, "1"), rec(CONFIRMED, "2")]
        stats = D.apply_decisions(records, dec)
        self.assertEqual(stats["exclude"], 1)
        self.assertEqual(records[0].status, EXCLUDED)
        self.assertEqual(D.sync_queue(path, records, dec), 0)       # решённое не дублируется
        self.assertEqual(D.load_decisions(path)["1"]["Решение"], "Исключить: текущий ремонт")

    def test_journal_marks_sent_and_detects_price_change(self):
        pkg = self.tmp / "nostroy_test.csv"
        pkg.write_text("ИНН члена СРО;Реестровый номер договора в ЕИС;Номер извещения;Дата заключения;Цена договора, ₽\n"
                       f"{MEMBER};A;;01.03.2026;15000000\n{MEMBER};B;;01.03.2026;15000000\n", encoding="utf-8-sig")
        journal = self.tmp / "peredano.csv"
        self.assertEqual(D.mark_sent(pkg, journal), 2)
        self.assertEqual(D.mark_sent(pkg, journal), 0)              # повторная отметка не дублирует
        records = [rec(CONFIRMED, "A"), rec(CONFIRMED, "B", price=20e6), rec(CONFIRMED, "C")]
        stats = D.apply_journal(records, D.load_journal(journal))
        self.assertEqual([r.status for r in records], [IN_NOSTROY, CONFIRMED, CONFIRMED])
        self.assertEqual(stats, {"sent": 1, "changed": 1})


class OpsTests(unittest.TestCase):
    def test_second_run_is_refused_and_stale_lock_is_ignored(self):
        tmp = Path(tempfile.mkdtemp())
        lock = tmp / "robot.lock"
        try:
            with run_lock(lock, "fetch"):
                self.assertTrue(lock.exists())
                # замок «чужого» живого процесса (родительский процесс точно жив)
                info = json.loads(lock.read_text(encoding="utf-8"))
                info["pid"] = os.getppid()
                lock.write_text(json.dumps(info), encoding="utf-8")
                with self.assertRaises(RobotBusy):
                    with run_lock(lock, "rebuild"):
                        pass
            lock.write_text(json.dumps({"pid": 999999, "what": "fetch", "since": "-"}), encoding="utf-8")
            with run_lock(lock, "fetch"):                             # замок упавшего процесса снимается
                pass
            self.assertFalse(lock.exists())
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class RetryTests(unittest.TestCase):
    """31.10: 64 ошибки ЕИС за прогон, дни догружали вручную (bat 9)."""

    def test_failed_days_are_fetched_again(self):
        from eis2nostroy import cli
        calls = []

        def fake(cfg, start, end, laws=None, **k):
            calls.append((start, end))
            if len(calls) == 1:
                return 5, ["2026-10-31 RGK регион 50: таймаут", "2026-10-30 RGK регион 77: 503"]
            return 1, ([] if start.day == 31 else ["2026-10-30 RGK регион 77: 503"])
        real, cli.cmd_fetch = cli.cmd_fetch, fake
        try:
            found, errors = cli.cmd_fetch_retry(None, date(2026, 10, 1), date(2026, 10, 31))
        finally:
            cli.cmd_fetch = real
        self.assertEqual(found, 7)
        self.assertEqual(errors, ["2026-10-30 RGK регион 77: 503"])
        self.assertEqual(calls[1:], [(date(2026, 10, 30), date(2026, 10, 30)), (date(2026, 10, 31), date(2026, 10, 31))])


class StoreTests(unittest.TestCase):
    """Хранилище больше не перечитывает файл при каждом добавлении — и при этом не теряет дублей."""

    def test_add_dedups_without_reloading_and_sees_other_process(self):
        from eis2nostroy.sources import ContractStore
        tmp = Path(tempfile.mkdtemp())
        try:
            st = ContractStore(tmp / "s.jsonl")
            a, b, c = rec(reg="A").primary, rec(reg="B").primary, rec(reg="C").primary
            self.assertEqual(st.add([a, a, b]), 2)                 # дубль внутри одной пачки
            self.assertEqual(st.add([a, b]), 0)                    # уже записанные
            loads = {"n": 0}
            real = st.load

            def counting():
                loads["n"] += 1
                return real()
            st.load = counting
            for _ in range(50):
                st.add([a])
            self.assertEqual(loads["n"], 0)                        # файл не перечитывался
            other = ContractStore(tmp / "s.jsonl")                 # «другой процесс» дописал договор
            self.assertEqual(other.add([c]), 1)
            self.assertEqual(st.add([c]), 0)                       # первый это заметил
            self.assertEqual(loads["n"], 1)
            st.load = real
            self.assertEqual(len(st.load()), 3)
            with open(tmp / "s.jsonl", "a", encoding="utf-8") as fh:   # окно закрыли посреди записи
                fh.write('{"source": "eis_44", "supplier_')
            fresh = ContractStore(tmp / "s.jsonl")
            self.assertEqual(len(fresh.load()), 3)
            self.assertEqual(fresh.add([rec(reg="D").primary]), 1)   # новая запись не склеилась с обрывком
            self.assertEqual(len(ContractStore(tmp / "s.jsonl").load()), 4)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class ContractVersions(unittest.TestCase):
    """04.11: пересборка оставляла первую версию договора, в пакет ушли цены до допсоглашений."""
    H = "A" * 32

    def _c(self, version, price, sign=date(2024, 1, 10), subject="Капремонт"):
        return Contract(source="eis_44", law="44", supplier_inn=MEMBER, reg_number="1" * 19, price=price, sign_date=sign,
                        subject=subject, raw_file=f"part1.zip:contract_{'1' * 19}_{version}_{self.H}.xml")

    def test_latest_version_wins_in_any_order(self):
        from eis2nostroy.sources import ContractStore
        with tempfile.TemporaryDirectory() as tmp:
            st = ContractStore(Path(tmp) / "s.jsonl")
            self.assertEqual(st.add([self._c(0, 100.0)]), 1)
            self.assertEqual(st.add([self._c(5, 500.0, sign=date(2024, 1, 11))]), 1)   # новая версия, другая дата
            self.assertEqual(st.add([self._c(3, 300.0)]), 0)                           # старая пришла позже
            self.assertEqual([c.price for c in st.load()], [500.0])
            self.assertEqual([c.price for c in ContractStore(st.path).load()], [500.0])

    def test_compact_leaves_one_row(self):
        from eis2nostroy.sources import ContractStore
        with tempfile.TemporaryDirectory() as tmp:
            st = ContractStore(Path(tmp) / "s.jsonl")
            st.add([self._c(0, 100.0)])
            st.add([self._c(2, 200.0)])
            self.assertEqual(st.compact(), 1)
            lines = st.path.read_text(encoding="utf-8").strip().splitlines()
            self.assertEqual(len(lines), 1)
            self.assertEqual(json.loads(lines[0])["price"], 200.0)
            self.assertEqual(st.add([self._c(2, 200.0)]), 0)

    def test_new_version_keeps_filled_subject(self):
        from eis2nostroy.sources import ContractStore
        with tempfile.TemporaryDirectory() as tmp:
            st = ContractStore(Path(tmp) / "s.jsonl")
            first = self._c(0, 100.0, subject="Из извещения")
            first.extra["subject_source"] = "извещение"
            st.add([first])
            st.add([self._c(1, 150.0, subject="")])
            got = st.load()[0]
            self.assertEqual((got.price, got.subject, got.extra["subject_source"]), (150.0, "Из извещения", "извещение"))


class VersionCheck(unittest.TestCase):
    H = "A" * 32

    @staticmethod
    def _xml(supplier: str, subject: str = "Капитальный ремонт школы", sub: str = "") -> bytes:
        block = f"<subContractorsInfo><subContractor><inn>{sub}</inn></subContractor></subContractorsInfo>" if sub else ""
        return (f"<contract><subject>{subject}</subject><customer><inn>{CUST}</inn></customer><suppliers><supplier>"
                f"<inn>{supplier}</inn><fullName>ООО Подрядчик</fullName></supplier></suppliers>{block}</contract>").encode()

    def test_verdicts(self):
        from eis2nostroy.versions import NOT_PARTY, REPAIR, HIDDEN, verdict
        self.assertEqual(verdict(self._xml(SUB), MEMBER, "44")[0], NOT_PARTY)
        self.assertEqual(verdict(self._xml(SUB, sub=MEMBER), MEMBER, "44")[0], NOT_PARTY)
        self.assertEqual(verdict(self._xml(MEMBER, "Текущий ремонт кровли"), MEMBER, "44")[0], REPAIR)
        self.assertEqual(verdict("<contract><subject>Капитальный ремонт</subject></contract>".encode(), MEMBER, "223")[0], HIDDEN)

    def test_only_newer_versions_reported(self):
        import zipfile
        from eis2nostroy.versions import newer_versions
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            day = t / "RGK" / "50" / "2026-09-02" / "contract"
            day.mkdir(parents=True)
            a, b = "1" * 19, "2" * 19
            with zipfile.ZipFile(day / "p.zip", "w") as z:
                z.writestr(f"contract_{a}_2_{self.H}.xml", self._xml(MEMBER))
                z.writestr(f"contract_{a}_5_{self.H}.xml", self._xml(SUB))
                z.writestr(f"contract_{b}_1_{self.H}.xml", self._xml(MEMBER))
            one = lambda reg, v: Contract(source="eis_44", law="44", supplier_inn=MEMBER, reg_number=reg,
                                          raw_file=f"p.zip:contract_{reg}_{v}_{self.H}.xml")
            got = newer_versions(t, {("RGK", a): one(a, 2), ("RGK", b): one(b, 1)})
            self.assertEqual(list(got), [a])
            self.assertEqual(got[a]["version"], 5)
            self.assertEqual(got[a]["verdict"], "Член больше не поставщик")


class NetworkRetry(unittest.TestCase):
    """05.11: обрыв посреди скачивания архива остановил загрузку 44-ФЗ."""

    def test_broken_download_is_retried(self):
        import requests
        from unittest import mock
        from eis2nostroy.eis_client import EisClient
        client = EisClient.__new__(EisClient)
        client.cfg = type("C", (), {"eis": {"net_wait_max_min": 1}})()
        calls = []

        def flaky():
            calls.append(1)
            if len(calls) == 1:
                raise requests.exceptions.ChunkedEncodingError("Response ended prematurely")
            return "ok"
        with mock.patch("eis2nostroy.eis_client.time.sleep"):
            self.assertEqual(client._net(flaky), "ok")
        self.assertEqual(len(calls), 2)


class Targeted44(unittest.TestCase):
    H = "A" * 32

    def _env(self, t):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.append(["Реестровый номер", "Дата заключения", "ИНН члена СРО", "Член СРО"])
        wb.active.append(["1" * 19, "01.02.2020", MEMBER, "ООО"])
        wb.active.append(["2" * 19, "01.02.2020", MEMBER, "ООО"])
        wb.save(t / "list.xlsx")
        (t / "m.csv").write_text("ИНН;Наименование\n" + MEMBER + ";ООО\n", encoding="utf-8")

        class Feed:
            law, kind, subsystem = "44", "contract", "RGK"

        class Cfg:
            feeds = [Feed()]

            def path(self, key):
                return {"members": t / "m.csv", "output_dir": t / "out"}.get(key)
        return Cfg()

    def _zip(self, path, reg, version, inn=MEMBER):
        import zipfile
        xml = (f"<contract><regNum>{reg}</regNum><subject>Капитальный ремонт школы</subject><suppliers><supplier><inn>{inn}</inn>"
               f"</supplier></suppliers></contract>").encode()
        with zipfile.ZipFile(path, "w") as z:
            z.writestr(f"contract_{reg}_{version}_{self.H}.xml", xml)
        return path

    def test_by_number(self):
        from eis2nostroy.sources import ContractStore
        from eis2nostroy.targeted44 import cmd_targeted44
        from openpyxl import load_workbook
        test = self
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            cfg = self._env(t)

            class Client:

                def unpurge_days(self, *a, **k):

                    return 0

                def prefetch_reestr(self, sub, numbers, law):
                    pass

                def fetch_by_reestr(self, sub, reg, law):
                    return [test._zip(t / f"{reg}.zip", reg, 3)] if reg[0] == "1" else []

                def prefetch_day(self, *a):
                    pass

                def fetch_day(self, feed, region, day):
                    return [test._zip(t / "d.zip", "2" * 19, 0)] if day.isoformat() == "2020-02-03" else []

            store = ContractStore(t / "s.jsonl")
            added, errors = cmd_targeted44(cfg, t / "list.xlsx", 3, store, lambda c, x: x, Client())
            self.assertEqual((added, errors), (2, []))
            self.assertEqual({c.reg_number for c in store.load()}, {"1" * 19, "2" * 19})
            rows = list(load_workbook(next((t / "out").glob("targeted44_*.xlsx"))).active.iter_rows(min_row=2, values_only=True))
            self.assertEqual({r[4] for r in rows}, {"да"})

    def test_falls_back_when_number_lookup_fails(self):
        from eis2nostroy.eis_client import EisError
        from eis2nostroy.sources import ContractStore
        from eis2nostroy.targeted44 import cmd_targeted44
        test = self
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            cfg = self._env(t)
            calls = []

            class Client:

                def unpurge_days(self, *a, **k):

                    return 0

                def prefetch_reestr(self, sub, numbers, law):
                    pass

                def fetch_by_reestr(self, sub, reg, law):
                    raise EisError("не поддерживается")

                def prefetch_day(self, feed, regions, day, errors):
                    calls.append(day)

                def fetch_day(self, feed, region, day):
                    return [test._zip(t / "d.zip", "1" * 19, 0)] if day.isoformat() == "2020-02-02" else []

            store = ContractStore(t / "s.jsonl")
            cmd_targeted44(cfg, t / "list.xlsx", 3, store, lambda c, x: x, Client())
            self.assertTrue(calls)
            self.assertEqual({c.reg_number for c in store.load()}, {"1" * 19})

    def test_version_check_reads_number_archives(self):
        from eis2nostroy.versions import newer_versions
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            reg = "3" * 19
            d = t / "byreestr" / "RGK" / reg
            d.mkdir(parents=True)
            self._zip(d / "part000.zip", reg, 7, SUB)
            c = Contract(source="eis_44", law="44", supplier_inn=MEMBER, reg_number=reg, raw_file=f"p:contract_{reg}_2_{self.H}.xml")
            self.assertEqual(newer_versions(t, {("RGK", reg): c})[reg]["version"], 7)


class TargetedTests(unittest.TestCase):
    """Точечная 223-ФЗ: только регион заказчика и дни после даты заключения, протоколы — только для списка."""

    def test_targeted_asks_only_needed_region_days(self):
        import zipfile
        from openpyxl import Workbook
        from eis2nostroy import cli, targeted
        from eis2nostroy.config import load_config
        tmp = Path(tempfile.mkdtemp())
        try:
            RebuildTests.setUp(self)                                  # конфиг, члены, пустое хранилище
            self.store.write_text("", encoding="utf-8")
            cfg_text = (self.tmp / "config.toml").read_text(encoding="utf-8").replace(
                'law = "44"', 'law = "223"').replace('endpoint_law = "44"', 'endpoint_law = "223"').replace(
                'subsystem = "RGK"', 'subsystem = "RD223"')
            (self.tmp / "config.toml").write_text(cfg_text, encoding="utf-8")
            cfg = load_config(self.tmp / "config.toml")
            reg = "57720518494170000010000"                           # заказчик ИНН 7720518494 → регион 77
            wb = Workbook()
            wb.active.append(["ИНН члена СРО", "Реестровый номер", "Дата заключения"])
            wb.active.append([MEMBER, reg, date(2018, 3, 1)])
            lst = self.tmp / "list.xlsx"
            wb.save(lst)
            asked = []
            arch = self.tmp / "a.zip"
            with zipfile.ZipFile(arch, "w") as zf:
                zf.writestr("c.xml", contract_xml(MEMBER, reg=reg))

            class FakeClient:

                def unpurge_days(self, *a, **k):

                    return 0

                def __init__(self, cfg):
                    pass

                def prefetch_day(self, feed, regions, day, errors=None):
                    asked.append((day, tuple(regions)))

                def fetch_day(self, feed, region, day):
                    return [arch] if day == date(2018, 3, 2) else []

            real = targeted.EisClient
            targeted.EisClient = FakeClient
            try:
                added, errors = targeted.cmd_targeted223(cfg, lst, 3, cli._store(cfg), cli._wins(cfg), cli._keep,
                                                         lambda *a: (lambda n: []))
            finally:
                targeted.EisClient = real
            self.assertEqual(added, 1)
            self.assertEqual(errors, [])
            self.assertEqual(sorted(d for d, _ in asked), [date(2018, 3, d) for d in (1, 2, 3, 4)])
            self.assertTrue(all(r == ("77",) for _, r in asked))
            # повторный запуск: договор уже в хранилище — в ЕИС не ходим
            asked.clear()
            targeted.EisClient = FakeClient
            try:
                targeted.cmd_targeted223(cfg, lst, 3, cli._store(cfg), cli._wins(cfg), cli._keep, lambda *a: (lambda n: []))
            finally:
                targeted.EisClient = real
            self.assertEqual(asked, [])
        finally:
            RebuildTests.tearDown(self)
            shutil.rmtree(tmp, ignore_errors=True)


class RebuildTests(unittest.TestCase):
    """02.11: прерванная пересборка оставила сверку из 29 записей вместо 471."""

    def setUp(self):
        import zipfile
        self.tmp = Path(tempfile.mkdtemp())
        (self.tmp / "data" / "work").mkdir(parents=True)
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.append(["ИНН", "Наименование"])
        wb.active.append([MEMBER, "ООО Тест"])
        wb.save(self.tmp / "data" / "members.xlsx")
        (self.tmp / "config.toml").write_text('''
[paths]
members = "data/members.xlsx"
cache_dir = "data/cache"
work_dir = "data/work"
output_dir = "data/out"
[eis]
regions = ["50"]
[eis.endpoints]
IP_44 = "x"
[[eis.feeds]]
law = "44"
endpoint_law = "44"
subsystem = "RGK"
document_types = ["contract"]
[grk]
only = false
''', encoding="utf-8")
        for i in range(6):
            folder = self.tmp / "data" / "cache" / "RGK" / "50" / f"2026-03-0{i + 1}" / "contract"
            folder.mkdir(parents=True)
            with zipfile.ZipFile(folder / "part000.zip", "w") as zf:
                zf.writestr("c.xml", contract_xml(MEMBER, reg=f"35000000000260000{i:02d}"))
        self.store = self.tmp / "data" / "work" / "eis_contracts.jsonl"
        self.store.write_text("old\n", encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_interrupted_rebuild_keeps_store_and_resumes(self):
        from eis2nostroy import cli
        from eis2nostroy.config import load_config
        cfg = load_config(self.tmp / "config.toml")
        real = cli.scan_archive
        calls = {"n": 0}

        def flaky(*a, **k):
            calls["n"] += 1
            if calls["n"] == 4:
                raise KeyboardInterrupt  # закрыли окно посреди пересборки
            return real(*a, **k)
        cli.scan_archive = flaky
        try:
            with self.assertRaises(KeyboardInterrupt):
                cli.cmd_rebuild(cfg)
        finally:
            cli.scan_archive = real
        self.assertEqual(self.store.read_text(encoding="utf-8"), "old\n")   # рабочий файл не тронут
        self.assertTrue((self.tmp / "data" / "work" / "rebuild_progress.json").exists())
        total = cli.cmd_rebuild(cfg)                                         # продолжение
        self.assertEqual(total, 6)
        self.assertFalse((self.tmp / "data" / "work" / "rebuild_progress.json").exists())
        self.assertEqual(len(self.store.read_text(encoding="utf-8").splitlines()), 6)
        self.assertTrue(list((self.tmp / "data" / "work").glob("*.bak.jsonl")))


class Enrich223(unittest.TestCase):
    def test_reg_number_is_contract_not_card(self):
        """В 223-ФЗ номер договора — contractRegNumber, а не номер карточки registrationNumber."""
        from eis2nostroy.parser import parse_contract_xml
        xml = """<ns:contractCutted xmlns:ns="http://zakupki.gov.ru/223fz/contract/1"><body><item><contractData>
          <registrationNumber>67715719854180002290001</registrationNumber>
          <customer><mainInfo><inn>7715719854</inn></mainInfo></customer>
          <contractRegNumber>67715719854180002290000</contractRegNumber>
          <contractDate>2018-06-22</contractDate>
          <supplierInfo><inn>5040078531</inn></supplierInfo>
          <contractPositions><contractPosition><okpd2><code>43.99.90</code></okpd2></contractPosition></contractPositions>
        </contractData></item></body></ns:contractCutted>""".encode()
        found = parse_contract_xml(xml, "223", {"5040078531"})
        self.assertEqual([c.reg_number for c in found], ["67715719854180002290000"])

    def test_subject_from_notice(self):
        """Один лот — предмет лота; несколько — общее название закупки."""
        from eis2nostroy.enrich import notice_subject
        one = """<n:purchaseNotice xmlns:n="u"><body><item><purchaseNoticeData><name>Закупка</name><lots><lot><lotData>
          <subject>Капитальный ремонт школы</subject></lotData></lot></lots></purchaseNoticeData></item></body></n:purchaseNotice>"""
        self.assertEqual(notice_subject(one.encode()), ("Капитальный ремонт школы", 1))
        two = one.replace("</lot></lots>", "</lot><lot><lotData><subject>Другое</subject></lotData></lot></lots>")
        self.assertEqual(notice_subject(two.encode()), ("Закупка", 2))
        self.assertEqual(notice_subject(b"<purchaseProtocol/>"), ("", 0))

    def test_audit_reasons(self):
        from eis2nostroy.enrich import reason_of
        base = {"suppliers": ["1"], "member": True, "grk": "Вероятно ГрК", "okpd2": ["43.2"], "okpd_name": "", "subject": ""}
        self.assertEqual(reason_of(None, "1")[0], "Нет в скачанных архивах")
        self.assertEqual(reason_of({**base, "suppliers": []}, "1")[0], "В ЕИС поставщик скрыт")
        self.assertEqual(reason_of({**base, "member": False}, "1")[0], "Поставщик не член СРО")
        self.assertEqual(reason_of({**base, "grk": "Не ГрК"}, "1")[0], "Не строительные работы")
        self.assertEqual(reason_of({**base, "subject": "Текущий ремонт кровли"}, "1")[0], "Текущий ремонт или содержание")
        self.assertEqual(reason_of(base, "1")[0], "Сохранён")


class Roles44(unittest.TestCase):
    H = "A" * 32

    @staticmethod
    def _xml(supplier: str, sub: str = "") -> bytes:
        block = f"<subContractorsInfo><subContractor><inn>{sub}</inn></subContractor></subContractorsInfo>" if sub else ""
        return (f"<contract><customer><inn>{CUST}</inn></customer><suppliers><supplier><inn>{supplier}</inn>"
                f"<fullName>ООО Генподрядчик</fullName></supplier></suppliers>{block}</contract>").encode()

    def test_roles(self):
        import xml.etree.ElementTree as ET
        from eis2nostroy.roles import inn_paths, role_of
        root = ET.fromstring(self._xml(SUB, MEMBER))
        self.assertEqual(role_of(inn_paths(root, MEMBER)), "Субподрядчик")
        self.assertEqual(role_of(inn_paths(root, SUB)), "Поставщик")
        self.assertEqual(role_of(inn_paths(root, CUST)), "Заказчик")
        self.assertEqual(role_of(inn_paths(root, "1")), "ИНН в договоре нет")

    def test_supplier_and_subcontractor_is_supplier(self):
        from eis2nostroy.roles import SUBCONTRACTOR, SUPPLIER, role_of
        self.assertEqual(role_of(["subContractorsInfo/subContractor", "suppliers/supplier"]), SUPPLIER)
        self.assertEqual(role_of(["subContractorsInfo/subContractor"]), SUBCONTRACTOR)

    def test_gone_keeps_contract_of_other_member(self):
        from eis2nostroy.roles import gone_contracts
        row = lambda inn, reg, law="44": {"law": law, "supplier_inn": inn, "reg_number": reg, "raw_file": ""}
        old = [row("1", "11"), row("2", "11"), row("3", "33"), row("4", "44", "223")]
        self.assertEqual(sorted(gone_contracts(old, [row("1", "11")])), ["11", "33"])
        self.assertEqual(gone_contracts(old, [row("1", "11"), row("2", "11"), row("3", "33")]), {})

    def test_report_from_cache(self):
        import zipfile
        from openpyxl import load_workbook
        from eis2nostroy.roles import cmd_roles44
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            day = t / "cache" / "RGK" / "50" / "2026-09-02" / "contract"
            day.mkdir(parents=True)
            reg = "3" * 19
            with zipfile.ZipFile(day / "part001.zip", "w") as z:
                z.writestr(f"contract_{reg}_0_{self.H}.xml", self._xml(MEMBER))         # старая версия
                z.writestr(f"contract_{reg}_1_{self.H}.xml", self._xml(SUB, MEMBER))    # член стал субподрядчиком
            work = t / "work"
            work.mkdir()
            line = {"law": "44", "supplier_inn": MEMBER, "reg_number": reg, "raw_file": "", "sign_date": "2026-09-01",
                    "price": 100.0, "customer_name": "Заказчик", "subject": "Капремонт"}
            miss = {**line, "reg_number": "4" * 19}
            (work / "eis_contracts_1.bak.jsonl").write_text(
                json.dumps(line, ensure_ascii=False) + "\n" + json.dumps(miss, ensure_ascii=False) + "\n", encoding="utf-8")
            store = work / "eis_contracts.jsonl"
            store.write_text("", encoding="utf-8")
            (t / "m.csv").write_text("ИНН;Наименование\n" + MEMBER + ";ООО Член\n", encoding="utf-8")

            class Cfg:
                def path(self, key):
                    return {"cache_dir": t / "cache", "output_dir": t / "out", "members": t / "m.csv"}.get(key)

            out = cmd_roles44(Cfg(), store, t / "нет.csv", lambda c: c)
            sheet = load_workbook(out)["Договоры"]
            rows = {r[2]: r for r in sheet.iter_rows(min_row=2, values_only=True)}
            self.assertEqual(rows[reg][0], "Субподрядчик")
            self.assertEqual(rows[reg][10], 1)
            self.assertIn(SUB, rows[reg][7])
            self.assertEqual(rows["4" * 19][0], "Нет в кэше")


class DoubtfulContracts(unittest.TestCase):
    def _run(self, *rows):
        from eis2nostroy.models import Contract, Member
        from eis2nostroy.reconcile import reconcile
        cs = [Contract(source="eis_223", law="223", supplier_inn="5001", reg_number=reg, notice_number="317",
                       customer_inn="77", customer_name="Заказчик", sign_date=date(2018, 1, 10), price=price,
                       subject="Капитальный ремонт школы", competitive=True,
                       extra={"grk": "Вероятно ГрК", "grk_note": ""}) for reg, price in rows]
        recs = reconcile({"5001": Member(inn="5001")}, cs, date(2017, 7, 1), date(2026, 3, 1))
        return {r.primary.reg_number: r.status for r in recs}

    def test_zero_price_goes_to_operator(self):
        self.assertEqual(set(self._run(("1" * 23, 0.0)).values()), {"Требует проверки"})

    def test_twins_go_to_operator(self):
        res = self._run(("1" * 23, 100.0), ("2" * 23, 100.0))
        self.assertEqual(sorted(res.values()), ["Требует проверки", "Требует проверки"])

    def test_plain_contract_passes(self):
        self.assertEqual(set(self._run(("1" * 23, 100.0)).values()), {"Подтверждён"})


class Suggest(unittest.TestCase):
    def test_groups(self):
        from eis2nostroy.suggest import suggest
        self.assertEqual(suggest("Капитальный ремонт школы", "", "Требует проверки")[0], "Передать")
        self.assertEqual(suggest("Устройство слоев износа на дороге М-9", "", "Требует проверки")[0], "Вопрос: дороги")
        self.assertEqual(suggest("Ремонт окон поликлиники", "", "Требует проверки")[0], "Вопрос: ремонт")
        self.assertEqual(suggest("Рамочное", "цена договора 0 ₽", "Требует проверки")[0], "Исключить: дубль или ошибка")
        self.assertEqual(suggest("Выполнение работ на объекте", "", "Требует проверки")[0], "Посмотреть")


if __name__ == "__main__":
    unittest.main()


class DayBatchTests(unittest.TestCase):
    """07.11: полная загрузка 44-ФЗ шла ~3,5 мин на день — дни теперь заказываются пачкой."""

    def _client(self, tmp):
        from eis2nostroy.eis_client import EisClient

        class Cfg:
            eis = {"min_interval_sec": 0, "download_interval_sec": 0, "post_workers": 4}
            token = "t"

            def path(self, key):
                return Path(tmp) / "cache"
        log = []

        class C(EisClient):
            def request_archives(self, feed, region, day, doc_type):
                log.append(("ask", day, region))
                return [f"u/{day}/{region}"]

            def download_many(self, groups):
                log.append(("download", len(groups)))
                for folder in groups:
                    folder.mkdir(parents=True, exist_ok=True)
                    (folder / ".done").write_text("x")
                return {}
        return C(Cfg()), log

    def test_days_ordered_and_downloaded_in_pipeline(self):
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory() as tmp:
            client, log = self._client(tmp)
            feed = SimpleNamespace(subsystem="RGK", document_types=["contract"])
            days = [date(2026, 6, d) for d in (1, 2, 3)]
            counts = client.prefetch_days(feed, ["50", "77"], days)
            self.assertEqual(counts, {d: 2 for d in days})
            # архивы каждого дня скачиваются отдельно, не дожидаясь заказа остальных дней
            self.assertEqual([x for x in log if x[0] == "download"], [("download", 2)] * 3)
            self.assertEqual(sum(1 for x in log if x[0] == "ask"), 6)
            log.clear()
            self.assertEqual(client.prefetch_days(feed, ["50", "77"], days), {})   # уже в кэше
            self.assertEqual(log, [])

    def test_downloads_run_in_parallel_threads(self):
        import threading
        from eis2nostroy.eis_client import EisClient
        with tempfile.TemporaryDirectory() as tmp:
            client, _ = self._client(tmp)
            seen = set()

            def fake(url, dest):
                seen.add(threading.get_ident())
                import time
                time.sleep(0.05)
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(b"PK")
                return True
            client._try_download = fake
            groups = {Path(tmp) / f"f{i}": [f"u{i}"] for i in range(8)}
            done = EisClient.download_many(client, groups)
            self.assertEqual(len(done), 8)
            self.assertGreater(len(seen), 1)


class InnPrefilterTests(unittest.TestCase):
    """07.11: XML без ИНН члена СРО не разбираем целиком — договор члена из него не получится."""

    def test_skips_parse_without_member_inn(self):
        from eis2nostroy import parser
        calls = []
        real = parser.parse_contract_xml
        parser.parse_contract_xml = lambda *a, **k: calls.append(1) or real(*a, **k)
        try:
            self.assertEqual(parser.scan_xml(contract_xml("5009998887"), "44", {MEMBER}, None)[0], [])
            self.assertEqual(calls, [])
            self.assertEqual(len(parser.scan_xml(contract_xml(MEMBER), "44", {MEMBER}, None)[0]), 1)
            self.assertEqual(calls, [1])
        finally:
            parser.parse_contract_xml = real


class ResourceSupplyTests(unittest.TestCase):
    """10.11: 527 договоров теплоснабжения «Т Плюс» без ОКПД2 попали в пакет к передаче."""

    def test_supply_without_okpd_is_not_grk(self):
        from eis2nostroy.grk import NOT_GRK, UNKNOWN, classify
        self.assertEqual(classify(subject="Снабжение тепловой энергией в горячей воде")[0], NOT_GRK)
        self.assertEqual(classify(subject="Услуги по горячему водоснабжению и отоплению")[0], NOT_GRK)
        self.assertEqual(classify(subject="Оказание услуг по освобождению участка")[0], UNKNOWN)
        self.assertEqual(classify(subject="Технологическое присоединение к сетям теплоснабжения")[0], UNKNOWN)

    def test_network_construction_stays_grk(self):
        from eis2nostroy.grk import NOT_GRK, classify
        ikz = "2" * 29 + "4321" + "414"   # 43.21, КВР 414
        self.assertNotEqual(classify(ikz, subject="Монтаж наружных сетей электроснабжения")[0], NOT_GRK)
        self.assertNotEqual(classify(subject="Строительство сетей водоснабжения", okpd_codes=["42.21"])[0], NOT_GRK)

    def test_unknown_goes_to_check(self):
        c = parse_contract_xml(contract_xml(MEMBER), "44", {MEMBER})[0]
        c.extra["grk"] = "Нет данных"
        recs = reconcile({MEMBER: Member(MEMBER, "ООО Тест")}, [c], date(2017, 7, 1), date(2026, 3, 1))
        self.assertEqual(recs[0].status, NEEDS_CHECK)
