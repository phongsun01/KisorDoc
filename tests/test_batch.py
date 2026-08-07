"""
tests/test_batch.py
────────────────────
Unit test cho kisorlib/batch.py.

Bao phủ:
  - IncrementalRunLogger: counts, record_result, record_ok_with_warning (LOG-01 fix),
    write_header/footer, log_event
  - _is_locked_error
  - _find_danh_muc_file
  - _make_progress_cb
  - run_batch: validation guards (option/package rỗng, template rỗng)
"""

import asyncio
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from kisorlib.batch import (
    IncrementalRunLogger,
    _find_danh_muc_file,
    _is_locked_error,
    _make_progress_cb,
    run_batch,
)
from kisorlib.generator import FileResult


# ── fixture ───────────────────────────────────────────────────────────────────

@pytest.fixture()
def logger(tmp_path):
    cfg = MagicMock()
    cfg.ProjectPath = str(tmp_path)
    return IncrementalRunLogger(cfg, "GT001", "OptA", "run")


# ══════════════════════════════════════════════
# _is_locked_error
# ══════════════════════════════════════════════

class TestIsLockedError:
    @pytest.mark.parametrize("msg,expected", [
        ("File đang mở trong Word",                True),
        ("file bị khóa bởi process",              True),
        ("PermissionError: access denied",         True),
        ("being used by another process",          True),
        ("RuntimeError: jinja parse fail",         False),
        ("FileNotFoundError: no such file",        False),
        ("SameFileError: src and dst are same",    False),
        ("",                                       False),
    ])
    def test_cases(self, msg, expected):
        assert _is_locked_error(msg) == expected


# ══════════════════════════════════════════════
# _find_danh_muc_file
# ══════════════════════════════════════════════

class TestFindDanhMucFile:

    def _cfg(self, tmp_path, danh_muc_file="DataFile"):
        cfg = MagicMock()
        cfg.data_path = tmp_path
        cfg.DanhMucFile = danh_muc_file
        return cfg

    def test_returns_matching_file(self, tmp_path):
        (tmp_path / "DataFile_2024.xlsx").touch()
        (tmp_path / "Other.xlsx").touch()
        result = _find_danh_muc_file(self._cfg(tmp_path))
        assert result is not None
        assert "DataFile" in result.name

    def test_case_insensitive_match(self, tmp_path):
        (tmp_path / "DATAFILE_v2.xlsx").touch()
        result = _find_danh_muc_file(self._cfg(tmp_path, "datafile"))
        assert result is not None

    def test_fallback_to_first_xlsx_when_no_match(self, tmp_path):
        (tmp_path / "AAA.xlsx").touch()
        (tmp_path / "ZZZ.xlsx").touch()
        result = _find_danh_muc_file(self._cfg(tmp_path, "NoMatch"))
        assert result is not None
        assert result.name == "AAA.xlsx"  # sorted → alphabetically first

    def test_returns_none_when_no_xlsx(self, tmp_path):
        result = _find_danh_muc_file(self._cfg(tmp_path))
        assert result is None

    def test_ignores_non_xlsx_files(self, tmp_path):
        (tmp_path / "data.csv").touch()
        (tmp_path / "data.xls").touch()
        result = _find_danh_muc_file(self._cfg(tmp_path))
        assert result is None


# ══════════════════════════════════════════════
# _make_progress_cb
# ══════════════════════════════════════════════

class TestMakeProgressCb:

    def test_calls_progress_with_fraction(self):
        calls = []
        cb = _make_progress_cb(lambda frac, msg: calls.append((frac, msg)), 0, 4, "lbl")
        cb({"level": "info", "message": "hello"})
        assert len(calls) == 1
        assert calls[0][0] == pytest.approx(1 / 4)

    def test_fraction_calculation(self):
        calls = []
        cb = _make_progress_cb(lambda f, m: calls.append(f), 2, 5, "lbl")
        cb({"level": "success", "message": "ok"})
        assert calls[0] == pytest.approx(3 / 5)

    def test_none_progress_cb_no_error(self):
        cb = _make_progress_cb(None, 0, 1, "lbl")
        cb({"level": "info", "message": "test"})  # không raise

    def test_ignores_non_standard_levels(self):
        calls = []
        cb = _make_progress_cb(lambda f, m: calls.append(m), 0, 1, "lbl")
        # level không phải info/warning/error/success thì không gọi progress_cb
        cb({"level": "debug", "message": "debug msg"})
        assert calls == []

    def test_all_visible_levels_trigger(self):
        calls = []
        cb = _make_progress_cb(lambda f, m: calls.append(m), 0, 4, "lbl")
        for level in ("info", "warning", "error", "success"):
            cb({"level": level, "message": f"msg-{level}"})
        assert len(calls) == 4


# ══════════════════════════════════════════════
# IncrementalRunLogger
# ══════════════════════════════════════════════

class TestIncrementalRunLogger:

    # ── khởi tạo ──

    def test_creates_log_file(self, logger, tmp_path):
        assert logger.filepath.exists()

    def test_initial_counts_zero(self, logger):
        assert logger.ok_count == 0
        assert logger.warning_count == 0
        assert logger.error_count == 0

    def test_mode_stored(self, tmp_path):
        cfg = MagicMock(); cfg.ProjectPath = str(tmp_path)
        lg = IncrementalRunLogger(cfg, "GT001", "Opt", "retry")
        assert lg.mode == "retry"

    # ── write_header / write_footer ──

    def test_write_header_no_crash(self, logger):
        logger.write_header("Option A", "Gói thầu 001", 10)
        content = logger.filepath.read_text(encoding="utf-8-sig")
        assert "Option A" in content
        assert "Gói thầu 001" in content

    def test_write_footer_closes_file(self, logger):
        logger.write_header("Opt", "Pkg", 5)
        logger.write_footer()
        assert logger._fh.closed

    def test_write_footer_records_elapsed(self, logger):
        logger.write_header("Opt", "Pkg", 5)
        logger.write_footer()
        content = logger.filepath.read_text(encoding="utf-8-sig")
        assert "Thời gian chạy" in content

    # ── log_event ──

    def test_log_event_writes_to_file(self, logger):
        logger.log_event("✅", "Template A", "extra info")
        content = logger.filepath.read_text(encoding="utf-8-sig")
        assert "Template A" in content
        assert "extra info" in content

    def test_log_event_no_extra(self, logger):
        logger.log_event("❌", "Template B")
        content = logger.filepath.read_text(encoding="utf-8-sig")
        assert "Template B" in content

    # ── record_result: success branches ──

    def test_record_result_success_increments_ok(self, logger):
        fr = FileResult("Tpl", success=True, output_path="/out/Tpl.docx")
        logger.record_result([], fr, "Tpl.docx")
        assert logger.ok_count == 1
        assert logger.warning_count == 0
        assert logger.error_count == 0

    def test_record_result_success_appends_checkmark(self, logger):
        results = []
        fr = FileResult("Tpl", success=True, output_path="/out/Tpl.docx")
        logger.record_result(results, fr, "Tpl.docx")
        assert results[0].startswith("✅")
        assert "Tpl.docx" in results[0]

    # ── LOG-01 fix: warning branch tăng cả ok_count lẫn warning_count ──

    def test_record_result_warning_increments_both_ok_and_warning(self, logger):
        """FIX LOG-01: file có warning vẫn là thành công → ok_count + warning_count cùng tăng."""
        fr = FileResult("Tpl", success=True, output_path="/out/Tpl.docx",
                        warnings=["Placeholder X không có data"])
        logger.record_result([], fr, "Tpl.docx")
        assert logger.ok_count == 1,      "ok_count phải tăng (file vẫn thành công)"
        assert logger.warning_count == 1, "warning_count phải tăng"
        assert logger.error_count == 0

    def test_record_result_warning_appends_warning_icon(self, logger):
        results = []
        fr = FileResult("Tpl", success=True, output_path="/out/Tpl.docx",
                        warnings=["Missing X", "Missing Y"])
        logger.record_result(results, fr, "Tpl.docx")
        assert results[0].startswith("⚠️")
        assert "Missing X" in results[0]

    def test_record_result_multiple_warnings_joined(self, logger):
        results = []
        fr = FileResult("Tpl", success=True, output_path="/out/Tpl.docx",
                        warnings=["W1", "W2"])
        logger.record_result(results, fr, "Tpl.docx")
        assert "W1" in results[0] and "W2" in results[0]

    # ── record_result: error branches ──

    def test_record_result_locked_error(self, logger):
        results = []
        fr = FileResult("Tpl", success=False, error="File đang mở trong Word")
        logger.record_result(results, fr, "Tpl.docx")
        assert logger.error_count == 1
        assert logger.ok_count == 0
        assert results[0].startswith("🔒")

    def test_record_result_other_error(self, logger):
        results = []
        fr = FileResult("Tpl", success=False, error="RuntimeError: jinja crash")
        logger.record_result(results, fr, "Tpl.docx")
        assert logger.error_count == 1
        assert results[0].startswith("❌")

    def test_record_result_none_error_message(self, logger):
        results = []
        fr = FileResult("Tpl", success=False, error=None)
        logger.record_result(results, fr, "Tpl.docx")
        assert logger.error_count == 1

    # ── Nhiều file: tổng counts nhất quán ──

    def test_multiple_files_count_summary(self, logger):
        """ok + warning + ok → ok_count=3, warning_count=1, error_count=0."""
        results = []
        logger.record_result(results,
            FileResult("T1", success=True, output_path="/o/T1.docx"), "T1.docx")
        logger.record_result(results,
            FileResult("T2", success=True, output_path="/o/T2.docx",
                       warnings=["W"]), "T2.docx")
        logger.record_result(results,
            FileResult("T3", success=True, output_path="/o/T3.docx"), "T3.docx")
        assert logger.ok_count == 3
        assert logger.warning_count == 1
        assert logger.error_count == 0

    def test_mixed_results_count(self, logger):
        """ok + warning + locked + other_error."""
        results = []
        logger.record_result(results, FileResult("T1", success=True,
            output_path="/o/T1.docx"), "T1.docx")
        logger.record_result(results, FileResult("T2", success=True,
            output_path="/o/T2.docx", warnings=["W"]), "T2.docx")
        logger.record_result(results, FileResult("T3", success=False,
            error="File đang mở trong Word"), "T3.docx")
        logger.record_result(results, FileResult("T4", success=False,
            error="RuntimeError"), "T4.docx")
        assert logger.ok_count == 2
        assert logger.warning_count == 1
        assert logger.error_count == 2

    # ── record_ok_with_warning (LOG-01 compat fix) ──

    def test_record_ok_with_warning_increments_both(self, logger):
        """Backward-compat: record_ok_with_warning cũng phải tăng ok_count (LOG-01 fix)."""
        results = []
        logger.record_ok_with_warning(results, "Tpl", "Tpl.docx", "warn msg")
        assert logger.ok_count == 1,      "ok_count phải tăng (LOG-01 fix)"
        assert logger.warning_count == 1
        assert logger.error_count == 0

    def test_record_ok_with_warning_appends_warning_icon(self, logger):
        results = []
        logger.record_ok_with_warning(results, "Tpl", "Tpl.docx", "placeholder missing")
        assert results[0].startswith("⚠️")
        assert "placeholder missing" in results[0]

    # ── record_ok (compat) ──

    def test_record_ok_increments_ok(self, logger):
        results = []
        logger.record_ok(results, "Tpl", "Tpl.docx")
        assert logger.ok_count == 1
        assert results[0].startswith("✅")

    # ── record_locked ──

    def test_record_locked_sets_flag(self, logger):
        results = []
        has_locked = [False]
        failed = []
        logger.record_locked(results, "Tpl", has_locked, failed)
        assert has_locked[0] is True
        assert "Tpl" in failed
        assert logger.error_count == 1
        assert results[0].startswith("🔒")

    # ── record_error ──

    def test_record_error_sets_flag(self, logger):
        results = []
        has_other = [False]
        failed = []
        logger.record_error(results, "Tpl", RuntimeError("oops"), has_other, failed)
        assert has_other[0] is True
        assert "Tpl" in failed
        assert logger.error_count == 1
        assert results[0].startswith("❌")


# ══════════════════════════════════════════════
# run_batch: validation guards
# ══════════════════════════════════════════════

def _collect_async(gen):
    """Thu thập tất cả item từ async generator."""
    results = []
    async def _run():
        async for item in gen:
            results.append(item)
    asyncio.run(_run())
    return results


class TestRunBatchValidation:
    """
    Chỉ test validation guards (option/package/template rỗng).
    Không cần KisorService thật hay file Excel.
    """

    def _service(self):
        svc = MagicMock()
        svc.config = MagicMock()
        svc.ds = MagicMock()
        return svc

    def test_empty_option_yields_warning(self):
        items = _collect_async(run_batch(self._service(), "", "Gói thầu A", ["T1"]))
        _, status, _ = items[-1]
        assert "quy trình" in status.lower() or "option" in status.lower()

    def test_blank_option_yields_warning(self):
        items = _collect_async(run_batch(self._service(), "   ", "Gói thầu A", ["T1"]))
        _, status, _ = items[-1]
        assert "⚠️" in status

    def test_empty_package_yields_warning(self):
        items = _collect_async(run_batch(self._service(), "OptA", "", ["T1"]))
        _, status, _ = items[-1]
        assert "gói thầu" in status.lower() or "⚠️" in status

    def test_blank_package_yields_warning(self):
        items = _collect_async(run_batch(self._service(), "OptA", "   ", ["T1"]))
        _, status, _ = items[-1]
        assert "⚠️" in status

    def test_empty_templates_yields_warning(self):
        items = _collect_async(run_batch(self._service(), "OptA", "Gói thầu A", []))
        _, status, _ = items[-1]
        assert "template" in status.lower() or "⚠️" in status

    def test_validation_yields_exactly_one_item(self):
        """Guard nên yield đúng 1 item rồi return."""
        items = _collect_async(run_batch(self._service(), "", "pkg", ["T"]))
        assert len(items) == 1

    def test_validation_third_element_is_none(self):
        """retry_state_data phải là None khi validation fail."""
        items = _collect_async(run_batch(self._service(), "", "pkg", ["T"]))
        assert items[0][2] is None
