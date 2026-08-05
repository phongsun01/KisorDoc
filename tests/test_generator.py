"""
tests/test_generator.py
────────────────────────
Unit test cho kisorlib/generator.py (Refactor v3.2).
"""

import pytest
from unittest.mock import MagicMock, patch, call
from pathlib import Path
from datetime import datetime

from kisorlib.generator import (
    build_context,
    generate_many,
    generate_one,
    generate_one_repeat,
    write_with_retry,
    FileResult,
    _emit,
    _make_jinja_env,
)


# ──────────────────────────────────────────────
# build_context
# ──────────────────────────────────────────────

class TestBuildContext:
    def test_basic_mapping(self):
        config_rows = [
            {"Key": "TenGoiThau", "Value": "TenCot"},
            {"Key": "NgayKy", "Value": "NgayKyCot"},
        ]
        pkg = {"TenCot": "Gói thầu ABC", "NgayKyCot": "2024-01-01"}
        ctx = build_context(pkg, config_rows)
        assert ctx["TenGoiThau"] == "Gói thầu ABC"
        assert ctx["NgayKy"] == "2024-01-01"

    def test_nan_becomes_empty(self):
        import math
        config_rows = [{"Key": "Field", "Value": "Col"}]
        pkg = {"Col": float("nan")}
        ctx = build_context(pkg, config_rows)
        assert ctx["Field"] == ""

    def test_none_becomes_empty(self):
        config_rows = [{"Key": "Field", "Value": "Col"}]
        pkg = {"Col": None}
        ctx = build_context(pkg, config_rows)
        assert ctx["Field"] == ""

    def test_datetime_formatted(self):
        config_rows = [{"Key": "NgayKy", "Value": "NgayKyCot"}]
        dt = datetime(2024, 3, 15)
        pkg = {"NgayKyCot": dt}
        ctx = build_context(pkg, config_rows)
        assert ctx["NgayKy"] == "15/03/2024"

    def test_extra_merged(self):
        config_rows = [{"Key": "Field", "Value": "Col"}]
        pkg = {"Col": "value"}
        ctx = build_context(pkg, config_rows, extra={"HoTen": "Nguyen Van A"})
        assert ctx["Field"] == "value"
        assert ctx["HoTen"] == "Nguyen Van A"

    def test_skips_empty_key_or_value(self):
        config_rows = [
            {"Key": "", "Value": "Col"},
            {"Key": "Field", "Value": ""},
            {"Key": "OK", "Value": "OKCol"},
        ]
        pkg = {"Col": "v1", "OKCol": "v3"}
        ctx = build_context(pkg, config_rows)
        assert "OK" in ctx
        assert len([k for k in ctx if k not in ("OK",)]) == 0

    def test_clean_config_key_applied(self):
        config_rows = [{"Key": "<<NgayKy.Date>>", "Value": "NgayKyCot"}]
        pkg = {"NgayKyCot": "2024-01-01"}
        ctx = build_context(pkg, config_rows)
        # clean_config_key strips <> and maps .Date → _Date
        assert "NgayKy_Date" in ctx


# ──────────────────────────────────────────────
# write_with_retry
# ──────────────────────────────────────────────

class TestWriteWithRetry:
    def test_success_first_try(self):
        called = []
        def func():
            called.append(1)
            return True, ""
        ok, err = write_with_retry(func, max_retries=3)
        assert ok is True
        assert len(called) == 1

    def test_returns_tuple(self):
        def func():
            return False, "some error"
        ok, err = write_with_retry(func)
        assert ok is False
        assert err == "some error"

    def test_returns_none_means_ok(self):
        def func():
            return None
        ok, err = write_with_retry(func)
        assert ok is True

    def test_retries_on_permission_error(self):
        import errno
        attempts = []
        def func():
            attempts.append(1)
            if len(attempts) < 3:
                e = PermissionError("being used by another process")
                e.errno = errno.EACCES
                raise e
            return True, ""
        ok, err = write_with_retry(func, max_retries=3, delay=0)
        assert ok is True
        assert len(attempts) == 3

    def test_raises_after_max_retries(self):
        import errno
        def func():
            e = PermissionError("being used by another process")
            e.errno = errno.EACCES
            raise e
        with pytest.raises(PermissionError):
            write_with_retry(func, max_retries=2, delay=0)

    def test_non_lock_permission_error_raises_immediately(self):
        def func():
            raise PermissionError("some other permission issue")
        with pytest.raises(PermissionError):
            write_with_retry(func, max_retries=3, delay=0)


# ──────────────────────────────────────────────
# FileResult
# ──────────────────────────────────────────────

class TestFileResult:
    def test_defaults(self):
        fr = FileResult(template_name="test")
        assert fr.success is False
        assert fr.output_path is None
        assert fr.error is None
        assert fr.warnings == []

    def test_success(self):
        fr = FileResult(template_name="test", success=True, output_path="/out/file.docx")
        assert fr.success
        assert fr.output_path == "/out/file.docx"


# ──────────────────────────────────────────────
# _emit
# ──────────────────────────────────────────────

class TestEmit:
    def test_calls_callback(self):
        events = []
        _emit(events.append, "info", "hello", step=1)
        assert len(events) == 1
        assert events[0]["level"] == "info"
        assert events[0]["message"] == "hello"
        assert events[0]["step"] == 1

    def test_none_callback_no_error(self):
        _emit(None, "info", "hello")  # should not raise

    def test_callback_exception_swallowed(self):
        def bad_cb(e):
            raise RuntimeError("oops")
        _emit(bad_cb, "info", "test")  # should not raise


# ──────────────────────────────────────────────
# _make_jinja_env
# ──────────────────────────────────────────────

class TestMakeJinjaEnv:
    def test_has_all_filters(self):
        env = _make_jinja_env()
        for name in ("date", "date_long", "number", "num2text", "day", "month",
                     "year", "add_days", "add_months", "date_diff", "quarter",
                     "weekday", "date_text"):
            assert name in env.filters, f"Missing filter: {name}"


# ──────────────────────────────────────────────
# generate_one (unit test với mock)
# ──────────────────────────────────────────────

class TestGenerateOne:
    def _make_cfg(self, tmp_path):
        cfg = MagicMock()
        cfg.output_path = tmp_path
        return cfg

    def test_dry_run_returns_context(self, tmp_path):
        src  = tmp_path / "Template.docx"
        src.write_bytes(b"fake")
        dst  = tmp_path / "Output.docx"
        ctx  = {"TenGoiThau": "ABC"}
        used = set()
        cfg  = self._make_cfg(tmp_path)

        result = generate_one(
            template_path  = src,
            template_name  = "TestTpl",
            output_path    = dst,
            nested_context = ctx,
            cfg            = cfg,
            goi_thau_id    = "GT001",
            tables_rows    = [],
            danh_muc_file  = None,
            key_id         = "ID",
            used_names     = used,
            dry_run        = True,
        )
        assert result.success is True
        assert result.dry_run_context == ctx

    def test_permission_error_captured(self, tmp_path):
        src = tmp_path / "Template.docx"
        src.write_bytes(b"fake")
        dst = tmp_path / "Output.docx"
        used = set()
        cfg  = self._make_cfg(tmp_path)

        with patch("kisorlib.generator.write_with_retry") as mock_retry:
            # First call (copy) succeeds, second call (merge) raises PermissionError
            mock_retry.side_effect = [
                None,  # copy OK (returns None)
                PermissionError("being used by another process"),
            ]
            result = generate_one(
                template_path  = src,
                template_name  = "TestTpl",
                output_path    = dst,
                nested_context = {},
                cfg            = cfg,
                goi_thau_id    = "GT001",
                tables_rows    = [],
                danh_muc_file  = None,
                key_id         = "ID",
                used_names     = used,
            )
        assert result.success is False
        assert "đang mở" in result.error or "Word" in result.error

    def test_general_exception_captured(self, tmp_path):
        src = tmp_path / "Template.docx"
        src.write_bytes(b"fake")
        dst = tmp_path / "Output.docx"
        used = set()
        cfg  = self._make_cfg(tmp_path)

        with patch("kisorlib.generator.write_with_retry") as mock_retry:
            mock_retry.side_effect = [
                None,
                RuntimeError("jinja error"),
            ]
            result = generate_one(
                template_path  = src,
                template_name  = "TestTpl",
                output_path    = dst,
                nested_context = {},
                cfg            = cfg,
                goi_thau_id    = "GT001",
                tables_rows    = [],
                danh_muc_file  = None,
                key_id         = "ID",
                used_names     = used,
            )
        assert result.success is False
        assert "RuntimeError" in result.error or "jinja" in result.error


# ──────────────────────────────────────────────
# generate_many (unit test với mock)
# ──────────────────────────────────────────────

class TestGenerateMany:
    def test_returns_results_in_order(self, tmp_path):
        files   = []
        results = []
        cfg     = MagicMock()
        cfg.output_path = tmp_path

        for i in range(3):
            p = tmp_path / f"tpl{i}.docx"
            p.write_bytes(b"fake")
            files.append((p, f"Tpl{i}"))

        with patch("kisorlib.generator.generate_one") as mock_one:
            mock_one.side_effect = [
                FileResult(template_name=f"Tpl{i}", success=True, output_path=str(tmp_path / f"out{i}.docx"))
                for i in range(3)
            ]
            results = generate_many(
                template_paths = files,
                nested_context = {},
                cfg            = cfg,
                goi_thau_id    = "GT001",
                tables_rows    = [],
                danh_muc_file  = None,
                key_id         = "ID",
            )

        assert len(results) == 3
        assert mock_one.call_count == 3
        assert results[0].template_name == "Tpl0"
        assert results[2].template_name == "Tpl2"

    def test_on_progress_called_per_file(self, tmp_path):
        events = []
        files  = [(tmp_path / f"t{i}.docx", f"T{i}") for i in range(2)]
        for p, _ in files:
            p.write_bytes(b"x")

        with patch("kisorlib.generator.generate_one") as mock_one:
            mock_one.side_effect = [
                FileResult(template_name="T0", success=True),
                FileResult(template_name="T1", success=True),
            ]
            generate_many(
                template_paths = files,
                nested_context = {},
                cfg            = MagicMock(),
                goi_thau_id    = "GT001",
                tables_rows    = [],
                danh_muc_file  = None,
                key_id         = "ID",
                on_progress    = events.append,
            )
        # Should have at least 2 "info" events (step announcements)
        info_events = [e for e in events if e.get("level") == "info"]
        assert len(info_events) >= 2

    def test_empty_template_list(self, tmp_path):
        results = generate_many(
            template_paths = [],
            nested_context = {},
            cfg            = MagicMock(),
            goi_thau_id    = "GT001",
            tables_rows    = [],
            danh_muc_file  = None,
            key_id         = "ID",
        )
        assert results == []
