"""A measurement that stops before writing a report used to leave only 'no fresh fixed rendering report': the launcher
threw the measurement's stderr away, so a runtime that did not match the pin, a missing browser and a profile error all
looked alike. The launcher now keeps stderr beside the report and shows its last line (absolute paths blanked)."""
import inspect

from web_check import stderr_detail
import web_check


def test_the_last_line_is_the_reason_and_absolute_paths_are_blanked(tmp_path):
    log = tmp_path / 'l2-0.stderr'
    log.write_text('warming up\n\nfixed runtime mismatch at /Users/someone/.cache/own-playwright/1.63.0/node_modules\n\n')
    detail = stderr_detail(log)
    assert detail.startswith('fixed runtime mismatch at <path>'), detail
    assert '/Users' not in detail and 'someone' not in detail


def test_a_long_line_is_cut_and_nothing_printed_is_none(tmp_path):
    log = tmp_path / 'l2-0.stderr'
    log.write_text('x' * 1000)
    assert len(stderr_detail(log)) <= 241
    log.write_text('\n  \n')
    assert stderr_detail(log) is None
    assert stderr_detail(tmp_path / 'absent') is None


def test_the_launcher_no_longer_discards_the_measurement_stderr_and_reports_it_when_there_is_no_report():
    source = inspect.getsource(web_check)
    assert 'stderr=subprocess.DEVNULL' not in source, 'the measurement stderr is still thrown away'
    assert source.count('stderr_detail(') >= 2, 'the last stderr line is not shown when there is no report'
