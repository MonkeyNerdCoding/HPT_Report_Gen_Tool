from pathlib import Path
import tempfile
import unittest

from extraction.edb360_metadata import build_edb360_text_mapping, extract_edb360_metadata
from extraction.edb360_assessment import build_edb360_assessment_mapping


SAMPLE_EDB360_ROOT = Path(r"D:\HPT\CGV\edb360_082026")


@unittest.skipUnless(SAMPLE_EDB360_ROOT.is_dir(), "local EDB360 sample is not available")
class Edb360MetadataTests(unittest.TestCase):
    def test_extracts_cover_metadata_from_edb360(self):
        metadata = extract_edb360_metadata(SAMPLE_EDB360_ROOT)

        self.assertEqual(metadata["database_name"], "GCGVDB")
        self.assertEqual(metadata["oracle_version"], "12.1.0.2.0")
        self.assertEqual(metadata["collection_date"], "Aug-2026")
        self.assertIn("gcgvdb", metadata["instance_summary"])

    def test_builds_defaults_for_manual_signature_fields(self):
        mapping = build_edb360_text_mapping({"database_name": "GCGVDB"})

        self.assertEqual(mapping["{{database_name}}"], "GCGVDB")
        self.assertIn("{{creator}}", mapping)
        self.assertIn("{{approver}}", mapping)

    def test_builds_dynamic_assessment_text_from_edb360(self):
        mapping = build_edb360_assessment_mapping(SAMPLE_EDB360_ROOT)

        self.assertIn("4 control files", mapping["{{assessment_control_file}}"])
        self.assertIn("redo log group đang được multiplexing", mapping["{{assessment_redo_log}}"])
        self.assertIn("SGA 180.000G/instance", mapping["{{assessment_memory_configuration}}"])
        self.assertIn("1146 bảng không có index", mapping["{{assessment_no_index}}"])
        self.assertIn("1678 bảng không có khóa chính", mapping["{{assessment_no_pk}}"])


class Edb360AssessmentRuleTests(unittest.TestCase):
    def test_reads_dynamic_backup_cpu_log_switch_and_asm_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            _write_table(
                root / "00065_edb360_2d_52_rman_backup.html",
                ["#", "SESSION_KEY", "INPUT_TYPE", "STATUS"],
                [["1", "5275", "ARCHIVELOG", "COMPLETED"], ["2", "5276", "DB FULL", "FAILED"]],
            )
            _write_table(
                root / "00070_edb360_2d_57_log_switch_frequency_for_instance_1.html",
                ["#", "SNAP_ID", "BEGIN_TIME", "END_TIME", "LOG_SWITCHES"],
                [
                    ["1", "1", "2026-08-17 01:00:00", "2026-08-17 02:00:00", "4"],
                    ["2", "2", "2026-08-17 02:00:00", "2026-08-17 03:00:00", "10"],
                    ["3", "3", "2026-08-17 03:00:00", "2026-08-17 04:00:00", "40"],
                ],
            )
            _write_table(
                root / "00081_edb360_3e_67_cpu_busy_and_idle_times_percent_for_instance_1.html",
                ["#", "SNAP_ID", "BUSY_TIME_PERC", "IDLE_TIME_PERC"],
                [["1", "1", "5"], ["2", "2", "7.5"]],
            )
            _write_table(
                root / "00063_edb360_2c_50_asm_disk_group.html",
                ["#", "GROUP_NUMBER", "NAME", "TOTAL_MB", "FREE_MB", "USABLE_FILE_MB"],
                [["1", "1", "DATA", "8192000", "216064", "216064"]],
            )
            _write_table(
                root / "00039_edb360_2a_26_scheduler_jobs.html",
                ["#", "OWNER", "JOB_NAME", "ENABL", "FAILURE_COUNT"],
                [["1", "SYS", "OK_JOB", "TRUE", "0"], ["2", "APP", "FAILED_JOB", "TRUE", "2"]],
            )

            mapping = build_edb360_assessment_mapping(root)

        self.assertIn("2 backup job", mapping["{{assessment_backup}}"])
        self.assertEqual(
            mapping["{{recommendation_backup}}"],
            "Khuyến nghị chuẩn bị môi trường thực hiện kiểm thử restore các bản backup. "
            "Việc không có môi trường khôi phục kiểm thử bản backup sẽ không đảm bảo bản backup có thể khôi phục thành công khi cần thiết.",
        )
        self.assertIn("dao động khoảng 1 - 40 lần/giờ", mapping["{{assessment_log_switch}}"])
        self.assertIn("trung bình 18 lần/giờ", mapping["{{assessment_log_switch}}"])
        self.assertIn("2026-08-17 03:00:00 - 2026-08-17 04:00:00 (40 lần/giờ)", mapping["{{assessment_log_switch}}"])
        self.assertIn("trung bình 6.2%", mapping["{{assessment_oracle_foreground_process}}"])
        self.assertIn("Disk group DATA chỉ còn trống 211GB", mapping["{{assessment_asm_disk_group}}"])
        self.assertEqual(mapping["{{assessment_sche_job}}"], "Một số job đang enable nhưng có ghi nhận lỗi trong quá trình chạy.")
        self.assertEqual(
            mapping["{{recommendation_sche_job}}"],
            "Kiểm tra lại các job đang enable và có FAILURE_COUNT > 0 để tránh ảnh hưởng đến hoạt động của hệ thống/ ứng dụng.",
        )


def _write_table(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    cells = ["<table class=\"sortable\"><tr>"]
    cells.extend(f"<th>{header}</th>" for header in headers)
    cells.append("</tr>")
    for row in rows:
        cells.append("<tr>")
        cells.extend(f"<td>{value}</td>" for value in row)
        cells.append("</tr>")
    cells.append("</table>")
    path.write_text("<html><body>" + "".join(cells) + "</body></html>", encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
