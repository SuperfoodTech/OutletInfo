import contextlib
import io
import json
import os
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest import mock

import pandas as pd

import export_and_upload_drive as exporter
import cronjobsot.weekly_batch_sync as weekly
from cronjobsot import export_to_sot
from GOFOOD import gofood_scraper
from GRAB import grab_merchant_scraper
from cronjobsot.sheet_syncer import STANDARD_HEADERS, SOT_FILLED_COLUMNS


class WeeklyDbrScrapeTests(unittest.TestCase):
    def test_sot_owner_picker_uses_dbr_not_vercel_or_masters(self):
        dbr = pd.DataFrame([
            {"Nama Pemilik": "DBR Owner", "Aplikator": "GoFood"},
        ])
        with tempfile.TemporaryDirectory() as temp_dir, \
             mock.patch.object(export_to_sot, "load_dbr_master_data", return_value=dbr) as load_dbr, \
             mock.patch.object(exporter, "load_vercel_data", side_effect=AssertionError("Vercel must not be read")), \
             mock.patch.object(exporter, "load_gofood_data", side_effect=AssertionError("local masters must not be read")), \
             mock.patch.object(exporter, "load_grab_data", side_effect=AssertionError("local masters must not be read")), \
             mock.patch.object(exporter, "load_shopee_data", side_effect=AssertionError("local masters must not be read")), \
             mock.patch.object(exporter, "OUTPUT_OWNERS_DIR", Path(temp_dir)):
            owners = exporter.get_owners_with_metadata(source="dbr", force_live=True)
        load_dbr.assert_called_once_with(force_live=True)
        self.assertEqual([owner["owner"] for owner in owners], ["DBR Owner"])

    def test_force_live_dbr_rejects_cache_and_does_not_persist_credentials(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_dir = Path(temp_dir)
            stale_cache = cache_dir / "dbr_sheet_cache.csv"
            stale_cache.write_text("Nama Pemilik,Aplikator\nStale Owner,GoFood\n", encoding="utf-8")
            with mock.patch.object(export_to_sot, "CACHE_DIR", cache_dir), \
                 mock.patch.object(export_to_sot, "GOOGLE_SHEET_DBR_URL", ""), \
                 mock.patch.object(export_to_sot, "SOT_APP_SCRIPT_URL", ""):
                self.assertTrue(export_to_sot.load_dbr_master_data(force_live=True).empty)

            live_csv = b"Nama Pemilik,Aplikator,Nama Portal,Kata Sandi\nLive Owner,GoFood,Portal,private-secret\n"
            with mock.patch.object(export_to_sot, "CACHE_DIR", cache_dir), \
                 mock.patch.object(export_to_sot, "GOOGLE_SHEET_DBR_URL", "https://example.test/dbr.csv"), \
                 mock.patch.object(export_to_sot, "SOT_APP_SCRIPT_URL", ""), \
                 mock.patch.object(export_to_sot.urllib.request, "urlopen", return_value=BytesIO(live_csv)):
                live_df = export_to_sot.load_dbr_master_data(force_live=True)
            self.assertEqual(live_df.iloc[0]["Nama Pemilik"], "Live Owner")
            self.assertIn("Stale Owner", stale_cache.read_text(encoding="utf-8"))

    def test_gofood_and_grab_load_credentials_from_private_dbr_snapshot(self):
        dbr = pd.DataFrame([
            {"Nama Pemilik": "Owner A", "Nama Brand": "Brand Go", "Aplikator": "GoFood", "Nama Portal": "Portal Go", "Email FoodMaster1": "go@example.test", "Kata Sandi": "go-secret"},
            {"Nama Pemilik": "Owner A", "Nama Brand": "Brand Grab", "Aplikator": "GrabFood", "Nama Portal": "Portal Grab", "Nama Pengguna": "grab-user", "Kata Sandi": "grab-secret"},
            {"Nama Pemilik": "Owner A", "Nama Brand": "Brand Shop", "Aplikator": "ShopeeFood", "Nama Portal": "Portal Shopee"},
        ])
        with tempfile.TemporaryDirectory() as temp_dir:
            snapshot = Path(temp_dir) / "dbr.csv"
            dbr.to_csv(snapshot, index=False)
            with mock.patch.dict(os.environ, {"WEEKLY_DBR_SNAPSHOT": str(snapshot)}):
                go_creds = gofood_scraper.get_credentials_from_sheet("dbr")
                grab_creds = grab_merchant_scraper.get_credentials_from_sheet("dbr")

        self.assertEqual([(cred["owner"], cred["portal"], cred["email"]) for cred in go_creds], [("Owner A", "Portal Go", "go@example.test")])
        self.assertEqual(go_creds[0]["password"], "go-secret")
        self.assertEqual([(cred["owner"], cred["name"], cred["username"]) for cred in grab_creds], [("Owner A", "Portal Grab", "grab-user")])
        self.assertEqual(grab_creds[0]["password"], "grab-secret")

    def test_dbr_targets_mask_accounts_and_scrape_only_changed_results(self):
        dbr = pd.DataFrame([
            {
                "Nama Pemilik": "Owner A", "Nama Brand": "Brand Go", "Aplikator": "GoFood",
                "Nama Portal": "Portal Go", "Email FoodMaster1": "go@example.test",
                "Kata Sandi": "go-secret",
            },
            {
                "Nama Pemilik": "Owner A", "Nama Brand": "Brand Grab", "Aplikator": "GrabFood",
                "Nama Portal": "Portal Grab", "Nama Pengguna": "grab-user",
                "Kata Sandi": "grab-secret",
            },
            {
                "Nama Pemilik": "Owner A", "Nama Brand": "Brand Shop", "Aplikator": "ShopeeFood",
                "Nama Portal": "Portal Shopee",
            },
        ])
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            go_dir, grab_dir, shopee_dir = root / "GOFOOD", root / "GRAB", root / "SHOPEE"
            for folder in (go_dir / "cache", grab_dir / "cache", shopee_dir):
                folder.mkdir(parents=True, exist_ok=True)

            def fake_subprocess(cmd, cwd, keywords, on_log, **kwargs):
                on_line = kwargs.get("on_line")
                if "gofood_scraper.py" in str(cmd[2]):
                    cache = go_dir / "cache" / "gofood_portal_test.json"
                    cache.write_text(json.dumps({
                        "owner": "Owner A", "portal": "Portal Go", "email": "go@example.test",
                        "outlets": [{"Nama Pemilik": "Owner A", "Nama Brand": "Brand Go", "Nama Listing": "Outlet Go", "Store ID": "G-1"}],
                    }), encoding="utf-8")
                    if on_line:
                        on_line("portal Portal Go selesai")
                    return 0
                if "grab_merchant_scraper.py" in str(cmd[2]):
                    if on_line:
                        on_line("Portal Portal Grab gagal login")
                    return 1
                output_path = Path(cmd[cmd.index("--output") + 1])
                pd.DataFrame([{
                    "Nama Pemilik": "Owner A", "Nama Brand": "Brand Shop", "Aplikator": "ShopeeFood",
                    "Nama Portal": "Portal Shopee", "Nama Listing": "Outlet Shop", "Store ID": "S-1",
                }]).to_excel(output_path, sheet_name="Listing", index=False)
                if on_line:
                    on_line("Portal Portal Shopee berhasil menarik outlet")
                return 0

            output = io.StringIO()
            with mock.patch.object(exporter, "GOFOOD_DIR", go_dir), \
                 mock.patch.object(exporter, "GRAB_DIR", grab_dir), \
                 mock.patch.object(exporter, "SHOPEE_DIR", shopee_dir), \
                 mock.patch.object(exporter, "run_subprocess_stream", side_effect=fake_subprocess), \
                 contextlib.redirect_stdout(output):
                result = exporter.run_live_scraping_for_owner("Owner A", dbr_df=dbr)

        self.assertEqual({row["Aplikator"] for row in result["live_rows"]}, {"GoFood", "ShopeeFood"})
        self.assertEqual([row["status"] for row in result["portal_results"]], ["success", "failed", "success"])
        self.assertEqual(result["portal_results"][1]["stage"], "autentikasi")
        self.assertIn("Akun=g***@example.test", output.getvalue())
        self.assertIn("Akun utama=allvbadmin", output.getvalue())
        for secret in ("go@example.test", "go-secret", "grab-user", "grab-secret"):
            self.assertNotIn(secret, output.getvalue())

    def test_listing_and_sot_use_the_same_live_rows(self):
        dbr = pd.DataFrame([
            {"Nama Pemilik": "Owner A", "Nama Brand": "Brand A", "Aplikator": "GoFood", "Nama Portal": "Portal OK", "Email FoodMaster1": "owner@example.test"},
            {"Nama Pemilik": "Owner A", "Nama Brand": "Brand A", "Aplikator": "GrabFood", "Nama Portal": "Portal Failed", "Nama Pengguna": "grab-user"},
        ])
        scraped = {
            "live_rows": [{
                "Nama Pemilik": "Owner A", "Nama Brand": "Brand A", "Aplikator": "GoFood",
                "Nama Portal": "Portal OK", "Nama Listing": "Live outlet", "Store ID": "LIVE-1",
                "Status Listing": "LIVE", "Alamat": "Address",
            }],
            "portal_results": [{
                "owner": "Owner A", "aplikator": "GoFood", "portal": "Portal OK",
                "status": "success", "stage": "detail outlet", "outlet_count": 1, "message": "1 outlet",
            }, {
                "owner": "Owner A", "aplikator": "GrabFood", "portal": "Portal Failed",
                "status": "failed", "stage": "autentikasi", "outlet_count": 0, "message": "login gagal",
            }],
        }
        captured = {}

        class DummyLock:
            def __init__(self, **_kwargs):
                pass
            def __enter__(self):
                return self
            def __exit__(self, *_args):
                return False

        def capture_listing(df, file_path, _headers):
            captured["listing"] = df.to_dict(orient="records")
            Path(file_path).write_bytes(b"mock workbook")

        with tempfile.TemporaryDirectory() as temp_dir, \
             mock.patch.object(exporter, "JobLock", DummyLock), \
             mock.patch.object(exporter, "OUTPUT_OWNERS_DIR", Path(temp_dir)), \
             mock.patch.object(exporter, "run_live_scraping_for_owner", return_value=scraped), \
             mock.patch.object(exporter, "save_owner_workbook", side_effect=capture_listing), \
             mock.patch.object(exporter, "load_vercel_data", side_effect=AssertionError("weekly DBR must not load Vercel")), \
             mock.patch.object(exporter, "load_gofood_data", side_effect=AssertionError("weekly DBR must not load stale masters")), \
             mock.patch.object(exporter, "load_grab_data", side_effect=AssertionError("weekly DBR must not load stale masters")), \
             mock.patch.object(exporter, "load_shopee_data", side_effect=AssertionError("weekly DBR must not load stale masters")):
            result = exporter.generate_for_owner_pipeline(
                "Owner A", source="dbr", dbr_df=dbr, upload=False, live_scrape=True,
            )

        self.assertTrue(result["success"])
        self.assertTrue(result["is_partial"])
        self.assertEqual([row["Store ID"] for row in captured["listing"]], ["LIVE-1"])
        self.assertEqual([row["Store ID"] for row in result["sot_rows"]], ["LIVE-1"])
        self.assertEqual(result["portal_results"][1]["portal"], "Portal Failed")

    def test_sot_result_serialization_excludes_credentials(self):
        rows = exporter.weekly_sot_rows(pd.DataFrame([{
            "Nama Pemilik": "Owner A", "Aplikator": "GoFood", "Nama Listing": "Outlet",
            "Store ID": "G-1", "Kata Sandi": "secret", "Nama Pengguna": "account",
        }]))
        self.assertEqual(rows[0]["Nama Listing"], "Outlet")
        self.assertNotIn("Kata Sandi", rows[0])
        self.assertNotIn("Nama Pengguna", rows[0])
        self.assertEqual(len(STANDARD_HEADERS), 44)
        self.assertNotIn("Nama Pemilik", SOT_FILLED_COLUMNS)
        sanitized = exporter._redact_weekly_log(
            "Authorization: Bearer token-value Cookie: session=cookie-value Kata Sandi: password-value"
        )
        for secret in ("token-value", "cookie-value", "password-value"):
            self.assertNotIn(secret, sanitized)


class WeeklyBatchSnapshotTests(unittest.TestCase):
    def test_resume_signature_detects_new_portal_without_fingerprinting_password(self):
        original = pd.DataFrame([{
            "Nama Pemilik": "Owner A", "Aplikator": "GoFood", "Nama Portal": "Portal A",
            "Nama Brand": "Brand A", "Kata Sandi": "old-secret",
        }])
        password_rotated = original.copy()
        password_rotated.loc[0, "Kata Sandi"] = "new-secret"
        new_portal = pd.concat([original, pd.DataFrame([{
            "Nama Pemilik": "Owner A", "Aplikator": "GrabFood", "Nama Portal": "Portal B",
            "Nama Brand": "Brand B", "Kata Sandi": "another-secret",
        }])], ignore_index=True)
        initial_signature = weekly._owner_dbr_signature(original, "Owner A")
        self.assertEqual(initial_signature, weekly._owner_dbr_signature(password_rotated, "Owner A"))
        self.assertNotEqual(initial_signature, weekly._owner_dbr_signature(new_portal, "Owner A"))

    def _run_fake_batch(self, portal_failure=True):
        dbr = pd.DataFrame([
            {"Nama Pemilik": "Owner A", "Aplikator": "GoFood"},
            {"Nama Pemilik": "Owner B", "Aplikator": "GrabFood"},
        ])
        temp = tempfile.TemporaryDirectory()
        root = Path(temp.name)
        cache_dir, reports_dir = root / "cache", root / "reports"
        cache_dir.mkdir()
        reports_dir.mkdir()
        snapshot_checks = []

        class FakeProcess:
            def __init__(self, cmd, env, **_kwargs):
                owner = cmd[cmd.index("--owner") + 1]
                snapshot = Path(env["WEEKLY_DBR_SNAPSHOT"])
                secure_snapshot = (snapshot.stat().st_mode & 0o777) == 0o600
                snapshot_df = pd.read_csv(snapshot)
                complete_snapshot = set(snapshot_df["Nama Pemilik"]) == {"Owner A", "Owner B"}
                snapshot_checks.append((secure_snapshot, complete_snapshot))
                if not portal_failure:
                    payload = {"owner": owner, "success": False, "portal_results": [], "sot_rows": []}
                elif owner == "Owner A" and portal_failure:
                    payload = {
                        "owner": owner, "success": True, "is_partial": True,
                        "portal_results": [
                            {"owner": owner, "aplikator": "GoFood", "portal": "Portal OK", "status": "success", "stage": "detail outlet", "outlet_count": 1, "message": "1 outlet"},
                            {"owner": owner, "aplikator": "GoFood", "portal": "Portal Error", "status": "failed", "stage": "autentikasi", "outlet_count": 0, "message": "login gagal"},
                        ],
                        "sot_rows": [{"Nama Pemilik": owner, "Aplikator": "GoFood", "Nama Listing": "Outlet live A", "Store ID": "A-1"}],
                    }
                elif owner == "Owner B":
                    payload = {
                        "owner": owner, "success": True, "is_partial": False,
                        "portal_results": [{"owner": owner, "aplikator": "GrabFood", "portal": "Portal B", "status": "success", "stage": "detail outlet", "outlet_count": 1, "message": "1 outlet"}],
                        "sot_rows": [{"Nama Pemilik": owner, "Aplikator": "GrabFood", "Nama Listing": "Outlet live B", "Store ID": "B-1"}],
                    }
                else:
                    payload = {"owner": owner, "success": False, "portal_results": [], "sot_rows": []}
                Path(env["WEEKLY_OWNER_RESULT_FILE"]).write_text(json.dumps(payload), encoding="utf-8")
                self.stdout = io.StringIO("[WEEKLY][PORTAL] masked progress\n")

            def wait(self):
                return 0

        captured = {}
        def export_sot(df):
            captured["rows"] = df.to_dict(orient="records")
            return True, {"status": "success", "total_updated": len(df), "total_inserted": 0, "totalRows": len(df)}

        with mock.patch.object(weekly, "CACHE_DIR", cache_dir), \
             mock.patch.object(weekly, "REPORTS_DIR", reports_dir), \
             mock.patch.object(weekly, "load_dbr_master_data", return_value=dbr), \
             mock.patch.object(weekly, "send_discord_start"), \
             mock.patch.object(weekly, "send_discord_summary"), \
             mock.patch.object(weekly, "cleanup_zombie_processes"), \
             mock.patch.object(weekly.time, "sleep"), \
             mock.patch.object(weekly.subprocess, "Popen", side_effect=lambda cmd, **kw: FakeProcess(cmd, **kw)), \
             mock.patch.object(weekly, "export_dataframe_to_sot", side_effect=export_sot):
            result = weekly.run_weekly_sync(resume=False, no_upload=True)

        report = json.loads((reports_dir / f"weekly_report_{__import__('datetime').date.today():%Y%m%d}.json").read_text(encoding="utf-8"))
        self.assertTrue(snapshot_checks)
        self.assertTrue(all(secure and complete for secure, complete in snapshot_checks))
        temp.cleanup()
        return result, captured, report

    def test_partial_portal_failure_syncs_only_successful_live_rows_and_reports_failure(self):
        result, captured, report = self._run_fake_batch(portal_failure=True)
        self.assertFalse(result)
        self.assertEqual({row["Store ID"] for row in captured["rows"]}, {"A-1", "B-1"})
        self.assertEqual([(item["portal"], item["stage"]) for item in report["portal_failures"]], [("Portal Error", "autentikasi")])

    def test_no_successful_rows_never_full_refreshes_sot(self):
        result, captured, report = self._run_fake_batch(portal_failure=False)
        self.assertFalse(result)
        self.assertNotIn("rows", captured)
        self.assertEqual(report["sheet_sync"]["status"], "skipped")
        self.assertIn("Tidak ada hasil live sukses", report["sheet_sync"]["message"])


if __name__ == "__main__":
    unittest.main()
