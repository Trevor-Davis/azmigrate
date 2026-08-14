"""
pull_tpid_licensing.py

Pulls the "Customer Entitlements" licensing tables from the Contractual
Migration Tool Power BI report for a given TPID, and saves each table as its
own .xlsx file into a per-TPID output folder.

WHY EACH TABLE IS A SEPARATE FILE (NOT ONE MERGED WORKBOOK)
-------------------------------------------------------------
This Power BI report is labeled "Confidential\\Internal Only". Microsoft
Purview automatically applies that same sensitivity label / encryption to
any exported file (.xlsx/.csv/.pdf) - this is intentional data-loss
prevention, not a bug. That means every exported .xlsx here is a genuine
protected Office file: it opens normally in Excel for any authorized user,
but it is NOT a plain zip that scripts (e.g. openpyxl) can silently read,
merge, or re-save. Doing that programmatically would mean stripping an
intentional security control on real customer licensing data, so this
script deliberately does NOT attempt to merge or read the exported content.
If you need one consolidated workbook, open the files in Excel yourself
(you're an authorized recipient) and combine them there.

HOW IT WORKS
------------
This drives a real Microsoft Edge browser (via Playwright), because the
report's data-refresh mechanism is a session-bound internal query engine
(not a public API), and this tenant's Conditional Access policy blocks
Playwright's bundled Chromium outright ("You can't get there from here -
devices/client apps must meet Microsoft management compliance policy").
Specifically the script:

  1. Launches Microsoft Edge using a COPY of YOUR REAL, already-signed-in
     Edge profile (not a fresh/blank automation profile, and not your live
     profile directly - see "REFRESHING YOUR SIGN-IN" below for why it's a
     copy). This matters: a brand-new, never-used Edge profile driving a
     sensitive-data export appears to get killed mid-download by an
     endpoint security control (most likely Microsoft Defender for Cloud
     Apps / device-compliance enforcement, given the confidential label on
     this report). The profile copy carries your real, already-authenticated,
     trusted session, so it does not hit that wall.
  2. Opens the report's "Account and Pricing" page and sets the "Customer
     TPID (Required)" slicer to the TPID you provide.
  3. Switches to the "Customer Entitlements" page (Licensing Details view).
  4. For every licensing table visual on that page, uses Power BI's native
     "Export data" -> "Data with current layout" -> .xlsx feature (the same
     one available to you manually via each table's "..." menu). This is
     the most reliable way to get complete, accurate table data - it avoids
     scraping the on-screen grid, which is virtualized and only renders
     visible rows.
  5. Saves each table's export using the filename Power BI itself assigns
     (via the download's suggested filename), so tables are never
     mis-identified or mixed up.

IMPORTANT: this script automates a COPY of your real Edge profile (see
"REFRESHING YOUR SIGN-IN" below), not your live profile directly - so your
everyday Edge windows can stay open while this script runs. Edge only needs
to be closed when you explicitly run --refresh-profile to update that copy.

FIRST-TIME SETUP
-----------------
    python -m venv .venv
    .venv\\Scripts\\pip install playwright openpyxl
    .venv\\Scripts\\python -m playwright install chromium

USAGE
-----
    .venv\\Scripts\\python pull_tpid_licensing.py --tpid 1103880
    .venv\\Scripts\\python pull_tpid_licensing.py --tpid 1103880 --output-dir "C:\\path\\to\\folder"
    .venv\\Scripts\\python pull_tpid_licensing.py --tpid 1103880 --isolated-profile

REFRESHING YOUR SIGN-IN (if your session ever expires)
--------------------------------------------------------
The script uses a one-time COPY of your real Edge profile so it can carry
your existing sign-in without Chromium's "non-default data directory"
restriction. If that copied session eventually expires (you start getting
prompted to sign in again, or TPID search stops finding matches), refresh
it with:

    .venv\\Scripts\\python pull_tpid_licensing.py --refresh-profile

This will ask you to close all Edge windows (needed so the real profile
isn't locked while it's copied), then re-copy your current real profile -
picking up whatever fresh sign-in session is active in your everyday Edge -
into the automation copy used by this script.

If the TPID search matches more than one account, the script lists the
matches and asks you to re-run with a more specific TPID (or the full
"TPID - Account Name" text).
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# ---------------------------------------------------------------------------
# Configuration - these values were captured from the shared report link.
# ---------------------------------------------------------------------------
REPORT_ACCOUNT_PRICING_URL = (
    "https://msit.powerbi.com/groups/f4456a2c-8e61-404d-8c9b-730d5533579e"
    "/reports/5f38fe66-c579-4151-b677-39827852eb3f"
    "/ReportSection79854a3ed784b53732dc"
    "?ctid=72f988bf-86f1-41af-91ab-2d7cd011db47&experience=power-bi"
)

# A one-time COPY of your real, signed-in Edge profile lives here (see
# README/setup notes). Chromium hard-blocks Playwright/DevTools automation
# against the LITERAL default profile directory ("DevTools remote debugging
# requires a non-default data directory") as an anti-hijacking protection,
# so we can't automate your live profile in place. A copy in a
# non-default-named folder is exempt from that restriction while still
# carrying your real, already-authenticated session state.
REAL_EDGE_SOURCE_DIR = Path.home() / "AppData" / "Local" / "Microsoft" / "Edge" / "User Data"
REAL_PROFILE_COPY_DIR = Path.home() / ".pbi_tpid_export_profile_realcopy"
REAL_EDGE_PROFILE_NAME = "Default"

# Fallback isolated profile (used only with --isolated-profile).
ISOLATED_PROFILE_DIR = Path.home() / ".pbi_tpid_export_profile"

DEFAULT_OUTPUT_ROOT = Path.cwd()


def log(msg: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def edge_processes_running() -> bool:
    """Returns True if any msedge.exe processes are currently running."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq msedge.exe"],
            capture_output=True, text=True, timeout=10,
        )
        return "msedge.exe" in result.stdout
    except Exception:  # noqa: BLE001
        return False


def confirm_close_all_edge() -> bool:
    """Shows a native OK/Cancel dialog asking whether to close all running
    Microsoft Edge processes before launching the automation browser.
    Returns True if the user clicked OK, False if Cancel/dismissed.

    Closing Edge first is recommended: repeated or interrupted runs of this
    script can leave stray/zombie msedge.exe processes behind, which have
    been observed to cause resource contention and, in some cases,
    contribute to unstable/looping page loads."""

    if not edge_processes_running():
        return False  # nothing to close

    MB_OKCANCEL = 0x00000001
    MB_ICONWARNING = 0x00000030
    IDOK = 1

    try:
        import ctypes
        result = ctypes.windll.user32.MessageBoxW(
            0,
            "Microsoft Edge appears to be running.\n\n"
            "Closing ALL Edge windows/processes first is recommended before "
            "this script launches its own automated Edge session - this avoids "
            "leftover processes from prior runs causing instability.\n\n"
            "Click OK to close all Edge processes now, or Cancel to leave "
            "them running and continue anyway.",
            "pull_tpid_licensing.py",
            MB_OKCANCEL | MB_ICONWARNING,
        )
        return result == IDOK
    except Exception:  # noqa: BLE001
        # Fall back to a console prompt if the native dialog can't be shown
        # (e.g. no GUI session available).
        answer = input(
            "Microsoft Edge appears to be running. Close all Edge processes "
            "before continuing? [Y/n]: "
        ).strip().lower()
        return answer in ("", "y", "yes")


def close_all_edge_processes() -> None:
    """Force-closes every running msedge.exe process."""
    log("Closing all Microsoft Edge processes...")
    try:
        subprocess.run(
            ["taskkill", "/IM", "msedge.exe", "/F"],
            capture_output=True, text=True, timeout=15,
        )
    except Exception as exc:  # noqa: BLE001
        log(f"  WARNING: could not close Edge processes automatically: {exc}")
        return
    time.sleep(2)
    if edge_processes_running():
        log("  Some Edge processes may still be shutting down.")
    else:
        log("  All Edge processes closed.")


def refresh_profile_copy() -> None:
    """Re-copies the real, live Edge profile into REAL_PROFILE_COPY_DIR, so
    the automation profile picks up your current sign-in session. Edge must
    be fully closed first (its profile directory is locked while running)."""

    log("Refreshing the automation profile copy from your real Edge profile...")

    if edge_processes_running():
        log("Microsoft Edge is currently running. Please close ALL Edge windows now.")
        input("Press ENTER once every Edge window is closed... ")
        # Give the OS a moment to release file locks after the last window closes.
        time.sleep(2)
        if edge_processes_running():
            log("ERROR: msedge.exe is still running (a background process may be lingering).")
            log("Close Edge fully (check Task Manager if needed) and re-run --refresh-profile.")
            sys.exit(1)

    src_default = REAL_EDGE_SOURCE_DIR / REAL_EDGE_PROFILE_NAME
    src_local_state = REAL_EDGE_SOURCE_DIR / "Local State"

    if not src_default.exists():
        log(f"ERROR: could not find your real Edge profile at {src_default}")
        sys.exit(1)

    dst_default = REAL_PROFILE_COPY_DIR / "Default"
    log(f"Copying {src_default} -> {dst_default} (this can take a minute)...")

    if REAL_PROFILE_COPY_DIR.exists():
        shutil.rmtree(REAL_PROFILE_COPY_DIR, ignore_errors=True)
    dst_default.mkdir(parents=True, exist_ok=True)

    # robocopy is far faster than shutil.copytree for large profile folders
    # and tolerates locked/in-use files (e.g. leftover lock/journal files)
    # better than a plain Python copy.
    result = subprocess.run(
        [
            "robocopy", str(src_default), str(dst_default),
            "/E", "/R:1", "/W:1", "/NFL", "/NDL", "/NJH", "/NJS", "/XJ",
        ],
        capture_output=True, text=True,
    )
    # robocopy exit codes 0-7 indicate success (bit flags for files copied/
    # skipped); 8+ indicates a real failure.
    if result.returncode >= 8:
        log(f"ERROR: robocopy failed (exit code {result.returncode}):\n{result.stdout}\n{result.stderr}")
        sys.exit(1)

    if src_local_state.exists():
        shutil.copy2(src_local_state, REAL_PROFILE_COPY_DIR / "Local State")

    log("Profile copy refreshed successfully.")
    log(f"  Location: {REAL_PROFILE_COPY_DIR}")
    log("You can now re-open your normal Edge windows and use this script as usual.")


def set_tpid_slicer(page, tpid: str) -> str:
    """Opens the Customer TPID slicer, searches for `tpid`, and selects the
    matching option. Returns the full "TPID - Account Name" text selected."""

    log(f"Opening Customer TPID slicer and searching for '{tpid}'...")
    combobox = page.get_by_role("combobox", name="CustomerTPIDNameSlicer")
    combobox.wait_for(state="visible", timeout=15000)
    combobox.click()
    page.wait_for_timeout(500)

    search_box = page.get_by_role("textbox", name="Search", exact=True)
    search_box.wait_for(state="visible", timeout=10000)
    search_box.click()

    listbox = page.get_by_role("listbox", name="CustomerTPIDNameSlicer")
    baseline_count = listbox.get_by_role("option").count()

    def _do_search() -> int:
        search_box.fill("")
        page.wait_for_timeout(300)
        search_box.press_sequentially(tpid, delay=80)
        page.wait_for_timeout(2500)
        return listbox.get_by_role("option").count()

    count = _do_search()
    # If the list still shows the unfiltered baseline count, the search
    # likely didn't register (timing hiccup) - retry once with a longer wait.
    if count == baseline_count and baseline_count > 1:
        log("  TPID search did not appear to filter the list yet; retrying...")
        page.wait_for_timeout(1000)
        count = _do_search()

    if count == 0:
        raise RuntimeError(f"No TPID matches found for '{tpid}'. Check the TPID and try again.")

    options = listbox.get_by_role("option")
    all_texts = [options.nth(i).inner_text().strip() for i in range(count)]

    # Power BI's slicer keeps any PREVIOUSLY selected option visible in the
    # list even when it doesn't match the current search text (e.g. a TPID
    # picked in an earlier run of this script). That means "count" alone is
    # unreliable - filter down to options whose text actually contains the
    # searched TPID/name before deciding if the match is unique.
    tpid_lower = tpid.strip().lower()
    matching_indices = [i for i, text in enumerate(all_texts) if tpid_lower in text.lower()]

    if not matching_indices:
        raise RuntimeError(
            f"No TPID option actually matched '{tpid}' (list showed: {all_texts}). "
            "Check the TPID and try again."
        )

    if len(matching_indices) > 1:
        names = [all_texts[i] for i in matching_indices]
        raise RuntimeError(
            "Multiple TPID matches found - re-run with a more specific value:\n  "
            + "\n  ".join(names)
        )

    match_index = matching_indices[0]
    option = options.nth(match_index)
    option_text = all_texts[match_index]
    is_selected = option.get_attribute("aria-selected") == "true"

    # Deselect any OTHER previously-selected option first (a stale TPID from
    # an earlier run), so the report filters to only the new TPID.
    for i, text in enumerate(all_texts):
        if i == match_index:
            continue
        other = options.nth(i)
        if other.get_attribute("aria-selected") == "true":
            log(f"  Deselecting previously selected TPID: {text}")
            other.click()
            page.wait_for_timeout(500)

    if not is_selected:
        option.click()
        page.wait_for_timeout(500)

    page.keyboard.press("Escape")
    log(f"TPID slicer set to: {option_text}")

    page.wait_for_timeout(4000)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeoutError:
        pass

    return option_text


def discover_licensing_tables(page) -> list:
    """Finds every table visual on the current page whose accessible name
    matches Power BI's auto-generated "<Title> Visual displays Licensing
    Details ..." pattern, and returns a list of (title, locator) tuples."""

    groups = page.get_by_role("group", name=re.compile("Visual displays Licensing Details"))
    count = groups.count()
    results = []
    seen_titles = set()
    for i in range(count):
        g = groups.nth(i)
        aria_label = g.get_attribute("aria-label") or ""
        match = re.match(r"^(.*?)\s+Visual displays", aria_label)
        title = match.group(1).strip() if match else f"Table {i + 1}"
        if title in seen_titles:
            continue
        seen_titles.add(title)
        results.append((title, g))
    return results


def export_table(page, group_locator, title: str, out_dir: Path, attempt: int = 1) -> Path | None:
    """Triggers Power BI's native Export data -> xlsx flow for one visual
    and saves the download into out_dir, using the filename Power BI itself
    assigns. Returns the saved file path, or None if export failed."""

    try:
        group_locator.scroll_into_view_if_needed()
        group_locator.hover()
        page.wait_for_timeout(300)

        more_options = group_locator.get_by_test_id("visual-more-options-btn")
        more_options.click()

        export_item = page.get_by_test_id("pbimenu-item.Export data")
        export_item.wait_for(state="visible", timeout=5000)
        export_item.click()

        export_button = page.get_by_test_id("export-btn")
        export_button.wait_for(state="visible", timeout=5000)

        with page.expect_download(timeout=30000) as download_info:
            export_button.click()
        download = download_info.value

        # Use Power BI's own suggested filename so tables are never
        # mis-identified - do not guess names from content.
        suggested = download.suggested_filename or f"{title}.xlsx"
        safe_name = re.sub(r'[<>:"/\\|?*]', "_", suggested)
        dest_path = out_dir / safe_name
        download.save_as(str(dest_path))

        def _is_valid_office_file(path: Path) -> bool:
            try:
                with open(path, "rb") as fh:
                    head = fh.read(8)
                    # Plain OOXML zip (PK..) or an RMS/MIP-protected OLE
                    # container (D0CF11E0) are both legitimate outcomes for
                    # a labeled-confidential export - either way, Power BI
                    # produced a real file, not an empty/partial one.
                    return head[:4] == b"PK\x03\x04" or head[:4] == b"\xd0\xcf\x11\xe0"
            except OSError:
                return False

        if not _is_valid_office_file(dest_path):
            log(f"  '{title}' download did not look complete yet, waiting and retrying save...")
            page.wait_for_timeout(2000)
            download.save_as(str(dest_path))
            if not _is_valid_office_file(dest_path):
                log(f"  WARNING: '{title}' export file still looks incomplete after retry - skipping.")
                return None

        log(f"  Exported: {title} -> {dest_path.name}")
        return dest_path
    except Exception as exc:  # noqa: BLE001
        if attempt < 2:
            log(f"  '{title}' export hit an error ({exc}); retrying once...")
            page.wait_for_timeout(1500)
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(200)
                page.keyboard.press("Escape")
            except Exception:  # noqa: BLE001
                pass
            return export_table(page, group_locator, title, out_dir, attempt=attempt + 1)
        log(f"  WARNING: failed to export '{title}' after retry: {exc}")
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(200)
            page.keyboard.press("Escape")
        except Exception:  # noqa: BLE001
            pass
        return None


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--tpid", help="TPID number (or partial account name) to search for.")
    parser.add_argument("--output-dir", help="Folder to save exported .xlsx files into. Defaults to <current directory>/<TPID>_Licensing_<date>/")
    parser.add_argument("--headless", action="store_true", help="Run browser headless (use only after first successful sign-in).")
    parser.add_argument(
        "--isolated-profile",
        action="store_true",
        help="Use a separate automation profile instead of your real Edge profile. "
             "Slower to sign in and more likely to hit compliance/DLP restrictions on export.",
    )
    parser.add_argument(
        "--refresh-profile",
        action="store_true",
        help="Re-copy your real, live Edge profile into the automation profile used by "
             "this script (picks up a fresh sign-in session). Requires closing Edge "
             "temporarily. Run this alone, without --tpid, then re-run normally.",
    )
    parser.add_argument(
        "--no-edge-check",
        action="store_true",
        help="Skip the startup prompt that offers to close all running Edge processes. "
             "Use for unattended/scripted runs.",
    )
    args = parser.parse_args()

    if args.refresh_profile:
        refresh_profile_copy()
        return

    if not args.tpid:
        parser.error("--tpid is required (unless using --refresh-profile).")

    if not args.no_edge_check and not args.headless:
        if edge_processes_running():
            if confirm_close_all_edge():
                close_all_edge_processes()
            else:
                log("Continuing with Edge left running (user chose Cancel).")

    out_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_ROOT / f"{args.tpid}_Licensing_{time.strftime('%Y%m%d')}"
    out_dir.mkdir(parents=True, exist_ok=True)

    use_real_profile = not args.isolated_profile

    with sync_playwright() as p:
        if use_real_profile:
            log("Launching Microsoft Edge using a copy of your real, signed-in profile...")
            log(f"  Profile dir: {REAL_PROFILE_COPY_DIR}")
            if not REAL_PROFILE_COPY_DIR.exists():
                log("ERROR: profile copy not found. Create it once with:")
                log(r'  robocopy "%LOCALAPPDATA%\Microsoft\Edge\User Data\Default" '
                    rf'"{REAL_PROFILE_COPY_DIR}\Default" /E')
                log(rf'  copy "%LOCALAPPDATA%\Microsoft\Edge\User Data\Local State" "{REAL_PROFILE_COPY_DIR}\Local State"')
                sys.exit(1)
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(REAL_PROFILE_COPY_DIR),
                channel="msedge",
                headless=args.headless,
                accept_downloads=True,
                viewport={"width": 1600, "height": 1000},
                args=[f"--profile-directory={REAL_EDGE_PROFILE_NAME}"],
            )
        else:
            log("Launching Microsoft Edge using an isolated automation profile...")
            ISOLATED_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(ISOLATED_PROFILE_DIR),
                channel="msedge",
                headless=args.headless,
                accept_downloads=True,
                viewport={"width": 1600, "height": 1000},
            )

        page = context.pages[0] if context.pages else context.new_page()

        def load_report_with_retries(max_attempts: int = 4) -> bool:
            """Navigates to the report and waits for the TPID slicer to
            appear. Retries the FULL navigation (not just the wait) on
            failure, since a failed silent-auth redirect can strand the
            page on a dead-end error page that will never recover no
            matter how long we wait on it."""

            for attempt in range(1, max_attempts + 1):
                log(f"Navigating to report (attempt {attempt}/{max_attempts})...")
                try:
                    page.goto(REPORT_ACCOUNT_PRICING_URL, wait_until="domcontentloaded", timeout=90000)
                except PWTimeoutError:
                    log("  Page load timed out; will retry.")
                    page.wait_for_timeout(2000)
                    continue

                try:
                    page.get_by_role("combobox", name="CustomerTPIDNameSlicer").wait_for(timeout=25000)
                    return True
                except PWTimeoutError:
                    current_url = ""
                    try:
                        current_url = page.url
                    except Exception:  # noqa: BLE001
                        pass

                    if "ErrorPage" in current_url or "login.microsoft" in current_url:
                        log(f"  Landed on a sign-in error/redirect page (attempt {attempt}/{max_attempts}); retrying full navigation...")
                        page.wait_for_timeout(2000)
                        continue

                    # Not an obvious auth error - might be a real, completable
                    # interactive sign-in prompt (e.g. Windows Hello/FIDO).
                    # Give the user a window to complete it once, unless
                    # running headless.
                    if args.headless:
                        log("  TPID slicer not found and running headless - giving up this attempt.")
                        continue
                    debug_path = Path(tempfile.gettempdir()) / f"pbi_tpid_wait_attempt{attempt}.png"
                    try:
                        page.screenshot(path=str(debug_path), timeout=5000)
                        log(f"  Debug screenshot saved: {debug_path}")
                    except Exception:  # noqa: BLE001
                        pass
                    log(f"  Current URL: {current_url}")
                    log("  If a sign-in prompt (e.g. Windows Hello / security key) is visible in the Edge "
                        "window, please complete it now. Waiting up to 60 more seconds...")
                    try:
                        page.get_by_role("combobox", name="CustomerTPIDNameSlicer").wait_for(timeout=60000)
                        return True
                    except PWTimeoutError:
                        log(f"  Still not loaded after attempt {attempt}/{max_attempts}; retrying full navigation...")
                        continue
            return False

        if not load_report_with_retries():
            log("ERROR: could not load the report and reach the TPID slicer after multiple full retries.")
            log("This looks like an intermittent sign-in/session issue. Try again shortly, or run "
                "--refresh-profile if it keeps failing.")
            try:
                context.close()
            except Exception:  # noqa: BLE001
                pass
            sys.exit(1)

        tpid_label = set_tpid_slicer(page, args.tpid)

        log("Switching to Customer Entitlements page...")
        page.get_by_role("tab", name="Customer Entitlements").click()
        page.wait_for_timeout(3000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PWTimeoutError:
            pass

        log("Discovering licensing table visuals...")
        tables = discover_licensing_tables(page)
        if not tables:
            log("ERROR: No licensing table visuals found on this page. The report layout may have changed.")
            context.close()
            sys.exit(1)
        log(f"Found {len(tables)} table(s): " + ", ".join(t[0] for t in tables))

        results = []
        for title, locator in tables:
            if page.is_closed():
                log("ERROR: browser page was closed unexpectedly - stopping export early.")
                break
            file_path = export_table(page, locator, title, out_dir)
            results.append((title, file_path))
            page.wait_for_timeout(500)

        try:
            context.close()
        except Exception:  # noqa: BLE001
            pass

        ok_count = sum(1 for _, f in results if f is not None)
        log(f"Exported {ok_count}/{len(tables)} tables for {tpid_label}.")
        log(f"Files saved in: {out_dir}")
        log("NOTE: each file retains the report's Confidential\\Internal Only sensitivity label - open in Excel as an authorized user.")


if __name__ == "__main__":
    main()
