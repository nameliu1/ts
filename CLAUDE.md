# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository purpose

This repository is a Windows-oriented active reconnaissance toolkit that chains three external scanners:

- `ts.exe` / `ts` for port and URL discovery
- `spray.exe` for directory scanning against discovered URLs
- `ehole.exe` / `ehole finger` for web fingerprinting

The Python code is primarily orchestration and result post-processing around those binaries. Most work in this repo is about keeping the scan pipeline, file handoff, and Excel/TXT outputs consistent.

## Environment assumptions

- The scripts are written to run on Windows even if they are edited from another environment.
- Batch files call `python`, `ts`, `spray.exe`, and `ehole` directly, so those commands must be available from the working directory or `PATH`.
- The scripts expect UTF-8 / `chcp 65001` style console behavior for Chinese output.
- `config.yaml` contains Windows paths such as `CfgPath` and `CfgOutPath`; `2.py` reads `CfgOutPath` and may copy generated `url.txt` / `port.txt` back into the repo root.

## Python dependencies

Install the libraries used by the processing scripts:

```bash
pip install pandas openpyxl psutil
```

`1.py` hard-fails if `psutil` or `pandas` are missing. `process_data.py` and `ppp.py` also depend on `openpyxl`.

## Common commands

Run these from the repository root.

### Full pipeline

```bat
python 2.py
python ppp.py
python 1.py
```

This is the same effective flow wrapped by `轮子top100.bat` and `轮子top1000.bat`, with the caveat that both batch files currently invoke the same Python entrypoints and do not themselves change the port set.

### Batch entrypoints

```bat
轮子top100.bat
轮子top1000.bat
top100仅端口.bat
top1000仅端口.bat
仅域名.bat
端口处理.bat
小字典.bat
```

### Port scanning only

```bat
python 2.py
python ppp.py
```

`2.py` runs `ts -hf ip.txt -portf ports.txt -np -m port,url` and normalizes `url.txt` / `port.txt`. `ppp.py` parses `port.txt` into an Excel report.

### Full URL / directory / fingerprint processing from existing inputs

```bat
python 1.py
```

`1.py` assumes `url.txt` already exists, launches `spray.exe`, processes `res.json` via `process_data.py`, filters status-200 URLs, then runs `ehole finger` on the filtered URL list.

### Re-process or beautify outputs

```bash
python process_data.py res.json output.xlsx
python process_data.py existing_ehole.xlsx existing_ehole.xlsx
python ppp.py
```

- Use the JSON form to turn Spray line-delimited JSON into Excel plus a TXT URL list.
- Use the Excel form to beautify an existing `ehole` workbook in place.
- Use `ppp.py` to regenerate the port scan Excel report from `port.txt`.

### Native scanner commands used by the scripts

```bat
ts -hf ip.txt -portf ports.txt -np -m port,url,js
ts -hf ip.txt -portf ports.txt -np -m port,url
ts -hf ip.txt -np -m port,url,js
ts -hf ip.txt -np -m port,url
spray.exe -l url.txt -d dirv2.txt -f res.json
ehole finger -l url.txt -o result.xlsx -t 10
```

These are documented in `命令.txt` and `AGENTS.md`; use them when debugging outside the Python wrappers.

## Key inputs and outputs

Inputs expected in the repo root:

- `ip.txt`: IP list, one target per line
- `ports.txt` / `port.txt`: port lists and/or scanner output, depending on stage
- `url.txt`: URL list produced by `ts` or extracted from `port.txt`
- `dirv2.txt` / `dirv3.txt`: directory brute-force dictionaries
- `finger.json`: fingerprint data for the external tooling
- `config.yaml`: scanner UI/runtime configuration, including `CfgOutPath` and `IpSelectedStrategy`

Generated outputs:

- `port_scan_report_YYYYMMDD_HHMMSS.xlsx` from `ppp.py`
- `res.json` from `spray.exe`
- `res_processed*.xlsx` and `res_processed*.txt` from `process_data.py`
- `MMDD/` date folders created by `1.py`
- `MMDD/spray_original_YYYYMMDD*.json`
- `MMDD/spray_processed_YYYYMMDD*.xlsx`
- `MMDD/YYYYMMDD_status200_urls_N*.txt`
- `MMDD/ehole_result_YYYYMMDD*.xlsx`

## Architecture

### 1. Discovery stage: `2.py`

`2.py` is the entrypoint for Tscan-based discovery.

- Reads `CfgOutPath` from `config.yaml`
- Runs `ts -hf ip.txt -portf ports.txt -np -m port,url`
- Verifies whether `url.txt` and `port.txt` were written locally or under the configured output path
- Copies those files back into the repo root when needed
- Falls back to extracting URLs from `port.txt` if `url.txt` was not produced
- Can generate a quick Excel summary of parsed URL data

The important design detail is that this script bridges between the scanner's configured output directory and the repository working directory.

### 2. Port report stage: `ppp.py`

`ppp.py` parses raw `port.txt` output into structured Excel.

It recognizes three line families:

- simple `IP:PORT STATUS` lines
- fingerprint lines such as protocol/component/version tuples
- URL lines containing status code, fingerprint, URL, and title

It writes a unified workbook with conditional formatting, per-source coloring, and frozen header panes. If parsing changes, this is the place to update regex handling.

### 3. Spray orchestration stage: `1.py`

`1.py` is the main pipeline coordinator after URLs exist.

Sequence:

1. Create a `MMDD` output directory
2. Delete transient files like `url.txt.stat` and previous processed text output
3. Run `spray.exe -l url.txt -d dirv2.txt -f res.json`
4. Wait for the process and the `res.json` artifact
5. Invoke `process_data.py` to convert Spray JSON into Excel and TXT outputs
6. Filter status-200 URLs into a dated TXT file
7. Move Spray raw/processed artifacts into the date folder
8. Run `ehole finger` against the filtered URL list
9. Re-run `process_data.py` on the resulting `ehole` workbook to beautify it

Notable implementation detail: `monitor_process()` has a special fast-finish path for `ehole.exe`, because the process may exit before polling catches it.

### 4. Result post-processing stage: `process_data.py`

`process_data.py` has two modes selected by input extension:

- `.json`: parse Spray line-delimited JSON into a DataFrame, drop `redirect_url`, normalize column order, extract status-200 URLs, write Excel plus TXT, and apply Excel formatting
- `.xlsx` / `.xls`: treat the file as an `ehole` result workbook and beautify it in place

This file is the main formatting and artifact-normalization layer in the repo.

## File and control-flow relationships

The effective data flow across scripts is:

`ip.txt` + `ports.txt` -> `2.py` / `ts` -> `url.txt` + `port.txt`

`port.txt` -> `ppp.py` -> `port_scan_report_*.xlsx`

`url.txt` + `dirv2.txt` -> `1.py` / `spray.exe` -> `res.json`

`res.json` -> `process_data.py` -> `res_processed*.xlsx` + `res_processed*.txt`

`res_processed*.xlsx` -> status-200 filtering in `1.py` -> dated URL TXT

status-200 URL TXT -> `ehole finger` -> `ehole_result_*.xlsx`

`ehole_result_*.xlsx` -> `process_data.py` -> beautified final workbook

## Maintenance notes for future edits

- Keep Windows process-launch behavior in mind: `1.py` uses `os.system("start cmd /c ...")`, not `subprocess.run`, to preserve native console behavior for external executables.
- Do not assume `url.txt` is always produced directly by `ts`; `2.py` intentionally reconstructs it from `port.txt` when necessary.
- `1.py` and `process_data.py` both assume Spray output semantics where URL data is effectively in column `E` and status code in column `J`; changes to upstream JSON shape can break filtering.
- `config.yaml` is not just documentation; `2.py` reads it at runtime, so changes to `CfgOutPath` can affect artifact discovery.
- Date-folder organization is part of the workflow, not just a convenience; preserve it unless intentionally changing output layout.
