"""Check tracked and unignored files without printing sensitive values.

This is a local guard, not a replacement for review or a VBA compiler.
Run with Python 3 from any directory inside this repository.
"""

import pathlib
import re
import subprocess
import sys


def git(*args, **kwargs):
    return subprocess.run(["git", *args], check=True, stdout=subprocess.PIPE, **kwargs).stdout


root = pathlib.Path(git("rev-parse", "--show-toplevel").decode().strip())
files = sorted(set(git("-C", str(root), "ls-files", "--cached", "--others",
                       "--exclude-standard", "-z").decode().split("\0")) - {""})
rules = {
    "private key": re.compile(rb"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    "provider token": re.compile(rb"(?:gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9_-]{20,}|AKIA[A-Z0-9]{16})"),
    "credential URL": re.compile(rb"[a-z][a-z0-9+.-]*://[^\s/:@]+:[^\s/@]+@", re.I),
    "quoted bank/account number": re.compile(rb'["\x27][0-9]{10,18}["\x27]'),
    "local user path": re.compile(rb"(?:/Users/|/home/)[A-Za-z0-9_.-]+/|[A-Z]:\\Users\\[A-Za-z0-9_.-]+\\"),
}
private_suffixes = {
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".csv", ".tsv",
    ".pem", ".key", ".p12", ".pfx", ".db", ".sqlite", ".sqlite3", ".sql",
    ".dump", ".bak", ".backup", ".old", ".copy", ".zip", ".rar", ".7z", ".exe",
}
failures = []
for name in files:
    path = root / name
    if not path.exists():
        failures.append((name, "tracked file missing; review index"))
        continue
    if path.is_symlink():
        failures.append((name, "symlink requires manual review"))
        continue
    if path.suffix.lower() in private_suffixes or (path.name.startswith(".env") and path.name != ".env.example"):
        failures.append((name, "private/generated file in publication set"))
    if path.stat().st_size > 100 * 1024 * 1024:
        failures.append((name, "file exceeds 100 MiB"))
        continue
    data = path.read_bytes()
    for label, pattern in rules.items():
        for match in pattern.finditer(data):
            line = data.count(b"\n", 0, match.start()) + 1
            failures.append((f"{name}:{line}", label))

# Check both ignored files that exist and representative future private files.
for name in [".env", ".env.production", "MacroFone.xlsm", "MacroPartFed.xlsm",
             "private/input.xlsx", "local/payroll.csv", "backup.sql", "private.key"]:
    result = subprocess.run(["git", "-C", str(root), "check-ignore", "--no-index", "-q", name])
    if result.returncode != 0:
        failures.append((name, "private file is not ignored"))

example = root / ".env.example"
expected = {"PARTFED_CUENTA_" + name for name in ("SERFIN", "BANCOMER", "BANAMEX", "BANORTE", "HSBC")}
entries = [line.strip().split("=", 1) for line in example.read_text().splitlines()
           if line.strip() and not line.lstrip().startswith("#")]
if (len(entries) != len(expected) or any(len(pair) != 2 for pair in entries)
        or {pair[0] for pair in entries} != expected
        or any(pair[1] != "your_account_here" for pair in entries)):
    failures.append((".env.example", "unexpected variables or non-placeholder values"))

for name in ("PartFed.bas", "fone.bas"):
    source = (root / name).read_text(encoding="ascii")
    if 'Do While Archivos <> ""' not in source or "Option Explicit" not in source:
        failures.append((name, "missing loop/variable declaration guard"))
    if "CleanUp:" not in source or "On Error GoTo Failed" not in source:
        failures.append((name, "missing error cleanup"))

for path, rule in failures:
    print(f"FAIL {path}: {rule}")
print(f"Checked {len(files)} publication files; {len(failures)} finding(s).")
print("VBA compilation and Excel execution require the manual checklist.")
sys.exit(bool(failures))
