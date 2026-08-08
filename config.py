"""Greenwich path resolution. One variable provisions a host; no module hardcodes a path.

Every store, reference model and output location resolves from AVIA_LOCAL_CACHE. Set that
one variable and nothing in the code changes between the Dev PC, the workstation and a
container. This is point 4 of the Avia tool standard, and the reason for it is on record:
three absolute paths from three different working sessions were left in this tree, and each
one broke silently on the next host.

Resolution order for the data root:
  1. AVIA_LOCAL_CACHE, if set
  2. AVIA_GREENWICH_DATA, if set (a per-tool override, for a host that splits its stores)
  3. the first of the known defaults that exists on this machine

Paths are found by landmark, never by counting folders up from __file__. find_reference()
searches the data root for a named file and reports every path it tried when it fails,
because a resolver that returns None in silence is how a neutral default gets substituted
for a missing model.

Author: Avia Solutions.
"""
import os
import sys

# Known data roots, in preference order. A host that uses none of these sets AVIA_LOCAL_CACHE.
_DEFAULT_ROOTS = (
    r"E:\Avia\greenwich",          # workstation and Dev PC: the single data root
    r"C:\Avia\greenwich",          # Dev PC fallback while stores are still on C:
    "/data/greenwich",             # container
)

_REFERENCE_DIRNAME = "reference"
_OUTPUT_DIRNAME = "outputs"


class ConfigError(RuntimeError):
    """Raised when a required path cannot be resolved. Never returns a default in silence."""


def data_root():
    """The Greenwich data root. Raises ConfigError listing what was tried, rather than guessing."""
    for var in ("AVIA_LOCAL_CACHE", "AVIA_GREENWICH_DATA"):
        v = os.environ.get(var)
        if v:
            if os.path.isdir(v):
                return v
            raise ConfigError(
                f"{var} is set to {v!r} but that directory does not exist. "
                "Point it at the Greenwich data root, or unset it to fall back to the defaults."
            )
    tried = []
    for p in _DEFAULT_ROOTS:
        tried.append(p)
        if os.path.isdir(p):
            return p
    raise ConfigError(
        "No Greenwich data root found. Set AVIA_LOCAL_CACHE to the folder holding "
        f"{_REFERENCE_DIRNAME}/. Tried: " + ", ".join(tried)
    )


def reference_dir():
    """Where the client and reference workbooks live. Never inside the repo."""
    d = os.path.join(data_root(), _REFERENCE_DIRNAME)
    if not os.path.isdir(d):
        raise ConfigError(
            f"Reference folder not found at {d}. It holds the client and reference "
            "workbooks, which are deliberately outside the repository."
        )
    return d


def output_dir():
    """Where builds are written. Created on demand; outputs never belong in the repo."""
    d = os.environ.get("AVIA_GREENWICH_OUT") or os.path.join(data_root(), _OUTPUT_DIRNAME)
    os.makedirs(d, exist_ok=True)
    return d


def find_reference(filename):
    """Locate a reference workbook by name, searching the reference folder and one level below.

    Reports every path tried on failure. Do not replace this with a relative path built by
    going up N levels from __file__: that is what put 'C:\\app' into a live path in Meridian.
    """
    root = reference_dir()
    tried = []
    direct = os.path.join(root, filename)
    tried.append(direct)
    if os.path.isfile(direct):
        return direct
    for entry in sorted(os.listdir(root)):
        sub = os.path.join(root, entry)
        if os.path.isdir(sub):
            cand = os.path.join(sub, filename)
            tried.append(cand)
            if os.path.isfile(cand):
                return cand
    raise ConfigError(
        f"Reference workbook {filename!r} not found. Tried:\n  " + "\n  ".join(tried)
    )


def python_exe():
    """The interpreter to invoke for a child step. Never the literal string 'python3'.

    The Dev PC runs py -3.12 and has no 'python3' on PATH; the workstation and any container
    do. sys.executable is correct on all three.
    """
    return sys.executable


def describe():
    """One line per resolved path, for check_env.py and for printing at the head of a run."""
    lines = []
    try:
        lines.append(f"data root      : {data_root()}")
    except ConfigError as e:
        lines.append(f"data root      : UNRESOLVED, {e}")
        return "\n".join(lines)
    for label, fn in (("reference dir  ", reference_dir), ("output dir     ", output_dir)):
        try:
            lines.append(f"{label}: {fn()}")
        except ConfigError as e:
            lines.append(f"{label}: UNRESOLVED, {e}")
    lines.append(f"interpreter    : {python_exe()}")
    return "\n".join(lines)


if __name__ == "__main__":
    print(describe())
