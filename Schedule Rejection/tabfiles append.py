"""Combine the interview__key -> interview__id columns of several Survey Solutions
.tab exports into one lookup file (out.tab) for schedule_action.py."""
import os
import sys
from pathlib import Path

import pandas as pd

# One .tab export per questionnaire (real paths: local_config.py -> TAB_FILES).
TAB_FILES = [
    "data/schedule_1A.tab",
    "data/schedule_1B.tab",
    "data/schedule_2.tab",
]
OUT_PATH = Path(__file__).resolve().parent / "out.tab"   # written next to this script

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from local_config import TAB_FILES  # noqa: F811
except ImportError:
    pass

# keep only interview__key and interview__id before merging
frames = [pd.read_csv(p, sep="\t")[["interview__key", "interview__id"]] for p in TAB_FILES]
out = pd.concat(frames, ignore_index=True)

OUT_PATH.parent.mkdir(parents=True, exist_ok=True)   # create the folder if it's missing
out.to_csv(OUT_PATH, sep="\t", index=False)
