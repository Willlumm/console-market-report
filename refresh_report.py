from datetime import date
from pathlib import Path
import re

# Find GfK and GSD files in downloads

gfk_fp = None
gfk_time = 0
gsd_fp = None
gsd_time = 0

for fp in Path(f"{Path.home()}/Downloads").iterdir():

    if re.fullmatch("^.*NEW_HW_DATA.txt$", str(fp)):
        time = fp.stat().st_ctime
        if time > gfk_time:
            gfk_fp = fp
            gfk_time = time

    elif re.fullmatch("^.*\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\.xlsx$", str(fp)):
        time = fp.stat().st_ctime
        if time > gsd_time:
            gsd_fp = fp
            gsd_time = time

# Replace input files

for fp in Path(".\\input_files").iterdir():
    if (re.fullmatch("^gfk_.*$", str(fp)) and gfk_fp) or (re.fullmatch("^gsd_.*$", str(fp)) and gsd_fp):
        fp.rename(f".\\input_files\\archive\\{fp.name}")
if gfk_fp:
    gfk_fp.rename(f".\\input_files\\gsd_{date.fromtimestamp(gfk_time)}.xlsx")
if gsd_fp:
    gsd_fp.rename(f".\\input_files\\gsd_{date.fromtimestamp(gsd_time)}.xlsx")

# Update main data file

# Update report

# Save report