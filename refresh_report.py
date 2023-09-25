from pathlib import Path
import re

# Find GfK and GSD files in downloads

gfk_fp = ""
gfk_time = 0
gsd_fp = ""
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

# Move files to archive and rename with date

# Update main data file

# Update report

# Save report