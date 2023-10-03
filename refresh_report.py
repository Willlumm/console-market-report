from datetime import date
from pathlib import Path
import pandas as pd
import re
import xlwings as xw

# Find GfK and GSD files in downloads

print("Seaching for new files in Downloads...")

gfk_fp = None
gfk_time = 0
gsd_fp = None
gsd_time = 0

for fp in Path(f"{Path.home()}\\Downloads").iterdir():

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

# Ingest data

classes = {
    "4 PRO":            "PRO",
    "4 SLIM":           "SLIM",
    "DIGITAL EDITION":  "DIGITAL EDITION",
    "SWITCH LITE":      "LITE",
    "OLED":             "OLED",
    "SWITCH 64 GB":     "OLED",
    "XBOX ONE S":       "ONE S",
    "XBOX ONE X":       "ONE X",
    "XBOX SERIES S":    "SERIES S",
    "XBOX SERIES X":    "SERIES X",
}
hdsizes = {
    "32 GB":                "32 GB",
    "64 GB":                "64 GB",
    "OLED":                 "64 GB",
    "250 GB":               "250 GB",
    "500 GB":               "500 GB",
    "500GB":                "500 GB",
    "SONY PLAYSTATION 4":   "500 GB",
    "512":                  "512 GB",
    "825":                  "825 GB",
    "1 TB":                 "1 TB",
    "1TB":                  "1 TB",    
    "2 TB":                 "2 TB",
    "2TB":                  "2 TB"
}
territories = {
    "GSA":      "SWITZERLAND",
    "BENE":     "BENELUX",
    "OCEANIA":  "ANZ",
    "ASIA":     "JAPAN",
}

date_ref = pd.read_csv(".\\ref\\DATES.csv")
dfs = []

if gfk_fp:

    print(f"Ingesting data from {gfk_fp}...")

    gfk = pd.read_csv(gfk_fp, sep="\t")

    gfk["Source"] = "GFK"
    gfk["CLASS"] = "ORIGINAL"
    gfk["HDSize"] = "UNKNOWN"

    gfk = gfk.merge(date_ref, how="left", left_on=["Year (W)", "Week (W)"], right_on=["YEAR", "WEEK"])

    gfk_extrap = pd.read_csv(".\\ref\\EXTRAPOLATION HW GFK.csv")
    gfk = gfk.merge(gfk_extrap, how="left", left_on=["Country", "FY", "Main Platform"], right_on=["Territory", "FY", "Format"])
    gfk[["Units Panel (W)", "Value Panel (W)", "Value Panel EUR (W)"]] = gfk[["Units Panel (W)", "Value Panel (W)", "Value Panel EUR (W)"]].replace(",", "", regex=True)
    gfk["Units 100%"] = gfk["Units Panel (W)"].astype(float) / gfk["Extrapolation"]
    gfk["Value Euro 100%"] = gfk["Value Panel EUR (W)"].astype(float) / gfk["Extrapolation"]
    gfk["Value Local 100%"] = gfk["Value Panel (W)"].astype(float) / gfk["Extrapolation"]

    gfk["Bundle"] = gfk["Bundle"].replace({0: "STANDALONE", 1: "BUNDLE"})
    gfk["Main Platform"] = gfk["Main Platform"].replace("NINTENDO SWITCH", "SWITCH")
    gfk["Territory"] = gfk["Territory"].str.upper()
    gfk.loc[gfk["Article Name"].isna(), "Article Name"] = gfk["Article Name (local)"]
    for string, tag in classes.items():
        gfk.loc[gfk["Article Name"].str.contains(string, na=False), "CLASS"] = tag
    for string, tag in hdsizes.items():
        gfk.loc[gfk["Article Name"].str.contains(string, na=False), "HDSize"] = tag
    
    gfk = gfk.rename(columns={"Article Name": "SKU", "Main Platform": "Platform", "Year (W)": "Year", "Week (W)": "Week", "Units Panel (W)": "Panel Units", "Value Panel (W)": "Panel Value EURO"})
    gfk = gfk[gfk["Platform"].isin(["PS4", "PS5", "SWITCH", "XBOX ONE", "XBOX SERIES"])]
    gfk = gfk[["Source", "SKU", "Platform", "Bundle", "HDSize", "CLASS", "Country", "Territory", "FY", "Year", "MONTH NEW", "Week", "Panel Units", "Panel Value EURO", "Extrapolation", "Units 100%", "Value Euro 100%", "Value Local 100%"]]

    dfs.append(gfk)
    # gfk_fp.rename(f".\\input_files\\gfk_{date.fromtimestamp(gfk_time)}.txt")
    
if gsd_fp:

    print(f"Ingesting data from {gsd_fp}...")

    gsd = pd.read_excel(gsd_fp)
    
    gsd = gsd[gsd["Country"] != "UNITED KINGDOM"]
    gsd = gsd[gsd["Platform"].isin(["PS4", "PS5", "SWITCH", "XBOX ONE", "XBOX SERIES"])]

    gsd["Source"] = "GSD"
    gsd["CLASS"] = "ORIGINAL"

    for string, tag in classes.items():
        gsd.loc[gsd["SKU"].str.contains(string), "CLASS"] = tag
    gsd["Territory"] = gsd["Territory"].replace(territories)

    gsd = gsd.merge(date_ref, how="left", left_on=["Year", "Week"], right_on=["YEAR", "WEEK"])

    gsd_extrap = pd.read_csv(".\\ref\\EXTRAPOLATION HW GSD.csv")
    gsd_extrap["Territory"] = gsd_extrap["Territory"].str.upper()
    gsd = gsd.merge(gsd_extrap, how="left", left_on=["Country", "FY", "Week", "Platform"], right_on=["Territory", "FY", "Week", "Format"])
    gsd["Units 100%"] = gsd["Units"] / gsd["Extrapolation"]
    gsd["Value Euro 100%"] = gsd["Values"] / gsd["Extrapolation"]
    gsd["Value Local 100%"] = ""
    
    # Rename and return data.
    gsd = gsd.rename(columns={"HD Size": "HDSize", "Territory_x": "Territory", "Units": "Panel Units", "Values": "Panel Value EURO"})
    gsd = gsd[["Source", "SKU", "Platform", "Bundle", "HDSize", "CLASS", "Country", "Territory", "FY", "Year", "MONTH NEW", "Week", "Panel Units", "Panel Value EURO", "Extrapolation", "Units 100%", "Value Euro 100%", "Value Local 100%"]]
    
    dfs.append(gsd)
    # gsd_fp.rename(f".\\input_file_archive\\gsd_{date.fromtimestamp(gsd_time)}.xlsx")
    
if dfs:
    df = pd.concat(dfs)
else:
    print("No new files found.")
    exit()

# Update report

df = df[df["Platform"].isin(["PS4", "PS5", "SWITCH", "XBOX ONE", "XBOX SERIES"])]
df = df[df.FY >= 2021]

print(f"Inserting {len(df):,} rows in report...")

report_fp = None
report_cywk = 0

for fp in Path(".\\reports").iterdir():
    match = re.fullmatch(".*EUA Weekly Console HW Report CY(\d\d) WK(\d\d).xlsx", str(fp))
    if match:
        cy, wk = match.groups()
        cywk = 100 * int(cy) + int(wk)
        if cywk > report_cywk:
            report_fp = fp
            report_cywk = cywk

wb = xw.Book(report_fp)
sheet = wb.sheets("weekly")
sheet.range("A:R").clear_contents()
sheet["A1"].options(index=False, header=True).value = df