import os
import pandas as pd


TARGET_SHEETS = ["SABERJET S", "SABERJET XP", "TITAN", "VOYAGER"]
COLUMN_ALIASES = {
    "Date": ["Date", "DATE"],
    "Division": ["Division"],
    "Operator": ["Operator"],
    "SLAB ID": ["SLAB ID"],
    "SF Utilized": ["SF Utilized", "SFT"],
    "Waste": ["Sf Waste", "Waste"],
    "Material Name": ["Marerial Name", "Material Name"],
    "Job Name": ["Job Name", "JOB NAME ", "JOB NAME"],
}


def normalize_columns(df, aliases):
    mapped = {}
    for canonical_name, variants in aliases.items():
        for variant in variants:
            if variant in df.columns:
                mapped[variant] = canonical_name
    return df.rename(columns=mapped)


def export_path_for(sheet_name):
    return os.path.join("exports", f"{sheet_name}.csv")


def combine_exports():
    frames = []
    for sheet in TARGET_SHEETS:
        df = pd.read_csv(export_path_for(sheet))
        df = normalize_columns(df, COLUMN_ALIASES)
        selected = [column for column in COLUMN_ALIASES if column in df.columns]
        df = df[selected]
        df["Source"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
