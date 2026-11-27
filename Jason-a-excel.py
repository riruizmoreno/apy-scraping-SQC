import json
from pathlib import Path

import pandas as pd

script_dir = Path(__file__).resolve().parent

with open(script_dir / "ruff_rules.json", "r", encoding="utf-16") as f:
    rules = json.load(f)

df = pd.json_normalize(rules)

df.to_csv(
    script_dir / "ruff_rules.csv",
    index=False,
    encoding="utf-8-sig"
)

df.to_excel(
    script_dir / "ruff_rules.xlsx",
    index=False
)

print(f"Reglas exportadas: {len(df)}")