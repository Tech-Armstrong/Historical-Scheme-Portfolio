"""Apply scheme_name_mapping.csv to an existing holdings CSV (fills amc_scheme_code, canonical scheme_name).
Usage: python apply_scheme_mapping.py <holdings.csv> [out.csv]"""
import sys, re, pandas as pd
from pathlib import Path
src = Path(sys.argv[1]); out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(src.stem + "_mapped.csv")
m = pd.read_csv(Path(__file__).with_name("scheme_name_mapping.csv"), dtype={"scheme_code": str})
key = lambda s: re.sub(r"[\s\^\*#]+$", "", " ".join(str(s).split()).upper())
mp = {}
for r in m.itertuples():
    mp[key(r.source_scheme_name)] = (r.scheme_code, r.canonical_scheme_name)
    mp.setdefault(key(re.sub(r"\s*\(.*$", "", r.source_scheme_name)), (r.scheme_code, r.canonical_scheme_name))
h = pd.read_csv(src, low_memory=False, dtype={"amc_scheme_code": str})
# Sundaram Value Fund files titled "Sundaram Diversified Equity" (no 'Fund') came through blank in older runs
h.loc[h.scheme_name.isna() & h.source_file.str.contains("Sundaram Value Fund", na=False), "scheme_name"] = "Sundaram Diversified Equity"
res = h.scheme_name.map(lambda s: mp.get(key(s)) or mp.get(key(re.sub(r"\s*\(.*$", "", str(s)))))
miss = h.loc[res.isna(), "scheme_name"].drop_duplicates().tolist()
h.loc[res.notna(), "amc_scheme_code"] = res[res.notna()].str[0]
h.loc[res.notna(), "scheme_name"] = res[res.notna()].str[1]
h.to_csv(out, index=False)
print(f"{len(h)} rows -> {out}; null amc_scheme_code: {h.amc_scheme_code.isna().sum()}; unmapped titles: {miss}")
