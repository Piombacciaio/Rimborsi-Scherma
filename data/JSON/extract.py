import json

with open("dt.json", "r", encoding="utf-8") as f:
    arbitri = json.load(f)

locs:list[str] = []
for x in arbitri:
    if x["Località"] not in locs:
        locs.append(x["Località"].capitalize())

locs.sort()
with open("città.json", "w", encoding="utf-8") as f:
    json.dump(locs, f, indent=4, ensure_ascii=False)