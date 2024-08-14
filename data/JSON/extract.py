import json

with open("dt.json", "r", encoding="utf-8") as f:
    arbitri = json.load(f)

locs:list[str] = []
for x in arbitri:
    if x["Località"] not in locs:
        locs.append(x["Località"])

with open("Città.json", "w") as f:
    json.dump(locs, f)