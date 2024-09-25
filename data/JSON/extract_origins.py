import json

with open("gsa.dt", "r", encoding="utf-8") as f:
    dt_dict = json.load(f)

locs:list[str] = []
for x in dt_dict["Arbitri"]:
    if x["Località"] not in locs:
        locs.append(x["Località"].capitalize())
locs.sort()

dt_dict["Città_origine"] = locs
with open("gsa.dt", "w", encoding="utf-8") as f:
    json.dump(dt_dict, f, indent=4, ensure_ascii=False)