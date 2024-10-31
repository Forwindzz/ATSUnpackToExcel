import json
import sys

uuidMapping = {}

with open("uuid_mapping.json","r",encoding="utf-8") as f:
    uuidMapping = json.load(f)

filenames = sys.argv[1:]
if len(filenames)==0:
    filenames=["text.txt"]
for fn in filenames:
    text=None
    with open(fn,"r",encoding="utf-8") as f:
        text = f.read()
    for k,v in uuidMapping.items():
        text = text.replace(k,v)
    fn+="_."+fn.split(".")[-1]
    with open(fn,"w",encoding="utf-8") as f:
        f.write(text)
    print("finish %s"%fn)