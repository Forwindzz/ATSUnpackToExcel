import json
import os

result={}
result_path={}
baseNameToFileName={}

for root, dirs, files in os.walk("."):
    for name in files:
        if name.endswith(".meta"):
            p = os.path.join(root, name)
            uuid = None
            fileName = os.path.basename(p)
            baseName = fileName.replace(".meta","")
            with open(p,"r",encoding="utf-8") as f:
                lines = f.readlines()
                for line in lines:
                    if line.startswith("guid: "):
                        uuid = line[5:].strip()
            if uuid is None:
                print("Error to find uuid for %s"%p)
                continue
            result[uuid]=baseName
            result_path[uuid]=p[:-5]
            if baseName in baseNameToFileName:
                print("Already exists %s :\n > %s\n > %s"%(baseName,baseNameToFileName[baseName],p))
            baseNameToFileName[baseName] = p
with open("uuid_mapping.json","w",encoding="utf-8") as f:
    json.dump(result,f,indent=4)
with open("uuid_path_mapping.json","w",encoding="utf-8") as f:
    json.dump(result_path,f,indent=4)
with open("filename_mapping.json","w",encoding="utf-8") as f:
    json.dump(baseNameToFileName,f,indent=4)
print("done : %d "%len(result))