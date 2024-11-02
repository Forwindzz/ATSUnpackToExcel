import json
import os
import yaml

import utils
from utils import *
results={}
def search_displayName_in_file(file_path):
    flag=False
    with open(file_path, 'r',encoding="utf-8") as file:
        for line in file:
            if "displayName" in line:
                flag=True
                continue
            if flag:
                if "key: " in line:
                    return line.strip()[5:]
    return None

targetLanguagePath = "./all_2/ExportedProject/Assets/Resources/texts/%s.json"%utils.getLangCode()
languageMapping = {}
extraMapping = {}
with open(targetLanguagePath, 'r',encoding="utf-8") as f:
    languageMapping = json.load(f)
with open(os.path.join("settings",utils.getLangCode(),"add_translate.json") , 'r',encoding="utf-8") as f:
    extraMapping = json.load(f)

for k,v in extraMapping.items():
    results[k]=v
for root, dirs, files in os.walk("."):
    for fileName in files:
        if fileName.endswith(".asset"):
            path = os.path.join(root,fileName)
            fileBaseName = os.path.basename(path)
            result = search_displayName_in_file(path) # display name place holder
            if result is None:
                continue
            if result in languageMapping:
                results[fileBaseName]=languageMapping[result]
                results[fileBaseName.replace(".asset","")]=languageMapping[result]
            else:
                results[fileBaseName]=result
                print("Cannot find translation: %s"%result)

for k,v in languageMapping.items():
    results[k]=v


with open("output/display_names.json","w",encoding="utf-8") as f:
    json.dump(results,f,indent=4)
print("done : %d "%len(results))