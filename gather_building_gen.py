import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

buildingsContainer = "BuildingsContainer.cs"
buildingsScriptUUID = invuuidMapping[buildingsContainer]

globalGather = {}
def gatherbuildingeInfo(fileName,yamlInfo):
    levels = yamlInfo["MonoBehaviour"]["buildings"]
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(buildingsScriptUUID,gatherbuildingeInfo)

processedGather = {}
for k,v in globalGather.items():
    buildingsGather={}
    for obj in v:
        obj["building"] = uuidMapping[obj["building"]["guid"]]
        lv = obj["level"]
        if lv not in buildingsGather:
            buildingsGather[lv]=[]
        buildingsGather[lv].append(obj)
    processedGather[k]=buildingsGather
with open("output/buildings_gen.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)