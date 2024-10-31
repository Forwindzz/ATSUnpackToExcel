import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

relicsContainer = "RelicsContainer.cs"
relicsScriptUUID = invuuidMapping[relicsContainer]

globalGather = {}
def gatherreliceInfo(fileName,yamlInfo):
    relics = yamlInfo["MonoBehaviour"]["relics"]
    globalGather[fileName] = relics

integrateSearchUseScriptUUID(relicsScriptUUID,gatherreliceInfo)
#searchUseScriptUUID(relicsScriptUUID,"Relic",gatherreliceInfo)

processedGather = {}
for k,v in globalGather.items():
    #level -> list of relic
    relicsGather={}
    for obj in v:
        obj["relic"] = uuidMapping[obj["relic"]["guid"]]
        lv = obj["level"]
        if lv not in relicsGather:
            relicsGather[lv]=[]
        relicsGather[lv].append(obj)
    processedGather[k]=relicsGather
with open("output/relics_gen.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)