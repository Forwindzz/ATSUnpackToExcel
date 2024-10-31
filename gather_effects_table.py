import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

#--------------------------------------------

effectTableContainer = "EffectsTable.cs"
effectTableScriptUUID = invuuidMapping[effectTableContainer]

globalGather = {}
def gatherreliceInfo(fileName,yamlInfo):
    effectTable = recursiveRemoveUUID(yamlInfo["MonoBehaviour"],uuidMapping)
    globalGather[fileName] = effectTable

integrateSearchUseScriptUUID(effectTableScriptUUID,gatherreliceInfo)
#searchUseScriptUUID(effectTableScriptUUID,"Relic",gatherreliceInfo)

processedGather = {}
for k,v in globalGather.items():
    processedGather[k]=v
with open("output/effects_table.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)