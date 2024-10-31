import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

buleprintGroupContainer = "BuildingsWeightedContainer.cs"
buleprintGroupScriptUUID = invuuidMapping[buleprintGroupContainer]

globalGather = {}
def gatherbuildingeInfo(fileName,yamlInfo):
    levels = utils.recursiveRemoveUUID(yamlInfo["MonoBehaviour"],uuidMapping)
    levels = utils.removeUnityInfo(levels)
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(buleprintGroupScriptUUID,gatherbuildingeInfo)

integrateSearchUseScriptUUID(invuuidMapping["BuildingsWeightedContainer.cs"],gatherbuildingeInfo)

processedGather = globalGather
with open("output/blueprints_gen.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)