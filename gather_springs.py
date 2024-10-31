import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

springsContainer = "SpringsContainer.cs"
springsScriptUUID = invuuidMapping[springsContainer]

globalGather = {}
def gatherspringeInfo(fileName,yamlInfo):
    levels = recursiveRemoveUUID(yamlInfo["MonoBehaviour"]["springs"],uuidMapping)
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(springsScriptUUID,gatherspringeInfo)

with open("output/springs_gen.json","w",encoding="utf-8") as f:
    json.dump(globalGather,f,indent=4)