import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

oresContainer = "OreContainer.cs"
oresScriptUUID = invuuidMapping[oresContainer]

globalGather = {}
def gatheroreeInfo(fileName,yamlInfo):
    levels = recursiveRemoveUUID(yamlInfo["MonoBehaviour"]["levels"],uuidMapping)
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(oresScriptUUID,gatheroreeInfo)

with open("output/ores_gen.json","w",encoding="utf-8") as f:
    json.dump(globalGather,f,indent=4)