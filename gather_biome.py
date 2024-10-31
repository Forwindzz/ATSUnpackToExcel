import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

biomesContainer = "BiomeModel.cs"
biomesScriptUUID = invuuidMapping[biomesContainer]

globalGather = {}
def gatherBiomeeInfo(fileName,yamlInfo):
    result=recursiveRemoveUUID(yamlInfo["MonoBehaviour"],uuidMapping)
    globalGather[fileName] = result

integrateSearchUseScriptUUID(biomesScriptUUID,gatherBiomeeInfo)

processedGather = globalGather

with open("output/biomes.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)