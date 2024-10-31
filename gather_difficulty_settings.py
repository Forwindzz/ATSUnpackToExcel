import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

difficultiesContainer = "BiomeDifficultyConfig.cs"
difficultiesScriptUUID = invuuidMapping[difficultiesContainer]

globalGather = {}
def gatherDifficultyInfo(fileName,yamlInfo):
    result=recursiveRemoveUUID(yamlInfo["MonoBehaviour"]["difficultiesData"],uuidMapping)
    globalGather[fileName] = result

integrateSearchUseScriptUUID(difficultiesScriptUUID,gatherDifficultyInfo)

processedGather = globalGather

with open("output/difficulties.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)