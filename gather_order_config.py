import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

ordersConfigContainer = "BiomeOrdersConfig.cs"
ordersConfigScriptUUID = invuuidMapping[ordersConfigContainer]

globalGather = {}
def gatherbuildingeInfo(fileName,yamlInfo):
    levels = utils.recursiveRemoveUUID(yamlInfo["MonoBehaviour"],uuidMapping)
    levels = utils.removeUnityInfo(levels)
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(ordersConfigScriptUUID,gatherbuildingeInfo)

processedGather = globalGather
with open("output/orders_config.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)