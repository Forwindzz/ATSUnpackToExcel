import json
import sys
import os
import yaml

import utils
from utils import *

uuidMapping,invuuidMapping = utils.loadUUIDMapping()

GladeFolderPath = "all_2/ExportedProject/Assets/TextAsset"

gladesRaw = {}
for root, dirs, files in os.walk(GladeFolderPath):
    for fileName in files:
        if fileName.endswith(".json"):
            filePath = os.path.join(root,fileName)
            baseName = os.path.basename(filePath)
            with open(filePath,"r",encoding="utf-8") as f:
                obj = json.load(f)
                if "deposits" in obj: # add on v1.4, filter out biome files
                    gladesRaw[baseName] = obj
                else:
                    print("filtered: "+fileName)

print("Load %d glades json"%len(gladesRaw))

with open("output/glades.json","w",encoding="utf-8") as f:
    json.dump(gladesRaw,f,indent=4)

print("finish glades\n")
# glades generation model


gladeGenModelsContainer = "BiomGenerationModel.cs"
gladeGenModelsScriptUUID = invuuidMapping[gladeGenModelsContainer]
globalGather = {}
def gatherreliceInfo(fileName,yamlInfo):
    gladeGenModels = recursiveRemoveUUID(yamlInfo["MonoBehaviour"],uuidMapping)
    globalGather[fileName] = gladeGenModels

integrateSearchUseScriptUUID(gladeGenModelsScriptUUID,gatherreliceInfo)

with open("output/glades_gen_model.json","w",encoding="utf-8") as f:
    json.dump(globalGather,f,indent=4)

print("finish gen model\n")
# extra glade mechaism

gladeGenModelsContainer = "ExtraGladeEffectModel.cs"
gladeGenModelsScriptUUID = invuuidMapping[gladeGenModelsContainer]
globalGather = {}
def gatherreliceInfo(fileName,yamlInfo):
    gladeGenModels = recursiveRemoveUUID(yamlInfo["MonoBehaviour"]["glade"],uuidMapping)
    globalGather[fileName] = gladeGenModels

integrateSearchUseScriptUUID(gladeGenModelsScriptUUID,gatherreliceInfo)

with open("output/glades_sp_group.json","w",encoding="utf-8") as f:
    json.dump(globalGather,f,indent=4)

print("finish sp group glades")
