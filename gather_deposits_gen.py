import json
import sys
import os
import yaml

import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()

depositsContainer = "DepositsContainer.cs"
depositsScriptUUID = invuuidMapping[depositsContainer]

globalGather = {}
def gatherDepositeInfo(fileName,yamlInfo):
    levels = yamlInfo["MonoBehaviour"]["levels"]
    globalGather[fileName] = levels

integrateSearchUseScriptUUID(depositsScriptUUID,gatherDepositeInfo)

processedGather = {}
for k,v in globalGather.items():
    #print("===================",k)
    depositsGather={}
    for obj in v:
        lv = obj["level"]
        #print("---------",lv)
        chances = []
        for x in obj["chances"]:
            amount = x["amount"]
            target = uuidMapping[x["deposit"]["guid"]]
            #print(x["deposit"]["guid"], uuidMapping[x["deposit"]["guid"]])
            chances.append({"amount":amount,"deposit":target})
        depositsGather[lv]=chances
    processedGather[k]=depositsGather
with open("output/deposites_gen.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)