import json
import sys
import os
import yaml
import re
import utils
from utils import *

# [uuid]=file name
uuidMapping,invuuidMapping = utils.loadUUIDMapping()
uuidPathMapping,invuuidPathMapping = utils.loadUUIDPathMapping()

def loadFile(path):
    with open(path,"r",encoding="utf-8") as f:
        return yaml.safe_load(f.read())
    
def showProgress(info,total,current,gap=16):
    if current % gap !=0:
        return
    percent = current*100.0/total
    gridShow = int(percent/10)
    gridHide = 10-gridShow
    bar = ">"*gridShow+" "*gridHide
    print("%s \t <%6d | %6d> %7.2f%% |%s|"%(info,total,current,percent,bar),end="\r")

def camel_to_snake(name):
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1).lower()
#--------------------------------------------

contentContainer = "Settings.cs"
contentScriptUUID = invuuidMapping[contentContainer]

globalGather = {}
def gatherreliceInfo(fileName,yamlInfo):
    content = yamlInfo["MonoBehaviour"]
    globalGather[fileName] = content

integrateSearchUseScriptUUID(contentScriptUUID,gatherreliceInfo)
#searchUseScriptUUID(contentScriptUUID,"Relic",gatherreliceInfo)

processedGather = None
if len(globalGather)>1:
    print("Settings asset file is more than 1 :(  ... We use the first file")
elif len(globalGather)==0:
    print("No settings found!")
    exit(-1)

for k,v in globalGather.items():
    processedGather = v

with open("output/content.json","w",encoding="utf-8") as f:
    json.dump(processedGather,f,indent=4)


def process(name:str,fileName:str = None):
    if fileName is None:
        fileName=camel_to_snake(name)
    effects = processedGather[name]
    processedEffects={}
    print("Process %s"%name)
    for i,effect in enumerate(effects):
        showProgress("Processing... ",len(effects),i)
        effectUUID = effect["guid"]
        effectAsset = utils.recursiveRemoveUUID(loadFile(uuidPathMapping[effectUUID]),uuidMapping)
        effectAsset = utils.recursiveRemoveFileID0(utils.removeUnityInfo(effectAsset["MonoBehaviour"]))
        name = uuidMapping[effectUUID]
        processedEffects[name]=effectAsset
    print("\nFinished Processing %s %d, save to file %s.json"%(name,len(effects), fileName))

    with open("output/%s.json"%fileName,"w",encoding="utf-8") as f:
        json.dump(processedEffects,f,indent=4)
        

for k,v in processedGather.items():
    if not isinstance(v,list):
        continue
    #try:
    process(k)
    #except BaseException as e:
    #    print("Error, ignore %s"%k)
    #    print(e)