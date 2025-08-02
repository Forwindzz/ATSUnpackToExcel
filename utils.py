import json
import sys
import os
import yaml

def any_constructor(loader, tag_suffix, node):
    if isinstance(node, yaml.MappingNode):
        return loader.construct_mapping(node)
    if isinstance(node, yaml.SequenceNode):
        return loader.construct_sequence(node)
    return loader.construct_scalar(node)

yaml.add_multi_constructor('', any_constructor, Loader=yaml.SafeLoader)

def loadJson(filePath):
    with open(filePath, 'r',encoding="utf-8") as f:
        return json.load(f)

def safeDel(obj,k):
    if k in obj:
        del obj[k]

def removeUnityInfo(obj):
    safeDel(obj,"m_ObjectHideFlags")
    safeDel(obj,"m_CorrespondingSourceObject")
    safeDel(obj,"m_PrefabInstance")
    safeDel(obj,"m_GameObject")
    safeDel(obj,"m_Enabled")
    safeDel(obj,"m_EditorHideFlags")
    safeDel(obj,"m_EditorClassIdentifier")
    safeDel(obj,"serializationData")
    safeDel(obj,"m_PrefabAsset")
    return obj

langSettings = loadJson("settings/lang.json")
def getLangSettings():
    return langSettings

def getLangCode():
    return langSettings["use_language"]

def recursiveRemoveFileID0(obj):
    if isinstance(obj,dict):
        result={}
        for k,v in obj.items():
            if isinstance(v,dict) and "fileID" in v:
                if v["fileID"]==0:
                    result[k] = None
                else:
                    result[k]=v
            else:
                result[k] = recursiveRemoveFileID0(v)
        return result
    else:
        return obj

def _safeGetUUIDMapping(key,uuidMapping):
    if key in uuidMapping:
        return uuidMapping[key]
    else:
        print("Warn: Cannot find uuid for "+key)
        return key

def recursiveRemoveUUID(obj,uuidMapping):
    if isinstance(obj,dict):
        result={}
        for k,v in obj.items():
            if isinstance(v,dict) and "guid" in v:
                result[k] = _safeGetUUIDMapping(v["guid"],uuidMapping)
            else:
                result[k] = recursiveRemoveUUID(v,uuidMapping)
        return result
    elif isinstance(obj,list):
        result=[]
        for v in obj:
            if isinstance(v,dict) and "guid" in v:
                result.append(_safeGetUUIDMapping(v["guid"],uuidMapping))
            else:
                result.append(recursiveRemoveUUID(v,uuidMapping))
        return result
    else:
        return obj
    
def search_uuid_in_file(file_path, uuid):
    with open(file_path, 'r',encoding="utf-8") as file:
        for line in file:
            if uuid in line:
                return True
    return False

def search_uuid_in_directory(directory, uuid):
    matching_files = []
    
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.endswith(".yaml"):
                file_path = os.path.join(root, filename)
                
                if search_uuid_in_file(file_path, uuid):
                    matching_files.append(file_path)
    
    return matching_files

def integrateSearchUseScriptUUID(uuid:str, callback:callable):
    for root, dirs, files in os.walk("."):
        for fileName in files:
            if fileName.endswith(".asset"):
                path = os.path.join(root,fileName)
                if search_uuid_in_file(path,uuid):
                    with open(path,"r",encoding="utf-8") as f:
                        try:
                            result = yaml.safe_load(f)
                            print("Try Detect: %s "%(path))
                            if uuid in result["MonoBehaviour"]["m_Script"]["guid"]:
                                print("!> Detect: %s"%path)
                                baseName = os.path.basename(path)
                                callback(baseName,result)
                        except BaseException as e:
                            print("Error :( %s"%str(e))
                            continue

# fast but may miss some files
def searchUseScriptUUID(uuid:str, fileNameContain, callback:callable):
    for root, dirs, files in os.walk("."):
        for fileName in files:
            if fileName.endswith(".asset") and fileNameContain in fileName:
                path = os.path.join(root,fileName)
                with open(path,"r",encoding="utf-8") as f:
                    try:
                        result = yaml.safe_load(f)
                        print("Try Detect: %s "%(path))
                        if uuid in result["MonoBehaviour"]["m_Script"]["guid"]:
                            print("!> Detect: %s"%path)
                            baseName = os.path.basename(path)
                            callback(baseName,result)
                    except BaseException as e:
                        print("Error :( %s"%str(e))
                        continue

def loadUUIDMapping():
    uuidMapping = None
    with open("uuid_mapping.json","r",encoding="utf-8") as f:
        uuidMapping = json.load(f)

    invuuidMapping = {}
    for k,v in uuidMapping.items():
        invuuidMapping[v.strip()]=k
    return uuidMapping,invuuidMapping

def loadUUIDPathMapping():
    uuidMapping = None
    with open("uuid_path_mapping.json","r",encoding="utf-8") as f:
        uuidMapping = json.load(f)

    invuuidMapping = {}
    for k,v in uuidMapping.items():
        invuuidMapping[v.strip()]=k
    return uuidMapping,invuuidMapping



def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb_color):
    return "{:02x}{:02x}{:02x}".format(rgb_color[0], rgb_color[1], rgb_color[2])

def interpolate_color(color1, color2, t):
    rgb1 = hex_to_rgb(color1)
    rgb2 = hex_to_rgb(color2)
    
    interpolated_rgb = (
        int(rgb1[0] + (rgb2[0] - rgb1[0]) * t),
        int(rgb1[1] + (rgb2[1] - rgb1[1]) * t),
        int(rgb1[2] + (rgb2[2] - rgb1[2]) * t)
    )
    
    return rgb_to_hex(interpolated_rgb)