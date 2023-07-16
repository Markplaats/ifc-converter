import json
import os
import ifcopenshell
import ifcopenshell.util.element as Element

fileList = []
for file in os.listdir('files'):
    if file.endswith('.ifc'):
        fileList.append(file)

for file in fileList:
    fileName = file.rsplit('.', 1)[0]
    ifcFile = ifcopenshell.open('./files/' + file)

    def getDictFromType(file: ifcopenshell.file, class_type: str):
        objects_data = []
        objects = file.by_type(class_type)
        for object in objects:
            objects_data.append({
                "ExpressID": object.id(),
                "GlobalID": object.GlobalId,
                "Class": object.is_a(),
                "PredefinedType": Element.get_predefined_type(object),
                "Name": object.Name,
                "Level": Element.get_container(object).Name
                if Element.get_container(object)
                else "",
                "ObjectType": Element.get_type(object).Name
                if Element.get_type(object)
                else "",
                "QuantitySets": Element.get_psets(object, qtos_only=True),
                "PropertySets": Element.get_psets(object, psets_only=True)
            })
        return objects_data

    data = getDictFromType(ifcFile, "IfcWall")

    sumDict = {}

    for line in data:
        if "PropertySets" in line and "ARTES Parameters" in line['PropertySets'] and "2.1 Artikel nr" in line['PropertySets']['ARTES Parameters']:
            artNum = line['PropertySets']['ARTES Parameters']['2.1 Artikel nr']
            if artNum in sumDict:
                sumDict[artNum] += line['PropertySets']['ARTES Parameters']['5.1 Oppervlakte']
            else:
                sumDict[artNum] = line['PropertySets']['ARTES Parameters']['5.1 Oppervlakte']

    with open('./export/' + fileName + '_export.json', 'w') as jsonFile:
        json.dump(data, jsonFile)

    with open('./export/' + fileName + '_oppervlakte.json', 'w') as jsonFile:
        json.dump(sumDict, jsonFile)