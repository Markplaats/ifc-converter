import json
import ifcopenshell
import ifcopenshell.util.element as Element
import tkinter as tk
from tkinter import filedialog, ttk
from excelwriter import writeExcel

def main():
    window = tk.Tk()
    window.title("IFC Converter")
    window.resizable(False, False)
    window.geometry('300x200')
    progBar = ttk.Progressbar(window, orient='horizontal', mode='determinate', length=200)
    progBar.pack()
    btn = ttk.Button(window, text="Convert", command=lambda: convertFiles(progBar))
    btn.pack(
        ipadx=5,
        ipady=5,
        expand=True
    )
    window.mainloop()

def convertFiles(progBar: ttk.Progressbar):
    filepath = filedialog.askopenfilenames(title="Select file(s) to convert", filetypes=[("IFC Files", "*.ifc")])
    if filepath == '':
        return
    outputpath = filedialog.askdirectory(title="Location to save files")
    if outputpath == '':
        return

    progPerFile = 100/len(filepath)

    for file in filepath:
        progBar.step(50)
        fileName = file.rsplit('/', 1)[1].rsplit('.', 1)[0]
        ifcFile = ifcopenshell.open(file)
        
        data = getDictFromType(ifcFile, "IfcWall")
        sumDict = {}

        for line in data:
            if "PropertySets" in line and "ARTES Parameters" in line['PropertySets'] and "2.1 Artikel nr" in line['PropertySets']['ARTES Parameters']:
                artNum = line['PropertySets']['ARTES Parameters']['2.1 Artikel nr']
                if artNum in sumDict:
                    sumDict[artNum] += line['PropertySets']['ARTES Parameters']['5.1 Oppervlakte']
                else:
                    sumDict[artNum] = line['PropertySets']['ARTES Parameters']['5.1 Oppervlakte']

        with open(f'{outputpath}/{fileName}_export.json', 'w') as jsonFile:
            json.dump(data, jsonFile)

        with open(f'{outputpath}/{fileName}_oppervlakte.json', 'w') as jsonFile:
            json.dump(sumDict, jsonFile)

        writeExcel(data, fileName, outputpath)

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

if __name__ == '__main__':
    main()