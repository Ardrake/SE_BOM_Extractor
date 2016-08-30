# -*- coding: utf-8 -*-
import pySqlOperation
import clr
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")
import System.Runtime.InteropServices as SRI

objOccurrences = None
objSubOccurenceSets = None

objApplication = SRI.Marshal.GetActiveObject("SolidEdge.Application")
objDocument = objApplication.ActiveDocument
objOccurenceSets = objDocument.Occurrences


def extractsuboccurrence(objOccurenceSets, level=1, nomenclature=[]):
    #Extract part object, in index 0, 2nd level of assembly Index 1, part name,
    # Index 2,3 and 4 are Length, Witdh and heigth
    #part = ["groupname", "order", "assembly", "part", "longueur", "largeur", "Epaisseur","Materiel"]

    for i in range(objOccurenceSets.Count):
        objsuboccurrences = objOccurenceSets.Item(i + 1)
        if not objsuboccurrences.Subassembly:
            try:
                variableset = objsuboccurrences.SubOccurrenceDocument.Variables
            except:
                variableset = None
            try:
                modelsset = objsuboccurrences.SubOccurrenceDocument.Models
            except:
                modelsset = None
            if CheckBody(modelsset):
                SE_long = 0
                SE_larg = 0
                SE_epai = 0
                for i in range(variableset.Count):
                    variables = variableset.Item(i + 1)
                    if variables.ExposeName == "SE_LONGUEUR":
                        SE_long = convert(variables.Value)
                    if variables.ExposeName == "SE_LARGEUR":
                        SE_larg = convert(variables.Value)
                    if variables.ExposeName == "SE_EPAISSEUR":
                        SE_epai = convert(variables.Value)
                SE_Material = objsuboccurrences.SubOccurrenceDocument.Properties.Item(7).Item(1).Value
                SE_DocName = objsuboccurrences.SubOccurrenceDocument.Properties.Item(5).Item(1).Value
                SE_TopName = findnamebylevel(objsuboccurrences,level-1)
                if SE_DocName == '':
                    SE_DocName = objsuboccurrences.Name
                part = [assigngroupname(SE_TopName)[0], assigngroupname(SE_TopName)[1], SE_TopName,
                        objOccurenceSets.Parent.Name, SE_DocName, (SE_epai, SE_larg, SE_long), SE_Material]
                nomenclature.append(part)
        if objsuboccurrences.SubOccurrences is not None:
            extractsuboccurrence(objsuboccurrences.SubOccurrences, level+1)

    return nomenclature


def convert(val, frac=0.03125):
    #convert decimal to inches
    return round((int((round(val, 8) / 0.0254) / frac))*frac, 4)


def BuildQuery(BOM):
    # Build SQL Query to insert list data in SQL Table
    keys_to_sql = "SE_groupname, SE_order, SE_top, SE_pere, SE_piece, SE_epaisseur, SE_largeur, SE_longueur, SE_Materiel"
    sqlquery = "INSERT INTO SE_BOM("+keys_to_sql+")"
    tot_rec = 0
    for item in BOM:
        vars_to_sql = []
        tot_rec += 1
        for i in item:
            value_type = type(i)
            if value_type == tuple:
                for n in i:
                    vars_to_sql.append("'"+str(n)+"'")
            if value_type <> tuple:
                vars_to_sql.append("'"+i+"'")
        final_to_sql = ','.join(vars_to_sql)
        if tot_rec == len(BOM):
            sqlquery = sqlquery + "SELECT " + final_to_sql + " GO"
        else:
            sqlquery = sqlquery + "SELECT " + final_to_sql + " UNION ALL "
    return sqlquery


def CheckBody(ModelSet):
    #Function to determine if Object has a body
    try:
        for Model in ModelSet:
            if Model.Body is not None:
                bodyvalid = True
    except:
        bodyvalid = False
    return bodyvalid


def findnamebylevel(root, level):
    currentelement = root
    for currentLevel in range(level):
        currentelement = currentelement.Parent
    # current element should now be the one we are searching for
    se_docname = currentelement.OccurrenceDocument.Properties.Item("ProjectInformation").Item("Document Number").Value
    if se_docname == "":
        return currentelement.name
    else:
        return se_docname


def assigngroupname(mur):
    se_commentaire = objDocument.Properties.Item("SummaryInformation").Item("Commentaires").Value
    listgroup = se_commentaire.splitlines()
    result = ""
    for idx, val in enumerate(listgroup):
        if mur in val:
            result = ("MUR " + val, str(idx+1))
    if result == '':
        result = ("MUR " + mur, "9")
    #print result
    return result


MP = extractsuboccurrence(objOccurenceSets)

#BuildQuery(MP)
pySqlOperation.ZapSE_BOM()
pySqlOperation.executeSQLQuery(BuildQuery(MP))





