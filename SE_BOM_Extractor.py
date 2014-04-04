# -*- coding: utf-8 -*-
import pySqlOperation
import clr
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")
import System.Runtime.InteropServices as SRI

objOccurenceSets = None
objOccurrences = None
objSubOccurenceSets = None
objSubOccurrences = None

objApplication = SRI.Marshal.GetActiveObject("SolidEdge.Application")
objDocument = objApplication.ActiveDocument
objOccurenceSets = objDocument.Occurrences
            
def ExtractSubOccurrence(objOccurenceSets, level=1, nomenclature = []):
    #Extract part object, in index 0, 2nd level of assembly Index 1, part name, Index 2,3 and 4 are Length, Witdh and heigth
    part = ["groupname", "order", "assembly", "part", "longueur", "largeur", "Epaisseur","Materiel"]
    
    for i in range(objOccurenceSets.Count):
        objSubOccurrences = objOccurenceSets.Item(i + 1)
        if objSubOccurrences.Subassembly == False:
            try:
                VariableSet = objSubOccurrences.SubOccurrenceDocument.Variables
            except:
                VariableSet = None
            try:    
                ModelsSet = objSubOccurrences.SubOccurrenceDocument.Models
            except:
                ModelsSet = None
            if CheckBody(ModelsSet):
                SE_long = 0
                SE_larg = 0 
                SE_epai = 0
                for i in range(VariableSet.Count):
                    Variables = VariableSet.Item(i + 1)
                    if Variables.ExposeName == "SE_LONGUEUR":
                        SE_long = convert(Variables.Value)
                    if Variables.ExposeName == "SE_LARGEUR":
                        SE_larg = convert(Variables.Value)
                    if Variables.ExposeName == "SE_EPAISSEUR":
                        SE_epai = convert(Variables.Value)
                SE_Material = objSubOccurrences.SubOccurrenceDocument.Properties.Item(7).Item(1).Value
                SE_DocName = objSubOccurrences.SubOccurrenceDocument.Properties.Item(5).Item(1).Value
                SE_TopName = findNameByLevel(objSubOccurrences,level-1)
                if SE_DocName == '':
                    SE_DocName = objSubOccurrences.Name
                part = [assignGroupName(SE_TopName)[0],assignGroupName(SE_TopName)[1], SE_TopName, objOccurenceSets.Parent.Name, SE_DocName, (SE_epai, SE_larg, SE_long),SE_Material]
                nomenclature.append(part)
        if objSubOccurrences.SubOccurrences != None: 
            ExtractSubOccurrence(objSubOccurrences.SubOccurrences, level+1)
    
    return nomenclature

            
def convert(val,frac=0.125):
    #convert decimal to inches
    return  round((int((round(val,8) / 0.0254) / frac))*frac,4)  

def BuildQuery(BOM):
    # Build SQL Query to insert list data in SQL Table
    keys_to_sql = "SE_groupname, SE_order, SE_top, SE_pere, SE_piece,SE_epaisseur,SE_largeur,SE_longueur,SE_Materiel"
    SQLQuery = "INSERT INTO SE_BOM("+keys_to_sql+")"
    tot_rec = 0
    for item in BOM:
        vars_to_sql = []
        tot_rec+=1
        for i in item:
            value_type = type(i)
            if value_type == tuple:
                for n in i:
                    vars_to_sql.append("'"+str(n)+"'")
            if value_type <> tuple:
                vars_to_sql.append("'"+i+"'")
        final_to_sql = ','.join(vars_to_sql)
        if tot_rec == len(BOM):
            SQLQuery = SQLQuery + "SELECT " + final_to_sql +" GO"
        else:
            SQLQuery = SQLQuery + "SELECT " + final_to_sql + " UNION ALL "
    return SQLQuery
        
def CheckBody(ModelSet):
    #Function to determine if Object has a body
    try:
        for Model in ModelSet:
            if Model.Body is not None:
                BodyValid = True     
    except:
        BodyValid = False
    return BodyValid

def findNameByLevel(root,level):
    currentElement = root;
    for currentLevel in range(level):
        currentElement = currentElement.Parent
    # current element should now be the one we are searching for
    SE_DocName = currentElement.OccurrenceDocument.Properties.Item("ProjectInformation").Item("Document Number").Value
    if SE_DocName == "":
        return currentElement.name
    else :
        return SE_DocName

def assignGroupName(mur):
    SE_Commentaire = objDocument.Properties.Item("SummaryInformation").Item("Commentaires").Value
    listgroup = SE_Commentaire.splitlines()
    result = ""
    for idx, val in enumerate(listgroup):
        if mur in val:
            result =  ("MUR "+ val, str(idx+1))
    if result == '':
        result = ("MUR "+ mur, "9")
    #print result
    return result   
   
MP = ExtractSubOccurrence(objOccurenceSets)  

#BuildQuery(MP)
pySqlOperation.ZapSE_BOM()
pySqlOperation.executeSQLQuery(BuildQuery(MP))    
    
           
           

    
