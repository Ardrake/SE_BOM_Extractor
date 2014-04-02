# -*- coding: utf-8 -*-
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
    #Extract part object, in index 0, 2nd level of assembly� Index 1, part name, Index 2,3 and 4 are Length, Witdh and heigth
    part = ["assembly", "part", "longueur", "largeur", "Epaisseur"]
    for i in range(objOccurenceSets.Count):
        objSubOccurrences = objOccurenceSets.Item(i + 1)
        if objSubOccurrences.Subassembly == False:
            VariableSet = objSubOccurrences.SubOccurrenceDocument.Variables
            SE_long = 0
            SE_larg = 0 
            SE_epai = 0
            for i in range(VariableSet.Count):
                Variables = VariableSet.Item(i + 1)
                if Variables.ExposeName == "SE_LONGUEUR":
                    SE_long = convert(Variables.Value)
                    #print "     - %s : %f " % (Variables.ExposeName, convert(Variables.Value))
                if Variables.ExposeName == "SE_LARGEUR":
                    SE_larg = convert(Variables.Value)
                if Variables.ExposeName == "SE_EPAISSEUR":
                    SE_epai = convert(Variables.Value)
            #This part needs to be optimised / refactored-------------**
            if level == 2:                    
                part = [objSubOccurrences.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            if level == 3:                    
                part = [objSubOccurrences.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            if level == 4:
                part = [objSubOccurrences.Parent.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            if level == 5:
                part = [objSubOccurrences.Parent.Parent.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            #----------------------------------------------------------**
            nomenclature.append(part)
            #print nomenclature
        if objSubOccurrences.SubOccurrences != None: 
            ExtractSubOccurrence(objSubOccurrences.SubOccurrences, level+1)
    return nomenclature

            
def convert(val,frac=0.125):
    #convert decimal to inches
    return  round((int((round(val,4) / 0.0254) / frac))*frac,4)  

def BuildQuery(BOM):
    # Build SQL Query to insert list data in SQL Table
    keys_to_sql = "SE_top, SE_pere, SE_piece,SE_epaisseur,SE_largeur,SE_longueur"
    SQLQuery = "INSERT INTO SE_BOM("+keys_to_sql+")"
    tot_rec = 0
    for item in BOM:
        vars_to_sql = []
        tot_rec+=1
        #print tot_rec
        for i in item:
            value_type = type(i)
            if value_type == tuple:
                for n in i:
                    #print (n , type(n))
                    vars_to_sql.append("'"+str(n)+"'")
            if value_type <> tuple:
                vars_to_sql.append("'"+i+"'")
        final_to_sql = ','.join(vars_to_sql)
        if tot_rec == len(BOM):
            SQLQuery = SQLQuery + "SELECT " + final_to_sql +" GO"
        else:
            SQLQuery = SQLQuery + "SELECT " + final_to_sql + " UNION ALL "
    return SQLQuery
        


MP = ExtractSubOccurrence(objOccurenceSets)    


print BuildQuery(MP)    
    
           
           
#print convert(0.3048+0.0015875)


assembly = []
for item in MP:
    assembly.append(item[0])
    #print item
#print ("Item count: ", len(MP))
#print ("Top level assemblies: ", set(assembly))     

#(SELECT %r UNION ALL)" % (keys_to_sql, tuple(vars_to_sql)) 
#print vars_to_sql
    
