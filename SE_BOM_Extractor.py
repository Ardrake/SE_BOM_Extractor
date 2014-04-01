import clr
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import System.Runtime.InteropServices as SRI

#from System import Console

objOccurenceSets = None
objOccurrences = None
objSubOccurenceSets = None
objSubOccurrences = None
objSubSubOccurenceSets = None
objSubSubOccurrences = None
strFormatBoMlvl1 = "[%s]"
strFormatBomLvl2 = " - %s"
strFormatBomLvl3 = "   - %s"
strFormat2 = "%s = %s (%s)"


objApplication = SRI.Marshal.GetActiveObject("SolidEdge.Application")
objDocument = objApplication.ActiveDocument
objOccurenceSets = objDocument.Occurrences

# for i in range(objOccurenceSets.Count):
#     objOccurrences = objOccurenceSets.Item(i + 1)
#     print(strFormatBoMlvl1 % objOccurrences.Name)
#   
#     objSubOccurenceSets = objOccurrences.SubOccurrences
#        
#     for i in range(objSubOccurenceSets.Count):
#         objSubOccurrences = objSubOccurenceSets.Item(i + 1)     
#         print(strFormatBomLvl2 % objSubOccurrences.Name)
#   
#         objSubSubOccurenceSets = objSubOccurrences.SubOccurrences
#            
#         for i in range(objSubSubOccurenceSets.Count):
#             objSubSubOccurrences = objSubSubOccurenceSets.Item(i + 1)     
#             print(strFormatBomLvl3 % objSubSubOccurrences.Name)
#             VariableSet = objSubSubOccurrences.SubOccurrenceDocument.Variables
#             for i in range(VariableSet.Count):
#                 Variables = VariableSet.Item(i + 1)
#                 if Variables.ExposeName[:2] == "SE":
#                     print "     - %s : %f = %f" % (Variables.ExposeName, Variables.Value, Variables.Value/0.0254) 

            
def ExtractSubOccurrence(objOccurenceSets, level=1, nomenclature = []):
    part = ["assembly", "part" "longueur", "largeur", "Epaisseur"]
    for i in range(objOccurenceSets.Count):
        objSubOccurrences = objOccurenceSets.Item(i + 1)
        prefix = " " * level
        #print ("%s-Assmbly :%s Part :%s" % (prefix, objOccurenceSets.Parent.Name, objSubOccurrences.Name))

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
                    
            if level == 3:                    
                part = [objSubOccurrences.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            if level == 4:
                part = [objSubOccurrences.Parent.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]
            if level == 5:
                part = [objSubOccurrences.Parent.Parent.Parent.Parent.Name, objOccurenceSets.Parent.Name,  objSubOccurrences.Name, (SE_epai, SE_larg, SE_long)]

            nomenclature.append(part)
            #print nomenclature
        if objSubOccurrences.SubOccurrences != None: 
            ExtractSubOccurrence(objSubOccurrences.SubOccurrences, level+1)
    return nomenclature

            
def convert(val,frac=0.125):
    return  round((int((round(val,4) / 0.0254) / frac))*frac,4)  
           
           
#print convert(0.3048+0.0015875)
MP = ExtractSubOccurrence(objOccurenceSets)

assembly = []
for item in MP:
    assembly.append(item[0])
    print item
print ("Item count: ", len(MP))
print ("Top level assemblies: ", set(assembly))

    
 

    
