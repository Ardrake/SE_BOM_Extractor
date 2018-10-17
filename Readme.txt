Solid Edge BOM Extractor

Utility that extracts the Bill of Material (BOM) from an open Solid Edge 3D assembly.

Requires 3 variable in each part (.par) SE_Longueur, SE_Epaisseur and SE_Largeur, these
are used to obtain the Width, height and length of the part.


BOM data saved to MSSQL data base using ODBC.
Reporting done with Crystal Reports XI on client side
_________________________________________________________________

The program requires Interop.SolidEdge.dll to be copied in its
folder. The DLL is part of Jason Newell's project and can be
downloaded from https://github.com/SolidEdgeCommunity/Interop.SolidEdge
Tested with IronPython 2.7.4 on .NET 4.0 and Solid Edge ST6.
_________________________________________________________________

Many thanks to 
http://www.arleedesign.com/Services/IronPythonForSETutorial.html


Autor : André Cooke
Email : andrecooke@hotmail.com