# VBA Project: **VbaGit**
## VBA Module: **[usefulSheetStuff](/libraries/usefulSheetStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 23/03/2015 10:33:27 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulSheetStuff

---
VBA Procedure: **firstCell**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function firstCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
inrange|Range|False|


---
VBA Procedure: **lastCell**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function lastCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
inrange|Range|False|


---
VBA Procedure: **isSheet**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Function isSheet(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
o|Object|False|


---
VBA Procedure: **findShape**  
Type: **Function**  
Returns: **Shape**  
Scope: **Public**  

*Public Function findShape(sName As String, Optional ws As Worksheet = Nothing) As Shape*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
ws|Worksheet|True| Nothing


---
VBA Procedure: **findRecurse**  
Type: **Function**  
Returns: **Shape**  
Scope: **Public**  

*Public Function findRecurse(target As String, co As GroupShapes) As Shape*  

*name*|*type*|*optional*|*default*
---|---|---|---
target|String|False|
co|GroupShapes|False|


---
VBA Procedure: **clearHyperLinks**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub clearHyperLinks(ws As Worksheet)*  

*name*|*type*|*optional*|*default*
---|---|---|---
ws|Worksheet|False|


---
VBA Procedure: **sheetExists**  
Type: **Function**  
Returns: **Worksheet**  
Scope: **Public**  

*Function sheetExists(sName As String, Optional complain As Boolean = True) As Worksheet*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
complain|Boolean|True| True


---
VBA Procedure: **wholeSheet**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function wholeSheet(wn As String) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
wn|String|False|


---
VBA Procedure: **wholeWs**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function wholeWs(ws As Worksheet) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
ws|Worksheet|False|


---
VBA Procedure: **wholeRange**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function wholeRange(r As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|


---
VBA Procedure: **cleanFind**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, Optional singlecell As Boolean = False) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
x|Variant|False|
r|Range|False|
complain|Boolean|True| False
singlecell|Boolean|True| False


---
VBA Procedure: **msglost**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Sub msglost(x As Variant, r As Range, Optional extra As String = "")*  

*name*|*type*|*optional*|*default*
---|---|---|---
x|Variant|False|
r|Range|False|
extra|String|True| ""


---
VBA Procedure: **SAd**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
rngIn|Range|False|
target|Range|True| Nothing
singlecell|Boolean|True| False
removeRowDollar|Boolean|True| False
removeColDollar|Boolean|True| False


---
VBA Procedure: **SAdOneRange**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
rngIn|Range|False|
target|Range|True| Nothing
singlecell|Boolean|True| False
removeRowDollar|Boolean|True| False
removeColDollar|Boolean|True| False


---
VBA Procedure: **AddressNoDollars**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Range|False|
doRow|Boolean|True| True
doColumn|Boolean|True| True


---
VBA Procedure: **isReallyEmpty**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Function isReallyEmpty(r As Range) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|


---
VBA Procedure: **toEmptyRow**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function toEmptyRow(r As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|


---
VBA Procedure: **toEmptyCol**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function toEmptyCol(r As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|


---
VBA Procedure: **toEmptyBox**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Function toEmptyBox(r As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|


---
VBA Procedure: **getLikelyColumnRange**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Public Function getLikelyColumnRange(Optional ws As Worksheet = Nothing) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
ws|Worksheet|True| Nothing


---
VBA Procedure: **deleteAllShapes**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Sub deleteAllShapes(r As Range, startingwith As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|
startingwith|String|False|


---
VBA Procedure: **makearangeofShapes**  
Type: **Function**  
Returns: **ShapeRange**  
Scope: **Public**  

*Function makearangeofShapes(r As Range, startingwith As String) As ShapeRange*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Range|False|
startingwith|String|False|


---
VBA Procedure: **nameExists**  
Type: **Function**  
Returns: **name**  
Scope: **Public**  

*Public Function nameExists(s As String) As name*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **whereIsThis**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Public Function whereIsThis(r As Variant) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Variant|False|


---
VBA Procedure: **pivotCacheRefreshAll**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Sub pivotCacheRefreshAll()*  

**no arguments required for this procedure**


---
VBA Procedure: **getLatFromDistance**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function getLatFromDistance(mLat As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
mLat|Double|False|
d|Double|False|
heading|Double|False|


---
VBA Procedure: **getLonFromDistance**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function getLonFromDistance(mLat As Double, mLon As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
mLat|Double|False|
mLon|Double|False|
d|Double|False|
heading|Double|False|


---
VBA Procedure: **earthRadius**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function earthRadius() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **toRadians**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function toRadians(deg)*  

*name*|*type*|*optional*|*default*
---|---|---|---
deg|Variant|False|


---
VBA Procedure: **fromRadians**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function fromRadians(rad) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
rad|Variant|False|


---
VBA Procedure: **min**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function min(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*
---|---|---|---
ParamArray|Variant|False|


---
VBA Procedure: **max**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function max(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*
---|---|---|---
ParamArray|Variant|False|


---
VBA Procedure: **toClipBoard**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function toClipBoard(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **importTabbed**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  

*Public Function importTabbed(fn As String, r As Range) As Range*  

*name*|*type*|*optional*|*default*
---|---|---|---
fn|String|False|
r|Range|False|
