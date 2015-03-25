# VBA Project: **VbaGit**
## VBA Module: **[usefulSheetStuff](/libraries/usefulSheetStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 25/03/2015 18:59:47 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulSheetStuff

---
VBA Procedure: **firstCell**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function firstCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
inrange|Range|False||


---
VBA Procedure: **lastCell**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function lastCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
inrange|Range|False||


---
VBA Procedure: **isSheet**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Function isSheet(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
o|Object|False||


---
VBA Procedure: **findShape**  
Type: **Function**  
Returns: **Shape**  
Scope: **Public**  
Description: ****  

*Public Function findShape(sName As String, Optional ws As Worksheet = Nothing) As Shape*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
ws|Worksheet|True| Nothing|


---
VBA Procedure: **findRecurse**  
Type: **Function**  
Returns: **Shape**  
Scope: **Public**  
Description: ****  

*Public Function findRecurse(target As String, co As GroupShapes) As Shape*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
target|String|False||
co|GroupShapes|False||


---
VBA Procedure: **clearHyperLinks**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub clearHyperLinks(ws As Worksheet)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|False||


---
VBA Procedure: **sheetExists**  
Type: **Function**  
Returns: **Worksheet**  
Scope: **Public**  
Description: ****  

*Function sheetExists(sName As String, Optional complain As Boolean = True) As Worksheet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
complain|Boolean|True| True|


---
VBA Procedure: **wholeSheet**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function wholeSheet(wn As String) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wn|String|False||


---
VBA Procedure: **wholeWs**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function wholeWs(ws As Worksheet) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|False||


---
VBA Procedure: **wholeRange**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function wholeRange(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **cleanFind**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, Optional singlecell As Boolean = False) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||
r|Range|False||
complain|Boolean|True| False|
singlecell|Boolean|True| False|


---
VBA Procedure: **msglost**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Sub msglost(x As Variant, r As Range, Optional extra As String = "")*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||
r|Range|False||
extra|String|True| ""|


---
VBA Procedure: **SAd**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngIn|Range|False||
target|Range|True| Nothing|
singlecell|Boolean|True| False|
removeRowDollar|Boolean|True| False|
removeColDollar|Boolean|True| False|


---
VBA Procedure: **SAdOneRange**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngIn|Range|False||
target|Range|True| Nothing|
singlecell|Boolean|True| False|
removeRowDollar|Boolean|True| False|
removeColDollar|Boolean|True| False|


---
VBA Procedure: **AddressNoDollars**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Range|False||
doRow|Boolean|True| True|
doColumn|Boolean|True| True|


---
VBA Procedure: **isReallyEmpty**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Function isReallyEmpty(r As Range) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyRow**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function toEmptyRow(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyCol**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function toEmptyCol(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyBox**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Function toEmptyBox(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **getLikelyColumnRange**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Public Function getLikelyColumnRange(Optional ws As Worksheet = Nothing) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|True| Nothing|


---
VBA Procedure: **deleteAllShapes**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Sub deleteAllShapes(r As Range, startingwith As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||
startingwith|String|False||


---
VBA Procedure: **makearangeofShapes**  
Type: **Function**  
Returns: **ShapeRange**  
Scope: **Public**  
Description: ****  

*Function makearangeofShapes(r As Range, startingwith As String) As ShapeRange*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||
startingwith|String|False||


---
VBA Procedure: **nameExists**  
Type: **Function**  
Returns: **name**  
Scope: **Public**  
Description: ****  

*Public Function nameExists(s As String) As name*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **whereIsThis**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Public Function whereIsThis(r As Variant) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Variant|False||


---
VBA Procedure: **pivotCacheRefreshAll**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Sub pivotCacheRefreshAll()*  

**no arguments required for this procedure**


---
VBA Procedure: **getLatFromDistance**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  
Description: ****  

*Public Function getLatFromDistance(mLat As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
mLat|Double|False||
d|Double|False||
heading|Double|False||


---
VBA Procedure: **getLonFromDistance**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  
Description: ****  

*Public Function getLonFromDistance(mLat As Double, mLon As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
mLat|Double|False||
mLon|Double|False||
d|Double|False||
heading|Double|False||


---
VBA Procedure: **earthRadius**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  
Description: ****  

*Public Function earthRadius() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **toRadians**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function toRadians(deg)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
deg|Variant|False||


---
VBA Procedure: **fromRadians**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  
Description: ****  

*Public Function fromRadians(rad) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rad|Variant|False||


---
VBA Procedure: **min**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function min(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **max**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function max(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **toClipBoard**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function toClipBoard(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **importTabbed**  
Type: **Function**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Public Function importTabbed(fn As String, r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fn|String|False||
r|Range|False||
