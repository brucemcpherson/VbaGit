# VBA Project: **VbaGit**
## VBA Module: **[usefulStuff](/libraries/usefulStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 23/03/2015 10:33:25 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulStuff

---
VBA Procedure: **OpenUrl**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function OpenUrl(url) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
url|Variant|False|


---
VBA Procedure: **deleteAllFromCollection**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Sub deleteAllFromCollection(co As Collection)*  

*name*|*type*|*optional*|*default*
---|---|---|---
co|Collection|False|


---
VBA Procedure: **UTF16To8**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function UTF16To8(ByVal UTF16 As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
ByVal|String|False|


---
VBA Procedure: **URLEncode**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function URLEncode( StringVal As String, Optional SpaceAsPlus As Boolean = False, Optional UTF8Encode As Boolean = True ) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
StringVal|String|False|
SpaceAsPlus|Boolean|True| False
UTF8Encode|Boolean|True| True


---
VBA Procedure: **cloneFormat**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub cloneFormat(b As Range, a As Range)*  

*name*|*type*|*optional*|*default*
---|---|---|---
b|Range|False|
a|Range|False|


---
VBA Procedure: **compareAsKey**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function compareAsKey(a As Variant, b As Variant, Optional asKey As Boolean = True) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Variant|False|
b|Variant|False|
asKey|Boolean|True| True


---
VBA Procedure: **SortColl**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  

*Function SortColl(ByRef coll As Collection, eorder As Long) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
ByRef|Collection|False|
eorder|Long|False|


---
VBA Procedure: **getHandle**  
Type: **Function**  
Returns: **Integer**  
Scope: **Public**  

*Public Function getHandle(sName As String, Optional readOnly As Boolean = False) As Integer*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
readOnly|Boolean|True| False


---
VBA Procedure: **afConcat**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function afConcat(arr() As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
arr|Variant|False|


---
VBA Procedure: **quote**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function quote(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **q**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function q() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **qs**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function qs() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **bracket**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function bracket(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **list**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function list(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
ParamArray|Variant|False|


---
VBA Procedure: **qlist**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function qlist(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
ParamArray|Variant|False|


---
VBA Procedure: **diminishingReturn**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function diminishingReturn(val As Double, Optional s As Double = 10) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
val|Double|False|
s|Double|True| 10


---
VBA Procedure: **makeKey**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function makeKey(v As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
v|Variant|False|


---
VBA Procedure: **Base64Encode**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function Base64Encode(sText)*  

*name*|*type*|*optional*|*default*
---|---|---|---
sText|Variant|False|


---
VBA Procedure: **Stream_StringToBinary**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function Stream_StringToBinary(Text)*  

*name*|*type*|*optional*|*default*
---|---|---|---
Text|Variant|False|


---
VBA Procedure: **Stream_BinaryToString**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function Stream_BinaryToString(Binary)*  

*name*|*type*|*optional*|*default*
---|---|---|---
Binary|Variant|False|


---
VBA Procedure: **Base64Decode**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function Base64Decode(ByVal base64String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
ByVal|Variant|False|


---
VBA Procedure: **openNewHtml**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function openNewHtml(sName As String, sContent As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
sContent|String|False|


---
VBA Procedure: **readFromFile**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function readFromFile(sName As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|


---
VBA Procedure: **arrayLength**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  

*Public Function arrayLength(a) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Variant|False|


---
VBA Procedure: **getControlValue**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function getControlValue(ctl As Object) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
ctl|Object|False|


---
VBA Procedure: **setControlValue**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function setControlValue(ctl As Object, v As Variant) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
ctl|Object|False|
v|Variant|False|


---
VBA Procedure: **isinCollection**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function isinCollection(vCollect As Variant, sid As Variant) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
vCollect|Variant|False|
sid|Variant|False|


---
VBA Procedure: **dimensionCount**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  

*Public Function dimensionCount(a As Variant) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Variant|False|


---
VBA Procedure: **encloseTag**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function encloseTag(tag As String, Optional newLine As Boolean = True, Optional tClass As String = vbNullString, Optional args As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
tag|String|False|
newLine|Boolean|True| True
tClass|String|True| vbNullString
args|Variant|True|


---
VBA Procedure: **scrollHack**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function scrollHack() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **escapeify**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function escapeify(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **unEscapify**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function unEscapify(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **basicStyle**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function basicStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **tableStyle**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function tableStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **is64BitExcel**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function is64BitExcel() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **includeJQuery**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function includeJQuery() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **includeGoogleCallBack**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function includeGoogleCallBack(c As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
c|String|False|


---
VBA Procedure: **jScriptTag**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function jScriptTag(Optional src As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
src|String|True|


---
VBA Procedure: **jDivAtMouse**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function jDivAtMouse()*  

**no arguments required for this procedure**


---
VBA Procedure: **biasedRandom**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function biasedRandom(possibilities, weights) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
possibilities|Variant|False|
weights|Variant|False|


---
VBA Procedure: **sleep**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub sleep(seconds As Long)*  

*name*|*type*|*optional*|*default*
---|---|---|---
seconds|Long|False|


---
VBA Procedure: **getDateFromTimestamp**  
Type: **Function**  
Returns: **Date**  
Scope: **Public**  

*Public Function getDateFromTimestamp(s As String) As Date*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **dateFromUnix**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function dateFromUnix(s As Variant) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|Variant|False|


---
VBA Procedure: **isSomething**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function isSomething(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
o|Object|False|


---
VBA Procedure: **tinyTime**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function tinyTime() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **applyDefaults**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function applyDefaults(value As Variant, defaultValue As Variant) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
value|Variant|False|
defaultValue|Variant|False|


---
VBA Procedure: **isUndefined**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Function isUndefined(value As Variant) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
value|Variant|False|


---
VBA Procedure: **conditionalAssignment**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Function conditionalAssignment(condition As Boolean, a As Variant, b As Variant) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
condition|Boolean|False|
a|Variant|False|
b|Variant|False|


---
VBA Procedure: **assignHelper**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function assignHelper(a As Variant) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Variant|False|


---
VBA Procedure: **getTimestampFromDate**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function getTimestampFromDate(Optional dt As Date = 0) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
dt|Date|True| 0


---
VBA Procedure: **checkOrCreateFolder**  
Type: **Function**  
Returns: **Object**  
Scope: **Public**  

*Public Function checkOrCreateFolder(path As String, Optional optCreate As Boolean = True) As Object*  

*name*|*type*|*optional*|*default*
---|---|---|---
path|String|False|
optCreate|Boolean|True| True


---
VBA Procedure: **recurseCreateFolder**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  

*Private Function recurseCreateFolder(fso As Object, cleanPath As String) As Object*  

*name*|*type*|*optional*|*default*
---|---|---|---
fso|Object|False|
cleanPath|String|False|


---
VBA Procedure: **writeToFolderFile**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function writeToFolderFile(folderName As String, fileName As String, content As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
folderName|String|False|
fileName|String|False|
content|String|False|


---
VBA Procedure: **getAllSubFolderPaths**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function getAllSubFolderPaths(folderName As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
folderName|String|False|


---
VBA Procedure: **readFromFolderFile**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function readFromFolderFile(folderName As String, fileName As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
folderName|String|False|
fileName|String|False|


---
VBA Procedure: **fileExists**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function fileExists(path As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
path|String|False|


---
VBA Procedure: **concatFolderName**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function concatFolderName(folderName As String, fileName As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
folderName|String|False|
fileName|String|False|
