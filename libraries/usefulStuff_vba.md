# VBA Project: **VbaGit**
## VBA Module: **[usefulStuff](/libraries/usefulStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:57 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulStuff

---
VBA Procedure: **OpenUrl**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function OpenUrl(url) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|Variant|False||


---
VBA Procedure: **deleteAllFromCollection**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub deleteAllFromCollection(co As Collection)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
co|Collection|False||


---
VBA Procedure: **UTF16To8**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function UTF16To8(ByVal UTF16 As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **URLEncode**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function URLEncode( StringVal As String, Optional SpaceAsPlus As Boolean = False, Optional UTF8Encode As Boolean = True ) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
StringVal|String|False||
SpaceAsPlus|Boolean|True| False|
UTF8Encode|Boolean|True| True|


---
VBA Procedure: **cloneFormat**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub cloneFormat(b As Range, a As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
b|Range|False||
a|Range|False||


---
VBA Procedure: **compareAsKey**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function compareAsKey(a As Variant, b As Variant, Optional asKey As Boolean = True) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||
b|Variant|False||
asKey|Boolean|True| True|


---
VBA Procedure: **SortColl**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SortColl(ByRef coll As Collection, eorder As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|Collection|False||
eorder|Long|False||


---
VBA Procedure: **getHandle**  
Type: **Function**  
Returns: **Integer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getHandle(sName As String, Optional readOnly As Boolean = False) As Integer*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
readOnly|Boolean|True| False|


---
VBA Procedure: **afConcat**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function afConcat(arr() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
arr|Variant|False||


---
VBA Procedure: **quote**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function quote(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **q**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function q() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **qs**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function qs() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **bracket**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function bracket(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **list**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function list(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **qlist**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function qlist(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **diminishingReturn**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function diminishingReturn(val As Double, Optional s As Double = 10) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
val|Double|False||
s|Double|True| 10|


---
VBA Procedure: **superTrim**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function superTrim(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **makeKey**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function makeKey(v As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **Base64Encode**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Base64Encode(sText)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sText|Variant|False||


---
VBA Procedure: **Stream_StringToBinary**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Stream_StringToBinary(Text)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Text|Variant|False||


---
VBA Procedure: **Stream_BinaryToString**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Stream_BinaryToString(Binary)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Binary|Variant|False||


---
VBA Procedure: **Base64Decode**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Base64Decode(ByVal base64String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **openNewHtml**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function openNewHtml(sName As String, sContent As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
sContent|String|False||


---
VBA Procedure: **readFromFile**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function readFromFile(sName As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||


---
VBA Procedure: **arrayLength**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function arrayLength(a) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **getControlValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getControlValue(ctl As Object) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ctl|Object|False||


---
VBA Procedure: **setControlValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setControlValue(ctl As Object, v As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ctl|Object|False||
v|Variant|False||


---
VBA Procedure: **isinCollection**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isinCollection(vCollect As Variant, sid As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vCollect|Variant|False||
sid|Variant|False||


---
VBA Procedure: **dimensionCount**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function dimensionCount(a As Variant) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **encloseTag**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function encloseTag(tag As String, Optional newLine As Boolean = True, Optional tClass As String = vbNullString, Optional args As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tag|String|False||
newLine|Boolean|True| True|
tClass|String|True| vbNullString|
args|Variant|True||


---
VBA Procedure: **scrollHack**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function scrollHack() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **escapeify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function escapeify(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **unEscapify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function unEscapify(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **basicStyle**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function basicStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **tableStyle**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tableStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **is64BitExcel**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function is64BitExcel() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **includeJQuery**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function includeJQuery() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **includeGoogleCallBack**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function includeGoogleCallBack(c As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
c|String|False||


---
VBA Procedure: **jScriptTag**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function jScriptTag(Optional src As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
src|String|True||


---
VBA Procedure: **jDivAtMouse**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function jDivAtMouse()*  

**no arguments required for this procedure**


---
VBA Procedure: **biasedRandom**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function biasedRandom(possibilities, weights) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
possibilities|Variant|False||
weights|Variant|False||


---
VBA Procedure: **sleep**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub sleep(seconds As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
seconds|Long|False||


---
VBA Procedure: **getDateFromTimestamp**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getDateFromTimestamp(s As String) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **dateFromUnix**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function dateFromUnix(s As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|Variant|False||


---
VBA Procedure: **isSomething**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isSomething(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
o|Object|False||


---
VBA Procedure: **tinyTime**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tinyTime() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **applyDefaults**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function applyDefaults(value As Variant, defaultValue As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
value|Variant|False||
defaultValue|Variant|False||


---
VBA Procedure: **isUndefined**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function isUndefined(value As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
value|Variant|False||


---
VBA Procedure: **conditionalAssignment**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function conditionalAssignment(condition As Boolean, a As Variant, b As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
condition|Boolean|False||
a|Variant|False||
b|Variant|False||


---
VBA Procedure: **assignHelper**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function assignHelper(a As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **getTimestampFromDate**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getTimestampFromDate(Optional dt As Date = 0) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dt|Date|True| 0|


---
VBA Procedure: **checkOrCreateFolder**  
Type: **Function**  
Returns: **Object**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function checkOrCreateFolder(path As String, Optional optCreate As Boolean = True) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
path|String|False||
optCreate|Boolean|True| True|


---
VBA Procedure: **recurseCreateFolder**  
Type: **Function**  
Returns: **Object**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function recurseCreateFolder(fso As Object, cleanPath As String) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fso|Object|False||
cleanPath|String|False||


---
VBA Procedure: **writeToFolderFile**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function writeToFolderFile(folderName As String, fileName As String, content As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
folderName|String|False||
fileName|String|False||
content|String|False||


---
VBA Procedure: **getAllSubFolderPaths**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getAllSubFolderPaths(folderName As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
folderName|String|False||


---
VBA Procedure: **readFromFolderFile**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function readFromFolderFile(folderName As String, fileName As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
folderName|String|False||
fileName|String|False||


---
VBA Procedure: **fileExists**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function fileExists(path As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
path|String|False||


---
VBA Procedure: **concatFolderName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function concatFolderName(folderName As String, fileName As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
folderName|String|False||
fileName|String|False||
