# VBA Project: **VbaGit**
## VBA Module: **[cJobject](/libraries/cJobject.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 23/03/2015 10:33:26 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cJobject

---
VBA Procedure: **backtrack**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Property Get backtrack() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **backtrack**  
Type: **Set**  
Returns: **void**  
Scope: **Public**  

*Public Property Set backtrack(back As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
back|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Property Get self() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **isValid**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get isValid() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **setValid**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let setValid(good As Boolean)*  

*name*|*type*|*optional*|*default*
---|---|---|---
good|Boolean|False|


---
VBA Procedure: **jString**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get jString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **fake**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get fake() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **childIndex**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get childIndex() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **childIndex**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let childIndex(p As Long)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Long|False|


---
VBA Procedure: **isArrayRoot**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get isArrayRoot() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **isArrayMember**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get isArrayMember() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **isArrayRoot**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let isArrayRoot(p As Boolean)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Boolean|False|


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Property Get parent() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Set**  
Returns: **void**  
Scope: **Public**  

*Public Property Set parent(p As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **isRoot**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get isRoot() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **clearParent**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub clearParent()*  

**no arguments required for this procedure**


---
VBA Procedure: **root**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Property Get root() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **key**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get key() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  

*Public Property Get value() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **setValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function setValue(p As Variant) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Variant|False|


---
VBA Procedure: **isVoid**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function isVoid(Optional childName As String = vbNullString) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **isObjValue**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function isObjValue(Optional childName As String = vbNullString) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **cValue**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function cValue(Optional childName As String = vbNullString) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **kids**  
Type: **Function**  
Returns: **Collection**  
Scope: **Public**  

*Public Function kids(Optional childName As String = vbNullString) As Collection*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **toString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function toString(Optional childName As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **getObject**  
Type: **Get**  
Returns: **Object**  
Scope: **Public**  

*Public Property Get getObject(Optional childName As String = vbNullString) As Object*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|True| vbNullString


---
VBA Procedure: **value**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let value(p As Variant)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Variant|False|


---
VBA Procedure: **children**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  

*Public Property Get children() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **children**  
Type: **Set**  
Returns: **void**  
Scope: **Public**  

*Public Property Set children(p As Collection)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Collection|False|


---
VBA Procedure: **hasChildren**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get hasChildren() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **deleteChild**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function deleteChild(childName As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
childName|String|False|


---
VBA Procedure: **valueIndex**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  

*Public Function valueIndex(v As Variant) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
v|Variant|False|


---
VBA Procedure: **toTreeView**  
Type: **Function**  
Returns: **Object**  
Scope: **Public**  

*Public Function toTreeView(tr As Object, Optional bEnableCheckBoxes As Boolean = False) As Object*  

*name*|*type*|*optional*|*default*
---|---|---|---
tr|Object|False|
bEnableCheckBoxes|Boolean|True| False


---
VBA Procedure: **treeViewPopulate**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function treeViewPopulate(tr As Object, cj As cJobject, Optional parent As cJobject = Nothing)*  

*name*|*type*|*optional*|*default*
---|---|---|---
tr|Object|False|
cj|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
parent|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function init(p As cJobject, Optional k As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
k|String|True| cNull
v|Variant|True| Empty


---
VBA Procedure: **child**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function child(s As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function insert(Optional s As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|True| cNull
v|Variant|True| Empty


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function add(Optional k As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
k|String|True| cNull
v|Variant|True| Empty


---
VBA Procedure: **addArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function addArray() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **childExists**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function childExists(s As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **unSplitToString**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function unSplitToString(a As Variant, delim As String, Optional startAt As Long = -999, Optional howMany As Long = -999, Optional startAtEnd As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Variant|False|
delim|String|False|
startAt|Long|True| -999
howMany|Long|True| -999
startAtEnd|Boolean|True| False


---
VBA Procedure: **findInArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function findInArray(prop As String, val As Variant, Optional asKey As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
prop|String|False|
val|Variant|False|
asKey|Boolean|True| True


---
VBA Procedure: **find**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function find(s As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **convertToArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function convertToArray() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **fullKey**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function fullKey(Optional includeRoot As Boolean = True) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
includeRoot|Boolean|True| True


---
VBA Procedure: **findByValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function findByValue(x As Variant) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
x|Variant|False|


---
VBA Procedure: **hasKey**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function hasKey() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **needsCurly**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function needsCurly() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **needsSquare**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function needsSquare() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **stringify**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function stringify(Optional blf As Boolean) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
blf|Boolean|True|


---
VBA Procedure: **serialize**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function serialize(Optional blf As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
blf|Boolean|True| False


---
VBA Procedure: **needsIndent**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get needsIndent() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **recurseSerialize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function recurseSerialize(job As cJobject, Optional soFar As cStringChunker = Nothing, Optional blf As Boolean = False) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
soFar|[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")|True| Nothing
blf|Boolean|True| False


---
VBA Procedure: **longestFullKey**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get longestFullKey() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **clone**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function clone() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **arrayAppend**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function arrayAppend(arrJob As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
arrJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **mergeArrayOfJobs**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function mergeArrayOfJobs() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **replaceBranch**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function replaceBranch(byThis As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
byThis|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **merge**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function merge(mergeThisIntoMe As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
mergeThisIntoMe|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **remove**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function remove() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **append**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function append(appendThisToMe As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
appendThisToMe|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **depth**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get depth(Optional l As Long = 0) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
l|Long|True| 0


---
VBA Procedure: **clongestFullKey**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  

*Private Function clongestFullKey(job As cJobject, Optional soFar As Long = 0) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
soFar|Long|True| 0


---
VBA Procedure: **formatData**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get formatData(Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
bDebug|Boolean|True| False


---
VBA Procedure: **cformatdata**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function cformatdata(job As cJobject, Optional soFar As String = "", Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
soFar|String|True| ""
bDebug|Boolean|True| False


---
VBA Procedure: **itemFormat**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function itemFormat(jo As cJobject, Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
jo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
bDebug|Boolean|True| False


---
VBA Procedure: **jdebug**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub jdebug()*  

**no arguments required for this procedure**


---
VBA Procedure: **quote**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function quote(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **parse**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function parse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|
jtype|eDeserializeType|True|
complain|Boolean|True| True


---
VBA Procedure: **deSerialize**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function deSerialize(s As String, Optional jtype As eDeserializeType = eDeserializeNormal, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|
jtype|eDeserializeType|True| eDeserializeNormal
complain|Boolean|True| True


---
VBA Procedure: **sever**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function sever() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **noisyTrim**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function noisyTrim(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **nullItem**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function nullItem(job As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **dsLoop**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function dsLoop(job As cJobject, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
complain|Boolean|True| True


---
VBA Procedure: **okWhat**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function okWhat(what As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
what|String|False|


---
VBA Procedure: **peekNextToken**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function peekNextToken() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **doNextToken**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function doNextToken() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dsProcess**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function dsProcess(job As cJobject, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
complain|Boolean|True| True


---
VBA Procedure: **nOk**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function nOk() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getvItem**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function getvItem(Optional whichQ As String = "", Optional nextToken = vbNullString) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
whichQ|String|True| ""
nextToken|Variant|True|


---
VBA Procedure: **peek**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function peek() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **peekBehind**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function peekBehind() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **toNumber**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function toNumber(sIn As String) As Variant*  

*name*|*type*|*optional*|*default*
---|---|---|---
sIn|String|False|


---
VBA Procedure: **pointedAt**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function pointedAt(Optional pos As Long = 0, Optional sLen As Long = 1) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
pos|Long|True| 0
sLen|Long|True| 1


---
VBA Procedure: **getQuotedItem**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function getQuotedItem(Optional whichQ As String = "") As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
whichQ|String|True| ""


---
VBA Procedure: **getNumericItem**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function getNumericItem() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **isQuote**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isQuote(s As String, Optional whichQ As String = "") As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|
whichQ|String|True| ""


---
VBA Procedure: **badJSON**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub badJSON(pWhatNext As String, Optional add As String = "", Optional complain As Boolean = True)*  

*name*|*type*|*optional*|*default*
---|---|---|---
pWhatNext|String|False|
add|String|True| ""
complain|Boolean|True| True


---
VBA Procedure: **ignoreNoise**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub ignoreNoise(Optional pos As Long = 0, Optional extraNoise As String = "")*  

*name*|*type*|*optional*|*default*
---|---|---|---
pos|Long|True| 0
extraNoise|String|True| ""


---
VBA Procedure: **isNoisy**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isNoisy(s As String, Optional extraNoise As String = "") As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|
extraNoise|String|True| ""


---
VBA Procedure: **isEscape**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isEscape(s As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **isUnicode**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isUnicode(s As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **q**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function q() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **qs**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function qs() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **anyQ**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function anyQ() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **sortByValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function sortByValue() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **needSwap**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function needSwap(cc As cJobject, e As eSort) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
cc|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
e|eSort|False|


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub tearDown()*  

**no arguments required for this procedure**
