# VBA Project: **VbaGit**
## VBA Module: **[cJobject](/libraries/cJobject.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:57 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cJobject

---
VBA Procedure: **backtrack**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get backtrack() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **backtrack**  
Type: **Set**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Set backtrack(back As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
back|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get self() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **isValid**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get isValid() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **setValid**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let setValid(good As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
good|Boolean|False||


---
VBA Procedure: **jString**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get jString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **fake**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get fake() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **childIndex**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get childIndex() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **childIndex**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let childIndex(p As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Long|False||


---
VBA Procedure: **isArrayRoot**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get isArrayRoot() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **isArrayMember**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get isArrayMember() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **isArrayRoot**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let isArrayRoot(p As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Boolean|False||


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Set**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Set parent(p As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **isRoot**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get isRoot() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **clearParent**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub clearParent()*  

**no arguments required for this procedure**


---
VBA Procedure: **root**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get root() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **key**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get key() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get value() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **setValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setValue(p As Variant) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|False||


---
VBA Procedure: **isVoid**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isVoid(Optional childName As String = vbNullString) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **isObjValue**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isObjValue(Optional childName As String = vbNullString) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **cValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function cValue(Optional childName As String = vbNullString) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **kids**  
Type: **Function**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function kids(Optional childName As String = vbNullString) As Collection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **toString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function toString(Optional childName As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **getObject**  
Type: **Get**  
Returns: **Object**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get getObject(Optional childName As String = vbNullString) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|True| vbNullString|


---
VBA Procedure: **value**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let value(p As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|False||


---
VBA Procedure: **children**  
Type: **Get**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get children() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **children**  
Type: **Set**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Set children(p As Collection)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Collection|False||


---
VBA Procedure: **hasChildren**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get hasChildren() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **deleteChild**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function deleteChild(childName As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
childName|String|False||


---
VBA Procedure: **valueIndex**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function valueIndex(v As Variant) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **toTreeView**  
Type: **Function**  
Returns: **Object**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function toTreeView(tr As Object, Optional bEnableCheckBoxes As Boolean = False) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tr|Object|False||
bEnableCheckBoxes|Boolean|True| False|


---
VBA Procedure: **treeViewPopulate**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function treeViewPopulate(tr As Object, cj As cJobject, Optional parent As cJobject = Nothing)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tr|Object|False||
cj|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
parent|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function init(p As cJobject, Optional k As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
k|String|True| cNull|
v|Variant|True| Empty|


---
VBA Procedure: **child**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function child(s As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function insert(Optional s As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|True| cNull|
v|Variant|True| Empty|


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function add(Optional k As String = cNull, Optional v As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
k|String|True| cNull|
v|Variant|True| Empty|


---
VBA Procedure: **addArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function addArray() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **childExists**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function childExists(s As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **unSplitToString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function unSplitToString(a As Variant, delim As String, Optional startAt As Long = -999, Optional howMany As Long = -999, Optional startAtEnd As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||
delim|String|False||
startAt|Long|True| -999|
howMany|Long|True| -999|
startAtEnd|Boolean|True| False|


---
VBA Procedure: **findInArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function findInArray(prop As String, val As Variant, Optional asKey As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
prop|String|False||
val|Variant|False||
asKey|Boolean|True| True|


---
VBA Procedure: **find**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function find(s As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **convertToArray**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function convertToArray() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **fullKey**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function fullKey(Optional includeRoot As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
includeRoot|Boolean|True| True|


---
VBA Procedure: **findByValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function findByValue(x As Variant) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||


---
VBA Procedure: **hasKey**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function hasKey() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **needsCurly**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function needsCurly() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **needsSquare**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function needsSquare() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **stringify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function stringify(Optional blf As Boolean) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
blf|Boolean|True||


---
VBA Procedure: **serialize**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function serialize(Optional blf As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
blf|Boolean|True| False|


---
VBA Procedure: **needsIndent**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get needsIndent() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **recurseSerialize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function recurseSerialize(job As cJobject, Optional soFar As cStringChunker = Nothing, Optional blf As Boolean = False) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
soFar|[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")|True| Nothing|
blf|Boolean|True| False|


---
VBA Procedure: **longestFullKey**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get longestFullKey() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **clone**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function clone() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **arrayAppend**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function arrayAppend(arrJob As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
arrJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **mergeArrayOfJobs**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function mergeArrayOfJobs() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **extractAndBlowArrayOfJobs**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function extractAndBlowArrayOfJobs() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **replaceBranch**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function replaceBranch(byThis As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
byThis|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **merge**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function merge(mergeThisIntoMe As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
mergeThisIntoMe|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **remove**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function remove() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **append**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function append(appendThisToMe As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
appendThisToMe|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **depth**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get depth(Optional l As Long = 0) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
l|Long|True| 0|


---
VBA Procedure: **clongestFullKey**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function clongestFullKey(job As cJobject, Optional soFar As Long = 0) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
soFar|Long|True| 0|


---
VBA Procedure: **formatData**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get formatData(Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
bDebug|Boolean|True| False|


---
VBA Procedure: **cformatdata**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function cformatdata(job As cJobject, Optional soFar As String = "", Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
soFar|String|True| ""|
bDebug|Boolean|True| False|


---
VBA Procedure: **itemFormat**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function itemFormat(jo As cJobject, Optional bDebug As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
bDebug|Boolean|True| False|


---
VBA Procedure: **jdebug**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub jdebug()*  

**no arguments required for this procedure**


---
VBA Procedure: **quote**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function quote(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **parse**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the branch header**  
Scope: **Public**  
Description: **this will take a JSON string and parse it into a cJobject branch under ME**  

*Public Function parse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True, Optional dealWithLikelyTimestamps As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||the string to be parsed
jtype|eDeserializeType|True||type of serialization
complain|Boolean|True| True|whether to complain in the event of an error
dealWithLikelyTimestamps|Boolean|True| True|whether to convert things that look like JS or unix timestamps to excel dates


---
VBA Procedure: **deSerialize**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the branch header**  
Scope: **Public**  
Description: **this will take a JSON string and deserialize into a cJobject branch under ME**  

*Public Function deSerialize(s As String, Optional jtype As eDeserializeType = eDeserializeNormal, Optional complain As Boolean = True, Optional dealWithLikelyTimestamps As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||the string to be parsed
jtype|eDeserializeType|True| eDeserializeNormal|type of serialization
complain|Boolean|True| True|whether to complain in the event of an error
dealWithLikelyTimestamps|Boolean|True| True|whether to convert things that look like JS or unix timestamps to excel dates


---
VBA Procedure: **sever**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function sever() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **noisyTrim**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function noisyTrim(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **nullItem**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function nullItem(job As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **dsLoop**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dsLoop(job As cJobject, complain As Boolean, dealWithLikelyTimestamps As Boolean) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
complain|Boolean|False||
dealWithLikelyTimestamps|Boolean|False||


---
VBA Procedure: **okWhat**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function okWhat(what As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
what|String|False||


---
VBA Procedure: **peekNextToken**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function peekNextToken() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **doNextToken**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function doNextToken() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dsProcess**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dsProcess(job As cJobject, complain As Boolean, dealWithLikelyTimestamps As Boolean) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
complain|Boolean|False||
dealWithLikelyTimestamps|Boolean|False||


---
VBA Procedure: **nOk**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function nOk() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getvItem**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getvItem(dealWithLikelyTimestamps As Boolean, Optional whichQ As String = "", Optional nextToken As String = vbNullString) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dealWithLikelyTimestamps|Boolean|False||
whichQ|String|True| ""|
nextToken|String|True| vbNullString|


---
VBA Procedure: **peek**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function peek() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **peekBehind**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function peekBehind() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **toNumber**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function toNumber(sIn As String, dealWithLikelyTimestamps As Boolean) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sIn|String|False||
dealWithLikelyTimestamps|Boolean|False||


---
VBA Procedure: **pointedAt**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function pointedAt(Optional pos As Long = 0, Optional sLen As Long = 1) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pos|Long|True| 0|
sLen|Long|True| 1|


---
VBA Procedure: **getQuotedItem**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getQuotedItem(Optional whichQ As String = "") As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
whichQ|String|True| ""|


---
VBA Procedure: **getNumericItem**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getNumericItem() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **isQuote**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isQuote(s As String, Optional whichQ As String = "") As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||
whichQ|String|True| ""|


---
VBA Procedure: **badJSON**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub badJSON(pWhatNext As String, Optional add As String = "", Optional complain As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pWhatNext|String|False||
add|String|True| ""|
complain|Boolean|True| True|


---
VBA Procedure: **ignoreNoise**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub ignoreNoise(Optional pos As Long = 0, Optional extraNoise As String = "")*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pos|Long|True| 0|
extraNoise|String|True| ""|


---
VBA Procedure: **isNoisy**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isNoisy(s As String, Optional extraNoise As String = "") As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||
extraNoise|String|True| ""|


---
VBA Procedure: **isEscape**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isEscape(s As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **isUnicode**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isUnicode(s As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **q**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function q() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **qs**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function qs() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **anyQ**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function anyQ() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **sortByValue**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function sortByValue() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **needSwap**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function needSwap(cc As cJobject, e As eSort) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cc|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
e|eSort|False||


---
VBA Procedure: **sortNorm**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function sortNorm(v As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**
