# VBA Project: **VbaGit**
## VBA Module: **[cStringChunker](/libraries/cStringChunker.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 26/03/2015 09:26:24 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cStringChunker

---
VBA Procedure: **size**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get size() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **content**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get content() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getLeft**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get getLeft(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||


---
VBA Procedure: **getRight**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get getRight(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||


---
VBA Procedure: **getMid**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get getMid(startPos As Long, Optional howMany As Long = -1) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|False||
howMany|Long|True| -1|


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Property Get self() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **clear**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function clear() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **uri**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function uri(addstring As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addstring|String|False||


---
VBA Procedure: **toString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function toString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function add(addstring As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addstring|String|False||


---
VBA Procedure: **addLine**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function addLine(Optional addstring As String = "") As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addstring|String|True| ""|


---
VBA Procedure: **addLines**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function addLines(Optional number As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
number|Long|True| 1|


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function insert(Optional insertString As String = " ", Optional insertBefore As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
insertString|String|True| " "|
insertBefore|Long|True| 1|


---
VBA Procedure: **overWrite**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function overWrite(Optional overWriteString As String = " ", Optional overWriteAt As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
overWriteString|String|True| " "|
overWriteAt|Long|True| 1|


---
VBA Procedure: **shift**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function shift(Optional startPos As Long = 1, Optional howManyChars As Long = 0, Optional replaceWith As String = vbNullString) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|True| 1|
howManyChars|Long|True| 0|
replaceWith|String|True| vbNullString|


---
VBA Procedure: **chop**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function chop(Optional n As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Long|True| 1|


---
VBA Procedure: **chopSuperTrim**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function chopSuperTrim() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **chopIf**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function chopIf(t As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
t|String|False||


---
VBA Procedure: **chopWhile**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  
Description: ****  

*Public Function chopWhile(t As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
t|String|False||


---
VBA Procedure: **maxNumber**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  
Description: ****  

*Private Function maxNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||
b|Long|False||


---
VBA Procedure: **minNumber**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  
Description: ****  

*Private Function minNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||
b|Long|False||


---
VBA Procedure: **adjustSize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Private**  
Description: ****  

*Private Function adjustSize(needMore As Long) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
needMore|Long|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
