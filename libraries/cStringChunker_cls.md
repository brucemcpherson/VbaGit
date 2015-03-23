# VBA Project: **VbaGit**
## VBA Module: **[cStringChunker](/libraries/cStringChunker.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 23/03/2015 10:33:26 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cStringChunker

---
VBA Procedure: **size**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get size() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **content**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get content() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getLeft**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get getLeft(howMany As Long) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
howMany|Long|False|


---
VBA Procedure: **getRight**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get getRight(howMany As Long) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
howMany|Long|False|


---
VBA Procedure: **getMid**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get getMid(startPos As Long, Optional howMany As Long = -1) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
startPos|Long|False|
howMany|Long|True| -1


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Property Get self() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **clear**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function clear() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **uri**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function uri(addstring As String) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
addstring|String|False|


---
VBA Procedure: **toString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function toString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function add(addstring As String) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
addstring|String|False|


---
VBA Procedure: **addLine**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function addLine(Optional addstring As String = "") As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
addstring|String|True| ""


---
VBA Procedure: **addLines**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function addLines(Optional number As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
number|Long|True| 1


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function insert(Optional insertString As String = " ", Optional insertBefore As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
insertString|String|True| " "
insertBefore|Long|True| 1


---
VBA Procedure: **overWrite**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function overWrite(Optional overWriteString As String = " ", Optional overWriteAt As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
overWriteString|String|True| " "
overWriteAt|Long|True| 1


---
VBA Procedure: **shift**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function shift(Optional startPos As Long = 1, Optional howManyChars As Long = 0, Optional replaceWith As String = vbNullString) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
startPos|Long|True| 1
howManyChars|Long|True| 0
replaceWith|String|True| vbNullString


---
VBA Procedure: **chop**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function chop(Optional n As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
n|Long|True| 1


---
VBA Procedure: **chopIf**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function chopIf(t As String) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
t|String|False|


---
VBA Procedure: **chopWhile**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Public**  

*Public Function chopWhile(t As String) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
t|String|False|


---
VBA Procedure: **maxNumber**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  

*Private Function maxNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Long|False|
b|Long|False|


---
VBA Procedure: **minNumber**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  

*Private Function minNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*
---|---|---|---
a|Long|False|
b|Long|False|


---
VBA Procedure: **adjustSize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Scope: **Private**  

*Private Function adjustSize(needMore As Long) As cStringChunker*  

*name*|*type*|*optional*|*default*
---|---|---|---
needMore|Long|False|


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
