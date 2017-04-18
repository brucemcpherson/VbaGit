# VBA Project: **VbaGit**
## VBA Module: **[cStringChunker](/libraries/cStringChunker.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cStringChunker

---
VBA Procedure: **size**  
Type: **Get**  
Returns: **Long**  
Return description: **the length of the current string**  
Scope: **Public**  
Description: **get the length of the current string**  

*Public Property Get size() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **content**  
Type: **Get**  
Returns: **String**  
Return description: **the current string**  
Scope: **Public**  
Description: **get the content of the current string**  

*Public Property Get content() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getLeft**  
Type: **Get**  
Returns: **String**  
Return description: **the current string**  
Scope: **Public**  
Description: **extract the leftmost portion of a string**  

*Public Property Get getLeft(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||the length of the string to return


---
VBA Procedure: **getRight**  
Type: **Get**  
Returns: **String**  
Return description: **the current string**  
Scope: **Public**  
Description: **extract the rightmost portion of a string**  

*Public Property Get getRight(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||the length of the string to return


---
VBA Procedure: **getMid**  
Type: **Get**  
Returns: **String**  
Return description: **the current string**  
Scope: **Public**  
Description: **extract a portion of a string**  

*Public Property Get getMid(startPos As Long, Optional howMany As Long = -1) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|False||the (1 base) start position to extraction from
howMany|Long|True| -1|the length of the string to return


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **a self reference (useful for inside with..)**  

*Public Property Get self() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **clear**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **clear a chunker (set string to null)**  

*Public Function clear() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **uri**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **encode a uri and add**  

*Public Function uri(url As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||the url to add


---
VBA Procedure: **toString**  
Type: **Function**  
Returns: **String**  
Return description: **the string**  
Scope: **Public**  
Description: **return the string**  

*Public Function toString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **add a string to the chunker**  

*Public Function add(addString As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addString|String|False||the string to add


---
VBA Procedure: **addLine**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **add a string to the chunker, followed by a new line**  

*Public Function addLine(Optional addString As String = "") As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addString|String|True| ""|the string to add


---
VBA Procedure: **addLines**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **add a number of new lines**  

*Public Function addLines(Optional howMany As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|True| 1|number of lines to add


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **insert a string at a particular position**  

*Public Function insert(Optional insertString As String = " ", Optional insertBefore As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
insertString|String|True| " "|the string to insert (default 1 space)
insertBefore|Long|True| 1|the position(base 1) before which to insert


---
VBA Procedure: **overWrite**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **overwrite a string at a particular position**  

*Public Function overWrite(Optional overWriteString As String = " ", Optional overWriteAt As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
overWriteString|String|True| " "|the string to insert (default 1 space)
overWriteAt|Long|True| 1|the position(base 1) to start overwriting at


---
VBA Procedure: **shift**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **shift the contents inside the chunker space**  

*Public Function shift(Optional startPos As Long = 1, Optional howManyChars As Long = 0, Optional replaceWith As String = vbNullString) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|True| 1|the start position (base 1) of the string to shift
howManyChars|Long|True| 0|the length of the string to shift (-ve means left, +ve right)_
replaceWith|String|True| vbNullString|what to replace the moved contents with


---
VBA Procedure: **chop**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **chop characters from the end of the content**  

*Public Function chop(Optional howMany As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|True| 1|number of characters to chop


---
VBA Procedure: **chopSuperTrim**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **trim \s type chars from beginning and end**  

*Public Function chopSuperTrim(Optional fromBeginning As Boolean = True, Optional fromEnd As Boolean = True) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fromBeginning|Boolean|True| True|trim the beginnging of the content
fromEnd|Boolean|True| True|trim the end of the content


---
VBA Procedure: **chopIf**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **trim chars from end of content**  

*Public Function chopIf(chopString As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
chopString|String|False||trim the beginnging of the content


---
VBA Procedure: **chopWhile**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Public**  
Description: **trim chars from end of content and keep doing it while it matches**  

*Public Function chopWhile(chopString As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
chopString|String|False||trim the beginnging of the content


---
VBA Procedure: **maxNumber**  
Type: **Function**  
Returns: **Long**  
Return description: **the bigger of a and b**  
Scope: **Private**  
Description: **local max function**  

*Private Function maxNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||first number to compare
b|Long|False||second number to compare


---
VBA Procedure: **minNumber**  
Type: **Function**  
Returns: **Long**  
Return description: **the smaller of a and b**  
Scope: **Private**  
Description: **local min function**  

*Private Function minNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||first number to compare
b|Long|False||second number to compare


---
VBA Procedure: **adjustSize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: **self**  
Scope: **Private**  
Description: **adjust the underlying chunker buffer size if its needed**  

*Private Function adjustSize(needMore As Long) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
needMore|Long|False||how many chars we want space for


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: **intialize some starting buffer**  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
