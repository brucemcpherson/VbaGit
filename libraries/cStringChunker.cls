'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:48:00 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/5102369/raw/cStringChunker.cls

Option Explicit
' v1.09  5102369
Private pContent As String
Private pSize As Long
' minimum amount to increment by each time
Const defaultChunkSize = 64
'/**
' * stringChunker class for VBA because string concat takes ages
' */

'/**
' * get the length of the current string
' *@return {long} the length of the current string
' */
Public Property Get size() As Long
    ' this is how much content is real
    size = pSize
End Property
'/**
' * get the content of the current string
' *@return {string} the current string
' */
Public Property Get content() As String
    ' return the real part of the content
    If pSize > 0 Then
        content = getLeft(size)
    Else
        content = vbNullString
    End If
End Property
'/**
' *extract the leftmost portion of a string
' *@param {} howMany the length of the string to return
' *@return {string} the current string
' */
Public Property Get getLeft(howMany As Long) As String
    ' return the left part of the content
    ' c.getLeft(howmany) is equivalent to left(c.content,howmany), but avoids extra assignment
    getLeft = getMid(1, howMany)
End Property
'/**
' *extract the rightmost portion of a string
' *@param {} howMany the length of the string to return
' *@return {string} the current string
' */
Public Property Get getRight(howMany As Long) As String
    ' return the right part of the content
    ' c.getRight(howmany) is equivalent to right(c.content,howmany), but avoids extra assignment
    getRight = getMid(pSize - howMany + 1, howMany)
End Property
'/**
' *extract a portion of a string
' *@param {} startPos the (1 base) start position to extraction from
' *@param {} howMany the length of the string to return
' *@return {string} the current string
' */
Public Property Get getMid(startPos As Long, _
            Optional howMany As Long = -1) As String
    ' extract from content
    ' c.getMid(startPos,howmany) is equivalent to mid(c.content,startPos, howmany), but avoids extra assignment
    Dim n As Long
    Debug.Assert startPos > 0 And startPos <= pSize
    n = howMany
    If n = -1 Then
        n = pSize - startPos + 1
    End If
    n = minNumber(pSize - startPos + 1, n)
    If n > 0 Then
        getMid = Mid(pContent, startPos, n)
    Else
        getMid = vbNullString
    End If
End Property
'/**
' *a self reference (useful for inside with..)
' *@return {} self
' */
Public Property Get self() As cStringChunker
    ' convenience for with in with
    Set self = Me
End Property
'/**
' *clear a chunker (set string to null)
' *@return {} self
' */
Public Function clear() As cStringChunker
    ' easy to clear out.. may as well keep the same buffer going
    pSize = 0
    Set clear = Me
End Function
'/**
' *encode a uri and add
' *@param {} url the url to add
' *@return {} self
' */
Public Function uri(url As String) As cStringChunker
    Set uri = add(URLEncode(url))
End Function
'/**
' *return the string
' *@return {} the string
' */
Public Function toString() As String
    toString = content()
End Function
'/**
' *add a string to the chunker
' *@param {} addstring the string to add
' *@return {} self
' */
Public Function add(addString As String) As cStringChunker
    Dim k As Long
    ' add some content to end
    k = Len(addString)
    If k > 0 Then
        adjustSize (k)
    
        Mid(pContent, size + 1, k) = addString
        pSize = size + k
    End If
    Set add = Me
End Function
'/**
' *add a string to the chunker, followed by a new line
' *@param {} addstring the string to add
' *@return {} self
' */
Public Function addLine(Optional addString As String = "") As cStringChunker
    Set addLine = add(addString).add(vbCrLf)
End Function
'/**
' *add a number of new lines
' *@param {} howMany number of lines to add
' *@return {} self
' */
Public Function addLines(Optional howMany As Long = 1) As cStringChunker
    Dim i As Long
    For i = 1 To howMany
        addLine
    Next i
   Set addLines = Me
End Function
'/**
' *insert a string at a particular position
' *@param {} insertString the string to insert (default 1 space)
' *@param {} insertBefore the position(base 1) before which to insert
' *@return {} self
' */
Public Function insert(Optional insertString As String = " ", _
                    Optional insertBefore As Long = 1) As cStringChunker
    'default position is at beginning, insert a space
    'c.insert("x",c.size+1) is equivalent to c.add("x")
    
    If insertBefore = pSize + 1 Then
        Set insert = add(insertString)
        
    Else
        ' 'todo .. how to handle programming errors?
        Debug.Assert insertBefore > 0 And insertBefore <= pSize
        
        ' regular string concatenation is better since there is overlap
        pContent = getLeft(insertBefore - 1) & insertString & getMid(insertBefore)
        pSize = Len(pContent)
        Set insert = Me
            
    End If
    Set insert = Me
End Function
'/**
' *overwrite a string at a particular position
' *@param {} overWriteString the string to insert (default 1 space)
' *@param {} overWriteAt the position(base 1) to start overwriting at
' *@return {} self
' */
Public Function overWrite(Optional overWriteString As String = " ", _
                    Optional overWriteAt As Long = 1) As cStringChunker
    'default position is at beginning, overwrite with a space
    Dim k As Long
    k = Len(overWriteString)
    If k > 0 Then
        ' 'todo .. how to handle programming errors?
        Debug.Assert overWriteAt >= 0
        '' we'll allow overwrite to extend past end, be greedy
        adjustSize (k)
        pSize = maxNumber(pSize, k + overWriteAt - 1)
        
        Mid(pContent, overWriteAt, k) = overWriteString
        
    End If
    Set overWrite = Me
End Function
'/**
' *shift the contents inside the chunker space
' *@param {} startPos the start position (base 1) of the string to shift
' *@param {} howManyChars the length of the string to shift (-ve means left, +ve right)_
' *@param {} replaceWith what to replace the moved contents with
' *@return {} self
' */
Public Function shift(Optional startPos As Long = 1, _
                Optional howManyChars As Long = 0, _
                Optional replaceWith As String = vbNullString) As cStringChunker
    ' shift by howmany chars .. negative= left, positive = right
    'TODO how to deal with programming errors? message, raise error, assert?
    Dim howMany As Long
    
    
    howMany = howManyChars
    If howMany = 0 Then
        howMany = Len(replaceWith)
    End If
        
    Debug.Assert howMany + startPos > 0
    Debug.Assert startPos <= pSize And startPos > 0
    
    ' make space
    If howMany <> 0 Then

        If howMany > 0 Then
        ' its a right shift, use insert
            Set shift = insert(Space(howMany), startPos)
        Else
            ' a left shift
            If startPos > 1 Then
                ' we can do an overwrite
                overWrite getMid(startPos, pSize + howMany), startPos + howMany
                pSize = pSize + howMany
            End If
        
        End If
    End If
    
    Set shift = Me
End Function
'/**
' *chop characters from the end of the content
' *@param {} howMany number of characters to chop
' *@return {} self
' */
Public Function chop(Optional howMany As Long = 1) As cStringChunker
    ' chop n charaters from end of content
    pSize = maxNumber(0, pSize - howMany)
    Set chop = Me
End Function
'/**
' *trim \s type chars from beginning and end
' *@param {} fromBeginning trim the beginnging of the content
' *@param {} fromEnd trim the end of the content
' *@return {} self
' */
Public Function chopSuperTrim(Optional fromBeginning As Boolean = True, _
    Optional fromEnd As Boolean = True) As cStringChunker
    
    Dim s As String, l As Long
    s = " " & vbCr & vbLf & vbTab

    If (fromEnd) Then
        Do While (pSize > 0)
            If (InStr(1, s, getRight(1)) = 0) Then Exit Do
            chop
        Loop
    End If
    
    l = 0
    If (fromBeginning) Then
        Do While (l < pSize)
            If (InStr(1, s, getMid(l + 1, 1)) = 0) Then Exit Do
            l = l + 1
        Loop
        ' l is the number of chars that are \s
        If (l > 0) Then shift l + 1, -l
    End If
    
    Set chopSuperTrim = Me
End Function
'/**
' *trim chars from end of content
' *@param {} chopString trim the beginnging of the content
' *@return {} self
' */
Public Function chopIf(chopString As String) As cStringChunker
    ' chop if content ends with chopString
    Dim k As Long
    k = Len(chopString)
    If k <= pSize Then
        If getRight(k) = chopString Then
            chop (k)
        End If
    End If
    Set chopIf = Me
End Function
'/**
' *trim chars from end of content and keep doing it while it matches
' *@param {} chopString trim the beginnging of the content
' *@return {} self
' */
Public Function chopWhile(chopString As String) As cStringChunker

    Dim k As Long, x As Long
    
    Set chopWhile = Me
    x = pSize
    While chopIf(chopString).size <> x
        x = pSize
    Wend

End Function
'/**
' *local max function
' *@param {} a first number to compare
' *@param {} b second number to compare
' *@return {} the bigger of a and b
' */
Private Function maxNumber(a As Long, b As Long) As Long
    If a > b Then
        maxNumber = a
    Else
        maxNumber = b
    End If
End Function
'/**
' *local min function
' *@param {} a first number to compare
' *@param {} b second number to compare
' *@return {} the smaller of a and b
' */
Private Function minNumber(a As Long, b As Long) As Long
    If a < b Then
        minNumber = a
    Else
        minNumber = b
    End If
End Function
'/**
' *adjust the underlying chunker buffer size if its needed
' *@param {} needMore how many chars we want space for
' *@return {} self
' */
Private Function adjustSize(needMore As Long) As cStringChunker
    Dim need As Long
    need = pSize + needMore
    If Len(pContent) < need Then
        pContent = pContent & Space(needMore + maxNumber(defaultChunkSize, Len(pContent)))
    End If
    Set adjustSize = Me
End Function
'/**
' *intialize some starting buffer
' */
Private Sub Class_Initialize()
    pSize = 0
    pContent = Space(defaultChunkSize)
End Sub






