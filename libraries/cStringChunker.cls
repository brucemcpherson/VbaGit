'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:48:00 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/5102369/raw/cStringChunker.cls
' stringChunker class for VBA because string concat takes ages
Option Explicit
' v1.07  5102369
Private pContent As String
Private pSize As Long
' minimum amount to increment by each time
Const defaultChunkSize = 64
Public Property Get size() As Long
    ' this is how much content is real
    size = pSize
End Property
Public Property Get content() As String
    ' return the real part of the content
    If pSize > 0 Then
        content = getLeft(size)
    Else
        content = vbNullString
    End If
End Property
Public Property Get getLeft(howMany As Long) As String
    ' return the left part of the content
    ' c.getLeft(howmany) is equivalent to left(c.content,howmany), but avoids extra assignment
    getLeft = getMid(1, howMany)
End Property
Public Property Get getRight(howMany As Long) As String
    ' return the right part of the content
    ' c.getRight(howmany) is equivalent to right(c.content,howmany), but avoids extra assignment
    getRight = getMid(pSize - howMany + 1, howMany)
End Property
Public Property Get getMid(startPos As Long, Optional howMany As Long = -1) As String
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
Public Property Get self() As cStringChunker
    ' convenience for with in with
    Set self = Me
End Property
Public Function clear() As cStringChunker
    ' easy to clear out.. may as well keep the same buffer going
    pSize = 0
    Set clear = Me
End Function
Public Function uri(addstring As String) As cStringChunker
    Set uri = add(URLEncode(addstring))
End Function
Public Function toString() As String
    toString = content()
End Function
Public Function add(addstring As String) As cStringChunker
    Dim k As Long
    ' add some content to end
    k = Len(addstring)
    If k > 0 Then
        adjustSize (k)
    
        Mid(pContent, size + 1, k) = addstring
        pSize = size + k
    End If
    Set add = Me
End Function
Public Function addLine(Optional addstring As String = "") As cStringChunker
    Set addLine = add(addstring).add(vbCrLf)
End Function
Public Function addLines(Optional number As Long = 1) As cStringChunker
    Dim i As Long
    For i = 1 To number
        addLine
    Next i
   Set addLines = Me
End Function
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
                overWrite getMid(startPos + howMany, pSize - startPos + 1), startPos
                pSize = pSize + howMany
            End If
        
        End If
    End If
    
    Set shift = Me
End Function
Public Function chop(Optional n As Long = 1) As cStringChunker
    ' chop n charaters from end of content
    pSize = maxNumber(0, pSize - n)
    Set chop = Me
End Function
Public Function chopIf(t As String) As cStringChunker
    ' chop if its t
    Dim k As Long
    k = Len(t)
    If k <= pSize Then
        If getRight(k) = t Then
            chop (k)
        End If
    End If
    Set chopIf = Me
End Function
Public Function chopWhile(t As String) As cStringChunker
    ' chop if its t
    Dim k As Long, x As Long
    
    Set chopWhile = Me
    x = pSize
    While chopIf(t).size <> x
        x = pSize
    Wend

End Function
Private Function maxNumber(a As Long, b As Long) As Long
    If a > b Then
        maxNumber = a
    Else
        maxNumber = b
    End If
End Function
Private Function minNumber(a As Long, b As Long) As Long
    If a < b Then
        minNumber = a
    Else
        minNumber = b
    End If
End Function
Private Function adjustSize(needMore As Long) As cStringChunker
    Dim need As Long
    need = pSize + needMore
    If Len(pContent) < need Then
        pContent = pContent & Space(needMore + maxNumber(defaultChunkSize, Len(pContent)))
    End If
    Set adjustSize = Me
End Function
Private Sub Class_Initialize()
    pSize = 0
    pContent = Space(defaultChunkSize)
End Sub






