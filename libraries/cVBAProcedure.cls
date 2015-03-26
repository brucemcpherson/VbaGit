Option Explicit
' one of these for each procedure/property in the VBE-IDE
' v2.4.1
Private pParent As cVBAmodule
Private pName As String
Private pProcKind As vbext_prockind
Private pArguments As Collection
Private pIndex As Long
Private pDescription As String
Private pReturnDoc  As String
Private Sub Class_Initialize()
    Set pArguments = New Collection
End Sub
Public Property Let description(p As String)
    pDescription = p
End Property
Public Property Get description() As String
    description = pDescription
End Property
Public Property Get returnDoc() As String
    returnDoc = pReturnDoc
End Property
Public Property Get name() As String
    name = pName
End Property
Public Property Get arguments() As Collection
    Set arguments = pArguments
End Property
Public Property Get parent() As cVBAmodule
    Set parent = pParent
End Property
Public Property Get procKind() As vbext_prockind
    procKind = pProcKind
End Property
'/**
' * checks to see if a given variable name is an argument of this procedure
' * @param {string} argName the name to check
' * @return {boolean} whether it is an argument
'*/
Public Function isAnArgument(argName As String) As Boolean
    Dim a As cVBAArgument
    For Each a In arguments
        If compareAsKey(argName, a.name) Then
            isAnArgument = True
            Exit Function
        End If
    Next a
    isAnArgument = False
End Function
Public Function init(m As cVBAmodule, pn As String, pk As vbext_prockind) As cVBAProcedure
    Set pParent = m
    pProcKind = pk
    pName = pn
    pParent.procedures.add Me, makeKey(name & "-" & procTextKind)
    pIndex = pParent.procedures.Count
    pDescription = superTrim(getDescription)
    dealWithArguments declaration
    pReturnDoc = getReturnDoc
    Set init = Me
End Function
Public Sub tearDown()
    Dim p As cVBAArgument
    For Each p In arguments
        p.tearDown
    Next p
    Set pParent = Nothing
End Sub
Public Property Get lineCount() As Long
    lineCount = codeModule.ProcCountLines(pName, pProcKind)
End Property
Public Property Get codeModule() As codeModule
    Set codeModule = pParent.vCom.codeModule
End Property
Public Property Get startLine() As Long
    startLine = codeModule.ProcBodyLine(pName, pProcKind)
End Property
Public Function removeComments(s As String) As String
    removeComments = getRidOfComments(s)
End Function
'/**
' * interprets jsdoc like procedure description
' * @return {string} the procedure description
'*/
Public Function getDescription() As String
    Dim matches As MatchCollection
    Set matches = getRx("\/\*{2}\s*[\n\r][\s'*]*(.*)").Execute(getTheCodePlusLeadingComments)

    getDescription = ""
    If (matches.Count > 0) Then
        getDescription = superTrim(matches(0).SubMatches(0))
    End If
End Function
Private Function getReturnDoc() As String
    Dim rx As String, s As String, a As Variant, i As Long, _
        p As cVBAArgument, j As Long, k As Long
    
    Dim matches As MatchCollection, m As match
    ' get any jdocs description of the arguments
    Set matches = getRx("@return\s+(?:{\w*})\s*(.*)$").Execute(getTheCodePlusLeadingComments)
    getReturnDoc = ""
    If matches.Count > 0 Then
        getReturnDoc = superTrim(matches(0).SubMatches(0))
    End If

End Function
Private Function dealWithArguments(dec As String) As cVBAProcedure
    Dim rx As String, s As String, a As Variant, i As Long, _
        p As cVBAArgument, j As Long, k As Long
    
    Dim matches As MatchCollection, m As match
    ' get any jdocs description of the arguments
    Set matches = getRx("@param\s+(?:{\w*})\s*(\w*)\s*(.*)$").Execute(getTheCodePlusLeadingComments)
    
    rx = "\((.*)\)"
    s = Trim(rxString(rx, dec))
    
    If Left(s, 1) = "(" Then s = Mid(s, 2)
    If Right(s, 1) = ")" Then s = Left(s, Len(s) - 1)
    a = Split(s, ",")
    For i = LBound(a) To UBound(a)
        Set p = New cVBAArgument
        p.init Me, CStr(a(i))
        arguments.add p, LCase(a(i))
        ' now find any doc matches
        If matches.Count > 0 Then
            For j = 0 To matches.Count - 1
                ' these 2 should match the arg name and its description
                If matches(j).SubMatches.Count = 2 Then
                    If (compareAsKey(matches(j).SubMatches(0), p.name)) Then
                        p.description = superTrim(matches(j).SubMatches(1))
                    End If
                Else
                    Debug.Print matches(j) & " had no description"
                End If
            Next j
        End If
    Next i
    Set dealWithArguments = Me
End Function
Public Property Get scope() As String
    Dim dec As String
    dec = declaration
    If rxTest("^Private", dec) Then
        scope = "Private"
    Else
        scope = "Public"
    End If
End Property
Private Function textKind(k As vbext_prockind) As String
    Dim dec As String
    
    Select Case k
        Case vbext_pk_Get
            textKind = "Get"
            
        Case vbext_pk_Set
            textKind = "Set"
            
        Case vbext_pk_Let
            textKind = "Let"
            
        Case vbext_pk_Proc
            dec = declaration
            If rxTest("function\s\w", dec) Then
                textKind = "Function"
            Else
                textKind = "Sub"
            End If
            
            
        Case Else
            Debug.Assert False
        
    End Select
End Function
Public Property Get procTextKind() As String
    procTextKind = textKind(pProcKind)
End Property
Public Property Get procReturns() As String
    ' return the type of thing it returns
    Dim dec As String, r As String
    'regex to get as xxx
    r = "(.*[^\w+$])(\w+$)"
    Select Case procTextKind
    
        Case "Get", "Function"
            dec = declaration
            
            If (rxTest(r, dec)) Then
                procReturns = rxReplace(r, dec, "$2")
            Else
                procReturns = "Variant"
            End If
            
        Case "Set", "Let", "Sub"
            procReturns = "void"
            
        Case Else
            Debug.Assert False
        
    End Select
End Property
Public Function getTheCode() As String
    ' returns the entire code for this procedure
    getTheCode = codeModule.Lines(startLine, _
        getFinishWithoutTrailingComments - startLine + 1)
End Function
Public Function getFinishWithoutTrailingComments() As Long
    ' the comments are attached to the end of the function, when they should belong to next
    Dim rx As RegExp, finish As Long
    Set rx = getTheEndRx
    finish = startLine + lineCount - 1

    While finish > startLine + 1 And Not rx.Test(codeModule.Lines(finish, 1))
        finish = finish - 1
    Wend
    
    getFinishWithoutTrailingComments = finish
    
End Function
Public Function getTheEndRx() As RegExp
    Set getTheEndRx = getRx("\bend\s*function|sub|property")
End Function
Public Function getTheCodePlusLeadingComments() As String
    ' returns the entire code for this procedure
    Dim start As Long, prev As cVBAProcedure, rx As RegExp, _
        lengthOfCode As Long, lengthOfLeadingComments As Long
    
    Set rx = getTheEndRx()
    
    ' if this is not the first procedure, the comments for it start at the end of the previous
    If (pIndex > 1) Then
        Set prev = pParent.procedures(pIndex - 1)
        ' we want the comments too (which get attached to the previous function)
        start = startLine
        While start > prev.startLine + 2 And Not rx.Test(codeModule.Lines(start - 1, 1))
            start = start - 1
        Wend
       
    Else
        start = codeModule.CountOfDeclarationLines + 1

    End If
    
    lengthOfCode = 1 + getFinishWithoutTrailingComments - startLine
    lengthOfLeadingComments = startLine - start
  
    getTheCodePlusLeadingComments = codeModule.Lines(start, _
       lengthOfLeadingComments + lengthOfCode)
End Function
Public Property Get declaration() As String
    ' return the declaration line, stripped of comments and consolidated onto one line
    Dim n As Long, s As String, t As String, doMore As Boolean
    
    n = 0
    t = vbNullString
    Do
        s = LTrim(RTrim(removeComments(codeModule.Lines(n + startLine, 1))))

        If (s <> vbNullString) Then
            doMore = Right(s, 1) = "_" And n < lineCount
            If (doMore) Then s = Left(s, Len(s) - 1)
            t = t & s
        Else
            doMore = False
        End If
        n = n + 1
        
    Loop While doMore
    declaration = t
    
End Property
