Option Explicit
' one of these for each procedure/property in the VBE-IDE
' v2.03
Private pParent As cVBAmodule
Private pName As String
Private pProcKind As vbext_prockind
Private pArguments As Collection
Private Sub Class_Initialize()
    Set pArguments = New Collection
End Sub
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
Public Function init(m As cVBAmodule, pn As String, pk As vbext_prockind) As cVBAProcedure
    Set pParent = m
    pProcKind = pk
    pName = pn
    pParent.procedures.add Me, name & "-" & procTextKind
    dealWithArguments declaration
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
Private Function dealWithArguments(dec As String) As cVBAProcedure
    Dim rx As String, s As String, a As Variant, i As Long, p As cVBAArgument
    ''rx = "(?:\w+\s*\()(.*)\)"
    rx = "\((.*)\)"
    s = Trim(rxString(rx, dec))
    
    If Left(s, 1) = "(" Then s = Mid(s, 2)
    If Right(s, 1) = ")" Then s = Left(s, Len(s) - 1)
    a = Split(s, ",")
    For i = LBound(a) To UBound(a)
        Set p = New cVBAArgument
        p.init Me, CStr(a(i))
        arguments.add p, LCase(a(i))
        
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
Public Property Get getTheCode() As String
    ' returns the entire code for this procedure
    getTheCode = codeModule.Lines(startLine, lineCount)
End Property
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
