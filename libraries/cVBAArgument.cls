Option Explicit
' one of these for each argument for each proc in the VBE-IDE
' v2.01 -
Private pParent As cVBAProcedure
Private pName As String
Private pArgType As String
Private pOptional As Boolean
Private pDefault As String
Private pDescription As String
Public Property Get name() As String
    name = pName
End Property
Public Property Get parent() As cVBAProcedure
    Set parent = pParent
End Property
Public Property Get argType() As String
    argType = pArgType
End Property
Public Property Get default() As String
    default = pDefault
End Property
Public Property Get isOptional() As Boolean
    isOptional = pOptional
End Property
Public Property Let description(p As String)
    pDescription = p
End Property
Public Property Get description() As String
    description = pDescription
End Property
Public Function init(m As cVBAProcedure, decString As String) As cVBAArgument
    Set pParent = m
    ' from the declaration item, figure out what this argument is
    Dim o As String, s As String, p As Long, np As Long, t As String, u As String, r As String
    Dim g As Long
    s = Trim(decString)
    o = "Optional"
    r = "(\w+)"
    t = rxGroup(r, s, 1)
    pOptional = (t = o)
    If pOptional Then
        r = "Optional\s+(\w+)"
        pName = rxGroup(r, s, 1)
    Else
        pName = t
    End If
    r = r + "\s+As\s+(\w+)"

    If rxTest(r, s) Then
        pArgType = Trim(rxGroup(r, s, 2))
        r = r + "\s+=(.*)"
        pDefault = rxGroup(r, s, 3)
    Else
        pArgType = "Variant"
    End If

    Set init = Me
End Function
Public Sub tearDown()
    Set pParent = Nothing
End Sub
