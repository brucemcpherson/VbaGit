Option Explicit
'v2.01
' one of these for each module/class
'
Private pvCom As VBComponent
Private pProcedures As Collection
Private pParent As cVBAProject
Public Property Get name() As String
    name = pvCom.name
End Property
Public Property Get parent() As cVBAProject
    Set parent = pParent
End Property
Public Property Get vCom() As VBComponent
    Set vCom = pvCom
End Property
Public Property Get procedures() As Collection
    Set procedures = pProcedures
End Property
Public Sub init(v As VBComponent, pj As cVBAProject)

    Set pParent = pj
    Set pvCom = v
End Sub
Public Sub tearDown()
    Dim proc As cVBAProcedure
    For Each proc In pProcedures
         proc.tearDown
    Next proc
    Set pProcedures = Nothing
    Set pParent = Nothing
End Sub
Public Property Get textKind() As String
    Select Case pvCom.Type
    
        Case vbext_ct_ActiveXDesigner
            textKind = "ActiveXDesigner"
            
        Case vbext_ct_ClassModule
            textKind = "ClassModule"
            
        Case vbext_ct_Document
            textKind = "Document"
        
        Case vbext_ct_MSForm
            textKind = "MSForm"
            
        Case vbext_ct_StdModule
            textKind = "StdModule"
            
        Case Else
            Debug.Assert False
            
    End Select
End Property


Private Sub Class_Initialize()
    Set pProcedures = New Collection
End Sub