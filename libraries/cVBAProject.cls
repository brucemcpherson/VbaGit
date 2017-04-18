Option Explicit
' v2.03
' one of these for each project in the workbook
Private pWorkBook As Workbook
Private pModules As Collection
Public Property Get theProject() As VBProject
    Set theProject = pWorkBook.VBProject
End Property

Public Property Get name() As String
    name = theProject().name
End Property

Public Property Get wBook() As Workbook
    Set wBook = pWorkBook
End Property

Public Property Get modules() As Collection
    Set modules = pModules
End Property
Public Function init(wb As Workbook) As cVBAProject
    Set pWorkBook = wb
    Set init = Me
End Function
Public Sub tearDown()
    Dim vm As cVBAmodule
    For Each vm In pModules
        vm.tearDown
    Next vm
    Set pModules = Nothing
End Sub

Private Sub Class_Initialize()
    Set pModules = New Collection
End Sub