Option Explicit
' v0.1 23.3.15
Public Function straightenOutContinuations(s As String) As String
    ' just get rid of continuations and move
    straightenOutContinuations = getRx("_\s*$\n").Replace(s, "")
End Function
Public Function getRidOfDims(s As String) As String
    'get rid of dims which may have locals matching function names, but have to leave if new mentioned
    getRidOfDims = getRx("\s*dim(?!.*\s*new\s*).*").Replace(s, "")
End Function
Public Function getRidOfQuoted(s As String) As String
    getRidOfQuoted = getRx("(""[^""]*"")").Replace(s, "")
End Function
Public Function getRidOfComments(s As String) As String
    getRidOfComments = getRx("("".*?"")|('.*$)").Replace(s, "$1")
End Function
Public Function getRx(pattern As String) As RegExp
    Dim rx As RegExp
    Set rx = New RegExp
    With rx
        .ignorecase = True
        .Global = True
        .MultiLine = True
        .pattern = pattern
    End With
    Set getRx = rx
End Function