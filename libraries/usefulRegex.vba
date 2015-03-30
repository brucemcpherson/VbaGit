Option Explicit
' v0.1.1 27.3.15
Public Function straightenOutContinuations(s As String) As String
    ' just get rid of continuations and move
    straightenOutContinuations = getRidOfMultipleSpaces( _
        getRx("_\s*$\n").Replace(s, ""))
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
Public Function getRidOfMultipleSpaces(s As String) As String
    getRidOfMultipleSpaces = getRx("[\t ]{2,}", False).Replace(s, " ")
End Function
Public Function getRx(pattern As String, Optional multi As Boolean = True) As RegExp
    Dim rx As RegExp
    Set rx = New RegExp
    With rx
        .ignorecase = True
        .Global = True
        .MultiLine = multi
        .pattern = pattern
    End With
    Set getRx = rx
End Function
'/**
' *@return {} get a regex that picks out the end of a sub/function
'*/
Public Function getTheEndRx() As RegExp
    Set getTheEndRx = getRx("\bend\s*function|sub|property")
End Function
'/**
' *@return {} get a regex that picks out all lines with Dim
'*/
Public Function getDimLinesRx() As RegExp
    Set getDimLinesRx = getRx("(^\s*dim\s+.*)$")
End Function
'/**
' *@return {} get a regex that picks out all locally defined variables from a dim
'*/
Public Function getDimLocalsRx() As RegExp
    Set getDimLocalsRx = getRx("dim|(?:\s+as\s+)(\w+)")
End Function