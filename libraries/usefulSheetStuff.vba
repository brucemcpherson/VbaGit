Option Explicit
' v0.1 23.3.15
Function firstCell(inrange As Range) As Range
    Set firstCell = inrange.Cells(1, 1)
End Function
Function lastCell(inrange As Range) As Range
    Set lastCell = inrange.Cells(inrange.rows.Count, inrange.columns.Count)
End Function
Function isSheet(o As Object) As Boolean
     Dim r As Range
     On Error GoTo handleError
        Set r = o.Cells
        isSheet = True
        Exit Function

handleError:
    isSheet = False
End Function
Public Function findShape(sName As String, Optional ws As Worksheet = Nothing) As Shape
    Dim s As Shape, t As Shape
    If ws Is Nothing Then Set ws = ActiveSheet
    For Each s In ws.Shapes
        If makeKey(s.name) = makeKey(sName) Then
            Set t = s
            Exit For
        End If
        If s.Type = msoGroup Then
            Set t = findRecurse(sName, s.GroupItems)
            If Not t Is Nothing Then
                Exit For
            End If
        End If
    Next s
    Set findShape = t
    
End Function
Public Function findRecurse(target As String, co As GroupShapes) As Shape
    Dim s As Shape, t As Shape
    ' only works one level down.. cant get .gtoupitems to work properly
    For Each s In co
        If makeKey(s.name) = makeKey(target) Then
            Set t = s
            Exit For
        End If
    Next s
    Set findRecurse = t
End Function

Public Sub clearHyperLinks(ws As Worksheet)
' delete all the hyperlinks on a sheet
    With ws
        While .Hyperlinks.Count > 0
           .Hyperlinks(1).Delete
        Wend
    End With
End Sub
Function sheetExists(sName As String, Optional complain As Boolean = True) As Worksheet
    
    On Error GoTo handleError
        Set sheetExists = Sheets(sName)
        Exit Function

handleError:
    If complain Then MsgBox ("Could not open sheet " & sName)
    Set sheetExists = Nothing

End Function
Function wholeSheet(wn As String) As Range
    ' return a range representing the entire used worksheet
    Set wholeSheet = wholeWs(sheetExists(wn))
End Function
Function wholeWs(ws As Worksheet) As Range
    Set wholeWs = ws.UsedRange
End Function
Function wholeRange(r As Range) As Range
    Set wholeRange = wholeWs(r.Worksheet)
End Function
Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, _
        Optional singlecell As Boolean = False) As Range
    ' does a normal .find, but catches where range is nothing
    Dim u As Range
    Set u = Nothing

    If r Is Nothing Then
        Set u = Nothing
    Else
        Set u = r.find(x, , xlValues, xlWhole)
    End If
    
    If singlecell And Not u Is Nothing Then
        Set u = firstCell(u)
    End If
 
    If complain And u Is Nothing Then
        Call msglost(x, r)
    End If
    
    Set cleanFind = u
    
End Function
Sub msglost(x As Variant, r As Range, Optional extra As String = "")

    MsgBox ("Couldnt find " & CStr(x) & " in " & SAd(r) & " " & extra)

End Sub
Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    Dim r As Range
    Dim u As Range
    
    ' creates an address including the worksheet name
    strA = ""
    For Each r In rngIn.Areas
        Set u = r
        If singlecell Then
            Set u = firstCell(u)
        End If
        strA = strA + SAdOneRange(u, target, singlecell, removeRowDollar, removeColDollar) & ","
    Next r
    SAd = Left(strA, Len(strA) - 1)
End Function
Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
                        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    
    ' creates an address including the worksheet name
    
    strA = AddressNoDollars(rngIn, removeRowDollar, removeColDollar)
    
    ' dont bother with worksheet name if its on the same sheet, and its been asked to do that
    
    If Not target Is Nothing Then
        If target.Worksheet Is rngIn.Worksheet Then
            SAdOneRange = strA
            Exit Function
        End If
    End If

    ' otherwise add the sheet name
    
    SAdOneRange = "'" & rngIn.Worksheet.name & "'!" & strA
        
End Function
Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String
' return address minus the dollars
    Dim st As String
    Dim p1 As Long, p2 As Long
    AddressNoDollars = a.Address
    
    If doRow And doColumn Then
        AddressNoDollars = Replace(a.Address, "$", "")
    Else
        p1 = InStr(1, a.Address, "$")
        p2 = 0
        If p1 > 0 Then
            p2 = InStr(p1 + 1, a.Address, "$")
        End If
        ' turn $A$1 into A$1
        If doColumn And p1 > 0 Then
            AddressNoDollars = Left(a.Address, p1 - 1) & Mid(a.Address, p1 + 1)
        
        ' turn $a$1 into $a1
        ElseIf doRow And p2 > 0 Then
            AddressNoDollars = Left(a.Address, p2 - 1) & Mid(a.Address, p2 + 1, p2 - p1)
    
        End If
    End If
    
    
End Function
Function isReallyEmpty(r As Range) As Boolean
    Dim b As Boolean
    b = (Application.CountBlank(r) = r.Cells.Count)

    isReallyEmpty = b
End Function
Function toEmptyRow(r As Range) As Range
    Dim o As Range, u As Range, w As Long
    ' returns to first blank row
    Set u = wholeRange(r)
    Set o = r
    w = lastCell(u).row + 1
    Do While True
        ' whats left in the sheet
        Set o = cleanFind(Empty, o.Resize(w, 1), True, True)
        If isReallyEmpty(o.Resize(1, r.columns.Count)) Then
            Exit Do
        Else
            Set o = o.Offset(1)
        End If
    Loop

    If (o.row > lastCell(r).row And r.rows.Count > 1) Then
        Set toEmptyRow = r
    Else
        If o.row > r.row Then
            Set toEmptyRow = r.Resize(o.row - r.row)
        Else
            MsgBox ("nothing on sheet")
            Set toEmptyRow = Nothing
        End If
    End If
    
End Function
Function toEmptyCol(r As Range) As Range

    Dim o As Range, u As Range, w As Long
    ' returns to first blank column
    Set u = wholeRange(r)
    Set o = r
    w = lastCell(u).column + 1
    Do While True
        Set o = cleanFind(Empty, o.Resize(1, w), True, True)
        If isReallyEmpty(toEmptyRow(o)) Then
            Exit Do
        Else
            Set o = o.Offset(, 1)
        End If
    Loop
    If (o.column > r.column) Then
        Set toEmptyCol = r.Resize(r.rows.Count, o.column - r.column)
    End If
End Function
Function toEmptyBox(r As Range) As Range
    Set toEmptyBox = toEmptyCol(toEmptyRow(r))
End Function
Public Function getLikelyColumnRange(Optional ws As Worksheet = Nothing) As Range
    ' figure out the likely default value for the refedit.
    Dim rstart As Range
    If ws Is Nothing Then
        Set rstart = wholeSheet(ActiveSheet.name)
    Else
        Set rstart = wholeSheet(ws.name)
    End If

    Set getLikelyColumnRange = toEmptyBox(rstart)
    
End Function
Sub deleteAllShapes(r As Range, startingwith As String)
   
    Dim l As Long
    With r.Worksheet
        For l = .Shapes.Count To 1 Step -1
            If Left(.Shapes(l).name, Len(startingwith)) = startingwith Then
                .Shapes(l).Delete
            End If
        Next l
    End With
    
End Sub
Function makearangeofShapes(r As Range, startingwith As String) As ShapeRange
   
    Dim s As Shape
    
    Dim n() As String, sz As Long
    With r.Worksheet
        For Each s In .Shapes
            If Left(s.name, Len(startingwith)) = startingwith Then
                sz = sz + 1
                ReDim Preserve n(1 To sz) As String
                n(sz) = s.name

            End If
        Next s
        Set makearangeofShapes = .Shapes.Range(n)
    End With
    
End Function
Public Function nameExists(s As String) As name
    On Error GoTo handle
    Set nameExists = ActiveWorkbook.Names(s)
    Exit Function
handle:
    Set nameExists = Nothing
End Function
Public Function whereIsThis(r As Variant) As Range
    Dim n As name
    
    If TypeName(r) = "range" Then
        Set whereIsThis = r
    Else
        Set n = nameExists(CStr(r))
        If Not n Is Nothing Then
            Set whereIsThis = n.RefersToRange
        Else
            Set whereIsThis = Range(r)
        End If
    End If
            
        
End Function
Sub pivotCacheRefreshAll()

    Dim pc As PivotCache
    Dim ws As Worksheet

    With ActiveWorkbook
        For Each pc In .PivotCaches
            pc.refresh
        Next pc
    End With

End Sub
'--- based on trig at http://www.movable-type.co.uk/scripts/latlong.html
Public Function getLatFromDistance(mLat As Double, d As Double, heading As Double) As Double
    Dim lat As Double
    ' convert ro radians
    lat = toRadians(mLat)
    getLatFromDistance = _
        fromRadians( _
            Application.WorksheetFunction.Asin(sIn(lat) * _
            Cos(d / earthRadius) + _
            Cos(lat) * _
            sIn(d / earthRadius) * _
            Cos(heading)))
End Function
Public Function getLonFromDistance(mLat As Double, mLon As Double, d As Double, heading As Double) As Double
    Dim lat As Double, lon As Double, newLat As Double
    ' convert ro radians
    lat = toRadians(mLat)
    lon = toRadians(mLon)
    newLat = toRadians(getLatFromDistance(mLat, d, heading))
    getLonFromDistance = _
        fromRadians( _
             (lon + Application.WorksheetFunction.Atan2(Cos(d / earthRadius) - _
            sIn(lat) * _
            sIn(newLat), _
            sIn(heading) * _
            sIn(d / earthRadius) * _
            Cos(lat))))
End Function
Public Function earthRadius() As Double
    ' earth radius in km.
    earthRadius = 6371
End Function
Public Function toRadians(deg)
    toRadians = Application.WorksheetFunction.Pi / 180 * deg
End Function
Public Function fromRadians(rad) As Double
    'convert radians to degress
    fromRadians = 180 / Application.WorksheetFunction.Pi * rad
End Function
Public Function min(ParamArray args() As Variant)
    min = Application.WorksheetFunction.min(args)
End Function
Public Function max(ParamArray args() As Variant)
    max = Application.WorksheetFunction.max(args)
End Function
Public Function toClipBoard(s As String) As String
    With New MSForms.DataObject
        .SetText s
        .PutInClipboard
    End With
End Function

Public Function importTabbed(fn As String, r As Range) As Range

    r.Worksheet.QueryTables.add(Connection:= _
        "TEXT;" + fn, Destination:=r).refresh BackgroundQuery:=False

    Set importTabbed = r
End Function
