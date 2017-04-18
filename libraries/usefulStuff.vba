'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:48:00 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/3414346/raw
Option Explicit
' v2.23  3414346

' Acknowledgement for the microtimer procedures used here to
' thanks to Charles Wheeler - http://www.decisionmodels.com/
' ---


#If VBA7 And Win64 Then

Private Declare PtrSafe Function getTickCount _
    Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Declare PtrSafe Function getFrequency _
    Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    
Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal fileName As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMaximizedFocus _
  ) As Longlong
  
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Longlong, ByVal dwflags As Longlong, _
    ByVal lpWideCharStr As Longlong, ByVal cchWideChar As Longlong, _
    ByVal lpMultiByteStr As Longlong, ByVal cchMultiByte As Longlong, _
    ByVal lpDefaultChar As Longlong, ByVal lpUsedDefaultChar As Longlong) As Longlong
    
    
#Else

Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal fileName As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMaximizedFocus _
  ) As Long
  
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwflags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
#End If

' note original execute shell stuff came from this post
' http://stackoverflow.com/questions/3166265/open-an-html-page-in-default-browser-with-vba
' thanks to http://stackoverflow.com/users/174718/dmr

Private Const CP_UTF8 = 65001
Public Const cFailedtoGetHandle = -1

Public Function OpenUrl(url) As Boolean
    #If VBA7 And Win64 Then
    Dim lSuccess As Longlong
    #Else
    Dim lSuccess As Long
    #End If
    lSuccess = ShellExecute(0, "Open", url)
    OpenUrl = lSuccess > 32
End Function

Sub deleteAllFromCollection(co As Collection)
    Dim o As Object, i As Long
    For i = co.Count To 1 Step -1
        co(i).Delete
    Next i
    
End Sub


Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
#If VBA7 And Win64 Then
    Dim lLength As Longlong
#Else
    Dim lLength As Long
#End If
If UTF16 <> "" Then
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    sBuffer = Space$(CLng(lLength))
    lLength = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, CLng(lLength - 1))
Else
    UTF16To8 = ""
End If
End Function
'end of utf16to8


Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False, _
   Optional UTF8Encode As Boolean = True _
) As String

Dim StringValCopy As String: StringValCopy = _
    IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For i = 1 To StringLen
    Char = Mid$(StringValCopy, i, 1)
    CharCode = Asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
    End Select
  Next i
  URLEncode = Join(result, "")

End If
End Function
Public Sub cloneFormat(b As Range, a As Range)
    
    ' this probably needs additional properties copied over
    With a.Interior
        .Color = b.Interior.Color
    End With
    With a.Font
        .Color = b.Font.Color
        .size = b.Font.size
    End With
    With a
        .HorizontalAlignment = b.HorizontalAlignment
        .VerticalAlignment = b.VerticalAlignment
        
    End With

End Sub
Public Function compareAsKey(a As Variant, b As Variant, Optional asKey As Boolean = True) As Boolean
    If (asKey And TypeName(a) = "String" And TypeName(b) = "String") Then
        compareAsKey = (makeKey(a) = makeKey(b))
    Else
        compareAsKey = (a = b)
    
    End If
End Function
' sort a collection
Function SortColl(ByRef coll As Collection, eorder As Long) As Long
    Dim ita As Long, itb As Long
    Dim va As Variant, vb As Variant, bSwap As Boolean
    Dim x As Object, y As Object
    
    For ita = 1 To coll.Count - 1
        For itb = ita + 1 To coll.Count
            Set x = coll(ita)
            Set y = coll(itb)
            bSwap = x.needSwap(y, eorder)
            If bSwap Then
                With coll
                    Set va = coll(ita)
                    Set vb = coll(itb)
                    .add va, , itb
                    .add vb, , ita
                    .remove ita + 1
                    .remove itb + 1
                End With
            End If
        Next
    Next
End Function
Public Function getHandle(sName As String, Optional readOnly As Boolean = False) As Integer
    Dim hand As Integer
    On Error GoTo handleError
        hand = FreeFile
        If (readOnly) Then
            Open sName For Input As hand
        Else
            Open sName For Output As hand
        End If
        getHandle = hand
        Exit Function

handleError:
    MsgBox ("Could not open file " & sName)
    getHandle = cFailedtoGetHandle
End Function
Function afConcat(arr() As Variant) As String
    Dim i As Long, s As String
    s = ""
    For i = LBound(arr) To UBound(arr)
        s = s & arr(i, 1) & "|"
    Next i
    afConcat = s
End Function
Public Function quote(s As String) As String
    quote = q & s & q
End Function
Public Function q() As String
    q = Chr(34)
End Function
Public Function qs() As String
    qs = Chr(39)
End Function
Public Function bracket(s As String) As String
    bracket = "(" & s & ")"
End Function
Public Function list(ParamArray args() As Variant) As String
    Dim i As Long, s As String
    s = vbNullString
    For i = LBound(args) To UBound(args)
        If s <> vbNullString Then s = s & ","
        s = s & CStr(args(i))
    Next i
    list = s
End Function

Public Function qlist(ParamArray args() As Variant) As String
    Dim i As Long, s As String
    s = vbNullString
    For i = LBound(args) To UBound(args)
        If s <> vbNullString Then s = s & ","
        s = s & quote(CStr(args(i)))
    Next i
    qlist = s
End Function
Public Function diminishingReturn(val As Double, Optional s As Double = 10) As Double
    diminishingReturn = Sgn(val) * s * (Sqr(2 * (Sgn(val) * val / s) + 1) - 1)
End Function
Public Function superTrim(s As String) As String
    Dim c As cStringChunker
    Set c = New cStringChunker
    superTrim = c.add(s).chopSuperTrim.toString
    
End Function
Public Function makeKey(v As Variant) As String
    makeKey = LCase(Trim(CStr(v)))
End Function
' The below is taken from http://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    ' function inserts line feeds so we need to get rid of them
    Base64Encode = Replace(oNode.Text, vbLf, "")
    Set oNode = Nothing
    Set oXML = Nothing
End Function
'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  base64String = Replace(base64String, vbLf, "")
   
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function
Public Function openNewHtml(sName As String, sContent As String) As Boolean
    Dim handle As Integer

    handle = getHandle(sName)
    If (handle <> cFailedtoGetHandle) Then
        Print #handle, sContent
        Close #handle
        openNewHtml = True
    End If

End Function
Public Function readFromFile(sName As String) As String
    Dim handle As Integer
    handle = getHandle(sName, True)
    If (handle <> cFailedtoGetHandle) Then
        readFromFile = Input$(LOF(handle), #handle)
        Close #handle
    End If
End Function
Public Function arrayLength(a) As Long
    arrayLength = UBound(a) - LBound(a) + 1
End Function
Public Function getControlValue(ctl As Object) As Variant
    Select Case TypeName(ctl)
        Case "Shape"
            getControlValue = ctl.TextFrame.Characters.Text
        Case "Label"
            getControlValue = ctl.Caption
        Case Else
            getControlValue = ctl.value
    End Select
End Function
Public Function setControlValue(ctl As Object, v As Variant) As Variant
    Select Case TypeName(ctl)
        Case "Shape"
            ctl.TextFrame.Characters.Text = v
        Case "Label"
            ctl.Caption = v
        Case Else
            ctl.value = v
    End Select
    setControlValue = v
End Function
Public Function isinCollection(vCollect As Variant, sid As Variant) As Boolean
    Dim v As Variant
    If Not vCollect Is Nothing Then
        On Error GoTo handle
        Set v = vCollect(sid)
        isinCollection = True
        Exit Function
    End If
handle:
    isinCollection = False
End Function

Public Function dimensionCount(a As Variant) As Long
' the only way I can figure out how to do this is to keep trying till it fails
    Dim n As Long, j As Long

    n = 1
    On Error GoTo allDone
    While True
        j = UBound(a, n)
        n = n + 1
    Wend
    Debug.Assert False
    Exit Function
    
allDone:
    dimensionCount = n - 1
    Exit Function
    
End Function

Public Function encloseTag(tag As String, Optional newLine As Boolean = True, _
                    Optional tClass As String = vbNullString, _
                    Optional args As Variant) As String
    
    Dim i As Long, t As cStringChunker
    Set t = New cStringChunker
    ' args can be an array or a single item
    If Not IsArray(args) Then
        With t
            .add("<").add (tag)
            If tClass <> vbNullString Then .add(" class=").add (tClass)
            .add (">")
            If newLine Then .add (vbCrLf)
            .add (CStr(args))
            If newLine Then .add (vbCrLf)
            .add("</").add(tag).add (">")
            If newLine Then .add (vbCrLf)
        End With
    Else
        ' recurse for array memmbers
        For i = LBound(args) To UBound(args)
            t.add encloseTag(tag, newLine, tClass, args(i))
        Next i
    End If
    encloseTag = t.content
End Function

Public Function scrollHack() As String
    'hack for IOS
    scrollHack = _
     "<div id='wrapper' style='width:100%;height:100%;overflow-x:auto;" & _
     "overflow-y:auto;-webkit-overflow-scrolling: touch;'>"
End Function

Public Function escapeify(s As String) As String
    escapeify = _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace(s _
                                    , q, "\" & q), _
                                "%", "\" & "%"), _
                            ">", "\>"), _
                        "<", "\<")
    

    
End Function
Public Function unEscapify(s As String) As String
    unEscapify = _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace( _
                                    s, "\" & q, q), _
                                 "\" & "%", "%"), _
                             "\>", ">"), _
                         "\<", "<")
    
End Function
Public Function basicStyle() As String
    With New cStringChunker
        .add ".viewdiv {}"
        .add ".hide {"
        .add "display:none;position:absolute;"
        .add "padding:5px;background:white;color:black;"
        .add "border-radius:5px;border:1px solid black;"
        .add "}"
        basicStyle = .content
    End With

End Function
' i adapted this from some table css I found - apologies I dont have the site for crediting.
Public Function tableStyle() As String
    Dim t As cStringChunker
    Set t = New cStringChunker
t.add _
 " table {" & _
    "font-family:Arial, Helvetica, sans-serif;" & _
    "color:#666;" & _
    "font-size:10px;" & _
    "background:#eaebec;" & _
    "margin:4px;" & _
    "border:#ccc 1px solid;" & _
    "-moz-border-radius:3px;" & _
    "-webkit-border-radius:3px;" & _
    "border-radius:3px;" & _
    "-moz-box-shadow: 0 1px 2px #d1d1d1;" & _
    "-webkit-box-shadow: 0 1px 2px #d1d1d1;" & _
    "box-shadow: 0 1px 2px #d1d1d1;" & _
    "}" & _
 "table th {" & _
    "padding:8px 9px 8px 9px;" & _
    "border-top:1px solid #fafafa;" & _
    "border-bottom:1px solid #e0e0e0;" & _
    "background: #ededed;" & _
    "background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));" & _
    "background: -moz-linear-gradient(top,  #ededed,  #ebebeb);" & _
    "}"
    
t.add _
 "table tr {" & _
    "text-align: left;" & _
    "padding-left:16px;" & _
    "}" & _
 "table td {" & _
    "padding:6px;" & _
    "border-top: 1px solid #ffffff;" & _
    "border-bottom:1px solid #e0e0e0;" & _
    "border-left: 1px solid #e0e0e0;" & _
    "background: #fafafa;" & _
    "}" & _
 "table tr.even td {" & _
    "background: #f6f6f6;" & _
    "}"


 
    tableStyle = t.content
End Function
Public Function is64BitExcel() As Boolean
#If VBA7 And Win64 Then
    is64BitExcel = True
#Else
    is64BitExcel = False
#End If
End Function
Public Function includeJQuery() As String
    ' include jquery source
    With New cStringChunker
        .addLine jScriptTag("http://www.google.com/jsapi")
        .addLine jScriptTag
        .addLine "google.load('jquery', '1');"
        .addLine "</script>"
        includeJQuery = .content
    End With
    
End Function
Public Function includeGoogleCallBack(c As String) As String
    ' include google call back
    With New cStringChunker
        .addLine jScriptTag
        .addLine "google.setOnLoadCallback("
        .addLine c
        .addLine ");"
        .addLine "</script>"
        includeGoogleCallBack = .content
    End With
    
End Function
Public Function jScriptTag(Optional src As String) As String
    With New cStringChunker
        .add "<script type='text/javascript'"
        If src <> vbNullString Then
            .add(" src='").add(src).addLine ("'></script>")
        Else
            .addLine ">"
        End If
        jScriptTag = .content
    End With
End Function
Public Function jDivAtMouse()
    With New cStringChunker
        .addLine "function() {"
        .add "$('a.viewdiv').mousemove("
        .addLine "function(e) {"
        .add "var targetdiv = $('#d'+this.id);"
        .add "targetdiv.css({left:(e.pageX + 20) + 'px',"
        .add "top: (Math.max(0,e.pageY - targetdiv.height()/2)) + 'px'}).show();"
        .addLine "});"
        .add "$('a.viewdiv').mouseout("
        .addLine "function(e) {"
        .add "$('#d'+this.id).hide();"
        .addLine "});"
        .addLine "}"
        jDivAtMouse = .content
    End With
End Function


Function biasedRandom(possibilities, weights) As String
    Dim w As Variant, a As Variant, p As Variant, _
        r As Double, i As Long
    ' comes in as 2 lists
    a = Split(weights, ",")
    p = Split(possibilities, ",")
    ReDim w(LBound(a) To UBound(a))

    ' create cumulative
    For i = LBound(w) To UBound(w)
        w(i) = CDbl(a(i))
        If i > LBound(w) Then w(i) = w(i - 1) + w(i)
    Next i
    
    ' get random index
    r = Rnd() * w(UBound(w))
    
    ' find its weighted position
    For i = LBound(w) To UBound(w)
        If (r <= w(i)) Then
            biasedRandom = p(i)
            Exit Function
        End If
    Next i
    
End Function

Public Sub sleep(seconds As Long)

    Application.Wait TimeSerial(hour(Now()), Minute(Now()), Second(Now()) + seconds)
End Sub
Public Function getDateFromTimestamp(s As String) As Date
    Dim d As Double
    
    If (Len(s) = 13) Then
        ' javaScript Time
        d = CDbl(Left(s, 10))
        ' may need to round for milliseconds
        If Int(Mid(s, 11, 3) >= 500) Then
            d = d + 1
        End If
        
    ElseIf (Len(s) = 10) Then
        ' unix Time
        d = CDbl(s)
    
    Else
        ' wtf time
        getDateFromTimestamp = 0
        Exit Function
    
    End If
    getDateFromTimestamp = DateAdd("s", d, DateSerial(1970, 1, 1))

End Function
Public Function dateFromUnix(s As Variant) As Variant
    Dim d As Date, sd As String
    sd = CStr(s)
    
    If (Len(sd) > 0) Then
        d = getDateFromTimestamp(sd)
        If d = 0 Then
            dateFromUnix = CVErr(xlErrValue)
        Else
            dateFromUnix = d
        End If
    Else
        dateFromUnix = Empty
    End If

End Function
Public Function isSomething(o As Object) As Boolean

    isSomething = Not o Is Nothing
End Function


Public Function tinyTime() As Double
' Returns seconds.
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    tinyTime = 0
' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency
' Get ticks.
    getTickCount cyTicks1
    If cyFrequency Then tinyTime = cyTicks1 / cyFrequency
End Function


Function applyDefaults(value As Variant, defaultValue As Variant) As Variant
    If (IsObject(defaultValue)) Then
        If isUndefined(value) Then
            Set applyDefaults = defaultValue
        Else
            Set applyDefaults = value
        End If
    Else
        If isUndefined(value) Then
            applyDefaults = defaultValue
        Else
            applyDefaults = value
        End If
    End If
End Function
Function isUndefined(value As Variant) As Boolean
    If (IsObject(value)) Then
        isUndefined = value Is Nothing
    Else
        If (IsMissing(value) Or IsEmpty(value)) Then
            isUndefined = True
        Else
            isUndefined = (value = vbNullString)
        End If
    End If
End Function
Function conditionalAssignment(condition As Boolean, a As Variant, b As Variant) As Variant
    If (condition) Then

        If IsObject(a) Then
            Set conditionalAssignment = a
        Else
            conditionalAssignment = assignHelper(a)
        End If

    Else
        If IsObject(b) Then
            Set conditionalAssignment = b
        Else
            conditionalAssignment = assignHelper(b)
        End If
    
    End If
End Function
Private Function assignHelper(a As Variant) As Variant
    If IsObject(a) Then
        Set assignHelper = a
    Else
        If Not isUndefined(a) Then
            assignHelper = a
        Else
            assignHelper = vbNullString
        End If
    End If
End Function
Public Function getTimestampFromDate(Optional dt As Date = 0) As Double
    Dim d As Double
    
    If (dt = 0) Then
        dt = Now()
    End If
    
    ' convert into time since the epoch
    d = DateDiff("s", DateSerial(1970, 1, 1) + TimeSerial(0, 0, 0), dt)
    
    ' convert to ms
    d = d * 1000#

    getTimestampFromDate = d

End Function

Public Function checkOrCreateFolder(path As String, Optional optCreate As Boolean = True) As Object
    ' doing late binding to avoid refernce for this
    
    Dim fso As Object, cleanPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    'fso not smart enough to create entire thing, so we need recurse for each
    If (optCreate) Then
        recurseCreateFolder fso, fso.GetAbsolutePathName(path)
    End If
    
    Set checkOrCreateFolder = fso.getFolder(path)
End Function
Private Function recurseCreateFolder(fso As Object, cleanPath As String) As Object
    Dim parentPathString As String

    If Not fso.FolderExists(cleanPath) Then
        ' need to create the parent first
        recurseCreateFolder fso, fso.GetParentFolderName(cleanPath)
        ' now we can do it
        fso.CreateFolder (cleanPath)
    End If

End Function
Public Function writeToFolderFile(folderName As String, fileName As String, content As String) As String
    
    Dim file As Object, fso As Object
    Dim path As String
    path = fileName
    ' create the folder if we need to
    
    If (folderName <> vbNullString) Then
        path = concatFolderName(folderName, path)
        checkOrCreateFolder folderName
    End If
    ' write the data
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(path, 2, True)
    file.write content
    writeToFolderFile = content
End Function
Public Function getAllSubFolderPaths(folderName As String) As String
    Dim folder As Object, subFolder As Object, c As cStringChunker
    Set c = New cStringChunker
    Set folder = checkOrCreateFolder(folderName, False)
    If (isSomething(folder)) Then
        For Each subFolder In folder.subFolders
            c.add(subFolder.path).add ","
        Next subFolder
    End If
    getAllSubFolderPaths = c.chopIf(",").toString
End Function
Public Function readFromFolderFile(folderName As String, fileName As String) As String
    Dim file As Object, fso As Object
    Dim path As String
    path = fileName
    
    If (folderName <> vbNullString) Then
        path = concatFolderName(folderName, fileName)
    End If
    ' read the data
    If (fileExists(path)) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(path, 1)
        If (file.AtEndOfStream) Then
            readFromFolderFile = ""
        Else
            readFromFolderFile = file.readAll()
        End If
        file.Close
    End If
    
End Function
Public Function fileExists(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(path)
End Function
Public Function concatFolderName(folderName As String, fileName As String) As String
    Dim c As cStringChunker
    Set c = New cStringChunker
    concatFolderName = c.add(folderName).chopIf("/").chopIf("\").add("/").add(fileName).toString
    
End Function
