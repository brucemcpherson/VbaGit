'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:46 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414365/raw/usefulcJobject.vba
'v2.16 11.5.15
Option Explicit

Public Function fromISODateTime(iso As String) As Date
    Dim rx As RegExp, matches As MatchCollection, d As Date, ms As Double, sec As Double
    Set rx = New RegExp
    With rx
        .ignorecase = True
        .Global = True
        .pattern = "(\d{4})-([01]\d)-([0-3]\d)T([0-2]\d):([0-5]\d):(\d*\.?\d*)Z"
    End With
    Set matches = rx.Execute(iso)
    
    ' TODO -- timeszone

    If matches.Count = 1 And matches.Item(0).SubMatches.Count = 6 Then

        With matches.Item(0)
            sec = CDbl(.SubMatches(5))
            ms = sec - Int(sec)
            d = DateSerial(.SubMatches(0), .SubMatches(1), .SubMatches(2)) + _
                TimeSerial(.SubMatches(3), .SubMatches(4), Int(sec)) + ms / 86400
        End With
    
    Else
        d = 0
    End If
    
    fromISODateTime = d
   
End Function

Public Function toISODateTime(d As Date) As String
    Dim s As String, ms As Double, adjustSecond As Long
    
    ' need to adjust if seconds are going to be rounded up
    ms = milliseconds(d)
    adjustSecond = 0
    If (ms >= 0.5) Then adjustSecond = -1
    
    ' TODO - timezone
    toISODateTime = Format(year(d), "0000") & "-" & Format(month(d), "00") & "-" & Format(day(d), "00T") & _
            Format(d, "hh:mm:") & Format(DateAdd("s", adjustSecond, d), "ss") & Format(ms, ".000Z")

    
End Function
Public Function milliseconds(d As Date) As Double
    ' extract the milliseconds from the time
    Dim t As Date
    t = (d - DateSerial(year(d), month(d), day(d)) - TimeSerial(hour(d), Minute(d), Second(d)))
    If t < 0 Then
        ' the millsecond rounded it up
        t = (d - DateSerial(year(d), month(d), day(d)) - TimeSerial(hour(d), Minute(d), Second(d) - 1))
    End If
    
    milliseconds = t * 86400
    
End Function
Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject
    Dim j As New cJobject
    Set JSONParse = j.init(Nothing).parse(s, jtype, complain)
    j.tearDown
End Function
Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String
    JSONStringify = j.stringify(blf)
End Function
Public Function jSonArgs(options As String) As cJobject
    ' takes a javaScript like options paramte and converts it to cJobject
    ' it can be accessed as job.child('argName').value or job.find('argName') etc.
    Dim job As New cJobject
    If options <> vbNullString Then
        Set jSonArgs = job.init(Nothing, "jSonArgs").deSerialize(options)
    End If
End Function
Public Function optionsExtend(givenOptions As String, _
            Optional defaultOptions As String = vbNullString) As cJobject
    Dim jGiven As cJobject, jDefault As cJobject, _
        jExtended As cJobject, cj As cJobject
    ' this works like $.extend in jQuery.
    ' given and default options arrive as a json string
    ' example -
    ' optionsExtend ("{'width':90,'color':'blue'}", "{'width':20,'height':30,'color':'red'}")
    ' would return a cJobject which serializes to
    ' "{width:90,height:30,color:blue}"
    Set jGiven = jSonArgs(givenOptions)
    Set jDefault = jSonArgs(defaultOptions)
    
    ' now we combine them
    If Not jDefault Is Nothing Then
        Set jExtended = jDefault
    Else
        Set jExtended = New cJobject
        jExtended.init Nothing
    End If
    
    ' now we merge that with whatever was given
    If Not jGiven Is Nothing Then
        jExtended.merge jGiven
    End If
    
    ' and its over
    Set optionsExtend = jExtended
End Function

'udfs to expose classes
Public Function ucJobjectMake(r As Variant) As cJobject
    Dim cj As New cJobject
    Set ucJobjectMake = cj.deSerialize(CStr(r))
End Function
Public Function ucJobjectChildValue(json As Variant, child As Variant) As String
    ucJobjectChildValue = ucJobjectMake(CStr(json)).child(CStr(child)).value
End Function
Public Function ucJobjectLint(json As Variant, Optional child As Variant) As String
    Dim cj As cJobject
    Set cj = ucJobjectMake(json)
    If Not IsMissing(child) Then
        Set cj = cj.child(CStr(child))
    End If
    ucJobjectLint = cj.serialize(True)
End Function
Public Function cleanGoogleWire(sWire As String) As String
    Dim jStart As String, p As Long, newWire As Boolean, e As Long, s As String, reg As RegExp, _
        match As match, matches As MatchCollection, v As Double, i As Long, _
        year As Long, month As Long, day As Long, hour As Long, min As Long, sec As Long, ms As Long, _
        t As cStringChunker, consumed As Long

    jStart = "table:"
    p = InStr(1, sWire, jStart)
    'there have been multiple versions of wire ...
    If p = 0 Then
        'try the other one
        jStart = q & ("table") & q & ":"
        p = InStr(1, sWire, jStart)
        newWire = True
    End If

    p = InStr(1, sWire, jStart)
    e = Len(sWire) - 1

    If p <= 0 Or e <= 0 Or p > e Then
        MsgBox " did not find table definition data"
        Exit Function
    End If
    
    If Mid(sWire, e, 2) <> ");" Then
        MsgBox ("incomplete google wire message")
        Exit Function
    End If
    ' encode the 'table:' part to a cjobject
    p = p + Len(jStart)
    s = "{" & jStart & "[" & Mid(sWire, p, e - p - 1) & "]}"
    ' google protocol doesnt have quotes round the key of key value pairs,
    ' and i also need to convert date from javascript syntax new Date()
    ' we'll force it to be a 13 digit timestamp, since cjobject knows how to make that into a date
    's = rxReplace("(new\sDate)(\()(\d+)(,)(\d+)(,)(\d+)(\))", s, "'$3/$5/$7'")
    'new\s+date\s*\(\s*(\d+)\s*(,\s*\d+)\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\)
    Set reg = New RegExp
    With reg
        .pattern = "new\s+Date\s*\(\s*(\d+)\s*(,\s*\d+)\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\)"
        .Global = True
    End With
    Set matches = reg.Execute(s)

    
    If matches.Count > 0 Then
        Set t = New cStringChunker
        consumed = 0
        For Each match In matches
            t.add Mid(s, consumed + 1, match.FirstIndex - consumed)
            consumed = consumed + match.FirstIndex - consumed
            With match
                If .SubMatches.Count >= 2 And .SubMatches.Count <= 7 Then
                    'these are the only valid number of args to a javascript new Date()
                    day = 1
                    hour = 0
                    min = 0
                    sec = 0
                    ms = 0
                    year = .SubMatches(0)
                    month = Replace(.SubMatches(1), ",", "") + 1
                    If .SubMatches.Count > 2 And Not IsEmpty(.SubMatches(2)) Then day = Replace(.SubMatches(2), ",", "")
                    If .SubMatches.Count > 3 And Not IsEmpty(.SubMatches(3)) Then hour = Replace(.SubMatches(3), ",", "")
                    If .SubMatches.Count > 4 And Not IsEmpty(.SubMatches(4)) Then min = Replace(.SubMatches(4), ",", "")
                    If .SubMatches.Count > 5 And Not IsEmpty(.SubMatches(5)) Then sec = Replace(.SubMatches(5), ",", "")
                    If .SubMatches.Count > 6 And Not IsEmpty(.SubMatches(6)) Then ms = Replace(.SubMatches(6), ",", "")
                    ' now convert to a date and format
                    t.add(q) _
                        .add(CStr(DateSerial(year, month, day) + TimeSerial(hour, min, sec) + CDbl(ms) / 86400)) _
                        .add (q)
                    consumed = consumed + match.Length
                End If
            End With
        Next match
        If consumed < Len(s) Then t.add Mid(s, consumed + 1)
        s = t.content
        Set t = Nothing
    End If
    If Not newWire Then s = rxReplace("(\w+)(:)", s, "'$1':")
    cleanGoogleWire = s
    
End Function

Public Function xmlStringToJobject(xmlString As String, Optional complain As Boolean = True) As cJobject
    Dim doc As Object
    ' parse xml

    Set doc = CreateObject("msxml2.DOMDocument")
    doc.LoadXML xmlString
    If doc.parsed And doc.parseError = 0 Then
        Set xmlStringToJobject = docToJobject(doc, complain)
        Exit Function
    End If

    Set xmlStringToJobject = Nothing
    If complain Then
        MsgBox ("Invalid xml string - xmlparseerror code:" & doc.parseError)
    End If
    
    Exit Function
    
End Function
Public Function docToJobject(doc As Object, Optional complain As Boolean = True) As cJobject
    ' convert xml document to a cjobject
    Dim node As IXMLDOMNode, job As cJobject
    Set job = New cJobject
    job.init Nothing
       
    Set docToJobject = handleNodes(doc, job)
End Function
Private Function isArrayRoot(parent As IXMLDOMNode) As Boolean
    
    Dim node As IXMLDOMNode, n As Long, node2 As IXMLDOMNode
    
    
    isArrayRoot = False
    If parent.NodeType = NODE_ELEMENT And parent.ChildNodes.Length > 1 Then
        For Each node2 In parent.ChildNodes
            If node2.NodeType = NODE_ELEMENT Then
                n = 0
                For Each node In parent.ChildNodes
                    If node.NodeType = NODE_ELEMENT And _
                        node2.nodeName = node.nodeName Then n = n + 1
                Next node
                If n > 1 Then
                    ' this shoudl be true, but for leniency i'll comment
                    'Debug.Assert n = parent.ChildNodes.Length
                    isArrayRoot = True
                    Exit Function
                End If
            End If
        Next node2
    End If

    
End Function
Private Function handleNodes(parent As IXMLDOMNode, job As cJobject) As cJobject
    Dim node As IXMLDOMNode, joc As cJobject, attrib As IXMLDOMAttribute, i As Long, _
         arrayJob As cJobject
    
    If isArrayRoot(parent) Then
        ' we need an array associated with this this node
        ' subsequent members will need to make space for themselves
        Set joc = job.add(parent.nodeName).addArray
    Else
        Set joc = handleNode(parent, job)
    End If
    
    ' deal with any attributes
    If Not parent.Attributes Is Nothing Then
        For Each attrib In parent.Attributes
            handleNode attrib, joc
        Next attrib
    End If
    
    ' do the children
    If Not parent.ChildNodes Is Nothing And parent.ChildNodes.Length > 0 Then
        For Each node In parent.ChildNodes
            handleNodes node, joc
        Next node
    End If
    
    ' always return the level at which we arrived
    Set handleNodes = job
    
End Function
Private Function handleNode(node As IXMLDOMNode, job As cJobject, Optional arrayHead As Boolean = False) As cJobject
    Dim key As cJobject
    '' not a comprehensive convertor
    Set handleNode = job
    Debug.Print node.nodeName & node.NodeType & node.NodeValue
    Select Case node.NodeType
        Case NODE_ATTRIBUTE
            ' we cant have an array of attributes - this will silently use the latest
            job.add node.nodeName, node.NodeValue
            
        Case NODE_ELEMENT
            If job.isArrayRoot Then
                Dim b As Boolean
                b = (node.ChildNodes.Length = 1)
                If (b) Then b = node.ChildNodes(0).NodeType = NODE_TEXT
                If (b) Then
                    Set handleNode = job.add.add
                Else
                    Set handleNode = job.add.add(node.nodeName)
                End If
            Else
                Set handleNode = job.add(node.nodeName)
            End If

        Case NODE_TEXT
            job.value = node.NodeValue

            
        Case NODE_DOCUMENT, NODE_CDATA_SECTION, NODE_ENTITY_REFERENCE, _
            NODE_ENTITY, NODE_PROCESSING_INSTRUCTION, NODE_COMMENT, NODE_DOCUMENT_TYPE, _
            NODE_DOCUMENT_FRAGMENT, NODE_NOTATION
            ' just ignore these for now

            
        Case Else
            Debug.Assert False
    End Select
    
End Function
'/**
'* this will deal with the problem of code copied from javascript, where JSON has no quotes round property names
'* @param {string} theString the string to be hacked
'* @return {string} the hacked string
'*/
Public Function hackJSObjectToJSON(theString As String) As String
    hackJSObjectToJSON = _
        rxReplace("({|,)(?:\s*)(?:')?([A-Za-z_$\.][A-Za-z0-9_ \-\.$]*)(?:')?(?:\s*):", theString, "$1""$2"":")

End Function

'/**
'* this will deal with the problem of code copied from javascript, where JSON has no quotes round property names, with a callback
'* @param {string} theString the string to be hacked
'* @return {string} the hacked string
'*/
Public Function hackJSONPObjectToJSON(theString As String) As String
    hackJSONPObjectToJSON = _
        hackJSObjectToJSON(rxReplace("\w+\s*\()(.*)\);*", theString, "$2"))

End Function
