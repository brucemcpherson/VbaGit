Option Explicit
' v0.1 23.3.15
' translation of cUrlResult - https://github.com/brucemcpherson/gasGit/blob/master/libraries/cUrlResult/Code.js

'/**
' * since I use this all the time,may as well make it a library
' * does UrlFetch() stuff and creates standard results
' */

'/**
'* execute a get
'* @param {string} url the url
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Public Function urlGet(url As String, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False) As cJobject
    
    Set urlGet = _
        urlExecute(url, "GET", _
            vbNullString, _
            optOptions, _
            optAccessToken, _
            optBasic)
        
End Function

'/**
'* execute a post
'* @param {string} url the url
'* @param {string} optMethod the http method
'* @param {string} optPayload any payload
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Public Function urlPost(url As String, _
    Optional optMethod As String = "POST", _
    Optional optPayload As Variant, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False) As cJobject
    Dim payload As String
    If (IsObject(optPayload)) Then
        payload = optPayload.stringify
    Else
        If (isUndefined(optPayload)) Then
            payload = vbNullString
        Else
            payload = optPayload
        End If
    End If
    Set urlPost = _
        urlExecute(url, _
            optMethod, _
            payload, _
            optOptions, _
            optAccessToken, _
            optBasic)
  
End Function
'/**
'* execute a urlfetch
'* @param {string} url the url
'* @param {string} optMethod the http method
'* @param {string} optPayload any payload
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Private Function urlExecute(url As String, _
    Optional optMethod As String = "GET", _
    Optional optPayload As String = vbNullString, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False)
    
    Dim job As cJobject
    
    ' we'll need some headers
    If (optOptions Is Nothing) Then
        Set optOptions = New cJobject
        optOptions.init Nothing
    End If
    
    If (optOptions.childExists("headers") Is Nothing) Then
        optOptions.add "headers"
    End If
    
    ' apply the access token/ basic auth if there is one
    If (Not isUndefined(optAccessToken)) Then
        optOptions.child("headers").add "authorization", _
            conditionalAssignment(optBasic, "Basic ", "Bearer ") & optAccessToken
    End If
    

    ' do the operation - we're using server http .. better for cors
    Dim ohttp As MSXML2.ServerXMLHTTP60
    Set ohttp = New MSXML2.ServerXMLHTTP60
    
    With ohttp
        ' this is for some MS bug .. cant remember which now
        .setOption 2, .getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID
        
        ' set it up
        .Open optMethod, url, False
        
        ' set the headers
        For Each job In optOptions.child("headers").children
            .setRequestHeader job.key, job.value
        Next job
        
        ' execute the thing
        .send optPayload
        
    End With
    
    ' turn results into standard
    Set urlExecute = makeResults(ohttp, url)
    
End Function
'/**
'* this is a standard result object to simply error checking etc.
'* @param {HTTPResponse} response the response from UrlFetchApp
'/ @param {string} optUrl the url if given
'* @return {object} the result object
'*/
Private Function makeResults(response As Object, Optional optUrl As String = vbNullString)
    Dim result As cJobject, job As cJobject, _
        rx As RegExp, matches As MatchCollection, match As match, i
    
    ' default result
    Set result = JSONParse("{" & _
        "'success':false," & _
        "'data':null," & _
        "'code':null," & _
        "'url':'" & optUrl & "'," & _
        "'extended':'failed to parse'," & _
        "'parsed':false }")

'   // process the result
    If (Not isUndefined(response)) Then
        result.add "code", response.Status
        result.add "headers"
        result.add "content", response.responseText
        result.add "success", (result.cValue("code") = 200 Or result.cValue("code") = 201)

        ' parse
        Set job = JSONParse(result.cValue("content"), , False)
        If (Not isUndefined(job) And job.isValid) Then
            result.child("data").setValue job
            result.add "parsed", True
        End If
        
        ' headers - MS doesnt do this for you
        Set rx = New RegExp
        With rx
            .MultiLine = True
            .Global = True
            .pattern = "^([^:]+)\s*:\s*(.+)*"
            Set matches = .Execute(response.getAllResponseHeaders())
        End With
        For i = 0 To matches.Count - 1
            result.child("headers").add _
                rxReplace("^\s*", matches.Item(i).SubMatches(0), ""), _
                rxReplace("\s*$", matches.Item(i).SubMatches(1), "")
        Next i
    End If
    Set makeResults = result

End Function
