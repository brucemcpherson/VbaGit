Option Explicit
' translation of GasGit - https://github.com/brucemcpherson/gasGit/blob/master/scripts/GasGit.js
'
Dim pEnums As cJobject
Dim pAccessToken As String

Public Function getEnums() As cJobject
    Set getEnums = pEnums
End Function
Private Sub throw(message As String)
    Err.Raise vbObjectError + 2000, , message
End Sub
Private Function getTokenFromBasic(basicHash As String, clientHash As String) As String
    'this is the flow for non web apps as described in https://developer.github.com/v3/oauth_authorizations/#create-a-new-authorization
    ' decode the basic hash into secret and id
    Dim a As Variant, b As Variant, result As cJobject, i As Long, p As cJobject
    
    a = Split(Base64Decode(clientHash), ":")
    If (arrayLength(a) <> 2) Then
        throw "basichash was invalid-" + basicHash
    End If

    Set p = New cJobject
    With p.init(Nothing)
        b = Split(getVGSettings().toString("GIT.SCOPES"), ",")
        With .add("scopes").addArray
            For i = LBound(b) To UBound(b)
                .add , CStr(b(i))
            Next i
        End With
        .add "note", "vbagit get token"
        .add "client_id", CStr(a(LBound(a)))
        .add "client_secret", CStr(a(UBound(a)))
    End With
    
    'we'll create a post request to get a token
    'POST /authorizations
    Set result = _
        urlPost( _
            apiBase() & "/authorizations", _
            "POST" _
            , p _
            , apiOptions() _
            , basicHash _
            , True)
        
    If (Not result.cValue("success")) Then
        throw "unable to get oauthtoken-" & result.stringify(True)
    End If
    getTokenFromBasic = result.getObject("data").toString("token")
End Function
Public Function isAccessToken() As Boolean
    isAccessToken = pAccessToken <> vbNullString
End Function
Public Function setAccessToken(basic As String, client As String) As cVbaGit
    If (isUndefined(basic) Or isUndefined(client)) Then
        Debug.Print "no access token provided"
    Else
        ' convert basic hash for an access token
        pAccessToken = getTokenFromBasic(basic, client)
        Set setAccessToken = Me
    End If
End Function
'  /**
'   * get all my repos
'   * @return {object} standard result object
'   */
Public Function getMyRepos() As cJobject
    Set getMyRepos = getUnpaged(apiBase() + "/user/repos", pAccessToken, apiOptions())
End Function

Public Function getSpecificRepo(owner As String, repoName As String) As cJobject
    Set getSpecificRepo = urlGet(apiBase() & "/repos" & "/" & owner & "/" & repoName, apiOptions(), pAccessToken)
End Function

'  /**
'   * get a file by path and repo
'   * @param {string} path a path
'   * @param {object} repoObject a repo
'   * @return {object} standard result object
'   */
Public Function getFileByPath(path As String, repoObject As cJobject) As cJobject
    Dim result As cJobject
    
    ' gets a file by path
    Set result = _
        urlGet( _
            Replace(repoObject.toString("contents_url"), "{+path}", path) _
            , contentOptions() _
            , pAccessToken)
    
'    // do the base64 connversion - replaces content with decoded
    If (result.cValue("success")) Then
        result.add "content", Base64Decode(result.getObject("data").toString("content"))
    End If
    
    Set getFileByPath = result
    
End Function
'  /**
'   * get intercept to deal with pagination
'   * @param {string} url
'   * @param {string} accessToken
'   * @param {object} options
'   * @param {Array.object} data so far
'   * @return {object} standard result object
'   */
Public Function getUnpaged(url As String, accessToken As String, _
                options As cJobject, Optional data As cJobject = Nothing) As cJobject
    
    Dim result As cJobject, h As String, rx As RegExp, _
        matches As MatchCollection, e As cJobject, newUrl As String
    ' first time in data will be empty
    If data Is Nothing Then
        Set data = New cJobject
        data.init(Nothing).addArray
    End If
    
    ' get next batch
    Set result = urlGet(url, options, accessToken)
'      // need to recurse for multiple pages
    If (result.cValue("success")) Then
        data.add , result.getObject("data")
        
        h = result.toString("headers.Link")
        If (h <> vbNullString) Then
            Set rx = New RegExp
            rx.pattern = "<([^>]*)>;\s?rel=""next"""
            
            Set matches = rx.Execute(h)
            If (matches.Count > 0) Then
                ' go round again
                newUrl = CStr(matches.Item(0).SubMatches(0))
                getUnpaged newUrl, accessToken, options, data
            End If
        End If
    End If
   
   ' substitute the accumulated data
    result.child("data").setValue data.mergeArrayOfJobs
    
    Set getUnpaged = result
End Function
'  /**
'   * create a repo
'   * @param {string} name repo name
'   * @param {object} optOptions any additional options
'   * @return {object} standard result object
'   */
Public Function createRepo(name As String, Optional optOptions As cJobject = Nothing) As cJobject
    Dim payload As cJobject, o As String
    If Not isUndefined(optOptions) Then o = optOptions.stringify
    Set payload = optionsExtend( _
        "{'name':'" & name & "'," & _
        "'description':'" & name & " created by VbaGit automation'," & _
        "'homepage': 'http://www.mcpher.com'," & _
        "'private': false," & _
        "'has_issues': true, " & _
        "'has_wiki': true, " & _
        "'has_downloads': true }", o)
    
    '// create it
    Set createRepo = _
        urlPost( _
            apiBase() & "/user/repos" _
            , _
            , payload _
            , apiOptions() _
            , pAccessToken)
            
End Function
' /**
'  * special options for the api
'  * @return {object} options
'  */
Private Function contentOptions() As cJobject
    Dim options As cJobject

    Set options = apiOptions()
    options.child("headers").add "Content-Type", getEnums().toString("MIMES.CONTENT")


    Set contentOptions = options
End Function

'/**
'  * special options for the api
'  * @return {object} options
'  */
Private Function apiOptions(Optional optOptions As cJobject = Nothing) As cJobject
    Dim ob As cJobject

    Set ob = applyDefaults(optOptions, New cJobject)

    With ob.init(Nothing)
        With .add("headers")
            .add "Accept", getEnums().toString("MIMES.API")
            .add "User-Agent", getVGSettings().toString("GIT.USERAGENT")
        End With
    End With
    Set apiOptions = ob
End Function
' /**
'  * function the api base url
'  * @return {string} the api base url
'  */
Private Function apiBase()
    apiBase = "https://api.github.com"
End Function
'  /**
'   * commit a file
'   * @param {string} path the file path
'   * @param {object} repo the repo object
'   * @param {string} message a committ message
'   * @param {string} content some content
'   * @return {object} standard result
'   */
Public Function commitFile(path As String, repoObject As cJobject, message As String, content As String)
    Dim payload As cJobject, job As cJobject, f As cJobject
    Set payload = New cJobject

    With payload.init(Nothing)
        
        ' we need to see if it exists and get its sha
        Set f = getFileByPath(path, repoObject)
        If (f.cValue("success")) Then
            .add "sha", f.getObject("data").toString("sha")
        End If
        f.tearDown
        
        .add "message", message
        With .add("committer")
            For Each job In getVGSettings().kids("GIT.COMMITTER")
                .add job.key, job.value
            Next job
        End With
        getVGSettings().toString ("GIT.COMMITTER")
        .add "content", Base64Encode(content)
    End With

'    // writes a text files to a repo
    Set commitFile = _
        urlPost( _
            Replace(repoObject.toString("contents_url"), "{+path}", path), _
            "PUT", _
            payload, _
            contentOptions(), _
            pAccessToken)
    
    payload.tearDown
    
End Function


Public Function tearDown()
    pEnums.tearDown
End Function

Private Sub Class_Initialize()
    Set pEnums = New cJobject
    With pEnums.init(Nothing)
        With .add("MIMES")
            .add "API", "application/vnd.github.v3+json"
            .add "CONTENT", "application/vnd.github.VERSION.raw"
            '' i think this is the correct one.. todo.. figure it out...
            .add "CONTENT", "application/json"
 
        End With
    End With
    pAccessToken = vbNullString

End Sub
