# VBA Project: **VbaGit**
## VBA Module: **[urlResult](/libraries/urlResult.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in urlResult

---
VBA Procedure: **urlGet**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **a standard response**  
Scope: **Public**  
Description: **since I use this all the time,may as well make it a library**  

*Public Function urlGet(url As String, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||the url
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|optional headers
optAccessToken|String|True||an optional access token
optBasic|Boolean|True| False|the access token is for basic auth


---
VBA Procedure: **urlPost**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **a standard response**  
Scope: **Public**  
Description: **execute a post**  

*Public Function urlPost(url As String, Optional optMethod As String = "POST", Optional optPayload As Variant, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||the url
optMethod|String|True| "POST"|the http method
optPayload|Variant|True||any payload
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|optional headers
optAccessToken|String|True||an optional access token
optBasic|Boolean|True| False|the access token is for basic auth


---
VBA Procedure: **urlExecute**  
Type: **Function**  
Returns: **Variant**  
Return description: **a standard response**  
Scope: **Private**  
Description: **execute a urlfetch**  

*Private Function urlExecute(url As String, Optional optMethod As String = "GET", Optional optPayload As String = vbNullString, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||the url
optMethod|String|True| "GET"|the http method
optPayload|String|True| vbNullString|any payload
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|optional headers
optAccessToken|String|True||an optional access token
optBasic|Boolean|True| False|the access token is for basic auth


---
VBA Procedure: **makeResults**  
Type: **Function**  
Returns: **Variant**  
Return description: **the result object**  
Scope: **Private**  
Description: **this is a standard result object to simply error checking etc.**  

*Private Function makeResults(response As Object, Optional optUrl As String = vbNullString)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
response|Object|False||the response from UrlFetchApp
optUrl|String|True| vbNullString|the url if given
