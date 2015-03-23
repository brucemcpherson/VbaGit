# VBA Project: **VbaGit**
## VBA Module: **[urlResult](/libraries/urlResult.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 23/03/2015 10:33:26 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in urlResult

---
VBA Procedure: **urlGet**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function urlGet(url As String, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
url|String|False|
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing
optAccessToken|String|True|
optBasic|Boolean|True| False


---
VBA Procedure: **urlPost**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function urlPost(url As String, Optional optMethod As String = "POST", Optional optPayload As Variant, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
url|String|False|
optMethod|String|True| "POST"
optPayload|Variant|True|
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing
optAccessToken|String|True|
optBasic|Boolean|True| False


---
VBA Procedure: **urlExecute**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function urlExecute(url As String, Optional optMethod As String = "GET", Optional optPayload As String = vbNullString, Optional optOptions As cJobject = Nothing, Optional optAccessToken As String, Optional optBasic As Boolean = False)*  

*name*|*type*|*optional*|*default*
---|---|---|---
url|String|False|
optMethod|String|True| "GET"
optPayload|String|True| vbNullString
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing
optAccessToken|String|True|
optBasic|Boolean|True| False


---
VBA Procedure: **makeResults**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function makeResults(response As Object, Optional optUrl As String = vbNullString)*  

*name*|*type*|*optional*|*default*
---|---|---|---
response|Object|False|
optUrl|String|True| vbNullString
