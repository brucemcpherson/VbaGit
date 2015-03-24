# VBA Project: **VbaGit**
## VBA Module: **[cVbaGit](/libraries/cVbaGit.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVbaGit

---
VBA Procedure: **getEnums**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function getEnums() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **throw**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub throw(message As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
message|String|False|


---
VBA Procedure: **getTokenFromBasic**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function getTokenFromBasic(basicHash As String, clientHash As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
basicHash|String|False|
clientHash|String|False|


---
VBA Procedure: **setAccessToken**  
Type: **Function**  
Returns: **[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")**  
Scope: **Public**  

*Public Function setAccessToken(basic As String, client As String) As cVbaGit*  

*name*|*type*|*optional*|*default*
---|---|---|---
basic|String|False|
client|String|False|


---
VBA Procedure: **getMyRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function getMyRepos() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **getSpecificRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function getSpecificRepo(owner As String, repoName As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
owner|String|False|
repoName|String|False|


---
VBA Procedure: **getFileByPath**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function getFileByPath(path As String, repoObject As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
path|String|False|
repoObject|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **getUnpaged**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function getUnpaged(url As String, accessToken As String, options As cJobject, Optional data As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
url|String|False|
accessToken|String|False|
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
data|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing


---
VBA Procedure: **createRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function createRepo(name As String, Optional optOptions As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
name|String|False|
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing


---
VBA Procedure: **contentOptions**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function contentOptions() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **apiOptions**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function apiOptions(Optional optOptions As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing


---
VBA Procedure: **apiBase**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function apiBase()*  

**no arguments required for this procedure**


---
VBA Procedure: **commitFile**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function commitFile(path As String, repoObject As cJobject, message As String, content As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
path|String|False|
repoObject|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
message|String|False|
content|String|False|


---
VBA Procedure: **tearDown**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
