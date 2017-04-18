# VBA Project: **VbaGit**
## VBA Module: **[cVbaGit](/libraries/cVbaGit.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVbaGit

---
VBA Procedure: **getEnums**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getEnums() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **throw**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub throw(message As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
message|String|False||


---
VBA Procedure: **getTokenFromBasic**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getTokenFromBasic(basicHash As String, clientHash As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
basicHash|String|False||
clientHash|String|False||


---
VBA Procedure: **isAccessToken**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isAccessToken() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **setAccessToken**  
Type: **Function**  
Returns: **[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setAccessToken(basic As String, client As String) As cVbaGit*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
basic|String|False||
client|String|False||


---
VBA Procedure: **getMyRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **standard result object**  
Scope: **Public**  
Description: **get all my repos**  

*Public Function getMyRepos() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **getSpecificRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getSpecificRepo(owner As String, repoName As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
owner|String|False||
repoName|String|False||


---
VBA Procedure: **getFileByPath**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **standard result object**  
Scope: **Public**  
Description: **get a file by path and repo**  

*Public Function getFileByPath(path As String, repoObject As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
path|String|False||a path
repoObject|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||a repo


---
VBA Procedure: **getUnpaged**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **standard result object**  
Scope: **Public**  
Description: **get intercept to deal with pagination**  

*Public Function getUnpaged(url As String, accessToken As String, options As cJobject, Optional data As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||'   * @param {string} accessToken
accessToken|String|False||
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||'   * @param {Array.object} data so far
data|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **createRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **standard result object**  
Scope: **Public**  
Description: **create a repo**  

*Public Function createRepo(name As String, Optional optOptions As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||repo name
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|any additional options


---
VBA Procedure: **contentOptions**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **options**  
Scope: **Private**  
Description: **special options for the api**  

*Private Function contentOptions() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **apiOptions**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **options**  
Scope: **Private**  
Description: **special options for the api**  

*Private Function apiOptions(Optional optOptions As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
optOptions|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **apiBase**  
Type: **Function**  
Returns: **Variant**  
Return description: **the api base url**  
Scope: **Private**  
Description: **function the api base url**  

*Private Function apiBase()*  

**no arguments required for this procedure**


---
VBA Procedure: **commitFile**  
Type: **Function**  
Returns: **Variant**  
Return description: **standard result**  
Scope: **Public**  
Description: **commit a file**  

*Public Function commitFile(path As String, repoObject As cJobject, message As String, content As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
path|String|False||the file path
repoObject|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
message|String|False||a committ message
content|String|False||some content


---
VBA Procedure: **tearDown**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
