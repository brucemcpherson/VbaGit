# VBA Project: **VbaGit**
## VBA Module: **[VbaGit](/scripts/VbaGit.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in VbaGit

---
VBA Procedure: **doEverything**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub doEverything()*  

**no arguments required for this procedure**


---
VBA Procedure: **doTheImport**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub doTheImport()*  

**no arguments required for this procedure**


---
VBA Procedure: **deleteThisAfterRunningOnce**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function deleteThisAfterRunningOnce()*  

**no arguments required for this procedure**


---
VBA Procedure: **getVGSettings**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  

*Public Function getVGSettings(Optional force As Boolean)*  

*name*|*type*|*optional*|*default*
---|---|---|---
force|Boolean|True|


---
VBA Procedure: **doImportFromGit**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub doImportFromGit(repoName As String, Optional projectName As String = vbNullString, Optional applyExcelReferences As Boolean = False)*  

*name*|*type*|*optional*|*default*
---|---|---|---
repoName|String|False|
projectName|String|True| vbNullString
applyExcelReferences|Boolean|True| False


---
VBA Procedure: **getCodeFromGit**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function getCodeFromGit(project As cJobject, git As cVbaGit, folder As String, info As cJobject, childName As String, repo As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|
folder|String|False|
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
childName|String|False|
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **replaceModule**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function replaceModule(project As cJobject, infoItem As cJobject, code As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
infoItem|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
code|String|False|


---
VBA Procedure: **getRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getRepo(git As cVbaGit, repoName As String, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|
repoName|String|False|
complain|Boolean|True| True


---
VBA Procedure: **doExtraction**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub doExtraction(repoName As String, Optional optListOfModules As String = vbNullString, Optional projectName As String = vbNullString)*  

*name*|*type*|*optional*|*default*
---|---|---|---
repoName|String|False|
optListOfModules|String|True| vbNullString
projectName|String|True| vbNullString


---
VBA Procedure: **doGit**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function doGit(Optional specificRepoName As String = vbNullString)*  

*name*|*type*|*optional*|*default*
---|---|---|---
specificRepoName|String|True| vbNullString


---
VBA Procedure: **getAllTheRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getAllTheRepos(git As cVbaGit) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|


---
VBA Procedure: **createRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function createRepos(git As cVbaGit, infos As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|
infos|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **writeTheSource**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function writeTheSource(git As cVbaGit, kids As Collection, folderName As String, repo As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|
kids|Collection|False|
folderName|String|False|
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **writeTheFiles**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function writeTheFiles(git As cVbaGit, fileId As String, fileName As String, repo As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False|
fileId|String|False|
fileName|String|False|
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **getAllInfoFiles**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getAllInfoFiles(Optional specificRepoName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
specificRepoName|String|True| vbNullString


---
VBA Procedure: **writeInfoFile**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function writeInfoFile(project As cJobject, infoJob As cJobject, Optional cross As cJobject = Nothing, Optional dependencyList As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
cross|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing


---
VBA Procedure: **writeToStagingArea**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function writeToStagingArea(infoJob As cJobject, dependencyList As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **getDependencyList**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getDependencyList(project As cJobject, name As String, Optional optListOfModules As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
name|String|False|
optListOfModules|String|True| vbNullString


---
VBA Procedure: **findProc**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Scope: **Private**  

*Private Function findProc(procs As Collection, targetName As String) As cVBAProcedure*  

*name*|*type*|*optional*|*default*
---|---|---|---
procs|Collection|False|
targetName|String|False|


---
VBA Procedure: **dependencyResolve**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function dependencyResolve(modules As cJobject, dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
modules|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **makeCrossReferenceJob**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function makeCrossReferenceJob(dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **registerExcelReferences**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub registerExcelReferences(project As cJobject, references As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
references|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **registerExcelReference**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function registerExcelReference(project As cJobject, job As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **makeExcelReferences**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function makeExcelReferences(project As cVBAProject, addHere As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cVBAProject](/libraries/cVBAProject_cls.md "cVBAProject")|False|
addHere|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **isModuleObj**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isModuleObj(ob As Object) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
ob|Object|False|


---
VBA Procedure: **getVbaAsJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getVbaAsJobject(Optional optProjectName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
optProjectName|String|True| vbNullString


---
VBA Procedure: **blowProcedures**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function blowProcedures(module As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
module|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **blowArguments**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function blowArguments(pob As cVBAProcedure, argOb As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
pob|[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")|False|
argOb|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **getProjects**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getProjects(Optional optProjectName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
optProjectName|String|True| vbNullString


---
VBA Procedure: **getProcList**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub getProcList(module As cJobject)*  

*name*|*type*|*optional*|*default*
---|---|---|---
module|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **getmoduleList**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function getmoduleList(project As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **makeInfoFile**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function makeInfoFile(project As cJobject, dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **modulesToInfo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function modulesToInfo(moduleJob As cJobject, infoJob As cJobject, extract As String, folderName As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
moduleJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
extract|String|False|
folderName|String|False|


---
VBA Procedure: **mdWrap**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function mdWrap()*  

**no arguments required for this procedure**


---
VBA Procedure: **makeCross**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function makeCross(cross As cJobject, info As cJobject) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
cross|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **makeReadMe**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function makeReadMe(info As cJobject) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **makeDependency**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function makeDependency(project As cJobject, info As cJobject) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **constructModLink**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function constructModLink(name As String, folder As String, fileName As String, hover As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
name|String|False|
folder|String|False|
fileName|String|False|
hover|String|False|


---
VBA Procedure: **makeArguments**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function makeArguments(modl As cVBAmodule, info As cJobject) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
modl|[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")|False|
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **findModLink**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function findModLink(modlName As String, info As cJobject, Optional hover As String = vbNullString, Optional fn As String = "docsName") As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
modlName|String|False|
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
hover|String|True| vbNullString
fn|String|True| "docsName"


---
VBA Procedure: **getFromVbaGitRegistry**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function getFromVbaGitRegistry(key) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
key|Variant|False|


---
VBA Procedure: **setVbaGitRegistry**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function setVbaGitRegistry(key, value) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
key|Variant|False|
value|Variant|False|


---
VBA Procedure: **getGitBasicCredentials**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function getGitBasicCredentials()*  

**no arguments required for this procedure**


---
VBA Procedure: **setGitBasicCredentials**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub setGitBasicCredentials(user As String, pass As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
user|String|False|
pass|String|False|


---
VBA Procedure: **setGitClientCredentials**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub setGitClientCredentials(clientId As String, clientSecret As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
clientId|String|False|
clientSecret|String|False|


---
VBA Procedure: **getGitClientCredentials**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  

*Private Function getGitClientCredentials()*  

**no arguments required for this procedure**
