# VBA Project: **VbaGit**
## VBA Module: **[VbaGit](/scripts/VbaGit.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:59 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in VbaGit

---
VBA Procedure: **doEverything**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: **example of exporting/importing a repos from github**  

*Public Sub doEverything()*  

**no arguments required for this procedure**


---
VBA Procedure: **doTheImport**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: **example of importing a repo from github and replaces the code in the companion wokbook**  

*Public Sub doTheImport()*  

**no arguments required for this procedure**


---
VBA Procedure: **deleteThisAfterRunningOnce**  
Type: **Function**  
Returns: **Variant**  
Return description: **the settings**  
Scope: **Private**  
Description: **sets up your credentials in the windows registry.**  

*Private Function deleteThisAfterRunningOnce()*  

**no arguments required for this procedure**


---
VBA Procedure: **getVGSettings**  
Type: **Function**  
Returns: **Variant**  
Return description: **the settings**  
Scope: **Public**  
Description: **sets up the settings object if its not already set up and returns it**  

*Public Function getVGSettings(Optional force As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
force|Boolean|True||whether to force a new set up


---
VBA Procedure: **doImportFromGit**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: **do the import from github and replace the modules in the companion workbook**  

*Public Sub doImportFromGit(repoName As String, Optional projectName As String = vbNullString, Optional applyExcelReferences As Boolean = False)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
repoName|String|False||the github reponame
projectName|String|True| vbNullString|the vbaproject name
applyExcelReferences|Boolean|True| False|whether to apply the excel references in dependency list


---
VBA Procedure: **getCodeFromGit**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: **get the code from git for a particular module**  

*Private Sub getCodeFromGit(project As cJobject, git As cVbaGit, folder As String, info As cJobject, childName As String, repo As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the project object
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||a handle to the cVbaGit object
folder|String|False||the folder to find the file in
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
childName|String|False||the branch of the project to work from (scripts/libraries)
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the repo object containing this file


---
VBA Procedure: **replaceModule**  
Type: **Function**  
Returns: **Boolean**  
Return description: **whether it was successful**  
Scope: **Private**  
Description: **get the code from git for a particular module**  

*Private Function replaceModule(project As cJobject, infoItem As cJobject, code As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the project object
infoItem|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the object from info.json for this file
code|String|False||the new code to use


---
VBA Procedure: **getRepo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the repo object**  
Scope: **Private**  
Description: **get the code from git for a particular module**  

*Private Function getRepo(git As cVbaGit, repoName As String, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||a cVbaGit handle
repoName|String|False||the name of the repo
complain|Boolean|True| True|whether to complain on failure


---
VBA Procedure: **doExtraction**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: **extract the files for a particular project and write them to the staging area**  

*Private Sub doExtraction(repoName As String, Optional optListOfModules As String = vbNullString, Optional projectName As String = vbNullString)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
repoName|String|False||the name of the repo
optListOfModules|String|True| vbNullString|list of main modules to use as starting point
projectName|String|True| vbNullString|the name of the vba project


---
VBA Procedure: **testmodulestuff**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub testmodulestuff()*  

**no arguments required for this procedure**


---
VBA Procedure: **doGit**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: **call this to commit all extracted projects to github them from the staging area**  

*Private Sub doGit(Optional specificRepoName As String = vbNullString)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
specificRepoName|String|True| vbNullString|the name of the repo - if blank it will do them all


---
VBA Procedure: **getAllTheRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **all the known repos**  
Scope: **Private**  
Description: **get all known repos belonging to the git logged in individual**  

*Private Function getAllTheRepos(git As cVbaGit) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||a handle to the cVbaGit api


---
VBA Procedure: **createRepos**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **all the known repos updated**  
Scope: **Private**  
Description: **create any repos in our list of info objects that don't exist**  

*Private Function createRepos(git As cVbaGit, infos As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||a handle to the cVbaGit api
infos|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||a list of info objects


---
VBA Procedure: **writeTheSource**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function writeTheSource(git As cVbaGit, kids As Collection, folderName As String, repo As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||
kids|Collection|False||
folderName|String|False||
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **writeTheFiles**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function writeTheFiles(git As cVbaGit, fileId As String, fileName As String, repo As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
git|[cVbaGit](/libraries/cVbaGit_cls.md "cVbaGit")|False||
fileId|String|False||
fileName|String|False||
repo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getAllInfoFiles**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getAllInfoFiles(Optional specificRepoName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
specificRepoName|String|True| vbNullString|


---
VBA Procedure: **writeInfoFile**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function writeInfoFile(project As cJobject, infoJob As cJobject, Optional cross As cJobject = Nothing, Optional dependencyList As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
cross|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **writeToStagingArea**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function writeToStagingArea(infoJob As cJobject, dependencyList As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getDependencyList**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getDependencyList(project As cJobject, name As String, Optional optListOfModules As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
name|String|False||
optListOfModules|String|True| vbNullString|


---
VBA Procedure: **findProc**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function findProc(procs As Collection, targetName As String) As cVBAProcedure*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
procs|Collection|False||
targetName|String|False||


---
VBA Procedure: **dependencyResolve**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dependencyResolve(modules As cJobject, dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
modules|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getPosProc**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: **get the pos object the the procedure that provoked ths dependency**  

*Private Function getPosProc(pos As cJobject, matchOb As match) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pos|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the position object for all the code of this module
matchOb|match|False||the regex match that found this dependency


---
VBA Procedure: **makeCrossReferenceJob**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeCrossReferenceJob(dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **registerExcelReferences**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub registerExcelReferences(project As cJobject, references As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
references|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **registerExcelReference**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function registerExcelReference(project As cJobject, job As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **makeExcelReferences**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeExcelReferences(project As cVBAProject, addHere As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cVBAProject](/libraries/cVBAProject_cls.md "cVBAProject")|False||
addHere|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **isModuleObj**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isModuleObj(ob As Object) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ob|Object|False||


---
VBA Procedure: **getVbaAsJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getVbaAsJobject(Optional optProjectName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
optProjectName|String|True| vbNullString|


---
VBA Procedure: **blowProcedures**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function blowProcedures(module As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
module|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **blowArguments**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function blowArguments(pob As cVBAProcedure, argOb As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pob|[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")|False||
argOb|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getProjects**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getProjects(Optional optProjectName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
optProjectName|String|True| vbNullString|


---
VBA Procedure: **getProcList**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub getProcList(module As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
module|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getmoduleList**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getmoduleList(project As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **makeInfoFile**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeInfoFile(project As cJobject, dependencyList As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
dependencyList|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **modulesToInfo**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function modulesToInfo(moduleJob As cJobject, infoJob As cJobject, extract As String, folderName As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
moduleJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
infoJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
extract|String|False||
folderName|String|False||


---
VBA Procedure: **mdWrap**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function mdWrap()*  

**no arguments required for this procedure**


---
VBA Procedure: **makeCross**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeCross(cross As cJobject, info As cJobject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cross|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **makeReadMe**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeReadMe(info As cJobject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **makeDependency**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeDependency(project As cJobject, info As cJobject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
project|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **constructModLink**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function constructModLink(name As String, folder As String, fileName As String, hover As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||
folder|String|False||
fileName|String|False||
hover|String|False||


---
VBA Procedure: **makeArguments**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeArguments(modl As cVBAmodule, info As cJobject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
modl|[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")|False||
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **findModLink**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function findModLink(modlName As String, info As cJobject, Optional hover As String = vbNullString, Optional fn As String = "docsName") As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
modlName|String|False||
info|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
hover|String|True| vbNullString|
fn|String|True| "docsName"|


---
VBA Procedure: **getFromVbaGitRegistry**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getFromVbaGitRegistry(key) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|Variant|False||


---
VBA Procedure: **setVbaGitRegistry**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setVbaGitRegistry(key, value) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|Variant|False||
value|Variant|False||


---
VBA Procedure: **getGitBasicCredentials**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getGitBasicCredentials()*  

**no arguments required for this procedure**


---
VBA Procedure: **setGitBasicCredentials**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub setGitBasicCredentials(user As String, pass As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
user|String|False||
pass|String|False||


---
VBA Procedure: **setGitClientCredentials**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub setGitClientCredentials(clientId As String, clientSecret As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
clientId|String|False||
clientSecret|String|False||


---
VBA Procedure: **getGitClientCredentials**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getGitClientCredentials()*  

**no arguments required for this procedure**
