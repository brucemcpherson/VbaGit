Option Explicit
' this is based on the ideas from http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready
' and is about getting your excel code to github
' VbaGit v0.2.3

' settings are in public var
Dim VGSettings As cJobject
'/**
' * example of exporting/importing a repos from github
' */
Public Sub doEverything()
    doExtraction "VbaGit", "VbaGit"
    doGit "VbaGit"
    ' these are the projects in this workbook i want to separate
   'doExtraction "cJobject", "cJobject"
   'doGit "cJobject"

   'doExtraction "vanillacJobject", "advancedcJobject"
    'doExtraction "emptycDataSet", "googleSheets,googleWireExample,oauthexamples,restLibrary"
    'doGit "emptycDataSet"
'
'    ' utilities
'    doExtraction "excelClassSerializer", "classSerializer"
'
'    ' example projects
    'doExtraction "excelRestLibraryExamples", "restLibraryExamples"
'     doExtraction "excelRoadmapper", "doRoadmapper"
'    doExtraction "excelGoogleSheets", "googleSheets,googleWireExample"
'    doExtraction "excelColor", "heatmapExamples,colorizing"
'    doExtraction "excelD3", "D3"
'    doExtraction "excelOauth2", "oAuthExamples"
'    doExtraction "excelParseCom", "parseCom"
'    doExtraction "excelProgressBar", "TestProgressBar"

 '   doExtraction "cChromeTraceVBA", "cChromeTraceVBA,testChromeTrace"
    'doExtraction "cVBAProject", "cVBAProject,cVBAProcedure,cVBAmodule,cVBAArgument"
    'doExtraction "cDataSet", "cDataSet"
    'doExtraction "excelRestLibrary", "restLibrary,cRest"
    ' now write them to git
    'doGit "emptycDataSet"
    'doExtraction "effex-demo-markers-excel", "effexTests,VBAMapsEffex"
    'doGit "effex-demo-markers-excel"
    
End Sub
'/**
' * example of importing a repo from github and replaces the code in the companion wokbook
' */
Public Sub doTheImport()
    ' this is the something I want to import into the companion workbook
    '//doImportFromGit "cDataSet"
    ''doImportFromGit "cJobject"
    
End Sub

' NOTES ON IMPORTING CODE FROM GITHUB
'
' since this is for public code,
' there is no need to register or create a github app (unlike for writing code)
' although if you do have credentials you can apply them if you want
'
' 1 .open the vbagit worksheet and the sheet to which you want to import the code
' 2 .set up do the import
'       first argument is the repoName to get the code from
'       second is the project name - you should probably leave that blank
'       third is whether to apply the excel references - the defaut is not to
' 3. run it and compile the result
' ----------------------------------------
' NOTES ON COMMITTING CODE TO GITHUB

' open the vbagit workbook and set up for your environment if you havent already done it
'
' 1. set your credentials in deleteThisAfterRunningOnce,
'       then either delete or obscure your credentials, you wont need to run this again
'       you need to have set up an app in github and got the api credentials

' 2. set getVGSettings for your environment
'       this describes who you are and where to do things
'       the main ones are
'           EXTRACT.TO (a staging area folder for all your files to be dumped)
'           GIT.COMMITTER
'           GIT.USERAGENT
'       you probably don't need to change any others unless you want to rename or reorganize the repo contents

' vbagit runs as a companion workbook
' you open vbagit
' you open the workbook containing the modules you want to commit
' switch to the vbagit workbook and set up what you want to do
'
' 3. run doextraction for each repo you want to create, naming the module(s) for the repo
'       you can specify a project name if you have more than 2 workbooks open (vbagit + the one containing the modules to be processed)
'       however most people leave their projects with the default name of VBAProject.
'       It will take the first one it finds with the given name - so its best just to have the right workbook open
'       the first arg wil be the reponame - this what it will be called on GIThub
'       the second is the list of modules that are to be part of this repo
'       you just need to name the main module(s). These will end up in the scripts folder
'       Any dependent modules/classes will be automatically detected and added to the libraries folder
'       Module Documentation is also automatically generated at this stage
'       Note that you can create multiple repos from a single workbook
'       see doEverything for an example
'       any shared libraries needed will be detected and committed to whichever repo as required

' 4. when ready you can run doGit to commit everything in the EXTRACT.TO folder up to GIT
'       actually, its the contents of info.json that decides what to copy up. anyting in EXTRACT.TO not created by
'       vbagit will be ignored
'       all repos will be commited at once, except if you specify a repoName to doGit. then it will only do one
'       your readme will only be committed if there is not already one in the Github repo
'       if you prefer to use the git client instead, you can make your EXTRACT.TO location your local git repo

'/**
' * sets up your credentials in the windows registry.
' * should be deleted or obfuscated after running once
' */
Private Function deleteThisAfterRunningOnce()
    ' substitute your git application clientid/secret
    setGitBasicCredentials "username git", "passwrod git"
    setGitClientCredentials "short", "long"
End Function
'/**
' * sets up the settings object if its not already set up and returns it
' * @param {boolean} force whether to force a new set up
' * @return {cJobject} the settings
' */
Public Function getVGSettings(Optional force As Boolean)
    
    ' get the settings - only bothers with the parse once
    If force Or isUndefined(VGSettings) Then
        If (isSomething(VGSettings)) Then
            VGSettings.tearDown
        End If
        Set VGSettings = New cJobject
        With VGSettings.init(Nothing)
            With .add("EXTRACT")
                .add "TO", "c:/users/bruce/documents/gas/Extraction/Scripts/"
            End With
            With .add("GIT")
                With .add("COMMITTER")
                    .add "name", "Bruce McPherson"
                    .add "email", "bruce@mcpher.com"
                End With
                .add "USERAGENT", "brucemcpherson"
                .add "SCOPES", "repo,gist"
                .add "OWNER", .toString("USERAGENT")
            End With
            With .add("REGISTRY")
                .add "root", "xLiberation"
                .add "app", "vbagit"
                .add "basic", "basichash"
                .add "client", "clienthash"
            End With
            With .add("APP")
                .add "VERSION", "0.2.4"
            End With
            With .add("FILES")
                .add "README", "README.md"
                .add "INFO", "info.json"
                .add "DEPENDENCIES", "dependencies.md"
                .add "CROSS", "cross.md"
            End With
            With .add("FOLDERS")
                .add "SCRIPTS", "scripts"
                .add "DEPENDENCIES", "libraries"
            End With
            With .add("PROJECT")
                .add "NAME", "VbaGitAddOn"
            End With
            With .add("VBA")
                With .add("TYPES")
                    .add "StdModule", 1
                    .add "ClassModule", 2
                End With
            End With
        End With
    End If
    Set getVGSettings = VGSettings
End Function
'/**
' * do the import from github and replace the modules in the companion workbook
' * @param {} repoName the github reponame
' * @param {} projectName the vbaproject name
' * @param {} applyExcelReferences whether to apply the excel references in dependency list
' */
Public Sub doImportFromGit(repoName As String, _
    Optional projectName As String = vbNullString, _
    Optional applyExcelReferences As Boolean = False)
    
    ' get all the projects in this workbook
    Dim projects As cJobject, settings As cJobject, project As cJobject
    Dim repo As cJobject, git As cVbaGit, result As cJobject, _
        job As cJobject, info As cJobject
        
    Set settings = getVGSettings(True)
    Set projects = getVbaAsJobject(projectName)
    Set project = projects.children(1)
    
    ' create dependency list - only do the first project for this scope
    If (project.getObject("project").name = settings.toString("PROJECT.NAME")) Then
        MsgBox "you need to open both vbagit and the workbook to set up: you cannot overwrite vbagit"
        Exit Sub
    End If
    
    ' check we are doing the right thing
    If MsgBox("You are 100% sure that you want me to import the code from repo " & repoName & _
        " from github into " & project.getObject("project").wBook.name, vbYesNo) <> vbYes Then
        Exit Sub
    Else
        Debug.Print "Importing the code from repo " & repoName & _
        " from github into " & project.getObject("project").wBook.name
        
        ' get a handle for git api
        Set git = New cVbaGit
        
        ' actually we are using basic authentication
        ' this optional, if you have already setup a github app
        ' it will also get you more quota
        ' if not then you can comment this out
        git.setAccessToken getGitBasicCredentials(), getGitClientCredentials()
        If (Not git.isAccessToken) Then
            Debug.Print "you are using an unauthenticated git connection with limited quota"
        End If
        
        ' get the repo
        Set repo = getRepo(git, repoName).getObject("data")
        
        ' get the info file
        Set result = git.getFileByPath(settings.toString("FILES.INFO"), repo)
        
        ' process the scripts
        Set info = JSONParse(result.toString("content"))
        
        ' the modules
        getCodeFromGit project, git, _
            settings.toString("FOLDERS.SCRIPTS"), info, _
            "modules", repo
        
        ' the libraries
        getCodeFromGit project, git, _
            settings.toString("FOLDERS.DEPENDENCIES"), info, _
            "dependencies", repo
        
        If (applyExcelReferences) Then
            Debug.Print "Apply excel references"
            registerExcelReferences project, info.child("excelReferences")
        End If
        
        info.tearDown
        
    End If

End Sub
'/**
' * get the code from git for a particular module
' * @param {} project the project object
' * @param {} git a handle to the cVbaGit object
' * @param {} folder the folder to find the file in
' * @param {} childName the branch of the project to work from (scripts/libraries)
' * @param {} repo the repo object containing this file
' */
Private Sub getCodeFromGit(project As cJobject, git As cVbaGit, _
        folder As String, info As cJobject, _
        childName As String, repo As cJobject)
    
    Dim job As cJobject, result As cJobject
    Debug.Print "Importing project " & childName
    For Each job In info.kids(childName)
        Set result = git.getFileByPath(folder & "/" & job.toString("fileName"), repo)
        Debug.Assert result.cValue("success")
        replaceModule project, job, result.toString("content")
        result.tearDown
    Next job

End Sub
'/**
' * get the code from git for a particular module
' * @param {} project the project object
' * @param {} infoItem the object from info.json for this file
' * @param {} code the new code to use
' * @return {} whether it was successful
' */
Private Function replaceModule(project As cJobject, infoItem As cJobject, code As String) As Boolean

    Dim jm As cJobject, module As VBComponent, m As cJobject, _
        t As Long, settings As cJobject, vm As cVBAmodule
        
    Set settings = getVGSettings
    
    Set m = project.child("modules").findInArray("name", infoItem.toString("name"))
    
    If (m Is Nothing) Then
        ' we need to create a new module
        Debug.Print "creating "; infoItem.toString("type"); " "; infoItem.toString("name")
        Set module = project.getObject("project") _
            .theProject _
            .VBComponents _
            .add(settings.cValue("VBA.TYPES." & infoItem.toString("type")))
        module.name = infoItem.toString("name")
    Else
        Set vm = m.parent.getObject("module")
        Set module = vm.vCom
        If (vm.textKind <> infoItem.toString("type")) Then
            Debug.Print "module " & vm.name & " is not the same type as on github: cannot replace it"
            Exit Function
        End If
        ' delete existing content
        Debug.Print "replacing "; infoItem.toString("type"); " "; infoItem.toString("name")
        
    End If
    
    'clear current contents
    If module.codeModule.CountOfLines > 0 Then
        ' remove Option Explict lines if it was added automatically
        module.codeModule.DeleteLines 1, module.codeModule.CountOfLines
    End If
    
    ' add the new code
    module.codeModule.AddFromString code
    
    replaceModule = True
End Function
'/**
' * get the code from git for a particular module
' * @param {} git a cVbaGit handle
' * @param {} repoName the name of the repo
' * @param {} complain whether to complain on failure
' * @return {} the repo object
' */
Private Function getRepo(git As cVbaGit, repoName As String, Optional complain As Boolean = True) As cJobject
    Dim settings As cJobject, result As cJobject
    Set settings = getVGSettings()
    
    Set result = git.getSpecificRepo(settings.toString("GIT.OWNER"), repoName)
    If (complain And Not result.cValue("success")) Then
        Err.Raise vbObjectError + result.cValue("code"), result.stringify
    End If
    Set getRepo = result
End Function
'/**
' * extract the files for a particular project and write them to the staging area
' * @param {} repoName the name of the repo
' * @param {} optListOfModules list of main modules to use as starting point
' * @param {} projectName the name of the vba project
' */
Private Sub doExtraction(repoName As String, _
    Optional optListOfModules As String = vbNullString, _
    Optional projectName As String = vbNullString)

    ' get all the projects in this workbook
    Dim projects As cJobject, dependencyList As cJobject, _
        infoJob As cJobject, settings As cJobject, project As cJobject
        
    Set settings = getVGSettings(True)
    Set projects = getVbaAsJobject(projectName)
    
    ' create dependency list - only do the first project for this scope
    If (projects.hasChildren) Then
        Set project = projects.children(1)
        Set dependencyList = getDependencyList(project, repoName, optListOfModules)
    Else
        ' no projects
        MsgBox "no projects detected for repo " & repoName
        Exit Sub
    End If
    
    Debug.Print "extracting repo "; repoName; " from project "; project.getObject("project").theProject.name
      
    ' create the info & cross references file
    Set infoJob = makeInfoFile(project, dependencyList)
    
    ' now we write all the scripts to some staging area
    writeToStagingArea infoJob, dependencyList
    
    ' mark as extracted
    infoJob.child("extracted").setValue True
    writeInfoFile project, infoJob, makeCrossReferenceJob(dependencyList), dependencyList


    ' clean up
    projects.tearDown
    dependencyList.tearDown
    infoJob.tearDown
    settings.tearDown
    Debug.Print "done extracting "; repoName
End Sub
Private Sub testmodulestuff()
    Dim job As cJobject, settings As cJobject, projects As cJobject, proc As cVBAProcedure
    Dim r As Range, rx As RegExp
    Set r = Range("sheet1!a1")
    Set settings = getVGSettings(True)
    Set projects = getVbaAsJobject()
    
    For Each job In projects.children(1).kids("modules")
        With job.getObject("module")
            For Each proc In .procedures
                With proc
                    Set r = r.Offset(1)
                    r.Offset(, 0).value = .name
                    r.Offset(, 1).value = .startLine
                    r.Offset(, 2).value = .lineCount
                    r.Offset(, 3).value = .getTheEndRx.Test( _
                        .codeModule.Lines(.startLine, _
                            .getFinishWithoutTrailingComments - .startLine + 1))
                    r.Offset(, 4).value = .getTheCode
                    r.Offset(, 5).value = .getTheCodePlusLeadingComments
                End With
            Next proc
        End With
    Next job
    projects.tearDown
End Sub
' these are all about committing to Git
'--------------------------------------------------------
'
'/**
' * call this to commit all extracted projects to github them from the staging area
' * @param {} specificRepoName the name of the repo - if blank it will do them all
' */
Private Sub doGit(Optional specificRepoName As String = vbNullString)

    Dim allInfoFiles As cJobject, settings As cJobject, git As cVbaGit, _
        repos As cJobject, result As cJobject
    
    
    Set settings = getVGSettings(True)
    
    ' the info files drive what needs to be written to Git
    Set allInfoFiles = getAllInfoFiles(specificRepoName)
    
    ' get a handle for git api
    Set git = New cVbaGit
    
    ' actually we are using basic authentication
    git.setAccessToken getGitBasicCredentials(), getGitClientCredentials()
    If (Not git.isAccessToken) Then
        MsgBox ("you cannot commit to git without authentication- please set up")
    Else
    ' get all the repos, creating any missing ones, and adding all the files
        Set repos = createRepos(git, allInfoFiles)
    End If
    ' clean up
    allInfoFiles.tearDown
    repos.tearDown
    git.tearDown
    
End Sub
'/**
' * get all known repos belonging to the git logged in individual
' * @param {} git a handle to the cVbaGit api
' * @return {} all the known repos
' */
Private Function getAllTheRepos(git As cVbaGit) As cJobject
    Dim result As cJobject
    Set result = git.getMyRepos
    If (result.cValue("success")) Then
        Set getAllTheRepos = result.getObject("data")
    Else
        MsgBox "failed to get all the repos " + result.stringify
        Exit Function
    End If
End Function
'/**
' * create any repos in our list of info objects that don't exist
' * @param {} git a handle to the cVbaGit api
' * @param {} infos a list of info objects
' * @return {} all the known repos updated
' */
Private Function createRepos(git As cVbaGit, infos As cJobject) As cJobject

    Dim repos As cJobject, info As cJobject, repo As cJobject, result As cJobject, _
        added As Long, settings As cJobject, job As cJobject
    Set settings = getVGSettings
    ' all my repos
    Set repos = getAllTheRepos(git)
    added = 0
    
    ' find any missing
    For Each info In infos.children
        ' we'll only do uncommitted or modified since last commit
        If (info.cValue("committedDate") < info.cValue("modifieddate")) Then
            If (isSomething(repos)) Then
                Set repo = repos.findInArray("name", info.toString("repo"))
            Else
                Set repo = Nothing
            End If
            ' need to create it
            If (repo Is Nothing) Then
                Set result = git.createRepo(info.toString("repo"))
                If (Not result.cValue("success")) Then
                    MsgBox "error creating " & info.toString("repo") & "-" & _
                        result.stringify
                    Exit Function
                End If
                Debug.Print "created repo for " & info.toString("repo")
                added = added + 1
            End If
        End If
    Next info
    
    ' get them again
    If (added > 0) Then
        Set repos = getAllTheRepos(git)
        Debug.Print "added ", added, " repos"
    End If
    
    ' now add any missing readmes
    For Each info In infos.children
        Set repo = repos.findInArray("name", info.toString("repo")).parent
        
        ' check of thereis a readme and create one if noe
        Set result = git.getFileByPath(settings.toString("FILES.README"), repo)
        If (Not result.cValue("success")) Then
            Set result = writeTheFiles(git, info.toString("readmeFileId"), settings.toString("FILES.README"), repo)
        End If
        
        ' the dependencies file
        Set result = writeTheFiles(git, info.toString("dependenciesFileId"), settings.toString("FILES.DEPENDENCIES"), repo)
    
        ' the references file
        Set result = writeTheFiles(git, info.toString("crossFileId"), settings.toString("FILES.CROSS"), repo)
        
        ' the scripts
        writeTheSource git, info.kids("modules"), settings.toString("FOLDERS.SCRIPTS"), repo
        
        ' the libraries
        writeTheSource git, info.kids("dependencies"), settings.toString("FOLDERS.DEPENDENCIES"), repo
        
        ' the info file
        writeTheFiles git, info.toString("fileId"), info.toString("fileName"), repo

    Next info
    
    Set createRepos = repos

End Function
Private Function writeTheSource(git As cVbaGit, kids As Collection, _
                                    folderName As String, repo As cJobject)
    Dim job As cJobject
    For Each job In kids
        ' the source
        writeTheFiles git, _
            job.toString("id"), _
            folderName & "/" & job.toString("fileName"), _
            repo
        
        ' the docs
        writeTheFiles git, _
            job.toString("docsId"), _
            folderName & "/" & job.toString("docsName"), _
            repo
            
    Next job
    
End Function
Private Function writeTheFiles(git As cVbaGit, fileId As String, fileName As String, repo As cJobject) As cJobject
    Dim result As cJobject
    Debug.Print "committing " & fileName & " for " & repo.toString("name")
    
    Set result = git.commitFile(fileName, _
        repo, "created by vbagit", readFromFolderFile("", fileId))
        
    If (Not result.cValue("success")) Then
        MsgBox "error creating " & fileId & "-" & _
            result.stringify
    End If

    Set writeTheFiles = result
End Function
'
' these are all about reading and writing to EXTRACT.TO
'------------------------------------------------------
Private Function getAllInfoFiles(Optional specificRepoName As String = vbNullString) As cJobject

    ' get all info files in the area
    Dim infos As cJobject, settings As cJobject, s As String, a As Variant, _
        i As Long, info As cJobject
    Set infos = New cJobject
    infos.init(Nothing).addArray
    Set settings = getVGSettings()
    
    s = getAllSubFolderPaths(settings.toString("EXTRACT.TO"))
    If (s <> vbNullString) Then
        a = Split(s, ",")
        For i = LBound(a) To UBound(a)

            If fileExists(concatFolderName(CStr(a(i)), settings.toString("FILES.INFO"))) Then
                Set info = JSONParse( _
                    readFromFolderFile(CStr(a(i)), settings.toString("FILES.INFO")))
                If (specificRepoName = vbNullString Or _
                    compareAsKey(specificRepoName, info.toString("title"))) Then
                    infos.add.arrayAppend info
                Else
                    info.tearDown
                End If
            End If
            
        Next i
    End If
    
    If (specificRepoName <> vbNullString And infos.children.Count <> 1) Then
        MsgBox "didn't find repo info file for " & specificRepoName
    End If
    
    Set getAllInfoFiles = infos
End Function
Private Function writeInfoFile(project As cJobject, _
            infoJob As cJobject, _
            Optional cross As cJobject = Nothing, _
            Optional dependencyList As cJobject = Nothing) As cJobject
    Dim settings As cJobject
    Set settings = getVGSettings()
   
   ' make sure we have the directory structure set up
    checkOrCreateFolder infoJob.toString("extract")
    
    infoJob.child("modifiedDate").setValue getTimestampFromDate
    ' write it out
    writeToFolderFile "", _
        infoJob.toString("fileId"), _
        infoJob.stringify
        
    ' also need to write a readme file if there isnt one
    If (Not fileExists(infoJob.toString("readmeFileId"))) Then
        writeToFolderFile "", _
            infoJob.toString("readmeFileId"), _
            makeReadMe(infoJob)
    End If
    
    ' and a cross reference file
    If (isSomething(cross)) Then
        writeToFolderFile "", _
            infoJob.toString("crossFileId"), _
            makeCross(cross, infoJob)
    End If
    
    ' and a dependency reference file
    If (isSomething(dependencyList)) Then
        writeToFolderFile "", _
            infoJob.toString("dependenciesFileId"), _
            makeDependency(project, infoJob)
    End If


    Set writeInfoFile = infoJob

End Function
Private Function writeToStagingArea(infoJob As cJobject, dependencyList As cJobject)
    
    Dim job As cJobject, modl As cVBAmodule, code As String
    For Each job In infoJob.child("dependencies").children
        
        ' we don't write out dependencies that are already in the scripts list
        If (dependencyList.child("scripts") _
            .findInArray("name", job.toString("name")) Is Nothing) Then
            
            ' this wasnt, so its ok to go
            Set modl = dependencyList.child("dependencies") _
                .findInArray("name", job.toString("name")) _
                .parent _
                .getObject("module")
                
            If (modl.vCom.codeModule.CountOfLines > 0) Then
                code = modl.vCom.codeModule.Lines(1, modl.vCom.codeModule.CountOfLines)
            Else
                code = "'No code for this referenced module " & modl.name & vbCrLf & _
                    "'could be a problem if the reference was not for a built in excel function" & vbCrLf & _
                    "'check the cross reference md file"

                Debug.Print code
            End If
            writeToFolderFile job.toString("folder"), job.toString("fileName"), code
            
            writeToFolderFile job.toString("folder"), job.toString("docsName"), _
                makeArguments(modl, infoJob)
            
        End If
    Next job
    
    For Each job In infoJob.child("modules").children
                
        Set modl = dependencyList.child("scripts") _
            .findInArray("name", job.toString("name")) _
            .parent _
            .getObject("module")
        
        writeToFolderFile job.toString("folder"), job.toString("fileName"), _
            modl.vCom.codeModule.Lines(1, modl.vCom.codeModule.CountOfLines)

        writeToFolderFile job.toString("folder"), job.toString("docsName"), _
                makeArguments(modl, infoJob)
    Next job

End Function
'
' these are all about resolving dependencies
'-----------------------------------------------------
Private Function getDependencyList(project As cJobject, name As String, _
    Optional optListOfModules As String = vbNullString) As cJobject

    ' this is a shot at figureing out dependencies from a given list of modules
    Dim c As New cStringChunker, a As Variant, job As cJobject, dependencyList As cJobject, _
        deps As cJobject, mods As cJobject, i As Long, m As cJobject
        
    ' dependencies - all modules neededs, scripts - the ones asked for
    Set dependencyList = New cJobject
    With dependencyList.init(Nothing)
        .add "name", name
        Set deps = .add("dependencies").addArray
        Set mods = .add("scripts").addArray
    End With
    
    ' default is all modules
    If (optListOfModules = vbNullString) Then
        For Each job In project.child("modules").children
            c.add(job.toString("name")).add (",")
        Next job
        optListOfModules = c.chopIf(",").toString
    End If
    
    ' now check that they all exist
    a = Split(optListOfModules, ",")
    For i = LBound(a) To UBound(a)
        Set m = project.child("modules").findInArray("name", (CStr(a(i))))
        If (isSomething(m)) Then
            With deps.add
                .add "module", m.parent.getObject("module")
                .add "name", m.parent.getObject("module").name
            End With
            With mods.add
                .add "module", m.parent.getObject("module")
                .add "name", m.parent.getObject("module").name
            End With
        Else
            MsgBox "module doesnt exist " & CStr(a(i))
        End If
    Next i
    
    ' now we have to find modules referenced that are not in the dependecy list
    Set getDependencyList = dependencyResolve(project.child("modules"), dependencyList)
    
    
End Function
Private Function findProc(procs As Collection, targetName As String) As cVBAProcedure
    Dim proc As cVBAProcedure
    For Each proc In procs
        If (proc.name = targetName) Then
            Set findProc = proc
        End If
    Next proc
End Function
Private Function dependencyResolve(modules As cJobject, dependencyList As cJobject) As cJobject
    
    ' create a regex of all known modules that haven't yet been identified
    Dim c As cStringChunker, job As cJobject, _
        s As cStringChunker, d As cJobject, matchMod As cVBAmodule, proc As cVBAProcedure, _
        matches As MatchCollection, e As cJobject, pos As cJobject, _
        jo As cJobject, match As match, recurse As Boolean, procs As Collection, _
        code As String, pName As String, ob As Object, alreadyThere As cJobject, _
        warned As Boolean, posProc As cJobject

    Set c = New cStringChunker
    Set s = New cStringChunker

    recurse = False
    warned = False
    
    ' these are all the modules in the dependency list
    For Each job In dependencyList.child("dependencies").children
        ' get the local code
        s.clear
        Set e = New cJobject
        Set pos = New cJobject
        e.init Nothing
        pos.init(Nothing).addArray
        c.clear
        Set procs = job.getObject("module").procedures
        For Each proc In procs
            ' this is the position this code starts at - we'll need it later for finding where it came from
            
            With pos.add
                ' clean up the code getting rid of dims and continuations as well as th declaration
                code = getRidOfDims( _
                         getRidOfComments( _
                          getRidOfQuoted( _
                           Replace( _
                            straightenOutContinuations(proc.getTheCode), proc.declaration, ""))))
                
                ' remember where this code is stored
                .add "start", s.size
                .add "length", Len(code)
                .add "proc", proc
                
                ' push the code for searching
                s.add (code)
            End With
           
        Next proc
        
        ' now make a regex that describes all the other procs not in the
        ' dependency list already and not in this module
        For Each jo In modules.children
            Set matchMod = jo.getObject("module")
            If (matchMod.name <> job.toString("name")) Then
                Set d = dependencyList.child("dependencies").findInArray("name", matchMod.name)
                If (d Is Nothing) Then
                    If matchMod.textKind = "StdModule" Then
                        For Each proc In matchMod.procedures
                            ' but of course private procs will not be visible outside anyway
                            If proc.scope = "Public" And findProc(procs, proc.name) Is Nothing Then
                                e.add proc.name, proc
                                c.add(proc.name).add "|"
                            End If
                        Next proc
                    ElseIf matchMod.textKind = "ClassModule" Then
                        c.add(matchMod.name).add "|"
                        e.add matchMod.name, matchMod
                    Else
                        If (Not warned) Then
                            ' lets just tell this story one time
                            Debug.Print matchMod.textKind & " " & matchMod.name & " is skipped - only doing class and stdmodules"
                            warned = True
                        End If
                    End If
                    
                End If
            End If
        Next jo
        
        ' if we still have some to do, then kick off the matching
        If c.size > 0 Then

            'this will match all references to particular procedures
            Set matches = getRx("\b(" & c.chopIf("|").toString & ")\b(?!\s*=)").Execute(s.toString)
        
            ' we know which module they are in from the cross reference from earlier
            If (matches.Count > 0) Then
                ' add to the dependency list
                For Each match In matches
                    pName = CStr(match.SubMatches(0))
                   

                    ' find who referenced it by the position at which it appeared
                    Set posProc = getPosProc(pos, match)
                    
                    ' ob refers to the thing being called
                    Set ob = e.getObject(pName)
                    
                    ' if its a class then we dont deal in individual procs
                    If (isModuleObj(ob)) Then
                        Set matchMod = ob
                    Else
                        Set matchMod = ob.parent
                        ' but if this is a local argument to this function,
                        ' we dont need to look for references
                        ' so treat it like it was never a match
                        If (posProc.getObject("proc").isAnArgument(pName)) Then
                            Set matchMod = Nothing
                        End If
                    End If
                    
                    ' if we have found a proc/module decide if it needs to be added to depend list
                    If isSomething(matchMod) Then
                        Set alreadyThere = dependencyList.child("dependencies").findInArray("name", matchMod.name)
                        
                        ' first time we've seen it
                        If (alreadyThere Is Nothing) Then
                            Set d = dependencyList.child("dependencies").add
                            With d
                                .add "module", matchMod
                                .add "name", matchMod.name
                                If (.childExists("cross") Is Nothing) Then
                                    .add("cross").addArray
                                End If
                            End With
                            recurse = True
                        Else
                        
                        ' we already know it
                            Set d = alreadyThere.parent
                        End If
                    
                        ' record the cross reference event
                        If (alreadyThere Is Nothing Or _
                            d.child("cross").findInArray("name", _
                                    posProc.getObject("proc").name) Is Nothing) Then
                        ' first time we're seeing a reference to it so add to cross reference
                            With d.child("cross").add
                                .add "proc", ob
                                .add "by", posProc.getObject("proc")
                                .add "name", posProc.getObject("proc").name
                            End With
                        End If

                    
                    End If
                Next match
            End If
        End If
        e.tearDown
        pos.tearDown
    Next job

    ' weve added something so do it all over again
    If (recurse) Then
        dependencyResolve modules, dependencyList
    End If
    
    Set dependencyResolve = dependencyList
End Function
'/**
' * get the pos object the the procedure that provoked ths dependency
' * @param {} pos the position object for all the code of this module
' * @param {} matchOb the regex match that found this dependency
' * @return the pos object branch with the match
' */
Private Function getPosProc(pos As cJobject, matchOb As match) As cJobject
    Dim jo As cJobject
    For Each jo In pos.children
        If (matchOb.FirstIndex >= jo.cValue("start") And _
              matchOb.FirstIndex + matchOb.Length <= _
              jo.cValue("start") + jo.cValue("length")) Then
            ' this is who is referencing me
            Set getPosProc = jo
            Exit Function
        End If
    Next jo
    MsgBox ("failed to find who provoked this dependency " & matchOb)
    
End Function
Private Function makeCrossReferenceJob(dependencyList As cJobject) As cJobject
    Dim cross As cJobject, settings As cJobject, job As cJobject, jo As cJobject
    Set cross = New cJobject
    Set settings = getVGSettings()
    With cross.init(Nothing).addArray
        For Each job In dependencyList.child("dependencies").children
            If isSomething(job.childExists("cross")) Then
                For Each jo In job.children("cross").children
                    With .add
                        .add "proc", jo.getObject("proc")
                        .add "by", jo.getObject("by")
                        If (isModuleObj(jo.getObject("proc"))) Then
                            .add "sortKey", jo.getObject("proc").name
                        Else
                            .add "sortKey", jo.getObject("proc").parent.name & _
                            Space(50 - Len(jo.getObject("proc").parent.name)) & _
                            jo.getObject("proc").name
                        End If

                    End With
                Next jo
            End If
        Next job
    End With
    Set makeCrossReferenceJob = cross.sortByValue()
    
End Function
Private Sub registerExcelReferences(project As cJobject, references As cJobject)
    
    Dim job As cJobject
    
    For Each job In references.children
        registerExcelReference project, job
        
    Next job
    
End Sub
Private Function registerExcelReference(project As cJobject, job As cJobject)

    ' add a reference (if its not already there)
  
    Dim r As Reference ' Reference
    On Error GoTo handle
    With project.getObject("project").theProject
        For Each r In .references
            If (r.name = job.cValue("name")) Then
                If (r.major < job.cValue("major") Or _
                    (r.major = job.cValue("major") And _
                    r.minor < job.cValue("minor")) And Not r.BuiltIn) Then
                    .references.AddFromGuid job.cValue("guid"), _
                        job.cValue("major"), job.cValue("minor")
                    .references.remove (r)
                End If
                Exit Function
            End If
        Next r
    ' if we get here then we need to add it
      .references.AddFromGuid job.cValue("guid"), job.cValue("major"), job.cValue("minor")
      Exit Function
    End With
    
handle:
    MsgBox ("warning - tried and failed to add reference to " & _
        job.cValue("name") & "v" & job.cValue("major") _
        & "." & job.cValue("minor"))
    Exit Function
    
End Function
Private Function makeExcelReferences(project As cVBAProject, addHere As cJobject) As cJobject
                                 
    Dim r As Reference
    
    ' get all refs in this workbook
    With addHere
        For Each r In project.theProject.references
            With .add
                .add "name", r.name
                .add "guid", r.Guid
                .add "major", r.major
                .add "minor", r.minor
                .add "description", r.description
            End With
        Next r
    End With
    
    Set makeExcelReferences = addHere
End Function
'
' these are all about handling interface to VBA IDE
'-----------------------------------------------------
Private Function isModuleObj(ob As Object) As Boolean
    Dim obModel As cVBAmodule
    Set obModel = New cVBAmodule
    isModuleObj = (TypeName(ob) = TypeName(obModel))
End Function
Private Function getVbaAsJobject(Optional optProjectName As String = vbNullString) As cJobject
    Dim project As cJobject, knownProjects As cJobject, _
        module As cJobject, settings As cJobject, wb As Workbook
    
    Set settings = getVGSettings
    
    ' default is the first project that's not vbagit
    If optProjectName = vbNullString Then
        For Each wb In Workbooks
            If wb.VBProject.name <> settings.toString("PROJECT.NAME") Then
                optProjectName = wb.VBProject.name
                Exit For
            End If
        Next wb
    End If
    
    ' we must be comitting the code for vbagit
    If optProjectName = vbNullString Then
        Debug.Print "working on " & settings.toString("PROJECT.NAME"); ""
        Debug.Print "if you wanted to do one of your projects, you should have opened another workbook as well"
        optProjectName = settings.toString("PROJECT.NAME")
    End If

    ' projects in this workbook
    Set knownProjects = getProjects(optProjectName)
    
    For Each project In knownProjects.children
        
        ' get all the known modules
        For Each module In getmoduleList(project).children
            ' get all the known procedures
            getProcList module
            ' now blow out the procedures
            blowProcedures module
        Next module

    Next project
    Set getVbaAsJobject = knownProjects
End Function
Private Function blowProcedures(module As cJobject) As cJobject
    Dim pob As cVBAProcedure
    ' need to pick out to a stringifiable object
    With module.add("procedures").addArray
        For Each pob In module.getObject("module").procedures
            With .add
                .add "name", pob.name
                .add "procedure", pob
                'add the arguments
                blowArguments pob, .add("arguments").addArray
            End With

        Next pob
    End With
    Set blowProcedures = module
End Function
Private Function blowArguments(pob As cVBAProcedure, argOb As cJobject) As cJobject
    Dim argument As cVBAArgument
    ' need to pick out to a stringifiable object
    With argOb
        For Each argument In pob.arguments
            With .add
                .add "name", argument.name
                .add "argument", argument
            End With
        Next argument
    End With
    Set blowArguments = argOb
End Function

' get all projects in a workbook
Private Function getProjects(Optional optProjectName As String = vbNullString) As cJobject
    Dim wb As Workbook
    Dim project As cVBAProject
    Dim knownProjects As New cJobject
    knownProjects.init(Nothing).addArray
    
    For Each wb In Workbooks
        If wb.VBProject.name = optProjectName Or optProjectName = vbNullString Then
            Set project = New cVBAProject
            project.init wb
            With knownProjects.add
                .add "name", project.name
                .add "project", project
            End With
        End If
    Next wb
    Set getProjects = knownProjects
    
End Function
 ' get every proc in a module
 Private Sub getProcList(module As cJobject)

    Dim lStart As Long, pName As String
    Dim n As Long, s As String, t As String, doMore As Boolean, countLines As Long
    Dim cm As codeModule
    Dim pk As vbext_prockind
    Dim procedure As cVBAProcedure
    
    Set cm = module.child("module").value.vCom.codeModule
    
    lStart = cm.CountOfDeclarationLines + 1
    While lStart <= cm.CountOfLines
        pName = cm.ProcOfLine(lStart, pk)
        countLines = cm.ProcCountLines(pName, pk)
        Set procedure = New cVBAProcedure
        procedure.init module.child("module").value, pName, pk
        
        lStart = cm.ProcStartLine(pName, pk) + countLines
    Wend

End Sub
' get all modules in a project
Private Function getmoduleList(project As cJobject) As cJobject

    Dim v As VBComponent, vs As VBComponents, wb As Workbook, vj As cVBAProject
    Dim bInc As Boolean, n As Long, ml As cJobject
    
    ' this is the project object
    Set vj = project.child("project").value
    
    ' get the module components
    Set vs = vj.wBook.VBProject.VBComponents
    Dim apm As cVBAmodule
    
    n = 0
    ' add a branch to the project for modules
    Set ml = project.add("modules").addArray
    With ml
        ' append each module
        For Each v In vs
            Set apm = New cVBAmodule
            apm.init v, vj
            n = n + 1
            With .add
                .add "name", apm.name
                .add "module", apm
                .add "kind", apm.textKind
            End With
        Next v
    End With
    Set getmoduleList = ml
End Function

'
' these are all about making JSON info content
'-----------------------------------------------
Private Function makeInfoFile(project As cJobject, dependencyList As cJobject) As cJobject
    Dim infoJob As cJobject, settings As cJobject, job As cJobject
    Set infoJob = New cJobject
    Set settings = getVGSettings()
    
    ' actually the dependency list needs cut down since it contains both scripts and dependencies
    Dim library As cJobject
    Set library = New cJobject
    
    With library.init(Nothing).addArray
        For Each job In dependencyList.child("dependencies").children
            ' this means it's not in the script list
            If (dependencyList.child("scripts").findInArray("name", job.toString("name")) Is Nothing) Then
                ' since we can add objects, nothing to stop that being another cjobject!
                .add , job
            End If
        Next job
    End With

    ' now we can use the library object to driver dependencies
    With infoJob.init(Nothing)
        'the info file name .. we'll try to mirror the structur eof the google apps script/drive version
        
        ' preamble

        .add "title", dependencyList.toString("name")
        .add "committedDate", 0
        .add "createdDate", getTimestampFromDate()
        .add "modifiedDate", getTimestampFromDate()
        .add "version", settings.toString("APP.VERSION")
        .add "noticed", getTimestampFromDate()
        .add "extract", settings.toString("EXTRACT.TO") & dependencyList.toString("name") & "/"
        
        .add "fileName", settings.toString("FILES.INFO")
        .add "fileId", .toString("extract") & .toString("fileName")
        
        'module list
        modulesToInfo dependencyList.child("scripts"), _
            .add("modules").addArray, _
            .toString("extract"), _
            settings.toString("FOLDERS.SCRIPTS")

        .add "extracted", False
        .add "repo", dependencyList.toString("name")
        
        'dependency list
        modulesToInfo library, _
            .add("dependencies").addArray, _
            .toString("extract"), _
            settings.toString("FOLDERS.DEPENDENCIES")
            
        ' add excel references
        makeExcelReferences project.getObject("project"), .add("excelReferences").addArray
        
        .add "readmeFileId", .toString("extract") & settings.toString("FILES.README")
        .add "dependenciesFileId", .toString("extract") & settings.toString("FILES.DEPENDENCIES")
        .add "crossFileId", .toString("extract") & settings.toString("FILES.CROSS")

    End With
    library.tearDown
    Set makeInfoFile = infoJob
End Function

Private Function modulesToInfo(moduleJob As cJobject, infoJob As cJobject, _
        extract As String, folderName As String) As cJobject
    Dim job As cJobject, jo As cJobject, modl As cVBAmodule, fileName As String
    
    With infoJob
        For Each jo In moduleJob.children
        
            ' this is if we are using an indirection for the library
            If jo.isObjValue() Then
                Set job = jo.getObject
            Else
                Set job = jo
            End If
            
            ' get the module
            Set modl = job.getObject("module")
            fileName = modl.name & _
                    conditionalAssignment(modl.textKind = "ClassModule", ".cls", ".vba")
            With .add
                .add "name", modl.name
                .add "type", modl.textKind
                .add "folder", concatFolderName(extract, folderName) & "/"
                .add "id", concatFolderName(.toString("folder"), fileName)
                .add "fileName", fileName
                .add "docsName", modl.name & _
                    conditionalAssignment(modl.textKind = "ClassModule", "_cls", "_vba") & ".md"
                .add "docsId", concatFolderName(.toString("folder"), .toString("docsName"))
            End With
        Next jo
    End With

    Set modulesToInfo = infoJob
End Function
Private Function mdWrap()
    mdWrap = "  " & vbLf
End Function
'
'-- these are about making the content for documentation files
'--------------------------------------------------------------
Private Function makeCross(cross As cJobject, info As cJobject) As String
    Dim c As cStringChunker, job As cJobject
    Set c = New cStringChunker
    
    c.add("# VBA Project: ").addLine (info.toString("title"))
    c.add("This cross reference list for repo (").add(info.toString("repo")).add(") was automatically created on ").add(CStr(Now())).add (" by VBAGit.")
    c.addLine ("For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready ""desktop liberation"")")
    c.add ("You can see [library and dependency information here](")
    c.add(getVGSettings().toString("FILES.DEPENDENCIES")).addLine(")").addLine ("")
    c.addLine ("###Below is a cross reference showing which modules and procedures reference which others")
    c.addLine "*module*|*proc*|*referenced by module*|*proc*"
    c.addLine "---|---|---|---"

    
    For Each job In cross.children
        
        
        ' the module being referenced
        If (isModuleObj(job.getObject("proc"))) Then
            c.add (job.getObject("proc").name)
        Else
            c.add (job.getObject("proc").parent.name)
        End If
        c.add ("|")
        
        ' the proc doing the referencing
        If (Not isModuleObj(job.getObject("proc"))) Then
            c.add (job.getObject("proc").name)
        End If
        c.add ("|")
        
        ' the module doing the referencing
        If (isModuleObj(job.getObject("by"))) Then
            c.add (job.getObject("by").name)
        Else
            c.add (job.getObject("by").parent.name)
        End If
        c.add ("|")
        
        ' the proc doing the referencing
        If (Not isModuleObj(job.getObject("by"))) Then
            c.add (job.getObject("by").name)
        End If
        c.addLine ("")
        

    Next job
    makeCross = c.toString
End Function
Private Function makeReadMe(info As cJobject) As String
    Dim c As cStringChunker
    Set c = New cStringChunker
    
    c.add("# VBA Project: ").addLine (info.toString("title"))
    c.add("This repo (").add(info.toString("repo")).add(") was automatically created on ").add(CStr(Now())).add (" by VBAGit.")
    c.addLine ("For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/vbagit ""desktop liberation"")")
    c.add ("you can see [library and dependency information here](")
    c.add(getVGSettings().toString("FILES.DEPENDENCIES")).addLine(")").addLine ("")
    c.add ("To get started with VBA Git, you can either create a workbook with the [code on gitHub](https://github.com/brucemcpherson/VbaGit ""VbaGit repo"")")
    c.add (", or use this premade [VbaBootStrap workbook](http://ramblings.mcpher.com/Home/excelquirks/downlable-items/VbaGitBootStrap.xlsm ""VbaBootStrap"")")
    c.add (mdWrap)
    c.add ("Now update manually with details of this project - this skeleton file is committed only when there is no README.md in the repo.")

    makeReadMe = c.toString
  
 
End Function

Private Function makeDependency(project As cJobject, info As cJobject) As String
    Dim c As cStringChunker, job As cJobject, settings As cJobject, jo As cJobject
    Set settings = getVGSettings(True)
    Set c = New cStringChunker
    
    c.add("# VBA Project: ").addLine (info.toString("title"))
    c.add("This repo (").add(info.toString("repo")).add (") was automatically created on ")
    c.add(CStr(Now())).add (" by VBAGit.")
    c.add ("For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready ""desktop liberation"")")
    c.add (" or [contact me on G+](https://plus.google.com/+BruceMcpherson ""Bruce McPherson - GDE"")")
    c.add (mdWrap)
    
    c.add("## Details for VBA project ").addLine (info.toString("title"))
    c.add ("Where possibile directly referenced or sub referenced library sources have been copied to this repository")
    c.add (mdWrap)
    
    c.add("### Modules of ").add(info.toString("title")).add(" included in this repo").addLine
    c.add("*name*|*type*|*source*|*docs*").add(mdWrap).add("---|---|---|---").add (mdWrap)
   
    
    For Each job In info.kids("modules")
    
        c.add(job.toString("name")).add("|").add(job.toString("type")).add ("|")
        c.add("[").add(job.toString("fileName")).add("](").add(settings.toString("FOLDERS.SCRIPTS")).add ("/")
        c.add(job.toString("fileName")).add(" ""script source"")").add ("|")
        c.add("[").add(job.toString("docsName")).add("](").add(settings.toString("FOLDERS.SCRIPTS")).add ("/")
        c.add(job.toString("docsName")).add (" ""script docs"")")
        c.add (mdWrap)
    
    Next job
    
    c.add(mdWrap).add("### All dependencies and sub dependencies in this repo").add (mdWrap)
    c.add("*name*|*type*|*source*|*docs*").add(mdWrap).add("---|---|---|---").add (mdWrap)
    
    For Each job In info.kids("dependencies")
        c.add(job.toString("name")).add("|").add(job.toString("type")).add ("|")
        c.add("[").add(job.toString("fileName")).add("](").add(settings.toString("FOLDERS.DEPENDENCIES")).add ("/")
        c.add(job.toString("fileName")).add(" ""library source"")").add ("|")
        c.add("[").add(job.toString("docsName")).add("](").add(settings.toString("FOLDERS.DEPENDENCIES")).add ("/")
        c.add(job.toString("docsName")).add (" ""library docs"")")
        c.add (mdWrap)
    Next job

    c.add(mdWrap).add("###Excel references").add (mdWrap)
    
    If (info.child("excelReferences").children.Count > 0) Then
        c.add ("####These references were detected in the workbook (")
        c.add (project.getObject("project").wBook.name)
        c.add (") this repo was created from. You may not need them all")
        c.add (mdWrap)
        
        ' do the table titles
        For Each job In info.child("excelReferences").kids(1)
            c.add("*").add(job.key).add ("*|")
        Next job
        c.chopIf("|").add (mdWrap)
        
        For Each job In info.child("excelReferences").kids(1)
            c.add ("---|")
        Next job
        c.chopIf("|").add (mdWrap)
        
        ' now the content
        For Each jo In info.kids("excelReferences")
            For Each job In jo.children
                c.add(job.cValue).add ("|")
            Next job
            c.chopIf("|").add (mdWrap)
        Next jo
        c.chopIf("|").add (mdWrap)
        
    Else
        c.add ("####No references were detected in the workbook (")
        c.add (project.getObject("project").wBook.name)
        c.add (") this repo was created from.")
        c.add (mdWrap)
    End If
    
    c.add (mdWrap)
    c.add ("You can see [full project info as json here](")
    c.add(info.toString("fileName")).add (")")

    makeDependency = c.toString
  
End Function
Private Function constructModLink(name As String, folder As String, fileName As String, hover As String)
    Dim c As cStringChunker
    Set c = New cStringChunker
    c.add("[").add(name).add ("](")
    c.add(folder).add ("/")
    c.add(fileName).add (" """)
    If hover <> vbNullString Then
        c.add hover
    Else
        c.add name
    End If
    c.add("""").add (")")
    constructModLink = c.toString
End Function
Private Function makeArguments(modl As cVBAmodule, info As cJobject) As String
    ' this will make a mardown string for all the procedures and arguments in this module
    Dim c As cStringChunker, proc As cVBAProcedure, a As cVBAArgument
    Set c = New cStringChunker
    
    c.add("# VBA Project: **").add(info.toString("title")).addLine ("**")
    c.add("## VBA Module: **").add(findModLink(modl.name, info, "source is here", "fileName")).addLine ("**")
    c.add("### Type: ").add(modl.textKind).add("  ").addLines (2)
    c.add("This procedure list for repo (").add (info.toString("repo"))
    c.add(") was automatically created on ").add(CStr(Now())).addLine (" by VBAGit.")
    c.addLine ("For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready ""desktop liberation"")")
    c.addLine.add("Below is a section for each procedure in ").add (modl.name)

    For Each proc In modl.procedures
        c.addLines(2).addLine ("---")
        c.add("VBA Procedure: **").add(proc.name).add("**").add("  ").addLine
        c.add("Type: **").add(proc.procTextKind).add("**").add("  ").addLine
        c.add("Returns: **").add(findModLink(proc.procReturns, info, , "docsName")).add("**").add("  ").addLine
        c.add("Return description: **").add(proc.returnDoc).add("**").add("  ").addLine
        c.add("Scope: **").add(proc.scope).add("**").add("  ").addLine
        c.add("Description: **").add(proc.description).add("**").add("  ").addLine

        c.addLine.add("*").add(proc.declaration).add("*").add("  ").addLines (2)
        
        If (proc.arguments.Count > 0) Then
            c.addLine "*name*|*type*|*optional*|*default*|*description*"
            c.addLine "---|---|---|---|---"
            
            For Each a In proc.arguments
                c.add(a.name).add ("|")
                c.add (findModLink(a.argType, info, , "docsName"))
                c.add("|").add(a.isOptional).add ("|")
                c.add(a.default).add ("|")
                c.addLine (a.description)
            Next a
        Else
            c.addLine ("**no arguments required for this procedure**")
        End If
    Next proc
    makeArguments = c.toString
End Function
Private Function findModLink(modlName As String, info As cJobject, Optional hover As String = vbNullString, _
    Optional fn As String = "docsName") As String
    
    Dim job As cJobject, settings As cJobject, c As cStringChunker
    Set c = New cStringChunker
    Set settings = getVGSettings(True)
    If hover = vbNullString Then
        hover = modlName
    End If
    For Each job In info.kids("modules")
        If (job.toString("name") = modlName) Then
            findModLink = constructModLink(modlName, "/" & settings.toString("FOLDERS.SCRIPTS"), job.toString(fn), hover)
            Exit Function
        End If
    Next job
    
    For Each job In info.kids("dependencies")
        If (job.toString("name") = modlName) Then
            findModLink = constructModLink(modlName, "/" & settings.toString("FOLDERS.DEPENDENCIES"), job.toString(fn), hover)
            Exit Function
        End If
    Next job
    
    findModLink = modlName
    
End Function
'
' these are all about handling credentials
'-----------------------------------------
Public Function getFromVbaGitRegistry(key) As String
    Dim j As cJobject
    Set j = getVGSettings().child("REGISTRY")
    getFromVbaGitRegistry = GetSetting(j.toString("root"), j.toString("app"), key)
    
End Function
Public Function setVbaGitRegistry(key, value) As String
    Dim j As cJobject
    Set j = getVGSettings().child("REGISTRY")
    SaveSetting j.toString("root"), j.toString("app"), key, value
End Function

Private Function getGitBasicCredentials()
    getGitBasicCredentials = getFromVbaGitRegistry( _
        getVGSettings().toString("REGISTRY.basic"))
End Function
Private Sub setGitBasicCredentials(user As String, pass As String)
    setVbaGitRegistry getVGSettings() _
        .toString("REGISTRY.basic"), Base64Encode(user & ":" & pass)
End Sub
Private Sub setGitClientCredentials(clientId As String, clientSecret As String)
    setVbaGitRegistry getVGSettings() _
        .toString("REGISTRY.client"), _
        Base64Encode(clientId & ":" & clientSecret)
End Sub
Private Function getGitClientCredentials()
    getGitClientCredentials = getFromVbaGitRegistry( _
        getVGSettings().toString("REGISTRY.client"))
End Function
