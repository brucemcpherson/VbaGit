# VBA Project: **VbaGit**
## VBA Module: **[cVBAProcedure](/libraries/cVBAProcedure.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVBAProcedure

---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **description**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let description(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **description**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get description() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **returnDoc**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get returnDoc() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **arguments**  
Type: **Get**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get arguments() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cVBAmodule*  

**no arguments required for this procedure**


---
VBA Procedure: **procKind**  
Type: **Get**  
Returns: **vbext_prockind**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get procKind() As vbext_prockind*  

**no arguments required for this procedure**


---
VBA Procedure: **isAnArgument**  
Type: **Function**  
Returns: **Boolean**  
Return description: **whether it is an argument**  
Scope: **Public**  
Description: **checks to see if a given variable name is an argument of this procedure**  

*Public Function isAnArgument(argName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
argName|String|False||the name to check


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function init(m As cVBAmodule, pn As String, pk As vbext_prockind) As cVBAProcedure*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
m|[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")|False||
pn|String|False||
pk|vbext_prockind|False||


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **lineCount**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get lineCount() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **codeModule**  
Type: **Get**  
Returns: **codeModule**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get codeModule() As codeModule*  

**no arguments required for this procedure**


---
VBA Procedure: **startLine**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get startLine() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **removeComments**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function removeComments(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getDescription**  
Type: **Function**  
Returns: **String**  
Return description: **the procedure description**  
Scope: **Public**  
Description: **interprets jsdoc like procedure description**  

*Public Function getDescription() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getReturnDoc**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getReturnDoc() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dealWithArguments**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dealWithArguments(dec As String) As cVBAProcedure*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dec|String|False||


---
VBA Procedure: **scope**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get scope() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **textKind**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function textKind(k As vbext_prockind) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
k|vbext_prockind|False||


---
VBA Procedure: **procTextKind**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get procTextKind() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **procReturns**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get procReturns() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheCode**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getTheCode() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getFinishWithoutTrailingComments**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getFinishWithoutTrailingComments() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheEndRx**  
Type: **Function**  
Returns: **RegExp**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getTheEndRx() As RegExp*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheCodePlusLeadingComments**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getTheCodePlusLeadingComments() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **declaration**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get declaration() As String*  

**no arguments required for this procedure**
