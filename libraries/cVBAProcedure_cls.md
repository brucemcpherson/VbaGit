# VBA Project: **VbaGit**
## VBA Module: **[cVBAProcedure](/libraries/cVBAProcedure.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 25/03/2015 18:59:47 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVBAProcedure

---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **description**  
Type: **Let**  
Returns: **void**  
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
Scope: **Public**  
Description: ****  

*Public Property Get description() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **arguments**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get arguments() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cVBAmodule*  

**no arguments required for this procedure**


---
VBA Procedure: **procKind**  
Type: **Get**  
Returns: **vbext_prockind**  
Scope: **Public**  
Description: ****  

*Public Property Get procKind() As vbext_prockind*  

**no arguments required for this procedure**


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
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
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **lineCount**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get lineCount() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **codeModule**  
Type: **Get**  
Returns: **codeModule**  
Scope: **Public**  
Description: ****  

*Public Property Get codeModule() As codeModule*  

**no arguments required for this procedure**


---
VBA Procedure: **startLine**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get startLine() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **removeComments**  
Type: **Function**  
Returns: **String**  
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
Scope: **Public**  
Description: **interprets jsdoc like procedure description**  

*Public Function getDescription() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dealWithArguments**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
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
Scope: **Public**  
Description: ****  

*Public Property Get scope() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **textKind**  
Type: **Function**  
Returns: **String**  
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
Scope: **Public**  
Description: ****  

*Public Property Get procTextKind() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **procReturns**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get procReturns() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheCode**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getTheCode() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getFinishWithoutTrailingComments**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Function getFinishWithoutTrailingComments() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheEndRx**  
Type: **Function**  
Returns: **RegExp**  
Scope: **Public**  
Description: ****  

*Public Function getTheEndRx() As RegExp*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheCodePlusLeadingComments**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getTheCodePlusLeadingComments() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **declaration**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get declaration() As String*  

**no arguments required for this procedure**
