# VBA Project: **VbaGit**
## VBA Module: **[cVBAProcedure](/libraries/cVBAProcedure.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVBAProcedure

---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **arguments**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  

*Public Property Get arguments() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")**  
Scope: **Public**  

*Public Property Get parent() As cVBAmodule*  

**no arguments required for this procedure**


---
VBA Procedure: **procKind**  
Type: **Get**  
Returns: **vbext_prockind**  
Scope: **Public**  

*Public Property Get procKind() As vbext_prockind*  

**no arguments required for this procedure**


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Scope: **Public**  

*Public Function init(m As cVBAmodule, pn As String, pk As vbext_prockind) As cVBAProcedure*  

*name*|*type*|*optional*|*default*
---|---|---|---
m|[cVBAmodule](/libraries/cVBAmodule_cls.md "cVBAmodule")|False|
pn|String|False|
pk|vbext_prockind|False|


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **lineCount**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get lineCount() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **codeModule**  
Type: **Get**  
Returns: **codeModule**  
Scope: **Public**  

*Public Property Get codeModule() As codeModule*  

**no arguments required for this procedure**


---
VBA Procedure: **startLine**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  

*Public Property Get startLine() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **removeComments**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function removeComments(s As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|


---
VBA Procedure: **dealWithArguments**  
Type: **Function**  
Returns: **[cVBAProcedure](/libraries/cVBAProcedure_cls.md "cVBAProcedure")**  
Scope: **Private**  

*Private Function dealWithArguments(dec As String) As cVBAProcedure*  

*name*|*type*|*optional*|*default*
---|---|---|---
dec|String|False|


---
VBA Procedure: **scope**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get scope() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **textKind**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  

*Private Function textKind(k As vbext_prockind) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
k|vbext_prockind|False|


---
VBA Procedure: **procTextKind**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get procTextKind() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **procReturns**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get procReturns() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getTheCode**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get getTheCode() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **declaration**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get declaration() As String*  

**no arguments required for this procedure**
