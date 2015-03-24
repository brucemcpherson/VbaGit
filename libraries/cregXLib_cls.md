# VBA Project: **VbaGit**
## VBA Module: **[cregXLib](/libraries/cregXLib.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cregXLib

---
VBA Procedure: **pattern**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get pattern() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **pattern**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let pattern(p As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|String|False|


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let name(p As String)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|String|False|


---
VBA Procedure: **ignorecase**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get ignorecase() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **ignorecase**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let ignorecase(p As Boolean)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Boolean|False|


---
VBA Procedure: **rGlobal**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  

*Public Property Get rGlobal() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **rGlobal**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  

*Public Property Let rGlobal(p As Boolean)*  

*name*|*type*|*optional*|*default*
---|---|---|---
p|Boolean|False|


---
VBA Procedure: **init**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  

*Public Sub init(sName As String, Optional spat As String = "", Optional bIgnoreSpaces As Boolean = True, Optional bIgnoreCase As Boolean = True, Optional bGlobal As Boolean = True)*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
spat|String|True| ""
bIgnoreSpaces|Boolean|True| True
bIgnoreCase|Boolean|True| True
bGlobal|Boolean|True| True


---
VBA Procedure: **getString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function getString(sFrom As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sFrom|String|False|


---
VBA Procedure: **getGroup**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function getGroup(sFrom As String, groupNumber As Long) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sFrom|String|False|
groupNumber|Long|False|


---
VBA Procedure: **getReplace**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Function getReplace(sFrom As String, sTo As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sFrom|String|False|
sTo|String|False|


---
VBA Procedure: **getTest**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Function getTest(sFrom As String) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
sFrom|String|False|
