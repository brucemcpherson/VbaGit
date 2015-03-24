# VBA Project: **VbaGit**
## VBA Module: **[regXLib](/libraries/regXLib.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in regXLib

---
VBA Procedure: **rxString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function rxString(sName As String, s As String, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
s|String|False|
ignorecase|Boolean|True| True


---
VBA Procedure: **rxGroup**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function rxGroup(sName As String, s As String, group As Long, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
s|String|False|
group|Long|False|
ignorecase|Boolean|True| True


---
VBA Procedure: **rxTest**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  

*Public Function rxTest(sName As String, s As String, Optional ignorecase As Boolean = True) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
s|String|False|
ignorecase|Boolean|True| True


---
VBA Procedure: **rxReplace**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function rxReplace(sName As String, sFrom As String, sTo As String, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
sFrom|String|False|
sTo|String|False|
ignorecase|Boolean|True| True


---
VBA Procedure: **rxPattern**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function rxPattern(sName As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|


---
VBA Procedure: **rxMakeRxLib**  
Type: **Function**  
Returns: **[cregXLib](/libraries/cregXLib_cls.md "cregXLib")**  
Scope: **Public**  

*Function rxMakeRxLib(sName As String) As cregXLib*  

*name*|*type*|*optional*|*default*
---|---|---|---
sName|String|False|
