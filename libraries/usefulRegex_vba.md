# VBA Project: **VbaGit**
## VBA Module: **[usefulRegex](/libraries/usefulRegex.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulRegex

---
VBA Procedure: **straightenOutContinuations**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function straightenOutContinuations(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRidOfDims**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRidOfDims(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRidOfQuoted**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRidOfQuoted(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRidOfComments**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRidOfComments(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRidOfMultipleSpaces**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRidOfMultipleSpaces(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRx**  
Type: **Function**  
Returns: **RegExp**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRx(pattern As String, Optional multi As Boolean = True) As RegExp*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pattern|String|False||
multi|Boolean|True| True|


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
VBA Procedure: **getDimLinesRx**  
Type: **Function**  
Returns: **RegExp**  
Return description: **get a regex that picks out all lines with Dim**  
Scope: **Public**  
Description: **@return {} get a regex that picks out all lines with Dim**  

*Public Function getDimLinesRx() As RegExp*  

**no arguments required for this procedure**


---
VBA Procedure: **getDimLocalsRx**  
Type: **Function**  
Returns: **RegExp**  
Return description: **get a regex that picks out all locally defined variables from a dim**  
Scope: **Public**  
Description: **@return {} get a regex that picks out all locally defined variables from a dim**  

*Public Function getDimLocalsRx() As RegExp*  

**no arguments required for this procedure**
