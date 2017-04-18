# VBA Project: **VbaGit**
## VBA Module: **[usefulcJobject](/libraries/usefulcJobject.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 4/18/2017 10:42:58 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulcJobject

---
VBA Procedure: **fromISODateTime**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function fromISODateTime(iso As String) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
iso|String|False||


---
VBA Procedure: **toISODateTime**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function toISODateTime(d As Date) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
d|Date|False||


---
VBA Procedure: **milliseconds**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function milliseconds(d As Date) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
d|Date|False||


---
VBA Procedure: **JSONParse**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||
jtype|eDeserializeType|True||
complain|Boolean|True| True|


---
VBA Procedure: **JSONStringify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
j|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
blf|Boolean|True||


---
VBA Procedure: **jSonArgs**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function jSonArgs(options As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
options|String|False||


---
VBA Procedure: **optionsExtend**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function optionsExtend(givenOptions As String, Optional defaultOptions As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
givenOptions|String|False||
defaultOptions|String|True| vbNullString|


---
VBA Procedure: **ucJobjectMake**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function ucJobjectMake(r As Variant) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Variant|False||


---
VBA Procedure: **ucJobjectChildValue**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function ucJobjectChildValue(json As Variant, child As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json|Variant|False||
child|Variant|False||


---
VBA Procedure: **ucJobjectLint**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function ucJobjectLint(json As Variant, Optional child As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json|Variant|False||
child|Variant|True||


---
VBA Procedure: **cleanGoogleWire**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function cleanGoogleWire(sWire As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sWire|String|False||


---
VBA Procedure: **xmlStringToJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function xmlStringToJobject(xmlString As String, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
xmlString|String|False||
complain|Boolean|True| True|


---
VBA Procedure: **docToJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function docToJobject(doc As Object, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
doc|Object|False||
complain|Boolean|True| True|


---
VBA Procedure: **isArrayRoot**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isArrayRoot(parent As IXMLDOMNode) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
parent|IXMLDOMNode|False||


---
VBA Procedure: **handleNodes**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function handleNodes(parent As IXMLDOMNode, job As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
parent|IXMLDOMNode|False||
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **handleNode**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function handleNode(node As IXMLDOMNode, job As cJobject, Optional arrayHead As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
node|IXMLDOMNode|False||
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
arrayHead|Boolean|True| False|


---
VBA Procedure: **hackJSObjectToJSON**  
Type: **Function**  
Returns: **String**  
Return description: **the hacked string**  
Scope: **Public**  
Description: **this will deal with the problem of code copied from javascript, where JSON has no quotes round property names, with a callback**  

*Public Function hackJSObjectToJSON(theString As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
theString|String|False||the string to be hacked


---
VBA Procedure: **hackJSONPObjectToJSON**  
Type: **Function**  
Returns: **String**  
Return description: **the hacked string**  
Scope: **Public**  
Description: ****  

*Public Function hackJSONPObjectToJSON(theString As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
theString|String|False||the string to be hacked
