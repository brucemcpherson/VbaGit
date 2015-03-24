# VBA Project: **VbaGit**
## VBA Module: **[usefulcJobject](/libraries/usefulcJobject.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VbaGit) was automatically created on 24/03/2015 10:59:10 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulcJobject

---
VBA Procedure: **fromISODateTime**  
Type: **Function**  
Returns: **Date**  
Scope: **Public**  

*Public Function fromISODateTime(iso As String) As Date*  

*name*|*type*|*optional*|*default*
---|---|---|---
iso|String|False|


---
VBA Procedure: **toISODateTime**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function toISODateTime(d As Date) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
d|Date|False|


---
VBA Procedure: **milliseconds**  
Type: **Function**  
Returns: **Double**  
Scope: **Public**  

*Public Function milliseconds(d As Date) As Double*  

*name*|*type*|*optional*|*default*
---|---|---|---
d|Date|False|


---
VBA Procedure: **JSONParse**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
s|String|False|
jtype|eDeserializeType|True|
complain|Boolean|True| True


---
VBA Procedure: **JSONStringify**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
j|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
blf|Boolean|True|


---
VBA Procedure: **jSonArgs**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function jSonArgs(options As String) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
options|String|False|


---
VBA Procedure: **optionsExtend**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function optionsExtend(givenOptions As String, Optional defaultOptions As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
givenOptions|String|False|
defaultOptions|String|True| vbNullString


---
VBA Procedure: **ucJobjectMake**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function ucJobjectMake(r As Variant) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
r|Variant|False|


---
VBA Procedure: **ucJobjectChildValue**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function ucJobjectChildValue(json As Variant, child As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
json|Variant|False|
child|Variant|False|


---
VBA Procedure: **ucJobjectLint**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function ucJobjectLint(json As Variant, Optional child As Variant) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
json|Variant|False|
child|Variant|True|


---
VBA Procedure: **cleanGoogleWire**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  

*Public Function cleanGoogleWire(sWire As String) As String*  

*name*|*type*|*optional*|*default*
---|---|---|---
sWire|String|False|


---
VBA Procedure: **xmlStringToJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function xmlStringToJobject(xmlString As String, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
xmlString|String|False|
complain|Boolean|True| True


---
VBA Procedure: **docToJobject**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  

*Public Function docToJobject(doc As Object, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
doc|Object|False|
complain|Boolean|True| True


---
VBA Procedure: **isArrayRoot**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  

*Private Function isArrayRoot(parent As IXMLDOMNode) As Boolean*  

*name*|*type*|*optional*|*default*
---|---|---|---
parent|IXMLDOMNode|False|


---
VBA Procedure: **handleNodes**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function handleNodes(parent As IXMLDOMNode, job As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
parent|IXMLDOMNode|False|
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|


---
VBA Procedure: **handleNode**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  

*Private Function handleNode(node As IXMLDOMNode, job As cJobject, Optional arrayHead As Boolean = False) As cJobject*  

*name*|*type*|*optional*|*default*
---|---|---|---
node|IXMLDOMNode|False|
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False|
arrayHead|Boolean|True| False
