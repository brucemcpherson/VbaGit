'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:47 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414836/raw/regXLib.vba
Option Explicit
' v2.02
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Public Function rxString(sName As String, s As String, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxString = rx.getString(s)

End Function
Public Function rxGroup(sName As String, s As String, group As Long, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxGroup = rx.getGroup(s, group)

End Function
Public Function rxTest(sName As String, s As String, Optional ignorecase As Boolean = True) As Boolean
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxTest = rx.getTest(s)

End Function
Public Function rxReplace(sName As String, sFrom As String, sTo As String, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
     ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' replace the string that matches the requested pattern
    rxReplace = rx.getReplace(sFrom, sTo)
    
End Function
Public Function rxPattern(sName As String) As String
    Dim rx As cregXLib
     ' create a new regx
    Set rx = rxMakeRxLib(sName)
    ' just returnthe pattern
    rxPattern = rx.pattern
    
End Function
 Function rxMakeRxLib(sName As String) As cregXLib
    Dim rx As cregXLib, s As String
    Set rx = New cregXLib
    ' normally sname points to a preselected regEX
    ' if not known, silently assume its a regex pattern
        s = Replace(UCase(sName), " ", "")
        Select Case s
            Case "POSTALCODEUK"
                rx.init s, _
                "(((^[BEGLMNS][1-9]\d?) | (^W[2-9] ) | ( ^( A[BL] | B[ABDHLNRST] | C[ABFHMORTVW] | D[ADEGHLNTY] | E[HNX] | F[KY] | G[LUY] | H[ADGPRSUX] | I[GMPV] |" & _
                " JE | K[ATWY] | L[ADELNSU] | M[EKL] | N[EGNPRW] | O[LX] | P[AEHLOR] | R[GHM] | S[AEGKL-PRSTWY] | T[ADFNQRSW] | UB | W[ADFNRSV] | YO | ZE ) \d\d?) |" & _
                " (^W1[A-HJKSTUW0-9]) | ((  (^WC[1-2])  |  (^EC[1-4]) | (^SW1)  ) [ABEHMNPRVWXY] ) ) (\s*)?  ([0-9][ABD-HJLNP-UW-Z]{2})) | (^GIR\s?0AA)"
            
            Case "POSTALCODESPAIN"
                rx.init s, _
                    "^([1-9]{2}|[0-9][1-9]|[1-9][0-9])[0-9]{3}$"
                    
            Case "PHONENUMBERUS"
                rx.init s, _
                "^\(?(?<AreaCode>[2-9]\d{2})(\)?)(-|.|\s)?(?<Prefix>[1-9]\d{2})(-|.|\s)?(?<Suffix>\d{4})$"
                
            Case "CREDITCARD" 'amex/visa/mastercard
                rx.init s, _
                "^((4\d{3})|(5[1-5]\d{2}))(-?|\040?)(\d{4}(-?|\040?)){3}|^(3[4,7]\d{2})(-?|\040?)\d{6}(-?|\040?)\d{5}"
                
            Case "NUMERIC"
                rx.init s, _
                    "[\0-9]"
            
            Case "ALPHABETIC"
                rx.init s, _
                    "[\a-zA-Z]"
                    
            Case "NONNUMERIC"
                rx.init s, _
                    "[^\0-9]"
                    
            Case "IPADDRESS"
                rx.init s, _
                "^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$"
            
            Case "SINGLESPACE"  ' should take a replace value of "$1 "
                rx.init s, _
                    "(\S+)\x20{2,}(?=\S+)"
            
            Case "EMAIL"
                rx.init s, _
                    "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$"
                    
            Case "EMAILINSIDE"
                rx.init s, _
                    "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b"
                    
            Case "NONPRINTABLE"
                rx.init s, "[\x00-\x1F\x7F]"
                
                
            Case "PUNCTUATION"
                rx.init s, "[^A-Za-z0-9\x20]+"

            Case Else
                rx.init "Adhoc", sName
        
        End Select
    
    Set rxMakeRxLib = rx
End Function




