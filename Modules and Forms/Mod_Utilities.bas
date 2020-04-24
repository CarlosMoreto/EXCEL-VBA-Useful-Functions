Attribute VB_Name = "Mod_Utilities"
'==========================================================================================================='
'------------------------------------------- Utilities Functions -------------------------------------------'
'==========================================================================================================='
'------------------------------------------- Translate Functions -------------------------------------------'
'Public Function Translate_Str(ByVal theStr As String, Optional ByVal transFrom As String = "pt", Optional ByVal transTo As String = "en") As String
'Public Function Refine_Trans(ByVal theStr As String) As String
'Public Function ConvertToGet(ByVal val As String)
'Public Function Clean(ByVal val As String)

'Public Function RegexExecute(ByVal str As String, ByVal reg As String, Optional ByVal matchIndex As Long, _
'                             Optional ByVal subMatchIndex As Long) As String
'-----------------------------------------------------------------------------------------------------------'

'-------------------------------------------- String Functions ---------------------------------------------'
'Public Function Standard_Str(ByVal theStr As String) As String
'Public Function Remove_Extra_Space(ByVal theStr As String) As String

'Public Function Regex_Match(ByVal theStr As String, ByVal theFilter As String, Optional ByVal matchSpace As Boolean = False, _
'                                  Optional ByVal matchCrLf As Boolean = False, Optional ByVal argGlobal As Boolean = True, _
'                                  Optional ByVal argCase As Boolean = True) As Variant

'Public Function Read_From_Pattern(ByVal theStr As String, ByVal thePat As Variant, Optional ByVal numPat As Boolean = False) As String
'Public Function Read_From_Pattern_X(ByVal theStr As String, ByVal thePat As String, Optional ByVal theOS As Integer = 0) As String
'Public Function Read_Until_Pattern_X(ByVal theStr As String, ByVal thePat As String, Optional ByVal theOS As Integer = 0) As String
'Public Function Read_Until_Pattern(ByVal theStr As String, ByVal thePat As Variant, Optional ByVal numPat As Boolean = False) As String

'Public Function Read_Delim_Str_X(ByVal theStr As String, ByVal startPat As String, ByVal endPat As String, Optional ByVal startOS As Integer = 0, _
'                                 Optional ByVal endOS As Integer = 0) As String

'Public Function Read_Delim_Str(ByVal theStr As String, ByVal startPat As Variant, ByVal endPat As Variant, _
'                               Optional ByVal numPat As Boolean = False) As String

'Public Function Isolate_Symbols(ByVal theStr As String) As Variant
'Public Function Isolate_Words(ByVal theStr As String) As Variant
'Public Function Isolate_Numbers(ByVal theStr As String) As Variant
'Public Function Trim_Left(ByVal theStr As String, theCount As Integer) As String
'Public Function Trim_Right(ByVal theStr As String, theCount As Integer) As String
'Public Function Count_Pattern_Occurrences(ByVal theStr As String, ByVal thePat As String) As Integer
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------- Document Functions --------------------------------------------'
'Public Sub Protect_Sheets(ByVal thePass As String, ByVal toggle As Integer)
'Public Sub Append_To_Sheet(ByVal theSht As String)

'Public Sub Reset_Workbook(ByVal theWB As String, Optional ByVal byCopy As Boolean = False, Optional ByVal copySht As String, _
'                          Optional ByVal refPos As String = "before", Optional ByVal refSht As String)

'Public Sub Remove_Sheet(ByVal theSht As String)

'Public Function Reset_Sheet(ByVal theSht As String, Optional ByVal byCopy As Boolean = False, Optional ByVal copySht As String = "", _
'                            Optional ByVal refPos As String = "before", Optional ByVal refSht As String, _
'                            Optional ByVal theWB As String = "") As Worksheet

'Public Function Check_Sheet(ByVal pName As String, Optional ByVal theWB As String = "") As Boolean

'Public Function Find_Doc_By_Pattern(ByVal docLvl As Integer, ByVal thePat As String, Optional ByVal dpSch As Boolean = False, _
'                                    Optional ByVal refStr As String = "", Optional ByVal excludeList As String = "") As Variant

'Public Function Count_Docs_By_Pattern(ByVal docLvl As Integer, ByVal thePat As String, Optional ByVal dpSch As Boolean = False, _
'                                      Optional ByVal refStr As String = "", Optional ByVal getNames As Boolean = False) As Variant

'Public Function Check_Workbook(ByVal theWB As String) As Boolean
'Public Function Count_Files_In_Folder(ByVal thePath As String, ByVal theExt As String) As Integer
'Public Sub Isolate_Sheet(ByVal theSht As Worksheet, Optional ByVal theWB As Workbook = Nothing)
'Public Function Get_Max_Row_Level(ByVal theSht As Worksheet) As Integer
'Public Function Get_Max_Col_Level(ByVal theSht As Worksheet) As Integer
'Function Get_Folder() As String
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------ Sheet Cells Functions ------------------------------------------'
'Public Function Check_Hidden_Columns(ByVal theSht As Worksheet) As String
'Public Sub Show_Columns(ByVal colList As String, ByVal theSht As Worksheet, ByVal showCols As Boolean)
'Public Function Get_Hidden_Column_Range(ByVal theSht As Worksheet, ByVal rngIdx As Integer) As String

'Public Function Find_N_Get_Cell(ByVal theSht As Worksheet, ByVal theStr As String, Optional ByVal ref As Variant = Empty, _
'                                Optional ByVal LkIn As Integer = xlValues, Optional ByVal LkAt As Integer = xlPart, _
'                                Optional ByVal mCase As Boolean = False) As Range
'Sub Add_Comment(ByVal theCell As Range, ByVal theStr As String, Optional ByVal cmtSets As Variant)
'-----------------------------------------------------------------------------------------------------------'

'--------------------------------------------- Array Functions ---------------------------------------------'
'Public Function Arr_To_Str(ByVal theArr As Variant) As String
'Public Function Find_In_Array(ByVal theArr As Variant, ByVal theElm As Variant, Optional ByVal matchWhole = True) As Integer
'Public Function Arr_Len(ByVal theArr As Variant) As Integer
'Public Sub Print_Array(ByVal theArr As Variant)
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------ Conversion Functions -------------------------------------------'
'Public Function To_Double(ByVal theStr As String, Optional ByVal digits As Integer = 0) As Double
'Public Function Hex_To_Dec(ByVal hex As String) As Long
'-----------------------------------------------------------------------------------------------------------'

'----------------------------------------- Miscellaneous Functions -----------------------------------------'
'Public Sub Blink(ByVal theCell As Range, Optional ByVal theColor As Long = vbRed, Optional ByVal blinkCount As Integer = 4, _
'                 Optional ByVal blinkTime As Integer = 100)

'Public Function Get_Float_Digits(ByVal theNum As Double) As Integer
'Public Sub Print_Var_Type(ByVal var As Variant)
'-----------------------------------------------------------------------------------------------------------'

'Option Private Module

Public Type st_rgxOut
    matchCount As Integer
    matchStr As String
    matchArr() As String
End Type

Public Const pxConv = 0.141
Public Const latStart = "\u00c0"
Public Const latEnd = "\u00ff"
Public Const latUpStart = "\u00c0"
Public Const latUpEnd = "\u00df"
Public Const latLowStart = "\u00e0"
Public Const latLowEnd = "\u00ff"
Public Const allChars = "A-Za-z\u00c0-\u00ff"
Public Const splitChar = "§"
Public Const wbLvl = 0
Public Const wsLvl = 1
Private Const frmHeadHt = 28.5
Private Const frmRightBorder = 13
Private Const MAX_BACKUP = 10
Public Const NOT_FOUND = -1
Public Const TYPE_EMPTY = 0
Public Const TYPE_NULL = 1
Public Const TYPE_INTEGER = 2
Public Const TYPE_LONG = 3
Public Const TYPE_SINGLE = 4
Public Const TYPE_DOUBLE = 5
Public Const TYPE_CURRENCY = 6
Public Const TYPE_DATE = 7
Public Const TYPE_STRING = 8
Public Const TYPE_OBJECT = 9
Public Const TYPE_ERROR = 10
Public Const TYPE_BOOLEAN = 11
Public Const TYPE_VARIANT = 12
Public Const TYPE_DATAOBJECT = 13
Public Const TYPE_DECIMAL = 14
Public Const TYPE_BYTE = 17
Public Const TYPE_LONGLONG = 20
Public Const TYPE_USERDEFINEDTYPE = 36
Public Const TYPE_ARRAY = 8192
Public Const dataSource = "VISION-SW3"
Public Const dataBase = "SGI"
Public Const usrName = "sa"
Public Const usrPassWord = "20Vision_10#"

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

'------------------------------------------- Translate Functions -------------------------------------------'
Public Function Translate_Str(ByVal theStr As String, Optional ByVal transFrom As String = "pt", Optional ByVal transTo As String = "en") As String
    Dim getParam As String, trans As String, translateFrom As String, translateTo As String
    translateFrom = transFrom
    translateTo = transTo
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    getParam = ConvertToGet(theStr)
    URL = "https://translate.google.pl/m?hl=" & translateFrom & "&sl=" & translateFrom & "&tl=" & translateTo & "&ie=UTF-8&prev=_m&q=" & getParam
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    
    If InStr(objHTTP.responseText, "div dir=""ltr""") > 0 Then
        trans = RegexExecute(objHTTP.responseText, "div[^""]*?""ltr"".*?>(.+?)</div>")
        trans = Refine_Trans(trans)
        Translate_Str = Clean(trans)
    Else
        Translate_Str = "Error"
    End If
End Function

Public Function Refine_Trans(theStr As String) As String
    theStr = Replace(theStr, "TC", "CT")
    theStr = Replace(theStr, "TP", "VT")
    theStr = Replace(theStr, "Qtd", "Qty")
    theStr = Replace(theStr, " arrester ", " surge arrester ")
    theStr = Replace(theStr, "In-light arresters ", " surge arrester ")
    theStr = Replace(theStr, "In-arrester ", " surge arrester ")
    theStr = Replace(theStr, "Isolator ", "Disconnector ")
    theStr = Replace(theStr, "Separator ", "Disconnector ")
    theStr = Replace(theStr, "description", "Description")
    theStr = Word_Replacement(theStr, "ônibus", "barramento")
    theStr = Word_Replacement(theStr, "protecção", "proteção")
    theStr = Word_Replacement(theStr, "conduíte", "condutor")
    theStr = Word_Replacement(theStr, "conduto", "condutor")
    theStr = Word_Replacement(theStr, "conduit", "eletroduto")
    theStr = Word_Replacement(theStr, "nominal chain", "rated current")
    theStr = Word_Replacement(theStr, "SECTIONERS", "DISCONNECTORS")
    
    If InStr(LCase(theStr), "default") <> 0 Then
        theStr = Replace(theStr, "Default", "")
        theStr = Replace(theStr, "default", "")
        theStr = theStr & " standard"
    End If
    
    Refine_Trans = theStr
End Function

Public Function Word_Replacement(ByVal theStr As String, ByVal theWord As String, ByVal newWord As String) As String
    theStr = Replace(theStr, LCase(theWord), LCase(newWord))
    theStr = Replace(theStr, UCase(theWord), UCase(newWord))
    theStr = Replace(theStr, UCase(Left(theWord, 1)) & LCase(Right(theWord, Len(theWord) - 1)), _
             UCase(Left(newWord, 1)) & LCase(Right(newWord, Len(newWord) - 1)))
    
    Word_Replacement = theStr
End Function

Public Function ConvertToGet(val As String)
    val = Replace(val, " ", "+")
    val = Replace(val, vbNewLine, "+")
    val = Replace(val, "(", "%28")
    val = Replace(val, ")", "%29")
    ConvertToGet = val
End Function

Public Function Clean(val As String)
    val = Replace(val, "&quot;", """")
    val = Replace(val, "%2C", ",")
    val = Replace(val, "&#39;", "'")
    Clean = val
End Function

Public Function RegexExecute(ByVal str As String, ByVal reg As String, Optional ByVal matchIndex As Long, _
                             Optional ByVal subMatchIndex As Long) As String
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = reg
    regex.Global = Not (matchIndex = 0 And subMatchIndex = 0) 'For efficiency
    If regex.Test(str) Then
        Set matches = regex.Execute(str)
        RegexExecute = matches(matchIndex).SubMatches(subMatchIndex)
        Exit Function
    End If
ErrHandl:
    RegexExecute = CVErr(xlErrValue)
End Function
'-----------------------------------------------------------------------------------------------------------'

'-------------------------------------------- String Functions ---------------------------------------------'
'Public Function Regex_Match(ByVal theStr As String, ByVal theFilter As String, Optional ByVal matchSpace As Boolean = True, _
'                                  Optional ByVal matchCrLf As Boolean = False, Optional ByVal argGlobal As Boolean = True, _
'                                  Optional ByVal argCase As Boolean = True) As Object
'    Set regex = CreateObject("VBScript.RegExp")
'
'    With regex
'      .ignoreCase = argCase
'      .Global = argGlobal
'      .MultiLine = True
'      .Pattern = theFilter
'    End With
'
'    'symStr = Isolate_Symbols(theStr)
'
'    Set matches = regex.Execute(theStr)
'    Set regex = Nothing
'
'    If Not matches Is Nothing And matches.Count <> 0 Then
'        Set Regex_Match = matches
'    Else
'        Set Regex_Match = Nothing
'    End If
'End Function

'Public Function Regex_Match(ByVal theStr As String, ByVal theFilter As String, _
'                            Optional ByRef rgxOut As st_rgxOut, _
'                            Optional ByVal outOpt As String = "") As Integer
'
'    Dim optArr() As String
'
'    Set rgxObj = CreateObject("VBScript.RegExp")
'
'    rgxOpt = Read_Until_Pattern(theFilter, ")", , True)
'
'    If Len(rgxOpt) = 0 Then rgxOpt = outOpt
'
'    If Len(rgxOpt) <> 0 And InStr(rgxOpt, "(") = 0 Then
'
'        theFilter = Replace(theFilter, rgxOpt, "")
'
'        For i = 1 To Len(rgxOpt): Array_Insert optArr, Mid(rgxOpt, i, 1): Next
'    End If
'
'    With rgxObj
'      .ignoreCase = True
'      .Global = True
'      .MultiLine = True
'      .Pattern = theFilter
'    End With
'
'    getCount = False
'    getMatch = False
'    getMatchList = False
'    rgxOut = ""
'
'    If Not IsEmpty(optArr(0)) Then
'        For Each opt In optArr
'            Select Case UCase(opt)
'                Case "I": rgxObj.ignoreCase = False
'                Case "M": rgxObj.MultiLine = False
'                Case "C": getCount = True
'                Case "O": getMatch = True
'                Case "L": getMatchList = True
'            End Select
'        Next
'    End If
'
'    'symStr = Isolate_Symbols(theStr)
'
'    Set matches = rgxObj.Execute(theStr)
'    Set rgxObj = Nothing
'
'    If Not matches Is Nothing And matches.Count <> 0 Then
'        output = InStr(theStr, matches(0))
'
'        If getCount Then rgxOut.matchCount = matches.Count
'        If getMatch Then rgxOut.matchStr = matches(0)
'
'        If getMatchList Then
'            For Each Match In matches: Array_Insert rgxOut.matchArr, Match: Next
'        End If
'    Else
'        output = 0
'    End If
'
'    Regex_Match = output
'End Function

'Function Regex_Replace(ByVal theStr As String, ByVal theFilter As String, _
'                       ByVal repPat As Variant) As String
'
'    Dim rgxOut As st_rgxOut
'
'    isMatch = Regex_Match(theStr, theFilter, rgxOut, "L)")
'
'    If isMatch Then
'        Print_Array rgxOut
'
'        If Not IsArray(repPat) Then
'            For Each Match In rgxOut: theStr = Replace(theStr, Match, repPat): Next Match
'        Else
'            Debug.Print
'        End If
'    End If
'
'    Regex_Replace = theStr
'End Function

Public Function Get_Regex_Object(Optional ByVal theFilt As String = "", _
                                 Optional ByVal rgxOpt As String = "")
    
    Dim rgxObj As New RegExp
    Dim optArr() As String
    
    With rgxObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With
    
    If Len(theFilt) Then
        rgxOpt = Read_Until_Pattern(theFilt, ")", , True)
        
        If Len(rgxOpt) = 0 Then rgxOpt = outOpt
        
        If Len(rgxOpt) <> 0 And InStr(rgxOpt, "(") = 0 Then
            theFilt = Replace(theFilt, rgxOpt, "")
            
            For i = 1 To Len(rgxOpt): Array_Insert optArr, Mid(rgxOpt, i, 1): Next
        End If
        
        If Arr_Len(optArr) Then
            For Each opt In optArr
                Select Case UCase(opt)
                    Case "I": rgxObj.IgnoreCase = False
                    Case "M": rgxObj.MultiLine = False
                End Select
            Next
        End If
    End If
    
    rgxObj.Pattern = theFilt
    
    Set Get_Regex_Object = rgxObj
End Function

Public Function Regex_Match(ByVal theStr As String, ByVal theFilt As String, _
                            Optional ByVal rgxOpt As String = "") As Object
    
    Dim rgxObj As New RegExp
    Dim optArr() As String
    
    Set Regex_Match = Nothing
    
    If Len(theFilt) = 0 Then Exit Function
    
    Set rgxObj = Get_Regex_Object(theFilt, rgxOpt)
    
    theFilt = rgxObj.Pattern
    
    If rgxObj.Test(theStr) Then Set Regex_Match = rgxObj.Execute(theStr)
End Function

Public Function Regex_Replace(ByVal theStr As String, ByVal theFilt As String, _
                              ByVal repStr As Variant) As String
                                  
    Dim rgxObj As New RegExp
    
    Regex_Replace = ""
    newStr = theStr
    
    If Len(theFilt) = 0 Then Exit Function
    
    Set rgxObj = Get_Regex_Object(theFilt, rgxOpt)
    
    If rgxObj.Test(theStr) Then
        If Not IsArray(repStr) Then
            newStr = rgxObj.Replace(theStr, repStr)
        Else
            Set matches = rgxObj.Execute(theStr)
            
            repMaxIdx = Arr_Len(repStr) - 1
            repIdx = 0
            
            For Each m In matches
                newStr = Replace(newStr, m, repStr(repIdx), Count:=1)
                repIdx = repIdx + 1
                
                If repIdx > repMaxIdx Then repIdx = 0
            Next
        End If
    End If
    
    Regex_Replace = newStr
End Function

Public Function Standard_Str(ByVal theStr As String) As String
'    Set matches = Regex_Match(theStr, "[0-9]+([\,\.]{1}[0-9]+)?")
'
'    If Not matches Is Nothing Then
'        For Each Match In matches
'            com2dot = Replace(Match, ",", ".")
'            theStr = Replace(theStr, Match, com2dot)
'        Next Match
'    End If
    
    theStr = Regex_Replace(theStr, "\,\s", " ")
    theStr = Regex_Replace(theStr, "\,", ".")
    theStr = Regex_Replace(theStr, "[\r\n]+", " ")
    theStr = Remove_Extra_Space(theStr)
    'Set matches = Regex_Match(theStr, "((?=[\W]).)+")
    
'    If Not matches is nothing Then
'        For Each Match In matches
'            theStr = Replace(theStr, Match, Replace(Match, " ", ""))
'        Next Match
'    End If
    
    'Acentos e caracteres especiais que serão buscados na theString
    'Você pode definir outros caracteres nessa variável, mas
    ' precisará também colocar a letra correspondente em codiB
    codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
     
    'Letras correspondentes para substituição
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
     
    'Armazena em temp a theString recebida
    temp = theStr
    
    'Loop que irá de andará a theString letra a letra
    For i = 1 To Len(temp)
     
        'InStr buscará se a letra indice i de temp pertence a
        ' codiA e se existir retornará a posição dela
        p = InStr(codiA, Mid(temp, i, 1))
         
        'Substitui a letra de indice i em codiA pela sua
        ' correspondente em codiB
        If p > 0 Then Mid(temp, i, 1) = Mid(codiB, p, 1)
    Next
    
    'Retorna a nova theString
    Standard_Str = LCase(temp)
End Function

Public Function Remove_Extra_Space(ByVal theStr As String) As String
    theStr = Replace(theStr, vbNewLine, " ")
    theStr = Replace(theStr, vbCr, " ")
    theStr = Replace(theStr, vbLf, " ")
    theStr = Replace(theStr, vbCrLf, " ")
    
    While InStr(theStr, "  ") <> 0
        theStr = Replace(theStr, "  ", " ")
    Wend
    
    While Right(theStr, 1) = " "
        theStr = Left(theStr, (Len(theStr) - 1))
    Wend
    
    While Left(theStr, 1) = " "
        theStr = Right(theStr, (Len(theStr) - 1))
    Wend
    
    Remove_Extra_Space = theStr
End Function

Public Function Read_From_Pattern(ByVal theStr As String, ByVal thePat As Variant, _
                                  Optional ByVal numPat As Boolean = False, _
                                  Optional ByVal includePat As Boolean = False) As String
    Read_From_Pattern = ""
    
    If Len(theStr) = 0 Then Exit Function
    
    If numPat And To_Int(thePat) <> 0 Then
        If thePat < 0 Or thePat > Len(theStr) Then Exit Function
        If includePat Then thePat = thePat - 1
        
        newStr = Trim_Left(theStr, thePat)
    Else
        patPos = InStr(theStr, thePat) - 1
        
        If Not includePat Then patPos = patPos + Len(thePat)
        
        newStr = Trim_Left(theStr, patPos)
    End If
    
    Read_From_Pattern = newStr
End Function

Public Function Read_From_Pattern_X(ByVal theStr As String, ByVal thePat As String, Optional ByVal theOS As Integer = 0) As String
    osCount = 0
    
    If theOS <> 0 Then
        auxStr = theStr
        startPos = 0
        
        While osCount <> theOS
            startPos = startPos + InStr(auxStr, thePat) + Len(thePat)
            
            If startPos <> 0 Then
                osCount = osCount + 1
    
            Else
                MsgBox ("Erro na contagem do padrão ou padrão não encontrado")
                
                Application.StatusBar = False

                End
            End If
            
            If Len(theStr) > startPos Then auxStr = Right(theStr, (Len(theStr) - startPos))
        Wend
        
    Else
        startPos = InStrRev(theStr, thePat) + Len(thePat)
    End If
    
    newStr = ""
    
    For i = startPos To Len(theStr)
        newStr = newStr & Mid(theStr, i, 1)
    Next
    
    Read_From_Pattern_X = newStr
End Function

Public Function Read_Until_Pattern_X(ByVal theStr As String, ByVal thePat As String, Optional ByVal theOS As Integer = 0) As String
    newStr = ""
    osCount = 0
    
    If theOS <> 0 Then
        For i = 1 To Len(theStr)
            If Mid(theStr, i, Len(thePat)) = thePat Then osCount = osCount + 1
            If osCount = theOS Then Exit For
            
            newStr = newStr & Mid(theStr, i, 1)
        Next
        
    Else
        endPos = InStrRev(theStr, thePat) - 1
        
        If endPos <> -1 Then
            newStr = Left(theStr, endPos)
            
        Else
            newStr = theStr
        End If
    End If
    
    Read_Until_Pattern_X = newStr
End Function

Public Function Read_Until_Pattern(ByVal theStr As String, ByVal thePat As Variant, _
                                   Optional ByVal numPat As Boolean = False, _
                                   Optional ByVal includePat As Boolean = False) As String
    Read_Until_Pattern = ""
    
    If Len(theStr) = 0 Then Exit Function
    
    If numPat And To_Int(thePat) <> 0 Then
        If thePat < 0 Or thePat > Len(theStr) Then Exit Function
        If Not includePat Then thePat = thePat - 1
        
        newStr = Left(theStr, thePat)
    Else
        patPos = InStr(theStr, thePat)
        
        If includePat Then patPos = patPos + Len(thePat)
        If patPos > 1 Then newStr = Left(theStr, patPos - 1)
    End If
    
    Read_Until_Pattern = newStr
End Function

Public Function Read_Delim_Str_X(ByVal theStr As String, ByVal startPat As String, ByVal endPat As String, Optional ByVal startOS As Integer = 0, Optional ByVal endOS As Integer = 0) As String
    endOS = endOS - startOS
    newStr = Read_From_Pattern_X(theStr, startPat, startOS)
    newStr = Read_Until_Pattern_X(newStr, endPat, endOS)
    
    Read_Delim_Str_X = newStr
End Function

Public Function Read_Delim_Str(ByVal theStr As String, ByVal startPat As Variant, ByVal endPat As Variant, Optional ByVal numPat As Boolean = False) As String
    newStr = ""
    
    If InStr(theStr, startPat) = 0 Then
        Read_Delim_Str = ""
        
        Exit Function
    End If
    
    If numPat = False Then
        startPos = InStr(theStr, startPat) + Len(startPat)
        
        For i = startPos To Len(theStr)
            If Mid(theStr, i, Len(endPat)) = endPat Then Exit For
            
            newStr = newStr & Mid(theStr, i, 1)
        Next
        
    ElseIf startPat < endPat And startPat > 0 Then
        newStr = Mid(theStr, startPat, endPat - startPat)
    End If
    
    Read_Delim_Str = newStr
End Function

Public Function Isolate_Symbols(ByVal theStr As String, Optional ByVal matchSpc As Boolean = False) As String
    theSpc = "\s"
    newStr = ""
    
    If matchSpc Then theSpc = ""
    
    Set matches = Regex_Match(theStr, "[\w" & latStart & "-" & latEnd & theSpc & "]*")
    
    If Not matches Is Nothing Then
        For Each Match In matches
            For i = 1 To Len(Match): newStr = Replace(newStr, Mid(Match, i, 1), ""): Next i
        Next Match
        
        If matchSpc = False Then Remove_Extra_Space (newStr)
        
        If newStr <> "" Then
            Isolate_Symbols = newStr
        Else
            Isolate_Symbols = Null
        End If
    Else
        Isolate_Symbols = theStr
    End If
End Function

Public Function Isolate_Letters(ByVal theStr As String, Optional ByVal minChars As Integer = 0, Optional ByVal getWords As Boolean = True, Optional ByVal trimStr As Boolean = True) As Variant
    Set matches = Regex_Match(theStr, "[0-9\W]")
    
    newStr = ""
    theSpc = ""
    
    If getWords Then theSpc = " "
    
    If Not matches Is Nothing Then
        newStr = theStr
        
        For Each Match In matches
            For i = 1 To Len(Match): newStr = Replace(newStr, Mid(Match, i, 1), theSpc): Next i
        Next Match
        
        If minChars Then
            Set matches = Regex_Match(newStr, "\b" & allChars & "{1," & minChars - 1 & "}\b")
            
            If Not matches Is Nothing Then
                For Each Match In matches
                    newStr = Replace(newStr, " " & Match & " ", " ")
                Next
            End If
        End If
        
        If Not getWords Then
            newStr = Replace(newStr, " ", "")
        ElseIf trimStr Then
            newStr = Remove_Extra_Space(newStr)
        End If
        
        If getWords Then newStr = Replace(newStr, " ", splitChar)
        
        If newStr <> "" Then
            Isolate_Letters = newStr
        Else
            Isolate_Letters = Null
        End If
    Else
        Isolate_Letters = theStr
    End If
End Function

Public Function Isolate_Numbers(ByVal theStr As String) As Variant
    If InStr(theStr, ",") And InStr(theStr, ".") Then
        theStr = Replace(theStr, ".", "")
        theStr = Replace(theStr, ",", ".")
    ElseIf InStr(theStr, ",") Then
        theStr = Replace(theStr, ",", ".")
    End If
    
    Set matches = Regex_Match(theStr, "\-*[0-9]+(\.[0-9]+\%?)?")
    
    newStr = ""
    
    If Not matches Is Nothing Then
        For Each Match In matches
            newStr = newStr & Match & splitChar
        Next Match
    
        If newStr <> "" Then
            newStr = Trim_Right(newStr, 1)
            Isolate_Numbers = Remove_Extra_Space(newStr)
        Else
            Isolate_Numbers = Null
        End If
    Else
        Isolate_Numbers = theStr
    End If
End Function

Public Function Remove_Numbers(ByVal theStr As String, Optional ByVal trimStr As Boolean = False) As Variant
    Set matches = Regex_Match(theStr, "[0-9]+(\.[0-9]+\%?)?")
    newStr = ""
    
    If Not matches Is Nothing Then
        newStr = theStr
        
        For Each Match In matches
            newStr = Replace(newStr, Match, "")
        Next Match
    
        If trimStr = True Then newStr = Remove_Extra_Space(newStr)
        
        If newStr <> "" Then
            newStr = Replace(newStr, " . ", " ")
            Remove_Numbers = newStr
        Else
            Remove_Numbers = Null
        End If
    Else
        Remove_Numbers = theStr
    End If
End Function

Public Function Remove_Symbols(ByVal theStr As String, Optional ByVal matchSpc As Boolean = True) As String
    sym = Isolate_Symbols(theStr, Not matchSpc)
    newStr = theStr
    
    If Not IsNull(sym) Then
        For i = 1 To Len(sym): newStr = Replace(newStr, Mid(sym, i, 1), "")
        
        If matchSpc Then newStr = Remove_Extra_Space(newStr)
        
        If newStr <> "" Then
            Remove_Symbols = newStr
        Else
            Remove_Symbols = Null
        End If
    Else
        Remove_Symbols = theStr
    End If
End Function

'Public Function Remove_Words(ByVal theStr As String) As String
'    set matches = Regex_Match(theStr, "((?!\s)(?=[\W]).)+")
'
'    If Not matches is nothing Then For Each Match In matches: theStr = Replace(theStr, Match, ""): Next Match
'
'    Remove_Words = theStr
'End Function
'
'Public Function Remove_Numbers(ByVal theStr As String) As String
'    set matches = Regex_Match(theStr, "((?!\s)(?=[\W]).)+")
'
'    If Not matches is nothing Then For Each Match In matches: theStr = Replace(theStr, Match, ""): Next Match
'
'    Remove_Numbers = theStr
'End Function

Public Function Trim_Left(ByVal theStr As String, ByVal theCount As Integer) As String
    If Len(theStr) < theCount Then theCount = Len(theStr)
    If theStr <> "" Then Trim_Left = Right(theStr, (Len(theStr) - theCount))
End Function

Public Function Trim_Right(ByVal theStr As String, ByVal theCount As Integer) As String
    If Len(theStr) < theCount Then theCount = Len(theStr)
    If theStr <> "" Then Trim_Right = Left(theStr, (Len(theStr) - theCount))
End Function

Public Function Count_Pattern_Occurrences(ByVal theStr As String, ByVal thePat As String) As Integer
    numPats = 0
    patOccur = InStr(theStr, thePat)
    
    Do While patOccur <> 0
        numPats = numPats + 1
        theStr = Read_From_Pattern_X(theStr, thePat, 1)
        patOccur = InStr(theStr, thePat)
    Loop
    
    Count_Pattern_Occurrences = numPats
End Function

Public Function Char_Match_Count(ByVal theStr As String, ByVal matchStr As String) As Integer
    matchCount = 0
    
    If InStr(theStr, " ") = 0 Then theStr = theStr & " " & "tmp"
    If InStr(matchStr, " ") = 0 Then matchStr = matchStr & " " & "tmp"
    
    theStrArr = Split(Remove_Extra_Space(theStr), " ")
    matchStrArr = Split(Remove_Extra_Space(matchStr), " ")
    
    For i = 0 To Min(theStrArr, matchStrArr) - 1
        strLen = Min(theStrArr(i), matchStrArr(i))
        
        For c = 1 To strLen
            If LCase(Mid(theStrArr(i), c, 1)) = LCase(Mid(matchStrArr(i), c, 1)) Then matchCount = matchCount + 1
        Next c
        
        matchCount = matchCount + 1
    Next i
    
    Char_Match_Count = matchCount
End Function
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------- Document Functions --------------------------------------------'
Public Sub Protect_Sheets(ByVal thePass As String, ByVal toggle As Boolean, Optional ByVal theWB As Workbook = Nothing)
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    If toggle Then
        For Each ws In theWB.Worksheets
            If Not ws.ProtectContents Then ws.Protect Password:=thePass, DrawingObjects:=True, contents:=True, AllowFiltering:=True
        Next
        
        theWB.Protect thePass
    Else
        For Each ws In theWB.Worksheets
            If ws.ProtectContents Then ws.Unprotect Password:=thePass
        Next
        
        theWB.Unprotect thePass
    End If
End Sub

Public Function Append_To_Sheet(ByVal wsName As String, Optional ByVal theWB As Workbook = Nothing) As Worksheet
    ''Application.ScreenUpdating = False

    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    If Check_Sheet(wsName) = True Then
        Set theWS = theWB.Worksheets(wsName)
    Else
        Set theWS = theWB.Worksheets.Add
        
        theWS.Name = wsName
    End If
    
    Set Append_To_Sheet = theWS
End Function

Public Function Reset_Workbook(ByVal wbName As String, _
                               Optional ByVal thePath As String = "") As Workbook
    
    If Regex_Match(wbName, "\.xl[a-z]{1,2}$") Is Nothing Then _
        wbName = wbName & ".xlsx"
    
    setResetADA = False
    
    If Application.DisplayAlerts Then setResetADA = True
    
    Application.DisplayAlerts = False
    
    For Each wb In Workbooks
        If wb.Name = wbName Then
            If thePath = "" Then thePath = wb.Path
            
            theName = wb.Name
            
            wb.Close
            
            Kill thePath & "\" & theName
        End If
    Next
    
    If thePath = "" Then thePath = Application.ActiveWorkbook.Path
    
    Set theWB = Workbooks.Add
    
    theWB.SaveAs thePath & "\" & wbName
    
    If setResetADA Then Application.DisplayAlerts = True
    
    Set Reset_Workbook = theWB
End Function

Public Sub Remove_Sheet(ByVal wsName As String, Optional ByVal theWB As Workbook)
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    If Check_Sheet(wsName, theWB) Then
        Application.DisplayAlerts = False
        
        If theWB.Worksheets.Count = 1 Then
            If Not Check_Sheet("TMP_SHT", theWB) Then
                Set tmpSht = theWB.Worksheets.Add
                
                tmpSht.Name = "TMP_SHT"
            End If
            
            theWB.Worksheets(wsName).Visible = xlSheetVisible
            theWB.Worksheets(wsName).Delete
        Else
            theWB.Worksheets(wsName).Visible = xlSheetVisible
            theWB.Worksheets(wsName).Delete
        End If
        
        Application.DisplayAlerts = True
    End If
End Sub

Public Function Reset_Sheet(ByVal wsName As String, Optional ByVal byCopy As Boolean = False, Optional ByVal copyWS As Worksheet = Nothing, _
                            Optional ByVal refPos As String = "before", Optional ByVal refWS As Worksheet = Nothing, Optional ByVal refWB As Workbook) As Worksheet
    ''Application.ScreenUpdating = False
    If refWB Is Nothing Then Set refWB = ActiveWorkbook
    
    If Check_Sheet(wsName, refWB) = True Then Remove_Sheet wsName, refWB
    
    If Not byCopy Then
        If Not refWS Is Nothing Then
            If LCase(refPos) = "before" Then
                Set theWS = refWB.Worksheets.Add(before:=refWS)
            ElseIf refPos = "after" Then
                Set theWS = refWB.Worksheets.Add(After:=refWS)
            Else
                MsgBox "Erro ao resetar a planilha " & wsName & vbCrLf & "Posição de referência inválida."
                
                End
            End If
        Else
            Set theWS = refWB.Worksheets.Add(After:=refWB.Worksheets(refWB.Worksheets.Count))
        End If
        
        theWS.Name = wsName
        
    ElseIf Not copyWS Is Nothing Then
        Application.DisplayAlerts = False
        
        If Not refWS Is Nothing Then
            If refPos = "before" Then
                copyWS.Copy before:=refWS
                
                Set theWS = ActiveSheet
            ElseIf refPos = "after" Then
                copyWS.Copy After:=refWS
                
                Set theWS = ActiveSheet
            Else
                MsgBox "Erro ao resetar a planilha " & wsName & vbCrLf & "Posição de referência inválida."
                
                End
            End If
        
        Else
            copyWS.Copy before:=copyWS
                
                Set theWS = ActiveSheet
        End If
        
        Application.DisplayAlerts = True
        theWS.Name = wsName
        
    Else
        MsgBox ("Erro ao resetar a planilha " & wsName & vbCrLf & "Não foi fornecido o nome da planilha a ser copiada.")

        End
    End If
    
    Remove_Sheet "TMP_SHT"
    
    Set Reset_Sheet = theWS
End Function

Public Function Check_Sheet(ByVal wsName As String, Optional ByVal theWB As Workbook) As Boolean
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    IsExist = False
    
    For Each ws In theWB.Worksheets
        If ws.Name = wsName Then
            IsExist = True
            
            Exit For
        End If
    Next ws
    
    Check_Sheet = IsExist
End Function

Public Function Find_Doc_By_Pattern(ByVal docLvl As Integer, ByVal thePat As String, Optional ByVal refDoc As Variant, _
                                    Optional excludeList As Variant, Optional ByVal hdlErr As Boolean = False, _
                                    Optional regexSearch As Boolean = False) As Variant
    Set theDoc = Nothing
    
    If Not IsMissing(excludeList) And Not IsArray(excludeList) Then
        MsgBox "Erro em ""Find_Doc_By_Pattern"": lista de exclusão não é um vetor"
        
        End
    End If
    
    If docLvl = wbLvl Then 'Workbook
        For Each wb In Workbooks
            nameMatch = False
            
            If (Not regexSearch And InStr(wb.Name, thePat) <> 0) Or (regexSearch And Not Regex_Match(wb.Name, thePat) Is Nothing) Then _
                nameMatch = True
                
            If nameMatch And (IsMissing(excludeList) Or Not IsFound(Find_In_Array(excludeList, wb))) Then
                Set theDoc = wb
                
                If Not IsMissing(refDoc) Then
                    Set theWS = Find_Doc_By_Pattern(wsLvl, refDoc, theDoc, excludeList)
                    
                    If theWS Is Nothing Then
                        Set theDoc = Nothing
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            End If
        Next
        
    Else 'Worksheet
        If IsMissing(refDoc) Then Set refDoc = ActiveWorkbook
        
        If Not IsObject(refDoc) Then
            Set refDoc = Find_Doc_By_Pattern(wbLvl, refDoc, thePat)
            
            If refDoc Is Nothing Then Set refDoc = ActiveWorkbook
        End If
        
        For Each ws In refDoc.Worksheets
            nameMatch = False
            
            If (Not regexSearch And InStr(ws.Name, thePat) <> 0) Or (regexSearch And Not Regex_Match(ws.Name, thePat) Is Nothing) Then _
                nameMatch = True
            
            If nameMatch And (IsMissing(excludeList) Or Not IsFound(Find_In_Array(excludeList, ws))) Then
                Set theDoc = ws
                
                Exit For
            End If
        Next
    End If
    
    If theDoc Is Nothing And hdlErr = True Then
        MsgBox "Erro: o padrão """ & thePat & """ não foi encontrado"
        Execute_Caller_Termination refDoc
        
        End
    End If
    
    Set Find_Doc_By_Pattern = theDoc
End Function

Public Function Count_Docs_By_Pattern(ByVal docLvl As Integer, ByVal thePat As String, Optional ByVal dpSch As Boolean = False, _
                                      Optional ByVal refStr As String = "", Optional ByVal getNames As Boolean = False) As Variant
    docList = ""
    docCount = 0
    
    Do
        Set theDoc = Find_Doc_By_Pattern(docLvl, thePat, dpSch, refStr, docList)
        
        If Not theDoc Is Nothing Then
            docList = docList & theDoc.Name & splitChar
            docCount = docCount + 1
        End If
    Loop While Not theDoc Is Nothing
    
    Count_Docs_By_Pattern = docCount
    
    If getNames = True Then Count_Docs_By_Pattern = Trim_Right(docList, 1)
End Function

Public Function Check_Workbook(ByVal theWB As String) As Boolean
    foundWB = False
    
    For Each wb In Workbooks
        If wb.Name = theWB Then
            foundWB = True
            
            Exit For
        End If
    Next wb
    
    Check_Workbook = foundWB
End Function

Public Function Count_Files_In_Folder(ByVal thePath As String, ByVal theExt As String) As Integer
    If Right(thePath, 1) <> "\" Then thePath = thePath & "\"
    
    fileList = Dir(thePath & theExt)
    fileCount = 0
    
    Do While fileList <> ""
        fileCount = fileCount + 1
        fileList = Dir
    Loop
    
    Count_Files_In_Folder = fileCount
End Function

Sub Isolate_Sheet(ByVal theWS As Worksheet, Optional ByVal theWB As Workbook = Nothing)
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    Application.DisplayAlerts = False
    
    For Each ws In theWB.Worksheets
        If ws.Name <> theWS.Name Then ws.Delete
    Next
    
    Application.DisplayAlerts = True
End Sub

Public Function Get_Max_Row_Level(ByVal theWS As Worksheet) As Integer
    maxLvl = 0
    
    For Each Row In theWS.UsedRange.Rows
        If Row.OutlineLevel > maxLvl Then maxLvl = Row.OutlineLevel
    Next Row
    
    Get_Max_Row_Level = maxLvl
End Function

Public Function Get_Max_Col_Level(ByVal theWS As Worksheet) As Integer
    maxLvl = 0
    
    For Each col In theWS.UsedRange.Columns
        If col.OutlineLevel > maxLvl Then maxLvl = col.OutlineLevel
    Next col
    
    Get_Max_Col_Level = maxLvl
End Function

Public Sub Expand_Sheet_Groups(ByVal theWS As Worksheet)
    theWS.Outline.ShowLevels RowLevels:=Get_Max_Row_Level(theWS), _
                             ColumnLevels:=Get_Max_Col_Level(theWS)
End Sub

Public Sub Shrink_Sheet_Groups(ByVal theWS As Worksheet)
    theWS.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Sub

Public Function Get_Dir_Unique_Files(ByVal thePath As String, ByVal theExt As String) As String
    If Right(thePath, 1) <> "\" Then thePath = thePath & "\"
    
    dirFile = Dir(thePath & theExt)
    fileList = ""
    dupFileList = ""
    uniFileList = ""
    fileCount = 0
    thisExt = Trim_Left(theExt, 1)
    
    Do While dirFile <> ""
        If theExt = "*.*" Then thisExt = "." & Read_From_Pattern_X(dirFile, ".")
        
        theFile = Replace(dirFile, thisExt, "")
        
        If InStr(fileList, theFile) Then
            dupFileList = dupFileList & theFile & splitChar
        Else
            fileList = fileList & theFile & splitChar
        End If
        
        dirFile = Dir
    Loop
    
    If fileList <> "" Then fileArr = Split(fileList, splitChar)
    
    For Each F In fileArr
        If InStr(dupFileList, F) = 0 Then uniFileList = uniFileList & F & vbCrLf
    Next F
    
    Get_Dir_Unique_Files = uniFileList
End Function

Public Function Get_File(Optional ByVal thePath As String = "", Optional ByVal extDesc As String = "All Files", Optional ByVal theExt As String = "*.*", _
                         Optional ByVal endOnCancel As Boolean = True) As String
    Dim fileDiag As FileDialog
    
    Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)

    If Len(thePath) = 0 Then thePath = ActiveWorkbook.Path
    
    fileDiag.InitialFileName = thePath
    fileDiag.InitialView = msoFileDialogViewList
    fileDiag.AllowMultiSelect = False
    
    fileDiag.Filters.Clear
    fileDiag.Filters.Add extDesc, theExt
    
    filePick = fileDiag.Show
    
    If filePick = 0 Then
        If Not endOnCancel Then Exit Function
        
        MsgBox "Processo abortado."
        
        End
    End If
    
    theFile = fileDiag.SelectedItems(1)
    
    Set fileDiag = Nothing
    
    Get_File = theFile
End Function

Public Function Get_Folder(Optional ByVal startFolder As String = "") As String
    Set folderObj = Application.FileDialog(msoFileDialogFolderPicker)
    
    If startFolder = "" Then startFolder = ActiveWorkbook.Path
    
    With folderObj
        .Title = "Selecione uma Pasta"
        .AllowMultiSelect = False
        .InitialFileName = startFolder
        
        If .Show <> -1 Then GoTo GET_FOLDER_END
        
        theFolder = .SelectedItems(1)
    End With
    
GET_FOLDER_END:
    Get_Folder = theFolder
    
    Set folderObj = Nothing
End Function

Private Sub Execute_Caller_Termination(ByVal theDoc As Variant)
    If theDoc.Parent.Name = "Microsoft Excel" Then
        Set theWB = theDoc
    Else
        Set theWB = theDoc.Parent
    End If
    
    'endSub = Find_VBA_Routine(theWB, "terminate.*program", True)
    endSub = Find_VBA_Routine(theWB, "Terminate_Program")
    
    If Len(endSub) Then Application.Run "'" & theWB.Name & "'!" & endSub, True, "Programa terminado por erro", True
End Sub

Public Function Find_VBA_Routine(ByVal theWB As Workbook, ByVal routName As String, _
                                 Optional ByVal regexSearch As Boolean = False) As String
    routFoundName = ""
    
    For Each comp In theWB.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            Set codeMod = comp.CodeModule
            
            lineNum = 1
            
            Do While lineNum < codeMod.CountOfLines
                procName = codeMod.ProcOfLine(lineNum, vbext_pk_Proc)
                
                If Len(procName) Then
                    If (InStr(procName, routName) And Not regexSearch) Or _
                    (regexSearch And Not Regex_Match(procName, routName) Is Nothing) Then
                        routFoundName = procName
                        
                        Exit Do
                    End If
                    
                    procOldName = procName
                    
                    Do While procName = procOldName And lineNum < codeMod.CountOfLines
                        lineNum = lineNum + 1
                        procName = codeMod.ProcOfLine(lineNum, vbext_pk_Proc)
                    Loop
                End If
                
                lineNum = lineNum + 1
            Loop
        End If
    Next
    
    Find_VBA_Routine = routFoundName
End Function

Public Function Backup_Workbook(Optional ByVal theWB As Workbook = Nothing) As String
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    wbBkpFullName = theWB.Path & "\BACKUP - " & theWB.Name
    
    Application.DisplayAlerts = False
    
    theWB.SaveCopyAs wbBkpFullName
    
    Application.DisplayAlerts = True
    
    Backup_Workbook = wbBkpFullName
End Function

Public Sub Delete_Workbook(ByVal theWB As Variant)
    If IsObject(theWB) Then
        wbFullName = theWB.FullName
    ElseIf InStr(theWB, "\") = 0 Then
        wbFullName = ActiveWorkbook.Path & "\" & theWB
    Else
        wbFullName = theWB
    End If
    
    wbName = Read_From_Pattern_X(wbFullName, "\")
    
    If Check_Workbook(wbName) Then Workbooks(wbName).Close False
    
    Application.DisplayAlerts = False
    
    Kill wbFullName
    
    Application.DisplayAlerts = True
End Sub
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------ Sheet Cells Functions ------------------------------------------'
Public Function Check_Hidden_Columns(ByVal theSht As Worksheet) As String
    hidColList = ""
    
    For Each col In theSht.UsedRange.Columns
        If col.Hidden = True Then
            theCol = Split(theSht.Cells(1, col.Column).Address, "$")(1)
            hidColList = hidColList & theCol & splitChar
        End If
    Next col
    
    Check_Hidden_Columns = hidColList
End Function

Public Sub Show_Columns(ByVal colList As String, ByVal theSht As Worksheet, ByVal showCols As Boolean)
    colArr = Split(Trim_Right(colList, 1), splitChar)
    
    For Each col In colArr
        theSht.Columns(col & ":" & col).Hidden = Not showCols
    Next col
End Sub

Public Function Get_Hidden_Column_Range(ByVal theSht As Worksheet, ByVal rngIdx As Integer) As String
    fstCol = theSht.UsedRange.Cells(1, 1).Column
    numCols = theSht.UsedRange.Columns.Count
    
    For i = 1 To rngIdx
        colHidStart = ""
        colHidEnd = ""
        j = fstCol
        
        Do
            theCol = Split(theSht.Cells(1, j).Address, "$")(1)
            hidCol = theSht.Columns(theCol).Hidden
            j = j + 1
        Loop While hidCol = False And j < numCols
        
        If hidCol = True Then colHidStart = theCol

        Do
            theCol = Split(theSht.Cells(1, j).Address, "$")(1)
            hidCol = theSht.Columns(theCol).Hidden
            j = j + 1
        Loop While hidCol = True And j < numCols
        
        If hidCol = False Then colHidEnd = theCol
        
        fstCol = Split(Cells(1, colHidEnd).Offset(0, 1).Address, "$")(1)
    Next i
    
    If colHidStart = "" Or colHidEnd = "" Then
        Get_Hidden_Column_Range = "null"
        
    Else
        Get_Hidden_Column_Range = colHidStart & ":" & colHidEnd
    End If
End Function

Public Function Find_N_Get_Cell(ByVal theStr As Variant, Optional ByVal theWS As Worksheet, _
                                Optional ByVal theRng As Range = Nothing, _
                                Optional ByVal ref As Variant, _
                                Optional ByVal rgxSch As Boolean = False, _
                                Optional ByVal LkIn As Integer = xlValues, _
                                Optional ByVal LkAt As Integer = xlPart, _
                                Optional ByVal SchDir As Integer = xlNext, _
                                Optional ByVal mCase As Boolean = False, _
                                Optional ByVal fixCol As Boolean = False, _
                                Optional ByVal findLast As Boolean = False, _
                                Optional hdlErr As Boolean = False) As Range
        
    If rgxSch Then
        theFilt = theStr
        
        If InStr(theFilt, "|") Then
            theStr = Split(theFilt, "|")
        Else
            theStr = Remove_Extra_Space(Regex_Replace(theStr, "[\(\)\^\$\+\?]|\.\*|\\s", " "))
        End If
    End If
    
    If IsArray(theStr) Then
        For Each s In theStr
            Set theCell = Find_N_Get_Cell(s, theWS, theRng, ref, rgxSch, LkIn, LkAt, SchDir, _
                          mCase, fixCol, findLast, hdlErr)
            
            If Not theCell Is Nothing Then
                Set Find_N_Get_Cell = theCell
                
                Exit Function
            End If
        Next
    End If
    
    Set theCell = Nothing
    
    If theWS Is Nothing Then Set theWS = ActiveSheet
    If theRng Is Nothing Then Set theRng = theWS.UsedRange.Cells
    If IsMissing(ref) Then Set ref = theRng.Cells(1, 1)
    
    If IsObject(ref) Then
        If ref.Cells.Count = 1 Then
            If fixCol = True Then
                theCol = Split(ref.Address, "$")(1)
            
                Set theRng = theWS.Columns(theCol & ":" & theCol)
            End If
        Else
            Set theRng = ref
        End If
        
        On Error Resume Next
        
        If rgxSch Then
            For Each cell In theRng
                Set cellMatch = Regex_Match(cell.Value, theFilt)
                
                If Not cellMatch Is Nothing Then
                    Set theCell = cell
                    
                    Exit For
                End If
            Next
        Else
            Set theCell = theRng.Find(What:=theStr, After:=ref, LookIn:=LkIn, LookAt:=LkAt, _
                                      MatchCase:=mCase, SearchDirection:=SchDir)
        End If
        
        On Error GoTo 0
        
    Else
        Set theCell = theRng.Find(What:=theStr, After:=theWS.Cells(1, 1), LookIn:=LkIn, _
                                  LookAt:=LkAt, MatchCase:=mCase, SearchDirection:=SchDir)
        
        ref = val(ref)
        idx = 1
        
        Set fstCell = theCell
        
        Do While idx < ref
            If rgxSch Then
                Set cellMatch = Regex_Match(theCell.Value, theFilt)
                
                If cellMatch Is Nothing Then
                    Set fstCell = theCell
                    
                    Do
                        Set theCell = theRng.FindNext(theCell)
                        Set cellMatch = Regex_Match(theCell.Value, theFilt)
                    Loop While cellMatch Is Nothing And theCell.Row <> fstCell.Row
                End If
            End If
            
            Set theCell = theRng.FindNext(theCell)
            
            idx = idx + 1
        Loop
    End If
    
    If theCell Is Nothing And hdlErr = True Then
        MsgBox "Erro:""" & theStr & """ não encontrado na planilha""" & theWS.Name & """."
        Execute_Caller_Termination theWS
        
        End
    ElseIf Not theCell Is Nothing And findLast Then
        startRow = theCell.Row
        
        Do
            Set lastCell = theCell
            Set theCell = theRng.FindNext(theCell)
        Loop While theCell.Row <> startRow
        
        Set theCell = lastCell
    End If
    
    Set Find_N_Get_Cell = theCell
End Function

Sub Add_Comment(ByVal theCell As Range, ByVal theStr As String, Optional ByVal cmtSets As Variant)
    If Not theCell.Comment Is Nothing Then theCell.Comment.Delete
    
    theCell.AddComment theStr
    
    theCell.Comment.Visible = False
    
    numSets = Arr_Len(cmtSets)
    s = 0
    
    Do While s < numSets - 1
        setName = LCase(cmtSets(s))
        setVal = cmtSets(s + 1)
        
        If IsNumeric(setVal) And VarType(setVal) <> vbBoolean Then setVal = To_Double(setVal)
        
        If Not IsNull(Isolate_Numbers(setName)) Then 'Or (Not IsNumeric(setVal) And VarType(setVal) <> vbBoolean) Then
            MsgBox "Error: settings array must contain a pair (""setting name"",""setting value"")"
            
            End
        End If
        
        With theCell.Comment
            Select Case setName
                Case "width": .Shape.Width = setVal
                Case "height": .Shape.Height = setVal
                Case "fbold": .Shape.TextFrame.Characters.Font.Bold = setVal
                Case "fcolor": .Shape.TextFrame.Characters.Font.Color = setVal
                Case "fsize": .Shape.TextFrame.Characters.Font.Size = setVal
                Case "fname": .Shape.TextFrame.Characters.Font.Name = setVal
                Case "autoscale":
                    If setVal = True Then
                        fsize = .Shape.TextFrame.Characters.Font.Size
                        strArr = Split(theStr, vbCrLf)
                        maxLen = 0
                        
                        For Each elm In strArr
                            If Len(elm) > maxLen Then maxLen = Len(elm)
                        Next elm
                        
                        .Shape.Width = fsize * maxLen / 1.9 '1.34
                        .Shape.Height = fsize * (Arr_Len(strArr) + 1) * 1.1
                    End If
            End Select
        End With
        
        s = s + 2
    Loop
End Sub

Public Function Get_Col_Char(ByVal colRef As Variant) As String
    If IsObject(colRef) Then
        Get_Col_Char = Split(colRef.Address, "$")(1)
    Else
        Get_Col_Char = Split(Cells(1, colRef).Address, "$")(1)
    End If
End Function

Public Function Sheet_Rows(Optional ByVal theWS As Worksheet = Nothing) As Integer
    If theWS Is Nothing Then Set theWS = ActiveSheet
    
    Sheet_Rows = theWS.UsedRange.Rows.Count
End Function

Public Function Sheet_Cols(Optional ByVal theWS As Worksheet = Nothing) As Integer
    If theWS Is Nothing Then Set theWS = ActiveSheet
    
    Sheet_Cols = theWS.UsedRange.Columns.Count
End Function

Public Function Sheet_Last_Col(Optional ByVal theWS As Worksheet = Nothing) As String
    If theWS Is Nothing Then Set theWS = ActiveSheet
    
    Sheet_Last_Col = Get_Col_Char(theWS.UsedRange.Columns.Count)
End Function

Public Sub Copy_Exact_Formulas()
    On Error Resume Next
    Set rng1 = Application.InputBox("Intervalo de células de origem", Type:=8)
    
    If Not IsEmpty(rng1) And rng1.Columns.Count = 1 Then
        Set rng2 = Application.InputBox("Intervalo de células de destino", Type:=8)
    Else
        End
    End If
    
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    If Not IsEmpty(rng2) And rng2.Columns.Count = 1 Then
        numCells1 = rng1.Rows.Count
        numCells2 = rng2.Rows.Count
        
        If numCells1 = numCells2 Then
            For i = 1 To numCells1
                rng2.Cells(i, 1).Formula = rng1.Cells(i, 1).Formula
            Next
        End If
    Else
        End
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Pronto!"
End Sub

Public Sub Copy_Values()
    On Error Resume Next
    
    Set srcRng = Application.InputBox("Origem:", Type:=8)
    
    If Not IsObject(srcRng) Then End
    
    Set destCell = Application.InputBox("Destino:", Type:=8)
    
    If Not IsObject(destCell) Then End
    
    Application.ScreenUpdating = False
    
    For Each cell In srcRng
        rowOS = cell.Row - srcRng.Cells(1, 1).Row
        colOS = cell.Column - srcRng.Cells(1, 1).Column
        
        On Error Resume Next
        
        If Len(cell.Value) Then
            destCell.Offset(rowOS, colOS + 3).Value = cell.Offset(0, 3).Value
            destCell.Offset(rowOS, colOS).Value = cell.Value
        End If
        
        On Error GoTo 0
    Next
    
    Application.ScreenUpdating = True
End Sub

Public Sub Copy_Formulas()
    On Error Resume Next
    
    Set srcRng = Application.InputBox("Origem:", Type:=8)
    
    If Not IsObject(srcRng) Then End
    
    Set destCell = Selection 'Application.InputBox("Destino:", Type:=8)
    
    If Not IsObject(destCell) Then End
    
    If destCell.Parent.Parent.Name = ThisWorkbook.Name Then
        MsgBox "Quase deu merda"
        
        End
    End If
    
    Application.ScreenUpdating = False
    
    For Each cell In srcRng
        rowOS = cell.Row - srcRng.Cells(1, 1).Row
        colOS = cell.Column - srcRng.Cells(1, 1).Column
        
        On Error Resume Next
        
        If Len(cell.Formula) Then
            destCell.Offset(rowOS, colOS - 4).Value = cell.Offset(0, -4).Value
            destCell.Offset(rowOS, colOS).Formula = cell.Formula
        End If
        
        On Error GoTo 0
    Next
    
    Application.ScreenUpdating = True
End Sub

Public Sub Copy_Values_N_Apply_Factor()
    On Error Resume Next
    
    Set srcRng = Application.InputBox("Origem:", Type:=8)
    
    If Not IsObject(srcRng) Then End
    
    Set destCell = Selection 'Application.InputBox("Destino:", Type:=8)
    
    If Not IsObject(destCell) Then End
    
    If destCell.Parent.Parent.Name = ThisWorkbook.Name Then
        MsgBox "Quase deu merda!"
        
        End
    End If
    
'    Set factorCell = Application.InputBox("Fator:", Type:=8)
'
'    If Not IsObject(factorCell) Then End
    
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each cell In srcRng
        rowOS = cell.Row - srcRng.Cells(1, 1).Row
        colOS = 3 'cell.Column - srcRng.Cells(1, 1).Column
        
        On Error Resume Next
        
        destCell.Offset(rowOS, 0).Value = cell.Value
        destCell.Offset(rowOS, colOS).Formula = "='[" & cell.Parent.Parent.Name & "]" & cell.Parent.Name & "'!" & cell.Offset(0, 2).Address & _
                                                "*'[Serviços em Campo - Eletronorte.xlsm]Resumo - SC SE XXX'!$B$36"
                                                '"*'[" & factorCell.Parent.Parent.Name & "]" & factorCell.Parent.Name & "'!" & factorCell.Address
        
        On Error GoTo 0
    Next
    
    Application.ScreenUpdating = True
    
    destCell.Offset(rowOS + 1, 0).Activate
End Sub

Private Function Search_Valid_Cell(ByVal refCell As Range, ByVal delimiter As Integer, _
                                   ByVal criteria As String, ByVal rgxSch As Boolean, _
                                   ByVal invCriteria As Boolean, ByVal invDir As Boolean, _
                                   ByVal schDim As Integer, ByVal secondPass As Boolean)
    Dim regexOutput As Variant
    
    rowOS = 0
    colOS = 0
    delimiter = delimiter + 1
    
    If schDim = 1 Then
        rowOS = 1
    Else
        colOS = 1
    End If
    
    If invDir Then
        rowOS = -rowOS
        colOS = -colOS
        delimiter = 1
    End If
    
    Set lastCell = refCell
    
    Do
        If Len(criteria) Then
            If rgxSch Then
                found = Not Regex_Match(lastCell.Value, criteria) Is Nothing
            Else
                found = InStr(lastCell.Value, criteria) <> 0
            End If
        Else
            found = Len(lastCell.Value) <> 0
        End If
        
        If invCriteria Then found = Not found
        If secondPass Then found = Not found
        
        If schDim = 1 Then
            cond = found Or (lastCell.Row - delimiter) = 0
        Else
            cond = found Or (lastCell.Column - delimiter) = 0
        End If
        
        If Not cond Then Set lastCell = lastCell.Offset(rowOS, colOS)
    Loop While Not cond
    
    If secondPass Then Set lastCell = lastCell.Offset(-rowOS, -colOS)
    
    Set Search_Valid_Cell = lastCell
End Function

Public Function Get_Last_Valid_Cell(ByVal refCell As Range, _
                                    Optional ByVal maxRows As Integer = 0, _
                                    Optional ByVal maxCols As Integer = 0, _
                                    Optional ByVal criteria As String = "", _
                                    Optional ByVal rgxSch As Boolean = False, _
                                    Optional ByVal invCriteria As Boolean = False, _
                                    Optional ByVal invDir As Boolean = False, _
                                    Optional ByVal getNext As Boolean = False) As Range
                                    
    Set lastCell = refCell
    
    rowOS = 1
    colOS = 1
    
    If invDir Then
        rowOS = -1
        colOS = -1
    End If
    
    If maxRows Then
        Set lastCell = Search_Valid_Cell(lastCell, maxRows, criteria, rgxSch, invCriteria, invDir, 1, False)
        Set lastCell = Search_Valid_Cell(lastCell, maxRows, criteria, rgxSch, invCriteria, invDir, 1, True)
        
        If getNext Then
            Set lastCell = Get_Last_Valid_Cell(lastCell.Offset(rowOS, 0), maxRows, , criteria, _
                                               rgxSch, invCriteria, invDir)
        End If
    End If
    
    If maxCols Then
        Set lastCell = Search_Valid_Cell(lastCell, maxCols, criteria, rgxSch, invCriteria, invDir, 2, False)
        Set lastCell = Search_Valid_Cell(lastCell, maxCols, criteria, rgxSch, invCriteria, invDir, 2, True)
        
        If getNext Then
            Set lastCell = Get_Last_Valid_Cell(lastCell.Offset(0, colOS), , maxCols, criteria, _
                                               rgxSch, invCriteria, invDir)
        End If
    End If
    
    Set Get_Last_Valid_Cell = lastCell
End Function
'-----------------------------------------------------------------------------------------------------------'

'--------------------------------------------- Array Functions ---------------------------------------------'
Public Function Arr_To_Str(ByVal theArr As Variant) As String
    theStr = ""
    
    For Each elm In theArr
        theStr = theStr & elm & splitChar
    Next elm
    
    Arr_To_Str = Trim_Right(theStr, 1)
End Function

Public Function Find_In_Array(ByVal theArr As Variant, ByVal theElm As Variant, _
                              Optional rgxSch As Boolean = False, _
                              Optional ByVal schWithArr As Boolean = False, _
                              Optional ByVal matchWhole As Boolean = False, _
                              Optional ByVal startIdx As Integer = 0, _
                              Optional ByVal endIdx As Integer = -1) As Variant
    On Error Resume Next
    
    numElm = UBound(theArr, 1) - LBound(theArr, 1)
    theIdx = NOT_FOUND
    
    On Error GoTo 0
    
    If IsEmpty(numElm) Then
        Find_In_Array = theIdx
        
        Exit Function
    End If
    
    'If rgxSch Then theElm = Replace(theElm, " ", "\s")
    
    If startIdx <= 0 Then
        startIdx = 0
    ElseIf startIdx > numElm Then
        startIdx = numElm
    End If
    
    If endIdx < 0 Or endIdx > numElm Then endIdx = numElm
    
    For i = startIdx To endIdx
        If Not schWithArr Then
            If IsObject(theArr(i)) Then
                Set haystack = theArr(i)
                Set needle = theElm
            Else
                haystack = theArr(i)
                needle = theElm
            End If
        Else
            If IsObject(theArr(i)) Then
                Set needle = theArr(i)
                Set haystack = theElm
            Else
                needle = theArr(i)
                haystack = theElm
            End If
        End If
        
        If IsObject(haystack) Then
            If haystack Is needle Then
                theIdx = i
                
                Exit For
            End If
        ElseIf rgxSch Then
            If Not Regex_Match(haystack, needle) Is Nothing Then
                theIdx = i
                
                Exit For
            End If
        Else
            If (haystack = needle And matchWhole) _
            Or (InStr(haystack, needle) And Not matchWhole) Then
                theIdx = i
                
                Exit For
            End If
        End If
    Next
    
    Find_In_Array = theIdx
End Function

Public Function IsFound(ByVal result As Integer) As Boolean
    If result > NOT_FOUND Then
        IsFound = True
    Else
        IsFound = False
    End If
End Function

Public Function Array_Insert(ByRef theArr As Variant, theItem As Variant, Optional ByVal theIdx As _
                        Integer = -1, Optional ByVal unique As Boolean = False) As Integer
                        
    If Not IsArray(theArr) Then
        MsgBox "Error in ""Array_Insert"" - arg #1:" & vbCrLf & _
               "Not an array"
        
        End
    End If
    
    Array_Insert = NOT_FOUND
    
    On Error Resume Next
    
    lastIdx = UBound(theArr, 1) - LBound(theArr, 1)
    
    On Error GoTo 0
    
    If IsEmpty(lastIdx) Or lastIdx < 0 Then 'Array is empty
        ReDim theArr(0)
        
        If IsObject(theArr(0)) Then
            Set theArr(0) = theItem
        Else
            theArr(0) = theItem
        End If
        
        Array_Insert = 0
        
        Exit Function
    End If
    
    If (IsObject(theArr(0)) And Not IsObject(theItem)) Or (Not IsObject(theArr(0)) _
    And IsObject(theItem)) Then
        MsgBox "Error in ""Array_Insert"" - args #1, #2:" & vbCrLf & _
               "Arguments types not matching"
        
        End
    End If
    
    If unique And IsFound(Find_In_Array(theArr, theItem)) Then Exit Function
    
    If theIdx >= 0 Then 'Index is valid
        Dim tmpArr() As Variant 'Temporary array for storing the array "theArr" elements
        ReDim tmpArr(0)
        
        tmpArr(0) = theArr(0)
        
        'Copy all elements before index "theIdx" to the temporary array "tmpArr"
        For i = 1 To theIdx - 1: Array_Insert tmpArr, theArr(i): Next
        
        Array_Insert tmpArr, theItem
        
        'Copy all elements after index "theIdx" to the temporary array "tmpArr"
        For i = theIdx To lastIdx: Array_Insert tmpArr, theArr(i): Next
        
        theArr = tmpArr
    ElseIf IsObject(theItem) Then
        If Not theArr(lastIdx) Is Nothing Then
            idx = lastIdx + 1
        Else
            idx = lastIdx
        End If
        
        ReDim Preserve theArr(idx)
            
        Set theArr(idx) = theItem
    Else
        If Not IsEmpty(theArr(lastIdx)) Then
            idx = lastIdx + 1
        Else
            idx = lastIdx
        End If
        
        ReDim Preserve theArr(idx)
            
        theArr(idx) = theItem
    End If
    
    
    Array_Insert = idx
End Function

Public Function Arr_Len(ByVal theArr As Variant) As Integer
    If IsArray(theArr) Then
        On Error Resume Next

        fstElm = theArr(0)

        On Error GoTo 0
        
        If Not IsEmpty(fstElm) Then
            Arr_Len = UBound(theArr, 1) - LBound(theArr, 1) + 1
        Else
            Arr_Len = 0
        End If
    Else
        Arr_Len = -1
    End If
End Function

Public Function Print_Array(ByVal theArr As Variant, Optional ByVal objOpt As String = "") As String
    If IsArray(theArr) Then
        elmCount = Arr_Len(theArr)
    ElseIf IsObject(theArr) Then
        elmCount = theArr.Count
    Else
        elmCount = 0
    End If
    
    arrPrint = ""
    
    For i = 0 To elmCount - 1
        If objOpt = "" Then
            elmVal = theArr(i)
        Else
            elmVal = Get_Object_Attribute(theArr(i), objOpt)
        End If
        
        arrPrint = arrPrint & "(" & i & "): " & elmVal & vbCrLf
    Next
    
    Debug.Print arrPrint
    
    Print_Array = arrPrint
End Function

Public Function Get_Object_Attribute(ByVal theElm As Object, ByVal objOpt As String) As String
    Select Case LCase(objOpt)
        Case "value": elmVal = theElm.Value
        Case "address": elmVal = theElm.Address
        Case "row": elmVal = theElm.Row
        Case "column": elmVal = theElm.Column
        Case Else: elmVal = ""
    End Select
    
    Get_Object_Attribute = elmVal
End Function
'-----------------------------------------------------------------------------------------------------------'

'------------------------------------------ Conversion Functions -------------------------------------------'
Public Function To_Int(ByVal theStr As Variant) As Variant
    If Not IsArray(theStr) Then
        On Error Resume Next
    
        intVal = CLng(theStr)
        
        If IsEmpty(intVal) Then intVal = 0
        
        On Error GoTo 0
        
        To_Int = intVal
    Else
        Dim intArr() As Double
        
        For Each elm In theStr: Array_Insert intArr, To_Double(elm): Next elm
        
        To_Int = intArr
    End If
End Function

Public Function To_Double(ByVal theStr As Variant, Optional ByVal digits As Integer = 0) As Variant
    Dim decDigits As Variant
    
    To_Double = 0
    decimSep = Application.DecimalSeparator
    thousSep = Application.ThousandsSeparator
    decDigits = ""
    
    If digits = 0 Then digits = 2
    
    If Not IsArray(theStr) Then
        Set numMatch = Regex_Match(theStr, "[\,\.\" & decimSep & "\" & thousSep & "]\d+$")
        
        If numMatch Is Nothing Then
            To_Double = To_Int(theStr)
            
            Exit Function
        End If
        
        decDigits = Isolate_Numbers(numMatch(0))
        theStr = Regex_Replace(theStr, "[\,\.\" & decimSep & "\" & thousSep & "]\d+$", splitChar)
        theStr = Regex_Replace(theStr, "[\,\.\" & decimSep & "\" & thousSep & "]", "")
        theStr = Replace(theStr, splitChar, ".") & decDigits
        theDbl = Round(val(theStr), digits)
        
        To_Double = theDbl
    Else
        Dim dblArr() As Double
        
        For Each elm In theStr: Array_Insert dblArr, To_Double(elm): Next elm
        
        To_Double = dblArr
    End If
End Function

Public Function Hex_To_Dec(ByVal hex As String) As Long
    Hex_To_Dec = CLng("&H" & hex)
End Function
'-----------------------------------------------------------------------------------------------------------'

'----------------------------------------- Miscellaneous Functions -----------------------------------------'
Public Function Gen_SGI_Table(ByVal sqlCodeOrHeader As String, _
                              Optional ByVal dbTable As String = "", Optional ByVal conditions As String = "", _
                              Optional ByVal sortList As String = "", Optional ByVal theWB As Workbook = Nothing) As Worksheet
    'Application.ScreenUpdating = False
    
    If theWB Is Nothing Then Set theWB = ActiveWorkbook
    
    If Len(sqlCodeOrHeader) <> 0 And Len(conditions) = 0 And Len(sortList) = 0 Then
        sqlCode = sqlCodeOrHeader
        
        If Len(dbTable) = 0 Then dbTable = Remove_Extra_Space(Read_Delim_Str(sqlCode, "FROM ", " "))
    Else
        tableHeader = sqlCodeOrHeader
        
        If sortList = "" Then sortList = Split(Replace(LCase(tableHeader), ",", ""))(0) & " asc"
        
        If conditions <> "" And InStr(LCase(conditions), "where") = 0 Then _
            conditions = " where " & conditions
            
        sqlCode = "select " & tableHeader & " from [SGI].[dbo]." & dbTable & conditions & _
                  " order by " & sortList
    End If
    
    'On Error GoTo ERROR_HANDLER
    
    'Criação de objetos para consulta
    Set objMyconn = New ADODB.Connection
    Set objMyCmd = New ADODB.Command
    Set objMyRecordset = New ADODB.Recordset
    
    'Abre conexão com o banco
    objMyconn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & dataSource & _
                                 ";Initial Catalog=" & dataBase & "; User Id = " & usrName & _
                                 "; Password = " & usrPassWord & ";"
    objMyconn.Open
    
    'Execução de comando de select
    Set objMyCmd.ActiveConnection = objMyconn
    
    objMyCmd.CommandText = sqlCode
    objMyCmd.CommandType = adCmdText
    
    objMyCmd.Execute

    'Abre Recordset - Result set - Dados pesquisados
    Set objMyRecordset.ActiveConnection = objMyconn
    
    objMyRecordset.Open objMyCmd

    'Trata dados do result set por código
    'Atribui nome de colunas para a primeira linha da planilha
    totalColunas = objMyRecordset.Fields.Count
    
    Set dbWS = Reset_Sheet(dbTable, refWB:=theWB)
    
    For i = 1 To totalColunas: dbWS.Cells(1, i).Value = objMyRecordset.Fields(i - 1).Name: Next
    
    'Copia o resultado da pesquisa para o excel
    dbWS.Cells(2, 1).CopyFromRecordset (objMyRecordset)

    'Fecha a conexão com o banco
    objMyconn.Close
    
    On Error GoTo 0
    
    Set objMyconn = Nothing
    Set objMyCmd = Nothing
    Set objMyRecordset = Nothing
    Set Gen_SGI_Table = dbWS
Exit Function

'Tratamento de erro
ERROR_HANDLER:
    MsgBox "Erro " & Err.Number & " - " & Err.Description
End Function

Private Function Get_Var_Sizes(ByVal var1 As Variant, ByVal var2 As Variant) As String
    varTypeNum = VarType(var1)
    
    If varTypeNum > TYPE_ARRAY Then varTypeNum = varTypeNum - VarType(var1(0))
    
    Select Case varTypeNum
        Case TYPE_STRING:
            size1 = Len(var1)
            size2 = Len(var2)
        Case TYPE_ARRAY:
            size1 = Arr_Len(var1)
            size2 = Arr_Len(var2)
        Case Else:
            If IsNumeric(var1) Then
                size1 = var1
                size2 = var2
            Else
                size1 = 0
                size2 = 0
            End If
    End Select
    
    Get_Var_Sizes = size1 & splitChar & size2
End Function

Public Function Min(ByVal var1 As Variant, ByVal var2 As Variant) As Integer
    sizes = Get_Var_Sizes(var1, var2)
    size1 = val(Split(sizes, splitChar)(0))
    size2 = val(Split(sizes, splitChar)(1))
    
    If size1 < size2 Then
        Min = size1
    Else
        Min = size2
    End If
End Function

Public Function Max(ByVal str1 As String, ByVal str2 As String) As Integer
    sizes = Get_Var_Sizes(var1, var2)
    size1 = Split(sizes, splitChar)(0)
    size2 = Split(sizes, splitChar)(1)
    
    If size1 > size2 Then
        Max = size1
    Else
        Max = size2
    End If
End Function

Public Sub Blink(ByVal theCell As Range, Optional ByVal theColor As Long = vbRed, Optional ByVal blinkCount As Integer = 4, Optional ByVal blinkTime As Integer = 100)
    originalColor = theCell.Interior.Color
    
    For i = 0 To blinkCount
        Sleep blinkTime
        theCell.Interior.Color = theColor
        Sleep blinkTime
        theCell.Interior.Color = originalColor
    Next
End Sub

Public Function Get_Float_Digits(ByVal theNum As Double) As Integer
    floatDigs = Read_From_Pattern(Replace(theNum, ",", "."), ".")
    Get_Float_Digits = Len(floatDigs)
End Function

Public Sub Print_Var_Type(ByVal var As Variant)
    Select Case VarType(var)
        Case 0: Debug.Print "vbEmpty - " & "Empty (uninitialized)"
        Case 1: Debug.Print "vbNull - " & "Null (no valid data)"
        Case 2: Debug.Print "vbInteger - " & "Integer"
        Case 3: Debug.Print "vbLong - " & "Long integer"
        Case 4: Debug.Print "vbSingle - " & "Single-precision floating-point number"
        Case 5: Debug.Print "vbDouble - " & "Double-precision floating-point number"
        Case 6: Debug.Print "vbCurrency - " & "Currency value"
        Case 7: Debug.Print "vbDate - " & "Date value"
        Case 8: Debug.Print "vbString - " & "String"
        Case 9: Debug.Print "vbObject - " & "Object"
        Case 10: Debug.Print "vbError - " & "Error value"
        Case 11: Debug.Print "vbBoolean - " & "Boolean value"
        Case 12: Debug.Print "vbVariant - " & "Variant (used only witharrays of variants)"
        Case 13: Debug.Print "vbDataObject - " & "A data access object"
        Case 14: Debug.Print "vbDecimal - " & "Decimal value"
        Case 17: Debug.Print "vbByte - " & "Byte value"
        Case 20: Debug.Print "vbLongLong - " & "LongLong integer (Valid on 64-bit platforms only.)"
        Case 36: Debug.Print "vbUserDefinedType - " & "Variants that contain user-defined types"
        Case 8192: Debug.Print "vbArray - " & "Array"
        Case Else: Debug.Print "Not Specified"
    End Select
End Sub

Sub Export_To_PDF(ByVal theWS As Worksheet, ByVal fileName As String, _
                  Optional ByVal filePath As String = "")
    
    If filePath = "" Then filePath = theWS.Parent.Path
    
    If Right(filePath, 1) <> "\" Then filePath = filePath & "\"
    
    theWS.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath & fileName & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub

Function Execute_String(ByVal code As String) As Variant
    Dim regexOutput As String
    Dim regexOutArr() As String
    ReDim regexOutArr(0)
    
    fooReturn = Empty
    signPos = Regex_Match(code, "C)\s\=\s", regexOutput)
    comparePos = Regex_Match(code, "[\>\<]|\<\>|[\=\>\<]\=|\sis\s")
    regexOutArr(0) = ""
    
    If To_Int(regexOutArr(0)) > 1 And comparePos = 0 Then
        MsgBox "Error in ""Execute_String"" - arg #1:" & vbCrLf & _
               "0 or 1 attributions expected"
        
        End
    End If
    
    If Len(regexOutput) Then regexOutArr = Split(regexOutput, splitChar)
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1)
    
    If (signPos > 0 And To_Int(regexOutArr(0)) = 1) Or comparePos > 0 Then
        argPos = Regex_Match(code, "O)^\b\w+\b", regexOutput)
        theArg = regexOutput
        
        If signPos > 0 And argPos > 0 Then
            formStr = "function foo() as variant" & vbCrLf & code & vbCrLf & "foo = " & theArg & vbCrLf & "End function"
        Else
            code = Replace(code, "==", "=")
            formStr = "function foo() as variant" & vbCrLf & "foo = " & code & vbCrLf & "end function"
        End If
        
        vbComp.CodeModule.AddFromString formStr
        
        fooReturn = Application.Run(vbComp.Name & ".foo")
    Else
        vbComp.CodeModule.AddFromString "sub foo()" & vbCrLf & code & vbCrLf & "End Sub"
        Application.Run vbComp.Name & ".foo"
    End If
    
    ThisWorkbook.VBProject.VBComponents.Remove vbComp
    
    Execute_String = fooReturn
End Function

Public Function Get_SubFolder_List(ByVal folder As String) As Object()
    Dim subFolderList() As Object
    
    Set fileObj = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fileObj.GetFolder(folder)
    
    For Each subF In folderObj.SubFolders: Array_Insert subFolderList, subF: Next
    
    Get_SubFolder_List = subFolderList
End Function

Private Function Get_Oldest_Backup(subFolderList() As Object) As String
    oldestDate = DateValue(subFolderList(0).DateLastModified)
    oldestTime = TimeValue(subFolderList(0).DateLastModified)
    fIdx = 0
    
    For F = 0 To Arr_Len(subFolderList) - 1
        theDate = DateValue(subFolderList(F).DateLastModified)
        theTime = TimeValue(subFolderList(F).DateLastModified)
        
        If theDate < oldestDate Then
            oldestDate = theDate
            fIdx = F
        ElseIf theTime < oldestTime Then
            oldestTime = theTime
            fIdx = F
        End If
    Next
    
    Get_Oldest_Backup = subFolderList(fIdx).Name
End Function

Private Function Create_Backup_Revision(ByVal rootFolder As String, ByVal theWB As Workbook)
    Dim subFolderList() As Object
    
    Set FSO = CreateObject("scripting.filesystemobject")
    
    todayD = Day(Now)
    todayM = Month(Now)
    todayY = Year(Now)
    
    If To_Int(todayD) < 10 Then todayD = "0" & todayD
    If To_Int(todayM) < 10 Then todayM = "0" & todayM
    
    wbRoot = rootFolder & "\" & Read_Until_Pattern(theWB.Name, ".xl")
    
    If Not FSO.FolderExists(wbRoot) Then MkDir wbRoot
    
    subRoot = wbRoot & "\" & todayY & "-" & todayM & "-" & todayD
    revFolderName = ""
    
    On Error Resume Next
    
    If Not FSO.FolderExists(subRoot) Then MkDir subRoot
        
    On Error GoTo 0
    
    If FSO.FolderExists(subRoot) Then
        subFolderList = Get_SubFolder_List(subRoot)
        
        If Arr_Len(subFolderList) >= MAX_BACKUP Then
            revFolderName = Get_Oldest_Backup(subFolderList)
        Else
            revFolderName = "Backup " & Arr_Len(subFolderList) + 1
        End If
        
        'revFolderName = "Backup " & (Arr_Len(subFolderList) Mod MAX_BACKUP) + 1
        revFolder = subRoot & "\" & revFolderName
        
        On Error Resume Next
        
        If FSO.FolderExists(revFolder) Then
            Kill revFolder & "\*.*"
        Else
            MkDir revFolder
        End If
        
        On Error GoTo 0
        
        Create_Backup_Revision = revFolder
    Else
        Create_Backup_Revision = "Error"
    End If
End Function

Private Function Get_Backup_Folder(Optional ByVal theWB As Workbook = Nothing) As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String
    
    If theWB Is Nothing Then Set theWB = ThisWorkbook
    
    bkpFolderName = "Macros\General Backup"
    
    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then SpecialPath = SpecialPath & "\"
        
    bkpFolderName = SpecialPath & bkpFolderName
    
    On Error Resume Next
    
    If Not FSO.FolderExists(bkpFolderName) Then MkDir bkpFolderName
        
    On Error GoTo 0
    
    If FSO.FolderExists(bkpFolderName) Then
        Get_Backup_Folder = Create_Backup_Revision(bkpFolderName, theWB)
    Else
        Get_Backup_Folder = "Error"
    End If
End Function

Public Sub Export_Modules(Optional ByVal theWB As Workbook = Nothing) '(Optional ByVal projPassword As String = "")
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim exportPath As String
    Dim fileName As String
    Dim comp As VBIDE.VBComponent
    
    If Application.UserName <> "Carlos Herculano - VISION" Then Exit Sub
    If theWB Is Nothing Then Set theWB = ThisWorkbook
    
    Set projWB = theWB
    
    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    backupFolder = Get_Backup_Folder(theWB)
    
    If backupFolder = "Error" Then
        MsgBox "Error while creating folder"
        
        End
    End If
    
    If backupFolder = "Full" Then Exit Sub
    
'    On Error Resume Next
'        Kill backupFolder & "\*.*"
'    On Error GoTo 0
    
    If projWB.VBProject.Protection = 1 Then
        If Len(projPassword) Then
        Else
            MsgBox "The VBA in this workbook is protected," & _
                   "not possible to export the code"
                
            End
        End If
    End If
    
    exportPath = backupFolder & "\"
    
    For Each comp In projWB.VBProject.VBComponents
        bExport = True
        
        'Concatenate the correct filename for export.
        Select Case comp.Type
            Case vbext_ct_ClassModule: fileName = comp.Name & ".cls"
            Case vbext_ct_MSForm: fileName = comp.Name & ".frm"
            Case vbext_ct_StdModule: fileName = comp.Name & ".bas"
            
            Case vbext_ct_Document 'This is a worksheet or workbook object. Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            'Export the component to a text file.
            comp.Export exportPath & fileName
            
            'Remove it from the project if you want
            'wkbSource.VBProject.VBComponents.Remove comp
        End If
    Next comp
    
    'MsgBox "Useful Functions Stored"
End Sub

Public Function Get_Custom_Logo() As String
    imgExt = Array("png", "bmp", "jpeg", "jpg")
    
    Do
        Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    
        thePath = "\\vision-fs\VISION_DADOS\Publico\Logotipos Fornecedores e Clientes\"
        
        fileDiag.InitialFileName = thePath
        fileDiag.InitialView = msoFileDialogViewList
        fileDiag.AllowMultiSelect = False
        
        theImg = fileDiag.Show
        
        If theImg = 0 Then
            MsgBox ("Processo abortado.")
            
            End
        End If
        
        logo = fileDiag.SelectedItems(1)
        
        If Not IsFound(Find_In_Array(imgExt, Read_From_Pattern_X(logo, "."))) Then _
            MsgBox "Erro: selecione um arquivo de imagem válido"
        
        Set fileDiag = Nothing
    Loop While Not IsFound(Find_In_Array(imgExt, Read_From_Pattern_X(logo, ".")))
    
    Get_Custom_Logo = logo
End Function

Public Sub Terminate_Program(Optional exitProg As Boolean = False, _
                             Optional ByVal endMsg As String = "Pronto!", _
                             Optional ByVal errDetect As Boolean = False)
    
    If Not errDetect Then Progress_Almost_Terminate
    
    Progress_Terminate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    If exitProg Then
        MsgBox endMsg
    
        End
    End If
End Sub

Public Sub AllInternalPasswords()
    ' Breaks worksheet and workbook structure passwords. Bob McCormick
    ' probably originator of base code algorithm modified for coverage
    ' of workbook structure / windows passwords and for multiple passwords
    '
    ' Modified 2003-Apr-04 by GBA: All msgs to PT-BR
    Const DBLSPACE As String = vbNewLine & vbNewLine
    Const AUTHORS As String = DBLSPACE & vbNewLine & _
    "Adaptado por: Guilherme Bicalho Andrade"
    Const HEADER As String = "AllInternalPasswords User Message"
    Const VERSION As String = DBLSPACE & "Version 1.1.3 Setembro/2014"
    Const REPBACK As String = DBLSPACE & "Se você encontrar alguma falha no" & _
    "código, gentileza consertar."
    Const ALLCLEAR As String = DBLSPACE & "As senhas devem ter sido removidas." & _
    " Seja feliz sem elas!!!" & _
    DBLSPACE & "O acesso e utilização de alguns dados pode ser proibidos. Em caso de dúvida, não use"
    Const MSGNOPWORDS1 As String = "Não foi encontrada nenhuma senha" & _
    "nas sheets, workbook structure ou windows." & AUTHORS & VERSION
    Const MSGNOPWORDS2 As String = "Não havia qualquer proteção por senha " & _
    "na estrutura da planilha." & DBLSPACE & _
    "Agora vamos quebrar a senha." & AUTHORS & VERSION
    Const MSGTAKETIME As String = "Este procedimento pode levar um bom " & _
    "tempo." & DBLSPACE & "O tempo necessário depende da quantidade " & _
    "de senhas existentes, da complexidade das senhas e " & _
    "da especificação da sua máquina." & DBLSPACE & _
    "Seja paciente! Vá tomar um café!" & AUTHORS & VERSION
    Const MSGPWORDFOUND1 As String = "Sua planilha tinha uma estrutura " & _
    "de planilha bloqueada por senha." & DBLSPACE & _
    "A senha encontra era: " & DBLSPACE & "$$" & DBLSPACE & _
    "Anote a senha, pois pode ser útil " & _
    "no futuro em outra planilha criada por essa mesma pessoa" & _
    DBLSPACE & "Now to check and clear other passwords." & AUTHORS & VERSION
    Const MSGPWORDFOUND2 As String = "Sua planilha era bloqueada por senha" & _
    DBLSPACE & "A senha encontra era: " & _
    DBLSPACE & "$$" & DBLSPACE & "Anote a senha, pois pode ser útil " & _
    "no futuro em outra planilha criada por essa mesma pessoa" & _
    DBLSPACE & "Now to check and clear " & _
    "other passwords." & AUTHORS & VERSION
    Const MSGONLYONE As String = "Somente a senha de bloqueio da estrutura " & _
    "da planilha foi encontrada." & _
    ALLCLEAR & AUTHORS & VERSION & REPBACK
    
    Dim w1 As Worksheet, w2 As Worksheet
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim m As Integer, n As Integer, i1 As Integer, i2 As Integer
    Dim i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
    Dim PWord1 As String
    Dim ShTag As Boolean, WinTag As Boolean
    
    Application.ScreenUpdating = False
    
    With ActiveWorkbook: WinTag = .ProtectStructure Or .ProtectWindows: End With
    
    ShTag = False
    
    For Each w1 In Worksheets: ShTag = ShTag Or w1.ProtectContents: Next w1
    
    If Not ShTag And Not WinTag Then
        MsgBox MSGNOPWORDS1, vbInformation, HEADER
        Exit Sub
    End If
    
    MsgBox MSGTAKETIME, vbInformation, HEADER
    
    If Not WinTag Then
        MsgBox MSGNOPWORDS2, vbInformation, HEADER
    Else
        On Error Resume Next
        
        Do 'dummy do loop
            For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
                For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
                    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
                        For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
                            With ActiveWorkbook
                                .Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                
                                If .ProtectStructure = False And .ProtectWindows = False Then
                                    PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                    
                                    MsgBox Application.Substitute(MSGPWORDFOUND1, "$$", PWord1), vbInformation, HEADER
                                    
                                    Exit Do 'Bypass all for...nexts
                                End If
                            End With
                        Next: Next: Next
                    Next: Next: Next
                Next: Next: Next
            Next: Next: Next
        Loop Until True
        
        On Error GoTo 0
        
    End If
    
    If WinTag And Not ShTag Then
        MsgBox MSGONLYONE, vbInformation, HEADER
        Exit Sub
    End If
    
    On Error Resume Next
        
    For Each w1 In Worksheets
        'Attempt clearance with PWord1
        w1.Unprotect PWord1
    Next w1
    
    On Error GoTo 0
    
    ShTag = False
    
    For Each w1 In Worksheets
        'Checks for all clear ShTag triggered to 1 if not.
        ShTag = ShTag Or w1.ProtectContents
    Next w1
    
     If ShTag Then
        For Each w1 In Worksheets
            With w1
                If .ProtectContents Then
                    On Error Resume Next
                    
                    Do 'Dummy do loop
                        For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
                            For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
                                For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
                                    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
                                        .Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                        
                                        If Not .ProtectContents Then 'Erro constante de loop infinito nesta instrução
                                            PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                            
                                            MsgBox Application.Substitute(MSGPWORDFOUND2, "$$", PWord1), vbInformation, HEADER
                                            
                                            'leverage finding Pword by trying on other sheets
                                            For Each w2 In Worksheets: w2.Unprotect PWord1: Next w2
                                            
                                            Exit Do 'Bypass all for...nexts
                                        End If
                                    Next: Next: Next
                                Next: Next: Next
                            Next: Next: Next
                        Next: Next: Next
                    Loop Until True
                    
                    On Error GoTo 0
                End If
            End With
        Next w1
    End If
    
    MsgBox ALLCLEAR & AUTHORS & VERSION & REPBACK, vbInformation, HEADER
End Sub
'-----------------------------------------------------------------------------------------------------------'
