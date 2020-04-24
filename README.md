# EXCEL-VBA-Useful-Functions
A library with many useful functions for assisting on EXCEL VBA coding

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
