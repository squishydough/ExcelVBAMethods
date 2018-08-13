'---------------------------------------------------------------------------
'   Description :   Adds a row to range and resizes the range
'---------------------------------------------------------------------------
Public Sub AddRowToRange( _
    TargetRange As Range _
)

    With TargetRange
        .Rows(.Rows.Count + 1).Insert CopyOrigin:=xlFormatFromLeftOrAbove
        .Resize(.Rows.Count + 1, .Columns.Count).Name = .Name.Name
    End With
End Sub

'---------------------------------------------------------------------------
'   Description :   Allows the user to browse the computer for a file
'---------------------------------------------------------------------------
Public Function BrowseForFile( _
    FilterTitle As String, _
    FilterTypes As String, _
    Optional DialogueTitle As String _
) As String
    
    '**
    '*  Optional parameter defaults
    '**
    If IsMissing(DialogueTitle) Then DialogueTitle = "Please browse to the location of the file."
    
    '**
    '*  Variables
    '**
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Title = DialogueTitle

        '**
        '*  Clear out the current filters, and add our own.
        '**
        .Filters.Clear
        .Filters.Add FilterTitle, FilterTypes

        '**
        '*  Show the dialog box. If the .Show method returns True, the
        '*  user picked at least one file. If the .Show method returns
        '*  False, the user clicked Cancel.
        '**
        If .Show = True Then
            BrowseForFile = .SelectedItems(1)
        Else
            BrowseForFile = "ERROR"
        End If
    End With
    
End Function

'----------------------------------------------------------------------------
'   Description     :   Opens a userform and centers it on the user's screen
'----------------------------------------------------------------------------
Public Sub CenterUserForm(frm As Object)
    With frm
        .StartUpPosition = 0
        .Left = Application.ActiveWindow.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show False
    End With
End Sub

'----------------------------------------------------------------------------
'   Description     :   Removes illegal worksheet name characters
'----------------------------------------------------------------------------
Public Function CleanSheetName( _
    DirtyString As String, _
    ReplaceChar As String _
) As String

    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[\<\>\*\\\/\?|]"
        CleanSheetName = .Replace(DirtyString, ReplaceChar)
    End With
End Function

' ---------------------------------------------------------------------------
'   Description     :   Returns the hex code of a color
' ---------------------------------------------------------------------------
Public Function ColorToHex( _
    TargetRange As Range _
) As String

    Dim sColor As String

    sColor = Right("000000" & Hex(TargetRange.Interior.Color), 6)
    ColorToHex = "#" & Right(sColor, 2) & Mid(sColor, 3, 2) & Left(sColor, 2)
End Function

' ---------------------------------------------------------------------------
'   Description     :   Returns an RGB color string
' ---------------------------------------------------------------------------
Public Function ColorToRGB( _
    TargetRange As Range, _
    Optional WhichVal As String _
) As String

    Dim C As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long

    C = TargetRange.Interior.Color
    R = C Mod 256
    G = C \ 256 Mod 256
    B = C \ 65536 Mod 256

    If WhichVal = "R" Then
        ColorToRGB = "R=" & R
    ElseIf WhichVal = "G" Then
        ColorToRGB = "G=" & G
    ElseIf WhichVal = "B" Then
        ColorToRGB = "B=" & B
    Else
        ColorToRGB = "R=" & R & ", G=" & G & ", B=" & B
    End If
End Function

'----------------------------------------------------------------------------
'   Description     :   Creates a named range in the workbook
'----------------------------------------------------------------------------
Public Sub CreateNamedRange( _
    RangeName As String, _
    SheetName As String, _
    TargetRange As Range _
)

    ThisWorkbook.Names.Add Name:=RangeName, RefersTo:="=" & "'" & SheetName & "'!" & TargetRange.Address
End Sub

'----------------------------------------------------------------------------
'   Description     :   Takes a passed column letter and returns the column number
'----------------------------------------------------------------------------
Public Function ColumnLetterToNumber( _
    ColumnLetter As String _
) As Long
    
    ColumnLetterToNumber = Range(ColumnLetter & 1).Column
End Function

'----------------------------------------------------------------------------
'   Description     :   Takes a passed column number and returns the column letter
'----------------------------------------------------------------------------
Public Function ColumnNumberToLetter( _
    ColumnNumber As Long _
) As String
    
    ColumnNumberToLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)
End Function

' ---------------------------------------------------------------------------
'   Description     :   Returns a collection of dates between two dates
'                       (including the passed dates) formatted in the passed 
'                       FormatType.
'
'                       Default format type is the English spelling of the day.
'                       Default ExcludedDays are Saturdays and Sundays
' ---------------------------------------------------------------------------
Public Function GetDatesBetweenDates( _
    FirstDate As Date, _
    LastDate As Date, _
    Optional FormatType As String, _
    Optional ExcludedDays As Collection _
) As Collection

    '**
    '*  If no format is passed, default to the English day
    '**
    If Len(FormatType) = 0 Then
        FormatType = "dddd"
    End If
    
    '**
    '*  If no ExcludedDays are passed, then default to no weekends
    '**
    If ExcludedDays Is Nothing Then
        Set ExcludedDays = New Collection
        ExcludedDays.Add "Saturday"
        ExcludedDays.Add "Sunday"
    End If
    
    '**
    '*  Variable declarations
    '**
    Dim LoopDate As Date                    '*  Stores the loop generated date
    Dim LoopDay As String                   '*  The English day of the current loop generated date
    Dim i As Integer                        '*  Used for looping through the ExcludedDays collection
    Dim TempDate As Date                    '*  Used for swapping the passed dates if needed
    Dim DatesCollection As New Collection   '*  Stores the dates to be returned
    
    '**
    '*  Error checking - if Last Date is greater than
    '*  FirstDate, swap them
    '**
    If LastDate < FirstDate Then
        TempDate = FirstDate
        FirstDate = LastDate
        LastDate = TempDate
    End If
    
    '**
    '*  Loop between the two dates, adding the days to the collection
    '**
    LoopDate = FirstDate
    Do While LoopDate <= LastDate
        '*  Check if the day should be added
        '*  If current day is excluded, go to next loop iteration
        If ExcludedDays.Count > 0 Then
            For i = 1 To ExcludedDays.Count Step 1
                LoopDay = Format(LoopDate, "dddd")
                If ExcludedDays.Item(i) = LoopDay Then GoTo NextLoop
            Next i
        End If
        
        '*  Add formatted date to collection
        DatesCollection.Add Format(LoopDate, FormatType)
NextLoop:
        '*  Increment the date
        LoopDate = DateAdd("d", 1, LoopDate)
    Loop
    
    Set GetDatesBetweenDates = DatesCollection
End Function

'----------------------------------------------------------------------------
'   Description     :   Returns the last used row in a column, range, or
'                       worksheet. Column can be passed as a string or integer.
'
'   Requires        :   ColumnLetterToNumber method
'----------------------------------------------------------------------------
Public Function GetLastUsedRow(SearchWhere As Variant, Optional TargetSheet As Worksheet) As Long
    '**
    '*  If TargetSheet is not provided, default to activesheet
    '**
    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet
    
    '**
    '*  Return value depending on passed argument type
    '**
    If TypeName(SearchWhere) = "Integer" Then
        With TargetSheet
            GetLastUsedRow = .Cells(.Rows.Count, SearchWhere).End(xlUp).Row
        End With
    ElseIf TypeName(SearchWhere) = "String" Then
        With TargetSheet
            GetLastUsedRow = .Cells(.Rows.Count, ColumnLetterToNumber(CStr(SearchWhere))).End(xlUp).Row
        End With
    ElseIf TypeName(SearchWhere) = "Range" Then
        With SearchWhere
            GetLastUsedRow = .Rows(.Rows.Count).Row
        End With
    ElseIf TypeName(SearchWhere) = "Worksheet" Then
        With SearchWhere
            GetLastUsedRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        End With
    Else
        MsgBox "Error - you passed an invalid type to GetLastUsedRow.", vbOKOnly + vbCritical, "Error!"
        GetLastUsedRow = 0
    End If
    
End Function

'----------------------------------------------------------------------------
'   Description     :   Returns the last used column in a row, range, or
'                       worksheet as a number.
'----------------------------------------------------------------------------
Public Function GetLastUsedColumnNumber(SearchWhere As Variant, Optional TargetSheet As Worksheet) As Long
    '**
    '*  If TargetSheet is not provided, default to activesheet
    '**
    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet
    
    '**
    '*  Return value depending on passed argument type
    '**
    If TypeName(SearchWhere) = "Integer" Then
        With TargetSheet
            GetLastUsedColumnNumber = .Cells(SearchWhere, .Columns.Count).End(xlToLeft).Column
        End With
    ElseIf TypeName(SearchWhere) = "Range" Then
        With SearchWhere
            GetLastUsedColumnNumber = .Columns(.Columns.Count).Column
        End With
    ElseIf TypeName(SearchWhere) = "Worksheet" Then
        With SearchWhere
            GetLastUsedColumnNumber = .UsedRange.Columns(.UsedRange.Columns.Count).Column
        End With
    Else
        MsgBox "Error - you passed an invalid type to GetLastUsedColumnNumber.", vbOKOnly + vbCritical, "Error!"
        GetLastUsedColumnNumber = 0
    End If
    
End Function
'----------------------------------------------------------------------------
'   Description     :   Returns the last used column in a row, range, or
'                       worksheet as a letter.
'
'   Requires        :   GetLastUsedColumnNumber, ColumnNumberToLetter
'----------------------------------------------------------------------------
Public Function GetLastUsedColumnLetter(SearchWhere As Variant, Optional TargetSheet As Worksheet) As String
    GetLastUsedColumnLetter = ColumnNumberToLetter(GetLastUsedColumnNumber(SearchWhere, TargetSheet))
End Function

' ---------------------------------------------------------------------------
'   Description     :   Checks if a named range exists
' ---------------------------------------------------------------------------
Public Function NamedRangeExists( _
    RangeName As String _
) As Boolean
    
    Dim LoopName As Name
    
    For Each LoopName In ThisWorkbook.Names
        If LoopName.Name = RangeName Then
            NamedRangeExists = True
            Exit Function
        End If
    Next LoopName
    
    NamedRangeExists = False
End Function

' ---------------------------------------------------------------------------
'   Description     :   Removes and re-creates a named range
' ---------------------------------------------------------------------------
Public Sub ReplaceNamedRange( _
    RangeName As String, _
    SheetName As String, _
    TargetRange As Range _
)
    If NamedRangeExists(RangeName) = True Then ThisWorkbook.Names(RangeName).Delete
    CreateNamedRange RangeName, SheetName, TargetRange
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RescopeNamedRangesToWorkbook
' Author    : Jesse Stratton
' Date      : 11/18/2013
' Purpose   : Rescopes the parent of worksheet scoped named ranges to the active workbook
' for each named range with a scope equal to the active sheet in the active workbook.
'---------------------------------------------------------------------------------------
Public Sub RescopeNamedRangesToWorkbook()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim objName As Name
    Dim sWsName As String
    Dim sWbName As String
    Dim sRefersTo As String
    Dim sObjName As String
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    sWsName = ws.Name
    sWbName = wb.Name

    'Loop through names in worksheet.
    For Each objName In ws.Names
    'Check name is visble.
        If objName.Visible = True Then
    'Check name refers to a range on the active sheet.
            If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
                sRefersTo = objName.RefersTo
                sObjName = objName.Name
    'Check name is scoped to the worksheet.
                If objName.Parent.Name <> sWbName Then
    'Delete the current name scoped to worksheet replacing with workbook scoped name.
                    sObjName = Mid(sObjName, InStr(1, sObjName, "!") + 1, Len(sObjName))
                    objName.Delete
                    wb.Names.Add Name:=sObjName, RefersTo:=sRefersTo
                End If
            End If
        End If
    Next objName
End Sub

' ---------------------------------------------------------------------------
'   Description     :   Checks if a worksheet exists
' ---------------------------------------------------------------------------
Public Function SheetExists(SheetName As String, Optional TargetBook As Workbook, Optional CheckCodeName = True) As Boolean
    Dim Sheet As Worksheet

    If TargetBook Is Nothing Then Set TargetBook = ActiveWorkbook

    For Each Sheet In TargetBook.Worksheets
        If SheetName = Sheet.Name Then
            SheetExists = True
            Exit Function
        End If
        
        If SheetName = Sheet.CodeName And CheckCodeName = True Then
            SheetExists = True
            Exit Function
        End If
    Next
    
    SheetExists = False
End Function