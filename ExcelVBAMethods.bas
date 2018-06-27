Attribute VB_Name = "ExcelVBAMethods"
'*****************************************************************************
'   Description         :   A collection of useful Excel methods
'   Methods             :   BrowseForFile
'                           CleanSheetName
'                           ColorToHex
'                           ColorToRGB
'                           ColumnLetterToNumber
'                           ColumnNumberToLetter
'                           CreateNamedRange
'                           NamedRangeExists
'                           ReplaceNamedRange
'                           RescopeNamedRangesToWorkbook
'*****************************************************************************






' ---------------------------------------------------------------------------
'   Description         :   Allows the user to browse the computer for a fil
' ---------------------------------------------------------------------------
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
'   Description     :   Removes illegal worksheet name characters
'----------------------------------------------------------------------------
Public Function CleanSheetName(DirtyString As String, ReplaceChar As String) As String
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
    '**
    '*  Variables
    '**
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
    '**
    '*  Variables
    '**
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
Public Sub CreateNamedRange(RangeName As String, SheetName As String, TargetRange As Range)
    ThisWorkbook.Names.Add Name:=RangeName, RefersTo:="=" & "'" & SheetName & "'!" & TargetRange.Address
End Sub

'----------------------------------------------------------------------------
'   Description     :   Takes a passed column letter and returns the column number
'----------------------------------------------------------------------------
Public Function ColumnLetterToNumber(ColumnLetter As String) As Long
    ColumnLetterToNumber = Range(ColumnLetter & 1).Column
End Function

'----------------------------------------------------------------------------
'   Description     :   Takes a passed column number and returns the column letter
'----------------------------------------------------------------------------
Public Function ColumnNumberToLetter(ColumnNumber As Long) As String
    ColumnNumberToLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)
End Function

' ---------------------------------------------------------------------------
'   Description     :   Get last used row in a worksheet
' ---------------------------------------------------------------------------
Public Function GetLastUsedRowInSheet(TargetSheet As Worksheet) As Long
    GetLastUsedRowInSheet = TargetSheet.UsedRange.Rows(TargetSheet.UsedRange.Rows.Count).Row
End Function

' ---------------------------------------------------------------------------
'   Description     :   Get last used column in a worksheet
' ---------------------------------------------------------------------------
Public Function GetLastUsedColumnInSheet(TargetSheet As Worksheet) As Long
    GetLastUsedColumnInSheet = TargetSheet.UsedRange.Columns(TargetSheet.UsedRange.Columns.Count).Column
End Function

' ---------------------------------------------------------------------------
'   Description     :   Checks if a named range exists
' ---------------------------------------------------------------------------
Public Function NamedRangeExists(RangeName As String) As Boolean
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
Public Sub ReplaceNamedRange(RangeName As String, SheetName As String, TargetRange As Range)
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
