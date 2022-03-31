Option Explicit

'Press button "Import Workplans" in another document's worksheet "Update workplan"
Sub Import_Workplans()
    Dim IW_StartTime As Double
    Dim IW_Sec As Long
    Dim LResult(1 To 6) As Date
    IW_StartTime = Timer
    'Stop updating screen and automatic calculation (for speed gains)
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    'Call subroutine for Importing Workplans and return the time needed to do so
    Call IW(LResult)
    'Re-establish updating screen and automatic calculation
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
' Import workplans to the other file
Sub IW(LResult)
    Dim Second_workbook, wbpath As Variant
    Dim Ldate As Date
    Dim Dash_row(1 To 3) As Integer
    Dim IW_i, IW_i_R As Integer
    IW_i_R = 1
    'Rows in the a sheet where the Sharepoint path is specified for each of the three workbooks we shall refer to as Dashboard
    Dash_row(1) = 6
    Dash_row(2) = 7
    Dash_row(3) = 8
    'Keep the name of the overall workplan
    Second_workbook = ThisWorkbook.Name
    'Remove filters
    With Workbooks(Second_workbook).Sheets("Update workplan")
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If
    End With
    'Delete contents of the previous workplan
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
    End With
    'Define path (as written in the Dashboard)
    wbpath = Workbooks(Second_workbook).Sheets("Dashboard").Cells(5, 2).Text
    'Import the three workbooks
    For Each IW_i In Dash_row
        Call IW_ImportWP(Second_workbook, wbpath, IW_i, LResult(IW_i_R))
        IW_i_R = IW_i_R + 1
    Next IW_i
    'Fix filters up to the end of the range
    Workbooks(Second_workbook).Sheets("Update workplan").AutoFilterMode = False

    'Format cells up to the last filled cell
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Range("A5:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).AutoFilter
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).RowHeight = 25
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.ThemeColor = 1
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.TintAndShade = -0.499984740745262
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.Weight = xlThin
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Font.Name = "Calibri"
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Font.Size = 8
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).HorizontalAlignment = xlCenter
    End With

    ' Calculate Formula date of updates 
    Ldate = Now()
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Cells(1, 4).Value = "Last update:" & Ldate
    End With
End Sub

Sub IW_ImportWP(Second_workbook, wbpath, IW_Dash_row_WP, LResult_WP)
    Dim IW_Name_WB_WP, IW_Name_WB_WP_path As String
    Dim FilePAth As String
    IW_Name_WB_WP = Workbooks(Second_workbook).Sheets("Dashboard").Cells(IW_Dash_row_WP, 2).Text
    IW_Name_WB_WP_path = wbpath & IW_Name_WB_WP
    FilePAth = wbpath
    Call Date_File(FilePAth, IW_Name_WB_WP, LResult_WP)
    Call OpenWB(IW_Name_WB_WP, IW_Name_WB_WP_path)
    Call Copy_wb(Second_workbook, IW_Name_WB_WP)
    Call CloseWB(IW_Name_WB_WP)
End Sub

Sub Date_File(File_Path, File_Name, LResult)
    Dim New_Path As String
    Dim filespec As String
    Dim fs, f As Variant
    On Error GoTo 20
    New_Path = SharePointURLtoUNC(File_Path)
    filespec = New_Path & File_Name
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    LResult = f.DateLastModified
    GoTo 30
20 LResult = "00:00:00"
30
End Sub

Sub OpenWB(WBFilename, WBFilename_path)
    Dim filetoopen As Variant
    On Error GoTo 10
    Workbooks.Open Filename:=WBFilename_path, Password:="", UpdateLinks:=0, ReadOnly:=True
    Exit Sub

10  filetoopen = Application.GetOpenFilename("Excel files (*.xl*), *.xls*")
    If filetoopen = False Then End
    Workbooks.Open filetoopen, Password:="", UpdateLinks:=0, ReadOnly:=True
    WBFilename = ActiveWorkbook.Name
End Sub

' Calculating the last row from which to copy
Sub Copy_wb(Second_workbook, WBFilename)
    With Workbooks(WBFilename).Sheets("Workplan")
        .AutoFilterMode = False
        .Columns("A:CW").EntireColumn.Hidden = False
        .Range("A7:CW" & Workbooks(WBFilename).Sheets("Workplan").Cells(.Rows.Count, 1).End(xlUp).Row).Copy
    End With

    ' Calculating the first empty row after the previous import
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteFormats
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    End With
End Sub

Sub CloseWB(WBFilename)
    Workbooks.Application.CutCopyMode = False
    Workbooks(WBFilename).Close SaveChanges:=False
    Workbooks.Application.CutCopyMode = True
End Sub

Public Function SharePointURLtoUNC(sURL)
    Dim bIsSSL As Boolean
    bIsSSL = InStr(1, sURL, "https:") > 0
    sURL = Replace(Replace(sURL, "/", "\"), "%20", " ")
    sURL = Replace(Replace(sURL, "https:", vbNullString), "http:", vbNullString)
    sURL = Replace(sURL, Split(sURL, "\")(2), Split(sURL, "\")(2) & "@SSL\DavWWWRoot")
    If Not bIsSSL Then sURL = Replace(sURL, "@SSL\", vbNullString)
    SharePointURLtoUNC = sURL
End Function



