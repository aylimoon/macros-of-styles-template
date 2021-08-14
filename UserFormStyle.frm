VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormStyle 
   Caption         =   "Styles Template"
   ClientHeight    =   1840
   ClientLeft      =   30
   ClientTop       =   -450
   ClientWidth     =   4695
   OleObjectBlob   =   "UserFormStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error GoTo mb

Dim oRng As Range
Dim text As Style
Dim iShape As InlineShape
Dim oTbl As Table

Select Case StyleName
    
Case "Main_text"
    Set oRng = Selection.Range
    Set text = ActiveDocument.Styles("Main_text_1")
    With oRng.Find
        .ClearFormatting
        .Style = wdStyleNormal
    With .Replacement
        .ClearFormatting
        .Style = text
    End With
    .Execute Wrap:=wdFindStop, Format:=True, Replace:=wdReplaceAll
    End With
    
Case "Picture_name"
    Set oRng = Selection.Range
    For Each iShape In oRng.InlineShapes
        iShape.Select
        Selection.Style = ActiveDocument.Styles("Picture_name_1")
        Selection.MoveDown
        Selection.Style = ActiveDocument.Styles("Picture_name_1")
    Next
    
Case "Table_text"
    For Each oTbl In Selection.Tables
        oTbl.Select
        Selection.Style = ActiveDocument.Styles(wdStyleNormalTable)
        On Error Resume Next
        With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderLeft)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderRight)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderHorizontal)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderVertical)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        On Error GoTo mb
        Selection.Style = ActiveDocument.Styles("Table_text_1")
    Next

Case "Table_header"
    For Each oTbl In Selection.Tables
        oTbl.Rows.First.Range.Select
        Selection.Style = ActiveDocument.Styles("Table_header_1")
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Rows.HeadingFormat = True
    Next
    
End Select

UserFormStyle.Hide
Exit Sub
mb:     If Err.Number = 5941 Then
        MsgBox "The requested style does not exist. Select the appropriate template of styles to apply the macro"
        Else
        MsgBox "Error" & Str(Err.Number) & ". " & Err.Description
        End If
End Sub
Private Sub CommandButton2_Click()
On Error GoTo mb

Dim oRng As Range
Dim text As Style
Dim iShape As InlineShape
Dim oTbl As Table

Select Case StyleName
    
Case "Main_text"
    Set oRng = Selection.Range
    Set text = ActiveDocument.Styles("Main_text_2")
    With oRng.Find
        .ClearFormatting
        .Style = wdStyleNormal
    With .Replacement
        .ClearFormatting
        .Style = text
    End With
    .Execute Wrap:=wdFindStop, Format:=True, Replace:=wdReplaceAll
    End With
    
Case "Picture_name"
    Set oRng = Selection.Range
    For Each iShape In oRng.InlineShapes
        iShape.Select
        Selection.Style = ActiveDocument.Styles("Picture_name_2")
        Selection.MoveDown
        Selection.Style = ActiveDocument.Styles("Picture_name_2")
    Next
    
Case "Table_text"
    For Each oTbl In Selection.Tables
        oTbl.Select
        Selection.Style = ActiveDocument.Styles(wdStyleNormalTable)
        On Error Resume Next
        With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderLeft)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderRight)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderHorizontal)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderVertical)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        On Error GoTo mb
        Selection.Style = ActiveDocument.Styles("Table_text_2")
    Next

Case "Table_header"
    For Each oTbl In Selection.Tables
        oTbl.Rows.First.Range.Select
        Selection.Style = ActiveDocument.Styles("Table_header_2")
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Rows.HeadingFormat = True
    Next
    
End Select

UserFormStyle.Hide
Exit Sub
mb:     If Err.Number = 5941 Then
        MsgBox "The requested style does not exist. Select the appropriate template of styles to apply the macro"
        Else
        MsgBox "Error" & Str(Err.Number) & ". " & Err.Description
        End If
End Sub
Private Sub CommandButton3_Click()
UserFormStyle.Hide
End Sub