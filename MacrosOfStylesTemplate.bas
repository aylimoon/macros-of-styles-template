Attribute VB_Name = "MacrosOfStylesTemplate"
Option Explicit
Public StyleName As String
Sub Template_style_main_text()
If Selection.Type = wdSelectionIP Then
   MsgBox "Highlight the text to apply the macro"
Else
    StyleName = "Main_text"
    UserFormStyle.Show
End If
End Sub
Sub Template_style_picture_name()
If Selection.Type = wdSelectionIP Then
   MsgBox "Highlight a picture / pictures to apply the macro"
Else
    StyleName = "Picture_name"
    UserFormStyle.Show
End If
End Sub
Sub Template_style_table_text()
If Selection.Type = wdSelectionIP Then
   MsgBox "Highlight the table / tables to apply the macro"
Else
    StyleName = "Table_text"
    UserFormStyle.Show
End If
End Sub
Sub Template_style_table_header()
If Selection.Type = wdSelectionIP Then
   MsgBox "Highlight the table / tables to apply the macro"
Else
    StyleName = "Table_header"
    UserFormStyle.Show
End If
End Sub