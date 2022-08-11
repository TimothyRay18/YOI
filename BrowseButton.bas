Attribute VB_Name = "BrowseButton"
Sub BomFile()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
'    dialogBox.InitialFileName = Range("A12").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B1").Value = dialogBox.SelectedItems(1)
    End If
End Sub
Sub QManFile()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
'    dialogBox.InitialFileName = Range("A12").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls;*.xlsb"
    
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B5").Value = dialogBox.SelectedItems(1)
    End If
End Sub
Sub TemplateFile()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
'    dialogBox.InitialFileName = Range("A12").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B9").Value = dialogBox.SelectedItems(1)
    End If
End Sub
