
Function Cover(gfilename As String)
    
    'Formats first slide in range of PPT.
    
    Set Logo = ActivePresentation.Slides(1).Shapes.AddPicture( _
    fileName:=gfilename, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, Left:=400, Top:=180, _
    Width:=-1, Height:=-1)
    
        'Sets height to 100px; best avg. fit.
        
        With Logo
            Logo.LockAspectRatio = msoTrue
            Logo.Height = 100
        End With
        
End Function

Function Master(gfilename As String)
    
    'Same process as above, just sets to .SlideMaster instead of .Slides().
    
    Set Logo = ActivePresentation.SlideMaster.Shapes.AddPicture( _
    fileName:=gfilename, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, Left:=600, Top:=30, _
    Width:=-1, Height:=-1)
    
        With Logo
            Logo.LockAspectRatio = msoTrue
            Logo.Height = 40
        End With
        
End Function

Function ErrorHandle()
    
    'Easier to signal need to update repository than iferror cases. 
    
    Dim errorVar As String
    errorVar = "\\Repository\error.png"
    
    Set Logo = ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
    fileName:=errorVar, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, Left:=400, Top:=125, _
    Width:=-1, Height:=-1)

        With Logo
            Logo.LockAspectRatio = msoTrue
            Logo.Height = 200
        End With
        
End Function

Public Sub PRJFormat()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim temp As String
    Dim ID As String
    Dim gfilename As String
    Dim fileName As Office.FileDialog
    
    Set fileName = Application.FileDialog(msoFileDialogFilePicker)
    
        With fileName
            .AllowMultiSelect = False
            .Filters.Clear
            If .Show = True Then
                tempVar = .SelectedItems.Item(1)
            End If
        End With
    
    'Opens WB based on excel file selection.
    'Pulls by company name as they were entered into the server, typos and all, based on cell value.
    
    Set wb = WorkBooks.Open(tempVar)
    ID = wb.Worksheets("Parameters").Range("B30").Value
    wb.Close Savechanges:=False
    
    If Len(ID) = 0 Then
        MsgBox "Null file selected."
        Exit Sub
    End If
    
    gfilename = "\\Repository\" & ID & ".png"
    
    If dir(gfilename) <> "" Then
        Call Cover(gfilename)
        Call Master(gfilename)
    Else
        Call ErrorHandle
        Exit Sub
    End If

End Sub
