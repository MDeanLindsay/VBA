Attribute VB_Name = "Module2"
Public Sub SlideHeaders()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim Claims1, Locations, Claims2, Claims3, BillReview, Pharmacy1, Pharmacy2, PatientMgmt, NurseTriage As Range
    Dim cell As Variant
    Dim fileName As Office.FileDialog
    
    'Open excel file with txt for EC slides.
    Set fileName = Application.FileDialog(msoFileDialogFilePicker)
    
        With fileName
            .AllowMultiSelect = False
            .Filters.Clear
            If .Show = True Then
                tempVar = .SelectedItems.Item(1)
            End If
        End With
    Set wb = WorkBooks.Open(tempVar)
    
    
    ''Text For Slides - Claims 1
    ''Slides 8 - 15

    Set Claims1 = wb.Worksheets("Text for Slides - Claims 1").Range("H5, F9, H13, E17, F21, F25, F29, F29, F33, F33")

    Dim A As Integer
    A = 8
    
    For Each cell In Claims1.cells

        Set Slide = ActivePresentation.Slides(A)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
            
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
            
        End With
        
        A = A + 1
        
    Next cell
    
    ''Text For Slides - Locations
    '' Slides 18 - 24
    
    Set Locations = wb.Worksheets("Text for Slides - Locations").Range("G5, F9, E13, G17, F21, E25, E29")
    
    Dim B As Integer
    B = 18
    
    For Each cell In Locations.cells

        Set Slide = ActivePresentation.Slides(B)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
            
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
        
        B = B + 1
        
    Next cell
    
    ''Text For Slides - Claims 2
    ''Slides 25 - 32
    
    Set Claims2 = wb.Worksheets("Text for Slides - Claims 2").Range("E5, E9, G13, G17, G21, G25, G29, F33")
    
    Dim C As Integer
    C = 25
    
    For Each cell In Claims2.cells

        Set Slide = ActivePresentation.Slides(C)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
            
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
            
        End With
         
        C = C + 1
        
    Next cell
    
    ''Text For Slides - Claims 3
    ''Slides 33 - 49
    
    Set Claims3 = wb.Worksheets("Text for Slides - Claims 3").Range("H5, G9, G13, F17, F21, F25, F29, F33, F37, F41, G45, F49, F53, F57, F61, E65, F69")
    
    Dim D As Integer
    D = 33
    
    For Each cell In Claims3.cells

        Set Slide = ActivePresentation.Slides(D)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
            
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
            
        End With
         
        D = D + 1
        
    Next cell
    
    ''Text For Slides = Bill Review
    ''Slides 51 - 62
    
    Set BillReview = wb.Worksheets("Text for Slides - Bill Review").Range("G2, G6, E10, E14, E18, E23, E27, F31, E35, F39, E43, E47")
    
    Dim E As Integer
    E = 51
    
    For Each cell In BillReview.cells

        Set Slide = ActivePresentation.Slides(E)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
        
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
         
        E = E + 1
        
    Next cell
    
    ''Text For Slides - Pharmacy
    ''Slides 64 - 79
    
    Set Pharmacy1 = wb.Worksheets("Text for Slides - Pharmacy").Range("G5, G9, E13, F18, G22, F26, E30, G34, F38, F42, G46, G50, G54, F58, F62, F68")
    
    Dim F As Integer
    F = 64
    
    For Each cell In Pharmacy1.cells

        Set Slide = ActivePresentation.Slides(F)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
        
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
         
        F = F + 1
        
    Next cell
    
    ''Pharmacy Text Slides [PBM Services]
    ''Slides 81 - 90
    
    Set Pharmacy2 = wb.Worksheets("Pharmacy Text Slides").Range("F5, F9, F13, F17, H21, H25, F29, I33, E41, F45")
    
    Dim G As Integer
    G = 81
    
    For Each cell In Pharmacy2.cells

        Set Slide = ActivePresentation.Slides(G)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
        
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
         
        G = G + 1
        
    Next cell
    
   ''Text For Slides - Patient Mgmt
   ''Slides 92 - 99
    
    Set PatientMgmt = wb.Worksheets("Text for Slides - Patient Mgmt").Range("E5, E5, E5, E9, F13, F17, F21, F25")
    
    Dim H As Integer
    H = 92
    
    For Each cell In PatientMgmt.cells

        Set Slide = ActivePresentation.Slides(H)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
        
            With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
         
        H = H + 1
        
    Next cell
    
   ''Text for Slides - Patient Mgmt [24/7 Nurse Triage]
   ''Slides 101 - 106
   
   Set NurseTriage = wb.Worksheets("Text for Slides - Patient Mgmt").Range("F29, F33, F37, F41, F45, F49")
   
    Dim I As Integer
    I = 101
    
    For Each cell In NurseTriage.cells

        Set Slide = ActivePresentation.Slides(I)
        With Slide.Shapes _
            .AddShape(msoShapeRectangle, 240, -115, 500, 120)
            .Fill.Transparency = 1
            .Line.Transparency = 1
            .TextFrame.TextRange.Text = cell.Value
        
             With .TextFrame.TextRange.Font
                .Size = 14
                .Name = "Gill Sans MT"
                .Color.RGB = RGB(0, 0, 0)
            End With
        
        End With
         
        I = I + 1
        
    Next cell
   
   
    wb.Close Savechanges:=False
    
End Sub
