Attribute VB_Name = "Module3"
Sub NewCreate()
'Open the Excel workbook. Change the filename here.
Dim OWB As New Excel.Workbook
Set OWB = Excel.Application.Workbooks.Open("C:\Users\ozgur\OneDrive\Masaüstü\final.xlsx")
'Grab the first Worksheet in the Workbook
Dim WS As Excel.Worksheet
Set WS = OWB.Worksheets(1)
'Loop through each used row in Column A
'Copy the first slide and paste at the end of the presentation
For i = 2 To WS.Range("A65536").End(xlUp).Row
    ActivePresentation.Slides(1).Copy
    Dim objPresentaion As Presentation
    Dim objSlide As Slide
    Dim objImageBox As Shape
    
    Set objPresentaion = ActivePresentation
    Set objSlide = objPresentaion.Slides.Item(2)
    imagePath = WS.Cells(i, 7).Value
    Set objImageBox = objSlide.Shapes.AddPicture(imagePath, msoCTrue, msoCTrue, 100, 100)
    ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Text = WS.Cells(i, 2).Value + " " + WS.Cells(i, 3).Value
    ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.Text = "Major: " + WS.Cells(i, 4).Value
    ActivePresentation.Slides(2).Shapes(3).TextFrame.TextRange.Text = "Minor: " + WS.Cells(i, 5).Value
    ActivePresentation.Slides(2).Shapes(4).TextFrame.TextRange.Text = "Year: " + WS.Cells(i, 6).Value
    
    
    
    ActivePresentation.Slides.Paste (2)
    Next
End Sub
