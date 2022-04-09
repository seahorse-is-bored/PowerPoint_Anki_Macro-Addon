Sub MoveToAnki()
    Dim oDestSlide As PowerPoint.Slide
    'setup oDestSlide as current presentation

    Dim content, myFile, FilePath As String
    Dim count As Integer
    'make some variables

    count = PowerPoint.ActivePresentation.Slides.count
    'count the total number of slides

    myFile = PowerPoint.ActivePresentation.Path
    'get the path of the presentation (must be saved)

    myFile = Replace(myFile, "Desktop", "") 
    myFile = Replace(myFile, "Documents", "")
    myFile = Replace(myFile, "Downloads", "")
    'Modify these for the appropriate path My example was made on the desktop

    myFile = myFile & ".anki" & "/"
    'add the .anki/ folder to the path
    
    For i = 1 To count Step 1
        'loop to create a new file for each slide notes

        Set oDestSlide = ActivePresentation.Slides(i)
        'set current slide to i

        content = oDestSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text
        'get the text from the notes section (this is unfortunately plane text, so you must use html formatting tags)

        FilePath = myFile & CStr(i) & ".csv"
        'set the file path to .anki/i.csv

        Open FilePath For Append As #1
        'open the file

        Print #1, content
        'write the content

        Close #1
        'close the file
    Next
End Sub
