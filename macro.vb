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

    myFile = myFile & ".anki-PPT_Macro" & "/"
    'add the .anki/ folder to the path
    
    Dim insertSlide As String
    insertSlide = ""
    Dim i As Integer
    For i = 1 To count Step 1
        'loop to create a new file for each slide notes

        Set oDestSlide = ActivePresentation.Slides(i)
        'set current slide to i
        
        Dim s As String * 8 'fixed length string with 8 characters
        Dim n As Integer
        Dim ch As Integer 'the character
        For n = 1 To Len(s) 'don't hardcode the length twice
            Do
                ch = Rnd() * 127 'This could be more efficient.
                '48 is '0', 57 is '9', 65 is 'A', 90 is 'Z', 97 is 'a', 122 is 'z'.
            Loop While ch < 48 Or ch > 57 And ch < 65 Or ch > 90 And ch < 97 Or ch > 122
            Mid(s, n, 1) = Chr(ch) 'bit more efficient than concatenation
        Next
        
        insertSlide = "<br><img src=" & Chr(39) & "Slide_" & s & CStr(i) & ".png" & Chr(39) & ">"
        'get string to append
        
        ActivePresentation.Slides(i).Export myFile & "Slide_" & s & CStr(i) & ".png", "PNG"
        'Export slide as PNG
        
        content = oDestSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text
        'get the text from the notes section (this is unfortunately plane text, so you must use html formatting tags)

        FilePath = myFile & CStr(i) & ".csv"
        'set the file path to .anki-PPT_Macro/i.csv

        Open FilePath For Append As #1
        'open the file

        Print #1, content & insertSlide
        'write the content

        Close #1
        'close the file
    Next
End Sub
