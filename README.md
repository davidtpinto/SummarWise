# Summarwise: PDF Summarizer

Summarwise is an innovative **Software as a Service (SaaS)** application designed to simplify the process of summarizing PDF documents. By leveraging advanced _natural language processing (NLP)_ and machine learning algorithms, Summarwise automatically condenses lengthy PDF files into concise and digestible summaries.

## Key Features

### 1. Automated Summarization

Summarwise uses cutting-edge _NLP algorithms_ to analyze PDF content and extract key information. The automated summarization process ensures users quickly grasp the main ideas without reading the entire document.

### 2. Customizable Summaries

Tailor summaries to your preferences. Whether you need a brief overview or a detailed summary, Summarwise adapts to your needs.

### 3. Time-Saving

Save valuable time by automating the summarization process, eliminating the need to manually extract information from lengthy PDFs.

### 4. Accurate Content Extraction

Advanced algorithms ensure accurate content extraction, focusing on the most relevant information within the document.

### 5. User-Friendly Interface

An intuitive and user-friendly interface enhances the user experience and facilitates efficient navigation.

### 6. Collaboration Features

Easily share summarized documents with team members or collaborators, promoting seamless communication and information sharing.

### 7. Security

Prioritizing the security and privacy of user data, Summarwise employs robust encryption measures to handle sensitive information within PDF documents.

### 8. Integration Capabilities

Summarwise integrates with popular cloud storage services, document management platforms, and productivity tools for a seamless workflow.

## Getting Started

### Installation

# Clone the repository

git clone https://github.com/davidtpinto/summarwise.git

# Change directory

cd summarwise

# Install dependencies

npm install

### Usage

# Run the application

npm start

Visit [http://localhost:3000](http://localhost:3000) in your browser to start using Summarwise.

## Contributing

We welcome contributions! Please follow our [Contribution Guidelines](CONTRIBUTING.md) for more information.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgments

Thank you to the open source community for their invaluable contributions.

---

```
--- ThisDocument

Private WithEvents App As Word.Application

Private Sub Document_Open()
Set App = Word.Application
End Sub

Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
MsgBox "Updating Excel File"
MyMacroToRunBeforeSave
End Sub

--- Module1

Sub MyMacroToRunBeforeSave()
    ' Your code here
    UpdateExcelWithRequirements
End Sub


Sub UpdateExcelWithRequirements()
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim WordRange As Range
    Dim ReqID As String
    Dim HeadingInfo As String
    Dim i As Long, lastRow As Long

    On Error GoTo ErrorHandler

    ' Open Excel workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Excel runs in the background
    Set xlBook = xlApp.Workbooks.Open("C:\Users\T0285664\Desktop\Book1.xlsm", ReadOnly:=False)
    Set xlSheet = xlBook.Sheets("Sheet2")

        ' Loop through each word in the Word document
    For Each WordRange In ActiveDocument.Words
        ReqID = ExtractRequirementID(WordRange.Text)
        HeadingInfo = GetHeadingInfo(WordRange)

        If ReqID <> "" Then
            ' Find the row in Excel with the requirement ID
            lastRow = xlSheet.Cells(xlSheet.Rows.Count, "B").End(-4162).Row
            For i = 1 To lastRow
                If xlSheet.Cells(i, "B").Value = ReqID Then ' Assuming Requirement ID is in column A
                    xlSheet.Cells(i, "E").Value = "Complete" ' Assuming status is in column C
                    xlSheet.Cells(i, "H").Value = HeadingInfo ' Assuming heading info is in column D
                    Exit For
                End If
            Next i
        End If
    Next WordRange

    ' Save and close the Excel workbook
    xlBook.Close SaveChanges:=True
    xlApp.Quit

    ' Clean up
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ' Ensure Excel is not left open in case of an error
    If Not xlBook Is Nothing Then
        xlBook.Close SaveChanges:=False
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub


Function ExtractRequirementID(Text As String) As String
    Dim ReqID As String
    ReqID = ""
        
    If Text Like "REQ####" Then
        ' Extract the matched requirement ID
        ReqID = Text
    End If

    ExtractRequirementID = ReqID
End Function

' Function to get heading information from the Word document
Function GetHeadingInfo(WordRange As Range) As String
    Dim HeadingInfo As String
    HeadingInfo = ""
    
    ' Define a variable to store the last encountered heading
    Dim LastHeading As String
    LastHeading = ""
    
    ' Loop through the document to find the last heading before the WordRange
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Start >= WordRange.Start Then
            Exit For ' Exit the loop when reaching or surpassing WordRange
        ElseIf para.Style.NameLocal = "Heading 1" Or _
               para.Style.NameLocal = "Heading 2" Or _
               para.Style.NameLocal = "Heading 3" Then
            ' Update LastHeading with the current heading text
            LastHeading = para.Range.Text
        End If
    Next para
    
    ' Set HeadingInfo to the last encountered heading
    HeadingInfo = LastHeading
    Debug.Print HeadingInfo
    GetHeadingInfo = HeadingInfo
End Function

```

```
Sub UpdateExcelWithRequirements()
    On Error GoTo ErrorHandler

    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim WordParagraph As Paragraph
    Dim ReqID As String
    Dim HeadingInfo As String
    Dim i As Long, lastRow As Long, insertRow As Long

    ' Open Excel workbook
    Set xlApp = New Excel.Application
    xlApp.Visible = False
    Set xlWorkbook = xlApp.Workbooks.Open("C:\Users\T0285664\Desktop\Book1.xlsm", ReadOnly:=False)
    Set xlSheet = xlWorkbook.Sheets("Sheet1")

    ' Loop through each paragraph in the Word document
    For Each WordParagraph In ActiveDocument.Paragraphs
        ' Use regular expressions to find REQs (e.g., REQ0001, REQ0002, etc.) in the paragraph
        Set RegEx = CreateObject("VBScript.RegExp")
        RegEx.Global = True
        RegEx.IgnoreCase = True
        RegEx.Pattern = "\[REQ\d+\]" ' Match REQ IDs in square brackets

        ' Find all matches in the paragraph
        Set Matches = RegEx.Execute(WordParagraph.Range.Text)
        
        ' Loop through all matched REQs in the paragraph
        For Each Match In Matches
            ReqID = Mid(Match.Value, 2, Len(Match.Value) - 2) ' Extract REQ ID without brackets
            HeadingInfo = WordParagraph.Range.Paragraphs(1).Range.Text ' Get the heading text
            
            ' Find the row in Excel with the requirement ID and update it
            lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
            For i = 1 To lastRow
                If xlSheet.Cells(i, "A").Value = ReqID Then ' Assuming Requirement ID is in column A
                    xlSheet.Cells(i, "C").Value = "Complete" ' Assuming status is in column C
                    xlSheet.Cells(i, "D").Value = HeadingInfo ' Assuming heading info is in column D
                    Exit For
                End If
            Next i
        Next Match
    Next WordParagraph


    ' Save and close the workbook
    xlWorkbook.Close SaveChanges:=True
    xlApp.Quit

    ' Clean up
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close SaveChanges:=False
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
End Sub


```
