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
    Dim i As Long, lastRow As Long

    On Error GoTo ErrorHandler

    ' Open Excel workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Excel runs in the background
    Set xlBook = xlApp.Workbooks.Open("C:\Users\T0285664\Desktop\Book1.xlsm", ReadOnly:=False)
    Set xlSheet = xlBook.Sheets("Sheet1")

    ' Loop through each word in the Word document
    For Each WordRange In ActiveDocument.Words
        ReqID = ExtractRequirementID(WordRange.Text)

        If ReqID <> "" Then
            ' Find the row in Excel with the requirement ID
            lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(-4162).Row
            For i = 1 To lastRow
                If xlSheet.Cells(i, 1).Value = ReqID Then
                    xlSheet.Cells(i, 3).Value = "Complete" ' Assuming status is in column 3
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

' Function to extract requirement ID from text
Function ExtractRequirementID(Text As String) As String
    Dim ReqID As String
    ReqID = ""
    
    ' Simple extraction based on a fixed format (e.g., "REQ001")
    If Text Like "*REQ###*" Then
        ReqID = Left(Text, 6) ' Extracts the first 6 characters
    End If

    ExtractRequirementID = ReqID
End Function




Function ExtractRequirementID(Text As String) As String
    Dim ReqID As String
    ReqID = ""
    
    ' Regular expression pattern to match requirement IDs
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "REQ_[\w\.]+"
    End With
    
    If regEx.test(Text) Then
        ' Extract the matched requirement ID
        ReqID = regEx.Execute(Text)(0)
    End If

    ExtractRequirementID = ReqID
End Function

```

```
Option Explicit ' Enforce variable declaration

Private WithEvents App As Word.Application

Private Sub Document_Open()
    Set App = Word.Application
End Sub

Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    MsgBox "Updating Excel File"
    MyMacroToRunBeforeSave
End Sub

Sub MyMacroToRunBeforeSave()
    UpdateExcelWithRequirements
End Sub

Sub UpdateExcelWithRequirements()
    On Error GoTo ErrorHandler

    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim WordRange As Range
    Dim ReqID As String
    Dim HeadingInfo As String
    Dim i As Long, lastRow As Long

    ' Open Excel workbook
    Set xlApp = New Excel.Application
    xlApp.Visible = False ' Excel runs in the background
    Set xlWorkbook = xlApp.Workbooks.Open("C:\Users\T0285664\Desktop\Book1.xlsm", ReadOnly:=False)
    Set xlSheet = xlWorkbook.Sheets("Sheet1")

    ' Loop through each word in the Word document
    For Each WordRange In ActiveDocument.Words
        ReqID = ExtractRequirementID(WordRange.Text)
        HeadingInfo = GetHeadingInfo(WordRange)

        If ReqID <> "" Then
            ' Find the row in Excel with the requirement ID
            lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
            For i = 1 To lastRow
                If xlSheet.Cells(i, "A").Value = ReqID Then ' Assuming Requirement ID is in column A
                    xlSheet.Cells(i, "C").Value = "Complete" ' Assuming status is in column C
                    xlSheet.Cells(i, "D").Value = HeadingInfo ' Assuming heading info is in column D
                    Exit For
                End If
            Next i
        End If
    Next WordRange

    ' Save and close the Excel workbook
    xlWorkbook.Close SaveChanges:=True
    xlApp.Quit

    ' Clean up
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    ' Ensure Excel is not left open in case of an error
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

' Function to extract requirement ID from text
Function ExtractRequirementID(Text As String) As String
    Dim ReqID As String
    ReqID = ""
    
    ' Regular expression pattern to match requirement IDs (REQ_0001, REQ_0002, etc.)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "REQ_\d{4}"
    End With
    
    If regEx.test(Text) Then
        ' Extract the matched requirement ID
        ReqID = regEx.Execute(Text)(0)
    End If

    ExtractRequirementID = ReqID
End Function

Function GetHeadingInfo(WordRange As Range) As String
    Dim HeadingInfo As String
    HeadingInfo = ""
    
    ' Check if the WordRange is within a heading
    If WordRange.Paragraphs.Count > 0 Then
        ' Get the paragraph style
        Dim paraStyle As String
        paraStyle = WordRange.Paragraphs(1).Style.NameLocal
        
        ' Check if it's a Heading 1, Heading 2, or Heading 3
        If paraStyle = "Heading 1" Or paraStyle = "Heading 2" Or paraStyle = "Heading 3" Then
            ' Get the text of the heading
            HeadingInfo = WordRange.Paragraphs(1).Range.Text
        End If
    End If

    GetHeadingInfo = HeadingInfo
End Function


```
