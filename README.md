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

Sub UpdateExcelWithRequirements()
    ' Your code here
    MsgBox "Updating Excel File"
    ExcelModule
End Sub


Function ExcelModule()
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim WordRange As Range
    Dim ReqID As String
    Dim HeadingInfo As String
    Dim i As Long, lastRow As Long
    
    ' Path of the Word document
    Dim wordDocPath As String
    wordDocPath = ActiveDocument.Path

    ' Construct the relative path for the Excel file
    ' For example, if the Excel file is in the same directory:
    Dim excelFilePath As String
    excelFilePath = wordDocPath & "\230700098_TDCD_CMCDRL_20240123_1.xlsm" ' Change '230700098_TDCD_CMCDRL_20240123_1.xlsm' to your Excel file's name
    
    ' Open Excel workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Excel runs in the background
    Set xlBook = xlApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set xlSheet = xlBook.Sheets("Compliance Matrix")

    ' Get the name of the current Word document
    Dim wordFileName As String
    wordFileName = ActiveDocument.Name

    ' Find the column for the current Word document or add a new one
    Dim wordFileColumn As Long
    wordFileColumn = FindOrAddWordFileColumn(xlSheet, wordFileName)
    
    ' Loop through each paragraph in the Word document
    For Each WordParagraph In ActiveDocument.Paragraphs
        ' Use regular expressions to find REQs (e.g., REQ0001, REQ0002, etc.) in the paragraph
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.Pattern = "\[REQ\d+\]" ' Match REQ IDs in square brackets

        ' Find all matches in the paragraph
        Set Matches = regEx.Execute(WordParagraph.Range.Text)
        
        ' Loop through all matched REQs in the paragraph
        For Each Match In Matches
            
            ReqID = Match.Value ' Extract REQ ID without brackets
            HeadingInfo = GetHeadingInfo(WordParagraph.Range) ' Get the heading text
                        
            ' Find the row in Excel with the requirement ID and update it
            lastRow = xlSheet.Cells(xlSheet.Rows.Count, "D").End(-4162).Row
            For i = 1 To lastRow
                If xlSheet.Cells(i, "D").Value = ReqID Then ' Assuming Requirement ID is in column D
                    xlSheet.Cells(i, "H").Value = "Complete" ' Assuming status is in column H
                     ' Append new heading info to existing data in column H
                        Dim existingData As String
                        existingData = xlSheet.Cells(i, wordFileColumn).Value
                        If existingData <> "" Then
                            xlSheet.Cells(i, wordFileColumn).Value = existingData & ", " & HeadingInfo
                        Else
                            xlSheet.Cells(i, wordFileColumn).Value = HeadingInfo
                        End If
                    Exit For
                End If
            Next i
        Next Match


 ' Loop through all matched REQs in the paragraph
    For Each Match In Matches
        ReqID = Match.Value ' Keep the brackets around the REQ ID
        HeadingInfo = GetHeadingInfo(WordParagraph.Range, Match.FirstIndex) ' Get the heading text

        ' Update the sections in the correct column for each requirement
        Dim currentRow As Long
        currentRow = FindRowForReqID(xlSheet, ReqID)
        If currentRow > 0 Then
            UpdateSectionInCell xlSheet.Cells(currentRow, wordFileColumn), HeadingInfo
        End If
    Next Match
    Next WordParagraph

    ' Save and close the Excel workbook
    xlBook.Close SaveChanges:=True
    xlApp.Quit

    ' Clean up
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    Exit Function

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
End Function


Function ExtractRequirementID(Text As String) As String
    Dim ReqID As String
    ReqID = ""
        
    If Text Like "REQ####" Then
        ' Extract the matched requirement ID
        ReqID = Text
    End If

    ExtractRequirementID = ReqID
End Function

Function FindRowForReqID(xlSheet As Excel.Worksheet, reqID As String) As Long
    Dim i As Long
    FindRowForReqID = 0 ' Default to 0 (not found)
    For i = 1 To xlSheet.Cells(xlSheet.Rows.Count, "A").End(-4162).Row ' -4162 corresponds to xlUp
        If xlSheet.Cells(i, "A").Value = reqID Then
            FindRowForReqID = i
            Exit Function
        End If
    Next i
End Function


' Function to get heading information from the Word document
Function GetHeadingInfo(WordRange As Range) As String
    Dim para As Word.Paragraph
    GetHeadingInfo = ""
     ' Loop through the document to find the last heading before the WordRange
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Start >= WordRange.Start Then
            Exit For ' Exit the loop when reaching or surpassing WordRange
        ElseIf para.Style.NameLocal <> "Normal" And para.Style.NameLocal <> "2-Legend" And para.Style.NameLocal <> "0-BODY TEXT" Then
            Debug.Print para.Style.NameLocal
            ' Update LastHeading with the current heading text
            If Not para.Range.ListFormat Is Nothing And para.Range.ListFormat.ListString <> "" Then
                GetHeadingInfo = para.Range.ListFormat.ListString & " " & para.Range.Text
            Else
                GetHeadingInfo = para.Range.Text
                End If
         End If
    Next para
  
End Function

Function GetHeadingInfo(rng As Word.Range, matchPosition As Integer) As String
    Dim para As Word.Paragraph
    GetHeadingInfo = ""
    For Each para In rng.Document.Paragraphs
        If para.Range.End <= matchPosition Then
            If para.Style <> "Normal" Then
                ' Check if there is numbering and include it
                If Not para.Range.ListFormat Is Nothing Then
                    If para.Range.ListFormat.ListString <> "" Then
                        GetHeadingInfo = para.Range.ListFormat.ListString & " " & para.Range.Text
                    Else
                        GetHeadingInfo = para.Range.Text
                    End If
                Else
                    GetHeadingInfo = para.Range.Text
                End If
            End If
        Else
            Exit For
        End If
    Next para
End Function


Function UpdateSectionInCell(cell As Excel.Range, section As String)
    Dim existingSections As String
    existingSections = cell.Value

    Dim sectionsArray() As String
    If existingSections <> "" Then
        sectionsArray = Split(existingSections, ", ")
        If IsInArray(section, sectionsArray) = False Then
            cell.Value = existingSections & ", " & section
        End If
    Else
        cell.Value = section
    End If
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function FindOrAddWordFileColumn(xlSheet As Excel.Worksheet, wordFileName As String) As Long
    Dim i As Long
    For i = 9 To xlSheet.Cells(5, xlSheet.Columns.Count).End(xlExcelToLeft).Column
        If xlSheet.Cells(5, i).Value = wordFileName Then
            FindOrAddWordFileColumn = i
            Exit Function
        End If
    Next i

    ' If the file name is not found, add it to the next available column
    FindOrAddWordFileColumn = i
    xlSheet.Cells(5, i).Value = wordFileName
End Function
```
