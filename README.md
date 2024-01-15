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
Private Sub DocumentBeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
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
    Set xlBook = xlApp.Workbooks.Open("C:\path\to\your\workbook.xlsm", ReadOnly:=False)
    Set xlSheet = xlBook.Sheets("Sheet1")

    ' Loop through each word in the Word document
    For Each WordRange In ActiveDocument.Words
        If VarType(WordRange) = vbObject Then
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
Function ExtractRequirementID(Text As Variant) As String
    Dim ReqID As String
    ReqID = ""
    
    If VarType(Text) = vbString Then
        If Text Like "*REQ###*" Then
            ReqID = Left(Text, 6) ' Extracts the first 6 characters
        End If
    End If

    ExtractRequirementID = ReqID
End Function

```
