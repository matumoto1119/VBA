Attribute VB_Name = "解析"
Option Explicit
Public ie As InternetExplorer
Public htdoc As HTMLDocument


Sub AnalyzeTagName()
Set ie = getIE
Dim book As Workbook
Dim s As Integer

Dim Tag As String
Tag = Application.InputBox("タグ名を入力")
If Tag = "False" Then
Exit Sub
End If

For Each book In Workbooks
    If book.name = "Book1" Then
    s = 1
    Exit For
    End If
Next

If s = 1 Then
Workbooks("Book1").Activate
Worksheets.Add after:=Worksheets(Worksheets.Count)
Else
Workbooks.Add
End If

Set htdoc = ie.document
Dim i As Long, r As Long
r = 1

For i = 0 To htdoc.all.Length - 1
    With htdoc.all(i)
    If .tagName = Tag Then
        Cells(r, 1) = r - 1
        Cells(r, 2) = .tagName
        Cells(r, 3) = .innerText
        r = r + 1
    End If
    End With
Next
Range("A:D").WrapText = False
Range("A:D").Columns.AutoFit

End Sub

Sub AnalyzeclassName()
Set ie = getIE
Dim book As Workbook
Dim s As Integer

Dim Class As String
Class = Application.InputBox("クラス名を入力")
If Class = "False" Then
Exit Sub
End If

For Each book In Workbooks
    If book.name = "Book1" Then
    s = 1
    Exit For
    End If
Next

If s = 1 Then
Workbooks("Book1").Activate
Worksheets.Add after:=Worksheets(Worksheets.Count)
Else
Workbooks.Add
End If

Set htdoc = ie.document
Dim i As Long, r As Long
r = 1

For i = 0 To htdoc.all.Length - 1
    With htdoc.all(i)
    If .className = Class Then
        Cells(r, 1) = r - 1
        Cells(r, 2) = .className
        Cells(r, 3) = .innerText
        r = r + 1
    End If
    End With
Next
Range("A:D").WrapText = False
Range("A:D").Columns.AutoFit

End Sub

Sub AnalyzeHTML()
Set ie = getIE
Dim htdoc As HTMLDocument
Set htdoc = ie.document

Dim ret As String
ret = htdoc.getElementsByTagName("HTML")(0).outerHTML & vbCrLf

Dim filename As String
filename = ThisWorkbook.Path & "解析用.txt"
Dim filenum As Integer
filenum = FreeFile

Open filename For Output As #filenum
    Print #filenum, ret

End Sub

Public Sub WaitBrowsing(ie As InternetExplorer)
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
End Sub

Public Function getIE() As InternetExplorer

Dim objshell As Object, objWindow As Object
Dim document_title As String
Dim i As Integer
i = 0
Set objshell = CreateObject("Shell.Application")

For Each objWindow In objshell.Windows
    On Error Resume Next
    document_title = objWindow.document.Title
    On Error GoTo 0
    If objWindow = "Internet Explorer" Then
        Set getIE = objWindow
        Exit For
    End If
Next

End Function
