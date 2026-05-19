Attribute VB_Name = "expertPPT"
Option Explicit
Option Base 1

Sub SplitPPTinFolder(filePath As String)
    Dim oPres As presentation, newPres As presentation
    Dim targetPath As String, targetName As String
    Dim sourceName As String
    Dim BaseName As String
    Dim i As Integer, j As Integer
    Dim slideCount As Integer

    On Error GoTo ErrorHandler
    Set oPres = Presentations.Open(filePath, WithWindow:=msoFalse)

    sourceName = oPres.path & "\" & oPres.Name
    slideCount = oPres.slides.count
    targetPath = oPres.path & "\��ҳ\" & GetBaseName(oPres.Name) & "_"

    For i = 1 To slideCount
        targetName = targetPath & i & ".pptx"
        FileCopy filePath, targetName
        Set newPres = Presentations.Open(targetName, WithWindow:=msoFalse)
        For j = newPres.slides.count To 1 Step -1
            If j <> i Then
                newPres.slides(j).Delete
            End If
        Next j
        newPres.Save
        newPres.Close
    Next i
    oPres.Close
    Exit Sub

    ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not newPres Is Nothing Then
        newPres.Close
    End If
    If Not oPres Is Nothing Then
        oPres.Close
    End If
End Sub


Sub SplitPPTinPPT()
    Dim oPres As presentation, newPres As presentation
    Dim targetPath As String, targetName As String
    Dim sourceName As String
    Dim i As Integer, j As Integer
    Dim slideCount As Integer
    Dim oSld As slide

    On Error GoTo ErrorHandler
    Set oPres = ActivePresentation
    If oPres.path = "" Then
        MsgBox "Please save the PPT file first!"
        Exit Sub
    End If

    sourceName = oPres.path & "\" & oPres.Name
    targetPath = oPres.path & "\��ҳ\" & GetBaseName(oPres.Name) & "_"

    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        For Each oSld In ActiveWindow.Selection.SlideRange
            targetName = targetPath & oSld.slideIndex & ".pptx"
            FileCopy sourceName, targetName
            Set newPres = Presentations.Open(targetName, WithWindow:=msoFalse)
            For j = newPres.slides.count To 1 Step -1
                If j <> oSld.slideIndex Then
                    newPres.slides(j).Delete
                End If
            Next j
            newPres.Save
            newPres.Close
        Next oSld
    Else
        For Each oSld In oPres.slides
            targetName = targetPath & oSld.slideIndex & ".pptx"
            FileCopy sourceName, targetName
            Set newPres = Presentations.Open(targetName, WithWindow:=msoFalse)
            For j = newPres.slides.count To 1 Step -1
                If j <> oSld.slideIndex Then
                    newPres.slides(j).Delete
                End If
            Next j
            newPres.Save
            newPres.Close
        Next oSld
    End If
    Exit Sub

    ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not newPres Is Nothing Then
        newPres.Close
    End If
End Sub

Sub exportAsImageInFolder(filePath As String, exportType As String)
    Dim oPres As presentation
    Dim targetPath As String, targetName As String
    Dim oSld As slide
    Dim validTypes As Variant

    validTypes = Array("JPG", "PNG", "GIF")
    If Not IsInArray(exportType, validTypes) Then
        MsgBox "Unsupported file type: " & exportType, vbExclamation, "Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Set oPres = Presentations.Open(filePath, WithWindow:=msoFalse)
    targetPath = oPres.path & "\ͼƬ\" & GetBaseName(oPres.Name) & "_"

    For Each oSld In oPres.slides
        targetName = targetPath & oSld.slideIndex & "." & exportType
        oSld.Export targetName, exportType, 1920, 1080
    Next oSld

    oPres.Close
    Exit Sub

    ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not oPres Is Nothing Then
        oPres.Close
    End If
End Sub

Sub exportAsImageInPPT(exportType As String)
    Dim oPres As presentation
    Dim targetPath As String, targetName As String
    Dim oSld As slide
    Dim validTypes As Variant

    validTypes = Array("JPG", "PNG", "GIF")
    If Not IsInArray(exportType, validTypes) Then
        MsgBox "Unsupported file type: " & exportType, vbExclamation, "Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Set oPres = ActivePresentation
    If oPres.path = "" Then
        MsgBox "Please save the PPT file first!"
        Exit Sub
    End If
    targetPath = oPres.path & "\ͼƬ\" & GetBaseName(oPres.Name) & "_"


    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        For Each oSld In ActiveWindow.Selection.SlideRange
            targetName = targetPath & oSld.slideIndex & "." & exportType
            oSld.Export targetName, exportType, 1920, 1080
        Next oSld
    Else
        For Each oSld In oPres.slides
            targetName = targetPath & oSld.slideIndex & "." & exportType
            oSld.Export targetName, exportType, 1920, 1080
        Next oSld
    End If
    Exit Sub

    ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function FolderExists(path As String) As Boolean
    FolderExists = (Dir(path, vbDirectory) <> "")
End Function

Function GetBaseName(fileName As String) As String
    Dim dotPosition As Integer
    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 0 Then
        GetBaseName = left(fileName, dotPosition - 1)
    Else
        GetBaseName = fileName
    End If
End Function
