Attribute VB_Name = "Module2"
Option Explicit


Public Sub getTodayReportsName()
    
    On Error GoTo ErrHandler
    
    Dim myFolder As Outlook.MAPIFolder
    Dim sFileTemplatePath, sFileEmailsPath, abcentReports, Reports As String
    Dim templateArray, emailsArray, Template, f As Variant
    Dim dtToday As Date
    Dim numberOfAbcentEmails, indexSubstr, i As Integer
    'ниже вставить свой путь до файла шаблона
    sFileTemplatePath = "C:\Git\python_projects\email_reader\templateReports.txt"
    'ниже вставить название своей папки, где хранятся отчеты
    Set myFolder = Application.Session.Folders.GetFirst.Folders.Item("Отчеты_автогенерация")

    dtToday = Date
   
    Set f = myFolder.Items
    
    f.Sort "[CreationTime]", False

    For i = f.count To 1 Step -1
        If Day(f(i).ReceivedTime) = Day(dtToday) And (Month(f(i).ReceivedTime) = Month(dtToday)) And (Year(f(i).ReceivedTime) = Year(dtToday)) Then
            Reports = Reports & f(i).Subject & " "
        Else: Exit For
        End If
    Next i
    
    templateArray = readAndTransformTextFileToArray(sFileTemplatePath)
    
    Set myFolder = Nothing
    
    For Each Template In templateArray
        indexSubstr = InStr(Reports, Template)
        If indexSubstr = 0 Then
            abcentReports = abcentReports & "- " & Template & Chr(10) & Chr(10)
            numberOfAbcentEmails = numberOfAbcentEmails + 1
        End If
    Next
    
    If abcentReports <> "" Then
        MsgBox "Отсутствуют следующие отчеты (" & CStr(numberOfAbcentEmails) & " шт.):" & Chr(10) & Chr(10) & abcentReports
    Else
        MsgBox "Все отчеты доставлены"
    End If
            
ErrHandler:
    Debug.Print Err.Description
End Sub


Function readAndTransformTextFileToArray(fileName As Variant)

    Dim objStream, strData
    Dim templateArray, emailsArray As Variant
    
    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (fileName)
    
    strData = objStream.ReadText()
    
    objStream.Close
    Set objStream = Nothing
    readAndTransformTextFileToArray = Split(strData, vbNewLine)

End Function
