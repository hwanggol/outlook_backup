Option Explicit

Private WithEvents InboxItems As Outlook.Items
Private WithEvents SentItems As Outlook.Items
Private Const BACKUP_BASE_PATH As String = "D:\Outlook_Backup\"

'===========================================
' Outlook 시작 시 자동 실행
'===========================================
Private Sub Application_Startup()
    InitializeEventHandler
End Sub

'===========================================
' 이벤트 핸들러 수동 초기화 (재시작 없이 즉시 활성화)
' 사용법: Alt+F8 → InitializeEventHandler 실행
'===========================================
Public Sub InitializeEventHandler()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' Inbox 이벤트 등록
    Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items

    ' Sent 이벤트 등록
    Set SentItems = ns.GetDefaultFolder(olFolderSentMail).Items

    Set ns = Nothing
    MsgBox "자동 백업 이벤트가 활성화되었습니다." & vbCrLf & _
           "Inbox 및 Sent 폴더가 모니터링됩니다.", vbInformation, "Outlook 백업 매크로"
End Sub

'===========================================
' 새 메일 수신 시 자동 저장
'===========================================
Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item

        SaveMailAsMSG mail, "Inbox"
        Set mail = Nothing
    End If

    Exit Sub
ErrorHandler:
    If TypeOf Item Is Outlook.MailItem Then
        LogError Item, "InboxItems_ItemAdd 에러: " & Err.Description
    End If
End Sub

'===========================================
' 메일 발송 시 자동 저장
'===========================================
Private Sub SentItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item

        SaveMailAsMSG mail, "Sent"
        Set mail = Nothing
    End If

    Exit Sub
ErrorHandler:
    If TypeOf Item Is Outlook.MailItem Then
        LogError Item, "SentItems_ItemAdd 에러: " & Err.Description
    End If
End Sub

'===========================================
' MSG 형식으로 저장
'===========================================
Sub SaveMailAsMSG(mail As Outlook.MailItem, folderType As String)
    On Error GoTo ErrorHandler

    Dim savePath As String
    Dim fileName As String
    Dim fullPath As String
    Dim mailTime As Date

    ' ReceivedTime 예외 처리 (임시저장/초안 메일 대응)
    If IsDate(mail.ReceivedTime) And mail.ReceivedTime > #1/1/1900# Then
        mailTime = mail.ReceivedTime
    ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
        mailTime = mail.CreationTime
    Else
        mailTime = Now
    End If

    ' 저장 경로 생성 (D:\Outlook_Backup\folderType\년도\월\)
    savePath = BACKUP_BASE_PATH & _
               folderType & "\" & _
               Format(mailTime, "yyyy") & "\" & _
               Format(mailTime, "mm") & "\"

    CreateFolderPath savePath

    ' 파일명 생성 (우선순위 기반 길이 제한)
    Dim dateTimePart As String
    Dim senderPart As String
    Dim subjectPart As String
    Dim maxPathLength As Integer
    Dim availableLength As Integer

    ' Windows 경로 제한 (260자)
    maxPathLength = 260

    ' 날짜_시간 부분 (15자, 필수)
    dateTimePart = Format(mailTime, "yyyymmdd_hhnnss")

    ' 사용 가능한 파일명 길이 계산
    ' = 최대경로 - 저장경로 - 날짜부분 - 언더스코어(2개) - 확장자(.msg=4자) - 여유(5자)
    availableLength = maxPathLength - Len(savePath) - Len(dateTimePart) - 2 - 4 - 5

    ' 발신자/수신자 결정 (folderType에 따라)
    Dim personName As String
    If folderType = "Sent" Then
        ' Sent: 수신자 사용 (첫 번째 To 수신자)
        If mail.Recipients.Count > 0 Then
            personName = mail.Recipients.Item(1).Name
        Else
            personName = "NoRecipient"
        End If
    Else
        ' Inbox: 발신자 사용
        personName = mail.SenderName
    End If

    ' 발신자/수신자 부분 (최대 50자)
    senderPart = CleanFileName(personName)
    If Len(senderPart) > 50 Then
        senderPart = Left(senderPart, 50)
    End If

    ' 제목 부분 (나머지 공간 사용)
    subjectPart = CleanFileName(mail.Subject)
    Dim remainingLength As Integer
    remainingLength = availableLength - Len(senderPart)

    If remainingLength < 20 Then
        ' 공간이 부족하면 발신자를 줄임
        senderPart = Left(senderPart, 30)
        remainingLength = availableLength - Len(senderPart)
    End If

    If Len(subjectPart) > remainingLength Then
        subjectPart = Left(subjectPart, remainingLength)
    End If

    ' 빈 제목 처리
    If Trim(subjectPart) = "" Then
        subjectPart = "NoSubject"
    End If

    ' 최종 파일명 조합 및 trailing 언더스코어 제거
    fileName = dateTimePart & "_" & senderPart & "_" & subjectPart
    Do While Right(fileName, 1) = "_"
        fileName = Left(fileName, Len(fileName) - 1)
    Loop

    fullPath = savePath & fileName & ".msg"

    ' MSG로 저장 (Outlook 내장 기능)
    mail.SaveAs fullPath, olMSG

    ' 저장 검증
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(fullPath) Then
        Dim fileSize As Long
        fileSize = fso.GetFile(fullPath).Size

        If fileSize < 100 Then
            LogError mail, "파일 크기 비정상: " & fullPath & " (" & fileSize & " bytes)"
        Else
            ' 저장 성공 로그 기록
            LogSuccess mail, fullPath, fileSize
        End If
    Else
        LogError mail, "파일 저장 실패: " & fullPath
    End If

    Set fso = Nothing

    Exit Sub
ErrorHandler:
    LogError mail, "SaveMailAsMSG 에러: " & Err.Description
End Sub

Sub CreateFolderPath(ByVal folderPath As String, Optional fso As Object = Nothing)
    Dim parentPath As String
    Dim needCleanup As Boolean

    ' FSO 객체가 전달되지 않았으면 새로 생성
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        needCleanup = True
    End If

    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    If fso.FolderExists(folderPath) Then
        If needCleanup Then Set fso = Nothing
        Exit Sub
    End If

    parentPath = fso.GetParentFolderName(folderPath)

    If Not fso.FolderExists(parentPath) Then
        CreateFolderPath parentPath, fso
    End If

    fso.CreateFolder folderPath

    ' 최상위 호출에서만 정리
    If needCleanup Then Set fso = Nothing
End Sub

'===========================================
' 파일명에 사용할 수 없는 문자 제거 (최적화)
'===========================================
Function CleanFileName(ByVal strFileName As String) As String
    Dim i As Integer
    Dim char As String
    Dim result As String
    Dim lastWasSpace As Boolean
    Dim lastWasUnderscore As Boolean

    result = ""
    lastWasSpace = False
    lastWasUnderscore = False

    For i = 1 To Len(strFileName)
        char = Mid(strFileName, i, 1)

        Select Case char
            Case "/", "\", ":", "*", "?", """", "<", ">", "|"
                ' 특수문자 → 언더스코어 (연속 방지)
                If Not lastWasUnderscore Then
                    result = result & "_"
                    lastWasUnderscore = True
                    lastWasSpace = False
                End If

            Case vbCr, vbLf
                ' 줄바꿈 문자 제거

            Case vbTab, " "
                ' 탭/공백 → 공백 (연속 방지)
                If Not lastWasSpace Then
                    result = result & " "
                    lastWasSpace = True
                    lastWasUnderscore = False
                End If

            Case "_"
                ' 언더스코어 (연속 방지)
                If Not lastWasUnderscore Then
                    result = result & "_"
                    lastWasUnderscore = True
                    lastWasSpace = False
                End If

            Case Else
                ' 일반 문자
                result = result & char
                lastWasSpace = False
                lastWasUnderscore = False
        End Select
    Next i

    CleanFileName = Trim(result)
End Function

'===========================================
' 수동 실행: 선택한 메일 저장
'===========================================
Public Sub SaveSelectedMailsAsMSG()
    On Error GoTo ErrorHandler
    
    Dim selectedItems As Outlook.Selection
    Dim Item As Object
    Dim mail As Outlook.MailItem
    Dim savedCount As Long
    Dim totalCount As Long
    
    Set selectedItems = Application.ActiveExplorer.Selection
    
    If selectedItems.Count = 0 Then
        MsgBox "저장할 메일을 선택해주세요.", vbExclamation
        Exit Sub
    End If
    
    totalCount = selectedItems.Count
    savedCount = 0
    
    If totalCount > 50 Then
        If MsgBox(totalCount & "개의 메일을 백업합니다. 진행하시겠습니까?", _
                  vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    For Each Item In selectedItems
        If TypeOf Item Is Outlook.MailItem Then
            Set mail = Item

            ' 메일이 속한 폴더 감지
            Dim folderName As String
            Dim folderType As String
            folderName = mail.Parent.Name

            ' Sent 폴더 변형 처리
            If InStr(1, folderName, "보낸", vbTextCompare) > 0 Or _
               InStr(1, folderName, "Sent", vbTextCompare) > 0 Then
                folderType = "Sent"
            Else
                folderType = "Inbox"
            End If

            SaveMailAsMSG mail, folderType
            savedCount = savedCount + 1
            DoEvents
        End If
    Next Item

    MsgBox savedCount & "개의 메일이 저장되었습니다." & vbCrLf & _
           "경로: " & BACKUP_BASE_PATH, vbInformation

    ' 메모리 정리
    Set mail = Nothing
    Set selectedItems = Nothing

    Exit Sub
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
    Set mail = Nothing
    Set selectedItems = Nothing
End Sub

'===========================================
' 로그 파일 경로 생성 (월별 로테이션)
'===========================================
Function GetLogFilePath(Optional logType As String = "error") As String
    Dim logFolder As String
    logFolder = BACKUP_BASE_PATH & "logs\"
    CreateFolderPath logFolder
    GetLogFilePath = logFolder & Format(Now, "yyyy-mm") & "_" & logType & ".log"
End Function

'===========================================
' 에러 로깅 함수
'===========================================
Sub LogError(mail As Outlook.MailItem, errMsg As String)
    On Error Resume Next
    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String
    Dim logEntry As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = GetLogFilePath()

    ' Append 모드로 열기 (파일 없으면 자동 생성)
    Set logFile = fso.OpenTextFile(logPath, 8, True)

    ' 로그 기록: [날짜시간] | 에러내용 | 발신자 | 제목
    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               errMsg & " | " & _
               mail.SenderName & " | " & _
               mail.Subject

    logFile.WriteLine logEntry
    logFile.Close

    Set logFile = Nothing
    Set fso = Nothing
End Sub

'===========================================
' 성공 로깅 함수
'===========================================
Sub LogSuccess(mail As Outlook.MailItem, filePath As String, fileSize As Long)
    On Error Resume Next
    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String
    Dim logEntry As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = GetLogFilePath("success")

    ' Append 모드로 열기 (파일 없으면 자동 생성)
    Set logFile = fso.OpenTextFile(logPath, 8, True)

    ' 로그 기록: [날짜시간] | SUCCESS | 파일경로 | 파일크기 | 발신자 | 제목
    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               "SUCCESS" & " | " & _
               filePath & " | " & _
               FormatNumber(fileSize, 0) & " bytes" & " | " & _
               mail.SenderName & " | " & _
               mail.Subject

    logFile.WriteLine logEntry
    logFile.Close

    Set logFile = Nothing
    Set fso = Nothing
End Sub
