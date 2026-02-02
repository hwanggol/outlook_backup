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

    ' 중복 저장 체크
    If IsMailAlreadySaved(mail) Then
        ' 이미 저장된 메일이면 스킵
        Exit Sub
    End If

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
            ' EntryID를 인덱스에 추가 (중복 저장 방지)
            AddEntryIDToIndex mail
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
' 저장된 메일 EntryID 인덱스 파일 경로
'===========================================
Function GetEntryIDIndexPath() As String
    Dim indexFolder As String
    indexFolder = BACKUP_BASE_PATH & "logs\"
    CreateFolderPath indexFolder
    GetEntryIDIndexPath = indexFolder & "saved_entries.txt"
End Function

'===========================================
' 메일이 이미 저장되었는지 확인
'===========================================
Function IsMailAlreadySaved(mail As Outlook.MailItem) As Boolean
    On Error Resume Next
    
    Dim entryID As String
    entryID = mail.EntryID
    
    ' EntryID가 없으면 저장되지 않은 것으로 간주
    If entryID = "" Then
        IsMailAlreadySaved = False
        Exit Function
    End If
    
    Dim fso As Object
    Dim indexFile As Object
    Dim indexPath As String
    Dim line As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    indexPath = GetEntryIDIndexPath()
    
    ' 인덱스 파일이 없으면 저장되지 않은 것으로 간주
    If Not fso.FileExists(indexPath) Then
        IsMailAlreadySaved = False
        Set fso = Nothing
        Exit Function
    End If
    
    ' 인덱스 파일에서 EntryID 검색
    Set indexFile = fso.OpenTextFile(indexPath, 1, False)
    
    Do While Not indexFile.AtEndOfStream
        line = Trim(indexFile.ReadLine)
        If line = entryID Then
            indexFile.Close
            Set indexFile = Nothing
            Set fso = Nothing
            IsMailAlreadySaved = True
            Exit Function
        End If
    Loop
    
    indexFile.Close
    Set indexFile = Nothing
    Set fso = Nothing
    IsMailAlreadySaved = False
End Function

'===========================================
' 저장된 메일 EntryID를 인덱스에 추가
'===========================================
Sub AddEntryIDToIndex(mail As Outlook.MailItem)
    On Error Resume Next
    
    Dim entryID As String
    entryID = mail.EntryID
    
    ' EntryID가 없으면 인덱스에 추가하지 않음
    If entryID = "" Then
        Exit Sub
    End If
    
    Dim fso As Object
    Dim indexFile As Object
    Dim indexPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    indexPath = GetEntryIDIndexPath()
    
    ' Append 모드로 열기 (파일 없으면 자동 생성)
    Set indexFile = fso.OpenTextFile(indexPath, 8, True)
    indexFile.WriteLine entryID
    indexFile.Close
    
    Set indexFile = Nothing
    Set fso = Nothing
End Sub

'===========================================
' 내부: 기간별 메일 백업 공통 로직
'===========================================
Private Sub BackupMailsByDateRange(ByVal startDate As Date, ByVal endDate As Date, ByVal periodLabel As String, ByVal errorSubName As String)
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Dim inboxFolder As Outlook.MAPIFolder
    Dim sentFolder As Outlook.MAPIFolder
    Dim inboxItems As Outlook.Items
    Dim sentItems As Outlook.Items
    Dim Item As Object
    Dim mail As Outlook.MailItem
    Dim filterString As String
    Dim savedCount As Long
    Dim skippedCount As Long
    Dim totalCount As Long
    
    Set ns = Application.GetNamespace("MAPI")
    Set inboxFolder = ns.GetDefaultFolder(olFolderInbox)
    Set sentFolder = ns.GetDefaultFolder(olFolderSentMail)
    
    savedCount = 0
    skippedCount = 0
    totalCount = 0
    
    ' Inbox 메일 백업
    Set inboxItems = inboxFolder.Items
    ' Outlook 날짜 필터 형식: "mm/dd/yyyy hh:nn AM/PM"
    filterString = "[ReceivedTime] >= '" & Format(startDate, "mm/dd/yyyy hh:nn AM/PM") & "' AND [ReceivedTime] <= '" & Format(endDate, "mm/dd/yyyy hh:nn AM/PM") & "'"
    
    On Error Resume Next
    Set inboxItems = inboxItems.Restrict(filterString)
    If Err.Number <> 0 Then
        ' 필터 실패 시 전체 아이템 사용 (날짜는 나중에 직접 비교)
        Set inboxItems = inboxFolder.Items
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If Not inboxItems Is Nothing Then
        On Error Resume Next
        inboxItems.Sort "[ReceivedTime]", True
        On Error GoTo ErrorHandler
        
        For Each Item In inboxItems
            If TypeOf Item Is Outlook.MailItem Then
                Set mail = Item
                
                ' 날짜 범위 확인 (필터가 실패한 경우를 대비)
                Dim mailDate As Date
                If IsDate(mail.ReceivedTime) And mail.ReceivedTime > #1/1/1900# Then
                    mailDate = mail.ReceivedTime
                ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
                    mailDate = mail.CreationTime
                Else
                    mailDate = #1/1/1900#
                End If
                
                ' 날짜 범위 내에 있는지 확인
                If mailDate >= startDate And mailDate <= endDate Then
                    totalCount = totalCount + 1
                    
                    ' 중복 체크
                    If IsMailAlreadySaved(mail) Then
                        skippedCount = skippedCount + 1
                    Else
                        SaveMailAsMSG mail, "Inbox"
                        savedCount = savedCount + 1
                    End If
                End If
                
                Set mail = Nothing
                DoEvents
            End If
        Next Item
    End If
    
    ' Sent 메일 백업
    Set sentItems = sentFolder.Items
    ' Outlook 날짜 필터 형식: "mm/dd/yyyy hh:nn AM/PM"
    filterString = "[SentOn] >= '" & Format(startDate, "mm/dd/yyyy hh:nn AM/PM") & "' AND [SentOn] <= '" & Format(endDate, "mm/dd/yyyy hh:nn AM/PM") & "'"
    
    On Error Resume Next
    Set sentItems = sentItems.Restrict(filterString)
    If Err.Number <> 0 Then
        ' 필터 실패 시 전체 아이템 사용 (날짜는 나중에 직접 비교)
        Set sentItems = sentFolder.Items
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If Not sentItems Is Nothing Then
        On Error Resume Next
        sentItems.Sort "[SentOn]", True
        On Error GoTo ErrorHandler
        
        For Each Item In sentItems
            If TypeOf Item Is Outlook.MailItem Then
                Set mail = Item
                
                ' 날짜 범위 확인 (필터가 실패한 경우를 대비)
                Dim mailSentDate As Date
                If IsDate(mail.SentOn) And mail.SentOn > #1/1/1900# Then
                    mailSentDate = mail.SentOn
                ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
                    mailSentDate = mail.CreationTime
                Else
                    mailSentDate = #1/1/1900#
                End If
                
                ' 날짜 범위 내에 있는지 확인
                If mailSentDate >= startDate And mailSentDate <= endDate Then
                    totalCount = totalCount + 1
                    
                    ' 중복 체크
                    If IsMailAlreadySaved(mail) Then
                        skippedCount = skippedCount + 1
                    Else
                        SaveMailAsMSG mail, "Sent"
                        savedCount = savedCount + 1
                    End If
                End If
                
                Set mail = Nothing
                DoEvents
            End If
        Next Item
    End If
    
    ' 결과 로그 기록 및 사용자에게 결과 표시
    If totalCount > 0 Then
        On Error Resume Next
        Dim fso As Object
        Dim logFile As Object
        Dim logPath As String
        Dim logEntry As String
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        logPath = GetLogFilePath("success")
        
        Set logFile = fso.OpenTextFile(logPath, 8, True)
        logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
                   "MANUAL_BACKUP (" & periodLabel & ")" & " | " & _
                   "총 " & totalCount & "개 중 " & savedCount & "개 저장, " & skippedCount & "개 건너뜀"
        
        logFile.WriteLine logEntry
        logFile.Close
        
        Set logFile = Nothing
        Set fso = Nothing
        
        ' 사용자에게 결과 표시
        Dim msgText As String
        msgText = periodLabel & " 메일 백업이 완료되었습니다." & vbCrLf & vbCrLf
        msgText = msgText & "총 " & totalCount & "개의 메일 중:" & vbCrLf
        msgText = msgText & "- 저장: " & savedCount & "개" & vbCrLf
        If skippedCount > 0 Then
            msgText = msgText & "- 건너뜀 (이미 저장됨): " & skippedCount & "개" & vbCrLf
        End If
        msgText = msgText & vbCrLf & "경로: " & BACKUP_BASE_PATH
        
        MsgBox msgText, vbInformation, periodLabel & " 메일 백업"
    Else
        MsgBox "백업할 메일이 없습니다." & vbCrLf & _
               "기간: " & Format(startDate, "yyyy-mm-dd") & " ~ " & Format(endDate, "yyyy-mm-dd"), _
               vbInformation, periodLabel & " 메일 백업"
    End If
    
    ' 메모리 정리
    Set inboxItems = Nothing
    Set sentItems = Nothing
    Set inboxFolder = Nothing
    Set sentFolder = Nothing
    Set ns = Nothing
    
    Exit Sub
ErrorHandler:
    ' 에러 발생 시에도 계속 진행 (로그만 기록)
    On Error Resume Next
    Dim fsoErr As Object
    Dim logFileErr As Object
    Dim logPathErr As String
    
    Set fsoErr = CreateObject("Scripting.FileSystemObject")
    logPathErr = GetLogFilePath("error")
    
    Set logFileErr = fsoErr.OpenTextFile(logPathErr, 8, True)
    logFileErr.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & errorSubName & " 에러: " & Err.Description
    logFileErr.Close
    
    Set logFileErr = Nothing
    Set fsoErr = Nothing
End Sub

'===========================================
' 수동 실행: 오늘 날짜 기준 이전 한 달치 메일 백업
' 사용법: Alt+F8 → BackupLastMonthMails 실행
'===========================================
Public Sub BackupLastMonthMails()
    Dim endDate As Date
    Dim startDate As Date
    endDate = Now
    startDate = DateAdd("m", -1, endDate)
    BackupMailsByDateRange startDate, endDate, "한 달치", "BackupLastMonthMails"
End Sub

'===========================================
' 수동 실행: 오늘 날짜 기준 이전 1년치 메일 백업
' 사용법: Alt+F8 → BackupLastYearMails 실행
'===========================================
Public Sub BackupLastYearMails()
    Dim endDate As Date
    Dim startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -1, endDate)
    BackupMailsByDateRange startDate, endDate, "1년치", "BackupLastYearMails"
End Sub

'===========================================
' 수동 실행: 오늘 날짜 기준 이전 2년치 메일 백업
' 사용법: Alt+F8 → BackupLast2YearsMails 실행
'===========================================
Public Sub BackupLast2YearsMails()
    Dim endDate As Date
    Dim startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -2, endDate)
    BackupMailsByDateRange startDate, endDate, "2년치", "BackupLast2YearsMails"
End Sub

'===========================================
' 수동 실행: 오늘 날짜 기준 이전 3년치 메일 백업
' 사용법: Alt+F8 → BackupLast3YearsMails 실행
'===========================================
Public Sub BackupLast3YearsMails()
    Dim endDate As Date
    Dim startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -3, endDate)
    BackupMailsByDateRange startDate, endDate, "3년치", "BackupLast3YearsMails"
End Sub

'===========================================
' 내부: PST 파일 기간별 백업 공통 로직 (대용량 PST 대응)
'===========================================
Private Sub BackupPSTByDateRange(ByVal startDate As Date, ByVal endDate As Date, ByVal periodLabel As String, ByVal errorSubName As String)
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Dim stores As Outlook.Stores
    Dim oStore As Outlook.Store
    Dim inboxFolder As Outlook.MAPIFolder
    Dim sentFolder As Outlook.MAPIFolder
    Dim folderItems As Outlook.Items
    Dim filterString As String
    Dim Item As Object
    Dim mail As Outlook.MailItem
    Dim savedCount As Long
    Dim skippedCount As Long
    Dim totalCount As Long
    Dim pstCount As Long
    Dim storePath As String
    Dim mailDate As Date
    Dim mailSentDate As Date
    
    Set ns = Application.GetNamespace("MAPI")
    Set stores = ns.Stores
    
    savedCount = 0
    skippedCount = 0
    totalCount = 0
    pstCount = 0
    
    For Each oStore In stores
        On Error Resume Next
        storePath = oStore.FilePath
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler
        
        If Len(storePath) = 0 Or InStr(1, UCase(storePath), ".PST") = 0 Then GoTo NextStore
        
        pstCount = pstCount + 1
        
        ' PST 받은편지함 (날짜 필터 적용)
        On Error Resume Next
        Set inboxFolder = oStore.GetDefaultFolder(olFolderInbox)
        If Err.Number <> 0 Then Set inboxFolder = Nothing: Err.Clear
        On Error GoTo ErrorHandler
        
        If Not inboxFolder Is Nothing Then
            ' PST는 Restrict 날짜 필터가 동작하지 않을 수 있으므로 전체 항목 순회 후 VBA에서 날짜 비교
            Set folderItems = inboxFolder.Items
            On Error Resume Next
            folderItems.Sort "[ReceivedTime]", True
            On Error GoTo ErrorHandler
            
            For Each Item In folderItems
                If TypeOf Item Is Outlook.MailItem Then
                    Set mail = Item
                    If IsDate(mail.ReceivedTime) And mail.ReceivedTime > #1/1/1900# Then
                        mailDate = mail.ReceivedTime
                    ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
                        mailDate = mail.CreationTime
                    Else
                        mailDate = #1/1/1900#
                    End If
                    If mailDate >= startDate And mailDate <= endDate Then
                        totalCount = totalCount + 1
                        If IsMailAlreadySaved(mail) Then
                            skippedCount = skippedCount + 1
                        Else
                            SaveMailAsMSG mail, "Inbox"
                            savedCount = savedCount + 1
                        End If
                    End If
                    Set mail = Nothing
                    DoEvents
                End If
            Next Item
            Set folderItems = Nothing
            Set inboxFolder = Nothing
        End If
        
        ' PST 보낸편지함 (날짜 필터 적용)
        On Error Resume Next
        Set sentFolder = oStore.GetDefaultFolder(olFolderSentMail)
        If Err.Number <> 0 Then Set sentFolder = Nothing: Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sentFolder Is Nothing Then
            ' PST는 Restrict 날짜 필터가 동작하지 않을 수 있으므로 전체 항목 순회 후 VBA에서 날짜 비교
            Set folderItems = sentFolder.Items
            On Error Resume Next
            folderItems.Sort "[SentOn]", True
            On Error GoTo ErrorHandler
            
            For Each Item In folderItems
                If TypeOf Item Is Outlook.MailItem Then
                    Set mail = Item
                    If IsDate(mail.SentOn) And mail.SentOn > #1/1/1900# Then
                        mailSentDate = mail.SentOn
                    ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
                        mailSentDate = mail.CreationTime
                    Else
                        mailSentDate = #1/1/1900#
                    End If
                    If mailSentDate >= startDate And mailSentDate <= endDate Then
                        totalCount = totalCount + 1
                        If IsMailAlreadySaved(mail) Then
                            skippedCount = skippedCount + 1
                        Else
                            SaveMailAsMSG mail, "Sent"
                            savedCount = savedCount + 1
                        End If
                    End If
                    Set mail = Nothing
                    DoEvents
                End If
            Next Item
            Set folderItems = Nothing
            Set sentFolder = Nothing
        End If
        
NextStore:
    Next oStore
    
    ' 결과 로그 및 메시지
    If totalCount > 0 Then
        On Error Resume Next
        Dim fso As Object
        Dim logFile As Object
        Dim logPath As String
        Dim logEntry As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        logPath = GetLogFilePath("success")
        Set logFile = fso.OpenTextFile(logPath, 8, True)
        logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
                   "MANUAL_BACKUP (PST " & periodLabel & ")" & " | " & _
                   "PST " & pstCount & "개, 총 " & totalCount & "개 중 " & savedCount & "개 저장, " & skippedCount & "개 건너뜀"
        logFile.WriteLine logEntry
        logFile.Close
        Set logFile = Nothing
        Set fso = Nothing
    End If
    
    Dim msgText As String
    If pstCount = 0 Then
        msgText = "연결된 PST 파일이 없습니다." & vbCrLf & "Outlook에서 PST를 추가한 뒤 다시 실행하세요."
    ElseIf totalCount = 0 Then
        msgText = "해당 기간(" & periodLabel & ")에 백업할 메일이 없습니다." & vbCrLf & "기간: " & Format(startDate, "yyyy-mm-dd") & " ~ " & Format(endDate, "yyyy-mm-dd") & vbCrLf & "PST 파일 수: " & pstCount & "개"
    Else
        msgText = "PST " & periodLabel & " 메일 백업이 완료되었습니다." & vbCrLf & vbCrLf
        msgText = msgText & "PST 파일 수: " & pstCount & "개" & vbCrLf
        msgText = msgText & "총 " & totalCount & "개의 메일 중:" & vbCrLf
        msgText = msgText & "- 저장: " & savedCount & "개" & vbCrLf
        If skippedCount > 0 Then msgText = msgText & "- 건너뜀 (이미 저장됨): " & skippedCount & "개" & vbCrLf
        msgText = msgText & vbCrLf & "경로: " & BACKUP_BASE_PATH
    End If
    MsgBox msgText, vbInformation, "PST " & periodLabel & " 메일 백업"
    
    Set stores = Nothing
    Set ns = Nothing
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Dim fsoErr As Object
    Dim logFileErr As Object
    Dim logPathErr As String
    Set fsoErr = CreateObject("Scripting.FileSystemObject")
    logPathErr = GetLogFilePath("error")
    Set logFileErr = fsoErr.OpenTextFile(logPathErr, 8, True)
    logFileErr.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & errorSubName & " 에러: " & Err.Description
    logFileErr.Close
    Set logFileErr = Nothing
    Set fsoErr = Nothing
    MsgBox "오류 발생: " & Err.Description, vbCritical, "PST 메일 백업"
End Sub

'===========================================
' 수동 실행: PST 오늘 기준 이전 1년치 백업
'===========================================
Public Sub BackupPSTLastYearMails()
    Dim endDate As Date, startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -1, endDate)
    BackupPSTByDateRange startDate, endDate, "1년치", "BackupPSTLastYearMails"
End Sub

'===========================================
' 수동 실행: PST 오늘 기준 이전 2년치 백업
'===========================================
Public Sub BackupPSTLast2YearsMails()
    Dim endDate As Date, startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -2, endDate)
    BackupPSTByDateRange startDate, endDate, "2년치", "BackupPSTLast2YearsMails"
End Sub

'===========================================
' 수동 실행: PST 오늘 기준 이전 3년치 백업
'===========================================
Public Sub BackupPSTLast3YearsMails()
    Dim endDate As Date, startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -3, endDate)
    BackupPSTByDateRange startDate, endDate, "3년치", "BackupPSTLast3YearsMails"
End Sub

'===========================================
' 수동 실행: PST 오늘 기준 이전 4년치 백업
'===========================================
Public Sub BackupPSTLast4YearsMails()
    Dim endDate As Date, startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -4, endDate)
    BackupPSTByDateRange startDate, endDate, "4년치", "BackupPSTLast4YearsMails"
End Sub

'===========================================
' 수동 실행: PST 오늘 기준 이전 5년치 백업
'===========================================
Public Sub BackupPSTLast5YearsMails()
    Dim endDate As Date, startDate As Date
    endDate = Now
    startDate = DateAdd("yyyy", -5, endDate)
    BackupPSTByDateRange startDate, endDate, "5년치", "BackupPSTLast5YearsMails"
End Sub

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
    
    Dim skippedCount As Long
    skippedCount = 0
    
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

            ' 중복 체크
            If IsMailAlreadySaved(mail) Then
                skippedCount = skippedCount + 1
            Else
                SaveMailAsMSG mail, folderType
                savedCount = savedCount + 1
            End If
            
            Set mail = Nothing
            DoEvents
        End If
    Next Item

    Dim msgText As String
    msgText = savedCount & "개의 메일이 저장되었습니다."
    If skippedCount > 0 Then
        msgText = msgText & vbCrLf & skippedCount & "개의 메일은 이미 저장되어 있어 건너뛰었습니다."
    End If
    msgText = msgText & vbCrLf & "경로: " & BACKUP_BASE_PATH
    
    MsgBox msgText, vbInformation

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
