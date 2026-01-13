# AGENT.md - Outlook Backup Project Overview

## Project Purpose
Microsoft Outlook 자동 메일 백업 시스템. VBA 매크로를 사용하여 수신/발신 메일을 MSG 파일로 자동 백업하는 솔루션.

## Project Structure
```
D:\Outlook_Backup/
├── script_msg.vba          # 메인 백업 매크로 (VBA)
├── README.md               # 사용자 설치/사용 가이드
├── AGENT.md                # AI 에이전트용 프로젝트 문서
├── .gitignore              # Git 무시 파일
├── Inbox/                  # 받은 메일 백업 (자동 생성)
│   └── [YYYY]/
│       └── [MM]/
│           └── *.msg
├── Sent/                   # 보낸 메일 백업 (자동 생성)
│   └── [YYYY]/
│       └── [MM]/
│           └── *.msg
└── logs/                   # 로그 디렉토리 (자동 생성)
    ├── yyyy-mm_success.log # 성공 로그
    └── yyyy-mm_error.log   # 에러 로그
```

## Core Functionality

### 1. **자동 백업 시스템**
- **Inbox 트리거**: 새 메일 수신 시 자동 실행 (`InboxItems_ItemAdd`)
- **Sent 트리거**: 메일 발송 시 자동 실행 (`SentItems_ItemAdd`)
- **이벤트 핸들러**: Outlook 시작 시 자동 초기화 (`Application_Startup`)
- **수동 초기화**: `InitializeEventHandler` 매크로 실행 (Alt+F8)

### 2. **파일 저장 구조**
```
파일명 형식: yyyymmdd_hhnnss_PersonName_Subject.msg
저장 경로 (Inbox): D:\Outlook_Backup\Inbox\YYYY\MM\
저장 경로 (Sent): D:\Outlook_Backup\Sent\YYYY\MM\
```

**파일명 생성 로직**:
- 날짜시간 (15자, 필수): `yyyymmdd_hhnnss`
- 발신자/수신자 (최대 50자):
  - **Inbox**: `CleanFileName(SenderName)` - 발신자 이름
  - **Sent**: `CleanFileName(Recipients.Item(1).Name)` - 첫 번째 수신자 이름
- 제목 (가변, 경로 제한 내): `CleanFileName(Subject)`
- Windows 경로 제한: 260자 준수
- 특수문자 자동 치환: `/\:*?"<>|` → `_`
- 연속 공백/언더스코어 방지

### 3. **핵심 함수**

#### `SaveMailAsMSG(mail As MailItem, folderType As String)`
메일을 MSG 파일로 저장하는 핵심 함수
- **파라미터**:
  - `mail`: 저장할 메일 객체
  - `folderType`: "Inbox" 또는 "Sent"
- 날짜 기반 폴더 구조 자동 생성
- folderType에 따른 발신자/수신자 자동 구분
- 파일명 길이 제한 및 최적화
- 저장 검증 (100바이트 미만 파일 경고)
- 성공 시 `LogSuccess` 호출, 실패 시 `LogError` 호출

#### `LogSuccess(mail As MailItem, filePath As String, fileSize As Long)`
정상 저장 로그 기록 (신규 추가)
- 로그 형식: `[yyyy-mm-dd hh:nn:ss] | SUCCESS | 파일경로 | 파일크기 | 발신자 | 제목`
- 로그 경로: `D:\Outlook_Backup\logs\yyyy-mm_success.log`

#### `LogError(mail As MailItem, errMsg As String)`
에러 로그 기록 (월별 로테이션)
- 로그 형식: `[yyyy-mm-dd hh:nn:ss] | 에러내용 | 발신자 | 제목`
- 로그 경로: `D:\Outlook_Backup\logs\yyyy-mm_error.log`

#### `CleanFileName(strFileName As String)`
파일명에서 특수문자 제거 및 정규화
- 연속 공백/언더스코어 방지
- 줄바꿈 문자 제거
- 탭 → 공백 변환

#### `CreateFolderPath(folderPath As String)`
재귀적 폴더 생성 (부모 폴더부터 순차 생성)
- FSO 객체 재사용 최적화
- Inbox, Sent, logs 폴더 자동 생성

#### `InitializeEventHandler()`
이벤트 핸들러 초기화
- Inbox 폴더 이벤트 등록
- Sent 폴더 이벤트 등록
- 활성화 확인 메시지 표시

#### `InboxItems_ItemAdd(Item As Object)`
Inbox 메일 수신 이벤트 핸들러
- `SaveMailAsMSG(mail, "Inbox")` 호출

#### `SentItems_ItemAdd(Item As Object)`
Sent 메일 발송 이벤트 핸들러 (신규 추가)
- `SaveMailAsMSG(mail, "Sent")` 호출

#### `SaveSelectedMailsAsMSG()`
수동 백업: Outlook에서 선택한 메일들을 일괄 저장
- 50개 초과 시 확인 메시지
- 메일이 속한 폴더 자동 감지 (Inbox/Sent 구분)
- 폴더명에 "보낸" 또는 "Sent" 포함 시 → "Sent"
- 그 외 → "Inbox"
- 진행 상황 표시 (`DoEvents`)

#### `GetLogFilePath(Optional logType As String = "error")`
로그 파일 경로 생성 (월별 로테이션)
- `logType`: "success" 또는 "error"
- logs 폴더 자동 생성

## Technical Details

### Constants
- `BACKUP_BASE_PATH`: `D:\Outlook_Backup\`
- 경로 최대 길이: 260자 (Windows 제한)
- 파일명 여유 공간: 5자

### Global Variables
- `InboxItems As Outlook.Items`: Inbox 이벤트 핸들러
- `SentItems As Outlook.Items`: Sent 이벤트 핸들러

### Error Handling
- 모든 주요 함수에 `On Error GoTo ErrorHandler` 구현
- `ReceivedTime` 예외 처리 (임시저장/초안 메일 대응)
- 파일 저장 실패 시 자동 로깅
- 파일 크기 검증 (100바이트 미만 경고)

### Date Handling Priority
1. `mail.ReceivedTime` (수신 시각)
2. `mail.CreationTime` (생성 시각, fallback)
3. `Now` (현재 시각, last resort)

### Encoding Considerations
- VBA 에디터는 시스템 인코딩(CP949) 사용
- MsgBox 등 사용자 메시지는 영문 권장
- 파일로 저장된 VBA 코드를 복사할 때 한글 주석 깨질 수 있음
- VBA 에디터에서 직접 수정 시 인코딩 문제 없음

## Development Context

### Language
- **Primary**: VBA (Visual Basic for Applications)
- **Target**: Microsoft Outlook (Windows)

### Dependencies
- `Outlook.Application`
- `Scripting.FileSystemObject`
- Outlook Object Model (MailItem, NameSpace, Items, Recipients)

### Installation
1. Outlook 실행 → Alt+F11 (VBA 편집기)
2. `ThisOutlookSession` 모듈에 `script_msg.vba` 코드 붙여넣기
3. 매크로 보안 설정: "알림을 표시하는 모든 매크로 사용"
4. BACKUP_BASE_PATH 경로 확인/수정 (필요 시)
5. Outlook 재시작 또는 `InitializeEventHandler` 수동 실행

### Usage
**자동 백업**:
- Inbox: 메일 수신 시 자동 실행
- Sent: 메일 발송 시 자동 실행
- 이벤트 핸들러 활성화 필요

**수동 백업**:
- 메일 선택 → Alt+F8 → `SaveSelectedMailsAsMSG` 실행
- 폴더 자동 감지 (Inbox/Sent 구분)

## Key Design Decisions

1. **MSG 형식 선택**: Outlook 네이티브 형식으로 첨부파일, 메타데이터 완전 보존
2. **폴더별 분리 구조**: Inbox/Sent 분리로 명확한 구분 및 관리 용이
3. **연도/월 폴더 구조**: 파일 관리 용이성 및 탐색 성능
4. **파일명 우선순위**: 날짜시간(필수) → 발신자/수신자(최대 50자) → 제목(가변)
5. **경로 길이 제한 준수**: Windows 260자 제한 대응
6. **FSO 재사용 최적화**: 재귀 호출 시 객체 전달로 성능 개선
7. **월별 로그 로테이션**: 로그 파일 크기 관리 (success/error 분리)
8. **자동 폴더 생성**: 사용자 개입 없이 필요한 모든 폴더 자동 생성

## Known Issues
1. **VBA 인코딩**: 파일 복사 시 한글 깨질 수 있음 → VBA 에디터에서 직접 수정 권장
2. **MsgBox 인코딩**: 영문 메시지 사용으로 해결

## Future Enhancement Ideas (Not Planned)
- 다중 계정 지원
- 첨부파일 별도 추출 옵션
- 데이터베이스 인덱싱
- 압축 백업 옵션
- 클라우드 동기화

### Code Style
- 명시적 주석 블록: `'===========================================`
- 함수 설명 주석 포함
- Error handling 패턴 일관성 유지
- 한글 주석 사용 (VBA 에디터 내에서)
- 영문 메시지 사용 (MsgBox 등)

### AI Agent Instructions
- VBA 문법 준수 (Option Explicit, 타입 선언)
- Outlook Object Model 이해 필요
- Windows 파일 시스템 제약 고려
- 기존 코드 스타일 및 패턴 유지
- README.md와 AGENT.md 동기화 유지
- 인코딩 문제 고려 (VBA 에디터 직접 수정 권장)

## Testing Checklist

### Installation Testing
- [ ] VBA 코드 복사/붙여넣기
- [ ] 매크로 보안 설정 확인
- [ ] InitializeEventHandler 실행 확인

### Functional Testing
- [ ] 메일 수신 시 Inbox 폴더 자동 백업 확인
- [ ] 메일 발송 시 Sent 폴더 자동 백업 확인
- [ ] 수동 백업 (Inbox 메일 선택)
- [ ] 수동 백업 (Sent 메일 선택)
- [ ] 파일명 특수문자 처리 확인
- [ ] 긴 제목 파일명 잘림 확인

### Log Testing
- [ ] 성공 로그 기록 확인 (logs\yyyy-mm_success.log)
- [ ] 에러 로그 기록 확인 (logs\yyyy-mm_error.log)
- [ ] 월별 로그 파일 로테이션 확인

### Edge Case Testing
- [ ] 제목 없는 메일 (NoSubject 처리)
- [ ] 수신자 없는 메일 (NoRecipient 처리)
- [ ] 임시저장 메일 (날짜 처리)
- [ ] 50개 이상 일괄 백업 (확인 메시지)
- [ ] 경로 길이 260자 근접 (자동 조정)

## File Manifest

| File | Purpose | Status |
|------|---------|--------|
| `script_msg.vba` | 메인 VBA 매크로 코드 | Production Ready |
| `README.md` | 사용자 가이드 | Production Ready |
| `AGENT.md` | AI 에이전트용 문서 | Production Ready |
| `.gitignore` | Git 무시 파일 | Production Ready |
| `Inbox/` | 받은 메일 백업 | Auto-generated |
| `Sent/` | 보낸 메일 백업 | Auto-generated |
| `logs/` | 로그 파일 | Auto-generated |
