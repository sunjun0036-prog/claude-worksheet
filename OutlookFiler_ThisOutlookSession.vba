'====================================================================
' OutlookFiler (1/2) - ThisOutlookSession  [이벤트 + 자동 분류]
'--------------------------------------------------------------------
' 이 코드는 VBA 편집기의
'   [Project1] > [Microsoft Outlook 개체] > ThisOutlookSession  에 넣습니다.
' 새 프로젝트 폴더 만들기 버튼(NewProject)은 '표준 모듈' 쪽
'   ( OutlookFiler_Module.vba ) 에 따로 넣습니다.
'
' 동작:
'   - 받은/보낸 메일의 본문·제목에서 프로젝트 코드를 찾아
'     SJLEE\Job\<연도>\<코드_회사명> 폴더로 '복사'
'   - 폴더 목록을 '실시간으로 읽어' 매핑하므로, 새 폴더만 만들면
'     코드 수정/재시작 없이 바로 인식됨 (모든 연도 폴더 검색)
'   - 폴더명에 없는 추가 키워드는 AliasList 로 보완 (예: HMT26-063 -> 260199)
'   - 모든 메일을 SJLEE\받은 편지함 / SJLEE\보낸 편지함으로 전체 백업
'     (웹에서 미리 읽어 Outlook 규칙이 동작하지 않는 경우도 포함)
'   - [수주통보] 메일은 수주통보 폴더에만, 백업·Remind 없음
'   - 받는사람에 내가 있으면 SJLEE\Remind 에도 복사 (WBlock·수주통보 제외)
'   - Outlook이 꺼져 있던 동안 쌓인 메일은 다음 실행 때 한 번에 처리
'
' 중복 방지:
'   - 처리한 메일에는 'FiledByMacro' 표식을 남겨 두 번 복사하지 않음
'   - 최초 설치 시점 이전의 기존 메일은 건드리지 않음(기준 시각만 기록)
'
' ※ 설정 상수(STORE_ROOT, JOB_FOLDER 등)와 폴더 헬퍼는 표준 모듈에 있음.
'====================================================================
Option Explicit

Private WithEvents mInboxItems As Outlook.Items
Private WithEvents mSentItems As Outlook.Items

'======================== 이벤트 진입점 ========================

Private Sub Application_Startup()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Set mInboxItems = ns.GetDefaultFolder(olFolderInbox).Items
    Set mSentItems = ns.GetDefaultFolder(olFolderSentMail).Items
    RefreshCache            ' 키워드->폴더 캐시 1회 생성
    CatchUp ns
End Sub

Private Sub Application_Quit()
    On Error Resume Next
    SaveSetting "OutlookFiler", "State", "LastRun", CStr(CDbl(Now))
End Sub

Private Sub mInboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    FileItem Item, False
End Sub

Private Sub mSentItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    FileItem Item, True
End Sub

'======================== 핵심 분류 로직 ========================

Private Sub FileItem(ByVal Item As Object, ByVal IsSent As Boolean)
    On Error Resume Next
    If Item Is Nothing Then Exit Sub
    If TypeName(Item) <> "MailItem" Then Exit Sub
    If IsStamped(Item) Then Exit Sub          ' 이미 처리한 메일

    ' 수주통보: 수주통보 폴더로 이동(복사 아님), 백업·Remind 건너뜀
    If InStr(LCase(Item.Subject), LCase(JUSU_KEYWORD)) > 0 Then
        Dim jtbFld As Outlook.Folder
        Set jtbFld = GetSub(GetRoot(), JUSU_FOLDER)
        If Not jtbFld Is Nothing Then DoMove Item, jtbFld
        Exit Sub
    End If

    Dim haystack As String
    haystack = LCase(Item.Subject & " " & Item.Body)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    ' 1) 캐시(키워드->폴더)를 순회하며 본문/제목에 키워드가 있으면 복사
    Dim map As Object
    Set map = ProjectMap()
    Dim key As Variant
    For Each key In map.Keys                   ' key 는 이미 소문자
        If InStr(haystack, key) > 0 Then
            Dim f As Outlook.Folder
            Set f = map(key)
            If Not f Is Nothing Then
                If Not seen.Exists(f.FolderPath) Then
                    seen.Add f.FolderPath, True
                    DoCopy Item, f
                End If
            End If
        End If
    Next key

    ' 2) 전체 백업: 키워드 매칭 여부와 무관하게 모든 메일 백업
    Dim bkName As String
    If IsSent Then bkName = BACKUP_SENT Else bkName = BACKUP_INBOX
    Dim bkFld As Outlook.Folder
    Set bkFld = GetSub(GetRoot(), bkName)
    If Not bkFld Is Nothing Then
        If Not seen.Exists(bkFld.FolderPath) Then
            seen.Add bkFld.FolderPath, True
            DoCopy Item, bkFld
        End If
    End If

    ' 3) Remind: 받은 메일, 받는사람에 내가 있고, WBlock 제외
    '    (수주통보는 위에서 이미 Exit Sub 했으므로 여기까지 오지 않음)
    If Not IsSent Then
        If InStr(LCase(Item.Subject), "wblock") = 0 Then
            If IsAddressedToMe(Item) Then
                Dim rmFld As Outlook.Folder
                Set rmFld = GetSub(GetRoot(), REMIND_FOLDER)
                If Not rmFld Is Nothing Then
                    If Not seen.Exists(rmFld.FolderPath) Then
                        seen.Add rmFld.FolderPath, True
                        DoCopy Item, rmFld
                    End If
                End If
            End If
        End If
    End If

    ' 표식은 DoCopy 안에서 (복사 전에) 원본에 찍힌다.
End Sub

'======================== 폴더 이동 / 복사 ========================

' 수주통보 전용: 원본을 목적지로 이동
Private Sub DoMove(ByVal Item As Object, ByVal target As Outlook.Folder)
    On Error Resume Next
    Stamp Item
    Item.Move target
End Sub

' 원본에 '먼저' 표식을 찍은 뒤 복사한다. 사본은 표식을 상속한 채 태어나므로,
' 감시 폴더에 잠깐 생기는 사본의 ItemAdd 이벤트가 어떤 타이밍에도 재처리되지 않는다.
' (사본 생성 후 표식을 찍으면 그 사이 틈에 이벤트가 끼어들어 중복 복사됨 -> 경쟁 상태)
Private Sub DoCopy(ByVal Item As Object, ByVal target As Outlook.Folder)
    On Error Resume Next
    Stamp Item
    Dim cpy As Object
    Set cpy = Item.Copy
    cpy.Move target
End Sub

'======================== 보조 함수 ========================

Private Function IsAddressedToMe(ByVal Item As Object) As Boolean
    On Error Resume Next
    Dim myAddr As String, myName As String
    myAddr = LCase(Application.Session.CurrentUser.Address)
    myName = LCase(Application.Session.CurrentUser.Name)
    Dim r As Outlook.Recipient
    For Each r In Item.Recipients
        If r.Type = olTo Then
            If LCase(r.Address) = myAddr Or LCase(r.Name) = myName Then
                IsAddressedToMe = True
                Exit Function
            End If
        End If
    Next r
End Function

'======================== 중복 방지 표식 ========================

Private Sub Stamp(ByVal Item As Object)
    On Error Resume Next
    Dim p As Outlook.UserProperty
    Set p = Item.UserProperties.Find(STAMP_NAME)
    If p Is Nothing Then Set p = Item.UserProperties.Add(STAMP_NAME, olText)
    p.Value = "1"
    Item.Save
End Sub

Private Function IsStamped(ByVal Item As Object) As Boolean
    On Error Resume Next
    Dim p As Outlook.UserProperty
    Set p = Item.UserProperties.Find(STAMP_NAME)
    IsStamped = (Not p Is Nothing)
End Function

'======================== 꺼져 있던 동안 쌓인 메일 처리 ========================

Private Sub CatchUp(ByVal ns As Outlook.NameSpace)
    On Error Resume Next
    Dim lastRun As String
    lastRun = GetSetting("OutlookFiler", "State", "LastRun", "")

    If lastRun = "" Then
        ' 최초 실행: 기존 메일은 건드리지 않고 기준 시각만 기록
        SaveSetting "OutlookFiler", "State", "LastRun", CStr(CDbl(Now))
        Exit Sub
    End If

    Dim since As Date
    since = CDate(CDbl(lastRun))
    ProcessSince ns.GetDefaultFolder(olFolderInbox), False, since
    ProcessSince ns.GetDefaultFolder(olFolderSentMail), True, since
    SaveSetting "OutlookFiler", "State", "LastRun", CStr(CDbl(Now))
End Sub

Private Sub ProcessSince(ByVal folder As Outlook.Folder, ByVal IsSent As Boolean, ByVal since As Date)
    On Error Resume Next
    ' 컬렉션으로 먼저 모은 뒤 처리 (반복 중 폴더 변경 방지)
    Dim col As New Collection
    Dim itm As Object
    For Each itm In folder.Items
        If TypeName(itm) = "MailItem" Then
            Dim t As Date
            If IsSent Then t = itm.SentOn Else t = itm.ReceivedTime
            If t > since Then col.Add itm
        End If
    Next itm
    For Each itm In col
        FileItem itm, IsSent
    Next itm
End Sub
