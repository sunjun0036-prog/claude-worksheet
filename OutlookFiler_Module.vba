'====================================================================
' OutlookFiler (2/2) - 표준 모듈  [설정 + 캐시 + 폴더 헬퍼 + 새 프로젝트 버튼]
'--------------------------------------------------------------------
' 넣는 위치:
'   VBA 편집기 > 메뉴 [삽입(I)] > [모듈(M)] 으로 새 표준 모듈을 만들고,
'   이 코드 전체를 붙여넣습니다.  (ThisOutlookSession 아님!)
'
' 캐시: 키워드->폴더 매핑을 한 번만 만들어 기억해두고 재사용한다.
'       폴더가 수백 개로 늘어도 메일당 비용이 일정. (메일마다 폴더 재탐색 X)
'       - Outlook 시작 시 1회 생성 (ThisOutlookSession의 Application_Startup)
'       - NewProject로 폴더를 만들면 자동 갱신(RefreshCache)
'       - 버튼 안 쓰고 폴더를 손으로 만들었다면, 다음 재시작 때 반영됨
'====================================================================
Option Explicit

' ===== 설정 (필요 시 여기만 수정) =====
Public Const STORE_ROOT As String = "SJLEE"        ' 목적지 폴더가 있는 저장소 이름
Public Const JOB_FOLDER As String = "Job"          ' 프로젝트 상위 폴더 (그 아래 연도 폴더들)
Public Const STAMP_NAME As String = "FiledByMacro" ' 중복 방지 표식 이름

' ===== 캐시 (모듈 수준에 보관) =====
Private mCache As Object        ' 소문자 키워드 -> Outlook.Folder
Private mBuilt As Boolean

' 폴더명에 없는 추가 키워드: [키워드, 대상폴더코드] 쌍
Public Function AliasList() As Variant
    AliasList = Array( _
        "HMT26-063", "260199")
End Function

'==================== 캐시 공개 함수 ====================

' 키워드->폴더 매핑(Dictionary)을 돌려준다. 아직 없으면 만든다.
Public Function ProjectMap() As Object
    If Not mBuilt Then BuildCache
    Set ProjectMap = mCache
End Function

' 캐시를 비우고 다시 만든다 (폴더가 새로 생겼을 때 호출).
Public Sub RefreshCache()
    mBuilt = False
    BuildCache
End Sub

Private Sub BuildCache()
    On Error Resume Next
    Set mCache = CreateObject("Scripting.Dictionary")

    ' 1) Job 아래 모든 연도 폴더 -> 프로젝트 폴더 스캔
    Dim job As Outlook.Folder
    Set job = GetJobFolder()
    If Not job Is Nothing Then
        Dim yf As Outlook.Folder, pf As Outlook.Folder
        For Each yf In job.Folders
            For Each pf In yf.Folders
                Dim code As String
                code = LCase(FolderCode(pf.Name))   ' '_' 앞부분 = 프로젝트 코드
                If Len(code) > 0 Then
                    If Not mCache.Exists(code) Then mCache.Add code, pf
                End If
            Next pf
        Next yf
    End If

    ' 2) 별칭: 이미 캐시된 대상 폴더를 추가 키워드로도 가리키게 함
    Dim ali As Variant, k As Long
    ali = AliasList()
    For k = 0 To UBound(ali) Step 2
        Dim aKey As String, tCode As String
        aKey = LCase(CStr(ali(k)))
        tCode = LCase(CStr(ali(k + 1)))
        If mCache.Exists(tCode) And Not mCache.Exists(aKey) Then
            mCache.Add aKey, mCache(tCode)
        End If
    Next k

    mBuilt = True
End Sub

'==================== 새 프로젝트 폴더 만들기 ====================
Public Sub NewProject()
    Dim code As String, company As String
    code = Trim(InputBox("프로젝트 번호를 입력하세요 (예: 260500)", "새 프로젝트 폴더"))
    If code = "" Then Exit Sub
    company = Trim(InputBox("회사명을 입력하세요", "새 프로젝트: " & code))
    If company = "" Then Exit Sub

    Dim job As Outlook.Folder
    Set job = GetJobFolder()
    If job Is Nothing Then
        MsgBox "'" & JOB_FOLDER & "' 폴더를 찾을 수 없습니다 (저장소: " & STORE_ROOT & ").", vbExclamation
        Exit Sub
    End If

    ' 올해 연도 폴더 (없으면 생성)
    Dim yearStr As String
    yearStr = CStr(Year(Date))
    Dim yf As Outlook.Folder
    Set yf = GetOrCreateSub(job, yearStr)
    If yf Is Nothing Then
        MsgBox "연도 폴더(" & yearStr & ") 생성에 실패했습니다.", vbExclamation
        Exit Sub
    End If

    ' 프로젝트 폴더
    Dim folderName As String
    folderName = code & "_" & company

    Dim pf As Outlook.Folder
    Set pf = GetSub(yf, folderName)
    If Not pf Is Nothing Then
        MsgBox "이미 있는 폴더입니다:" & vbCrLf & JOB_FOLDER & "\" & yearStr & "\" & folderName, vbInformation
        Exit Sub
    End If

    On Error Resume Next
    Set pf = yf.Folders.Add(folderName)
    On Error GoTo 0
    If pf Is Nothing Then
        MsgBox "폴더 생성에 실패했습니다: " & folderName, vbExclamation
    Else
        RefreshCache    ' 새 폴더를 즉시 인식하도록 캐시 갱신
        MsgBox "생성 완료:" & vbCrLf & JOB_FOLDER & "\" & yearStr & "\" & folderName & vbCrLf & vbCrLf & _
               "이제부터 이 번호가 들어간 메일은 자동으로 이 폴더에 복사됩니다.", vbInformation
    End If
End Sub

'==================== 공용 폴더 헬퍼 ====================

' 목적지 저장소(SJLEE) 루트
Public Function GetRoot() As Outlook.Folder
    On Error Resume Next
    Set GetRoot = Application.GetNamespace("MAPI").Folders(STORE_ROOT)
End Function

' SJLEE\Job
Public Function GetJobFolder() As Outlook.Folder
    On Error Resume Next
    Dim r As Outlook.Folder
    Set r = GetRoot()
    If r Is Nothing Then Exit Function
    Set GetJobFolder = r.Folders(JOB_FOLDER)
End Function

' 하위 폴더 가져오기 (없으면 Nothing)
Public Function GetSub(ByVal parent As Outlook.Folder, ByVal name As String) As Outlook.Folder
    On Error Resume Next
    Set GetSub = parent.Folders(name)
End Function

' 하위 폴더 가져오기 (없으면 생성)
Public Function GetOrCreateSub(ByVal parent As Outlook.Folder, ByVal name As String) As Outlook.Folder
    On Error Resume Next
    Dim f As Outlook.Folder
    Set f = parent.Folders(name)
    If f Is Nothing Then Set f = parent.Folders.Add(name)
    Set GetOrCreateSub = f
End Function

' 폴더 이름의 '_' 앞부분(프로젝트 코드). '_' 가 없으면 빈 문자열
Public Function FolderCode(ByVal folderName As String) As String
    Dim p As Long
    p = InStr(folderName, "_")
    If p > 1 Then FolderCode = Trim(Left(folderName, p - 1))
End Function
