# ============================================================
# Copy-OutlookRulesToSent.ps1
# 받은 메세지 규칙(본문 조건)을 보낸 메세지 규칙으로 복제
# 실행 방법: Outlook이 열린 상태에서 PowerShell로 이 스크립트를 실행하세요.
# ============================================================

# Outlook COM 연결 (이미 열려 있는 인스턴스 사용)
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    Write-Host "[OK] 실행 중인 Outlook에 연결했습니다." -ForegroundColor Green
} catch {
    Write-Host "[오류] Outlook이 실행 중이지 않습니다. Outlook을 먼저 열고 다시 실행하세요." -ForegroundColor Red
    exit 1
}

$rules = $outlook.Session.DefaultStore.GetRules()
$totalRules = $rules.Count
Write-Host "총 $totalRules 개의 규칙을 발견했습니다.`n"

$createdCount = 0
$skippedCount = 0

# 기존 규칙 이름 목록 미리 수집 (중복 방지)
$existingNames = @{}
for ($i = 1; $i -le $rules.Count; $i++) {
    $existingNames[$rules.Item($i).Name] = $true
}

for ($i = 1; $i -le $totalRules; $i++) {
    $rule = $rules.Item($i)

    # 받은 메세지 규칙만 처리 (olRuleReceive = 0)
    if ($rule.RuleType -ne 0) { continue }

    $bodyCondition = $rule.Conditions.Body

    # 본문 조건이 없는 규칙은 건너뜀
    if (-not $bodyCondition.Enabled) {
        Write-Host "[건너뜀] '$($rule.Name)' - 본문 조건 없음"
        $skippedCount++
        continue
    }

    $copyAction = $rule.Actions.CopyToFolder

    # 폴더 복사 액션이 없는 규칙은 건너뜀
    if (-not $copyAction.Enabled -or $null -eq $copyAction.Folder) {
        Write-Host "[건너뜀] '$($rule.Name)' - 폴더 복사 액션 없음"
        $skippedCount++
        continue
    }

    $targetFolder = $copyAction.Folder
    $sentRuleName = $rule.Name + "_SENT"

    # 이미 동일한 이름의 규칙이 있으면 건너뜀
    if ($existingNames.ContainsKey($sentRuleName)) {
        Write-Host "[건너뜀] '$sentRuleName' - 이미 존재함"
        $skippedCount++
        continue
    }

    # 보낸 메세지 규칙 생성 (olRuleSend = 1)
    try {
        $newRule = $rules.Create($sentRuleName, 1)
        $newRule.Conditions.Body.Text = $bodyCondition.Text
        $newRule.Conditions.Body.Enabled = $true
        $newRule.Actions.CopyToFolder.Folder = $targetFolder
        $newRule.Actions.CopyToFolder.Enabled = $true
        $newRule.Enabled = $true

        $existingNames[$sentRuleName] = $true
        Write-Host "[생성] '$sentRuleName' -> '$($targetFolder.Name)'" -ForegroundColor Cyan
        $createdCount++
    } catch {
        Write-Host "[오류] '$sentRuleName' 생성 실패: $_" -ForegroundColor Red
    }
}

# 규칙 저장
try {
    $rules.Save()
    Write-Host "`n[완료] 규칙이 저장되었습니다." -ForegroundColor Green
} catch {
    Write-Host "`n[오류] 규칙 저장 실패: $_" -ForegroundColor Red
}

Write-Host "생성: $createdCount 개 / 건너뜀: $skippedCount 개"
Write-Host "`nOutlook에서 홈 > 규칙 > 규칙 및 알림 관리를 열어 확인하세요."
