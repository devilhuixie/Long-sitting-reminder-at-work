<# :
@echo off
:: 将当前脚本交给 PowerShell 以纯净模式执行
chcp 65001 >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "$script = Get-Content -LiteralPath '%~f0' -Encoding UTF8 -Raw; Invoke-Expression $script"
exit /b
#>

$ErrorActionPreference = "Stop"
$targetDir = "$env:APPDATA\StandUpReminder"

function Show-Menu {
    while ($true) {
        Clear-Host
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "             ⏰ 久坐提醒安装/卸载工具 (工作日优享版)" -ForegroundColor Yellow
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "  提醒时间: 10:00, 11:30, 13:30, 15:00, 16:30"
        Write-Host "  提醒方式: 电脑右下角原生弹窗 (纯后台运行无黑框)"
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "  1. 安装提醒任务 (可选择每天或仅工作日)"
        Write-Host "  2. 卸载提醒任务"
        Write-Host "  3. 测试提醒弹窗 (查看效果)"
        Write-Host "  4. 退出"
        Write-Host "==================================================" -ForegroundColor Cyan
        $choice = Read-Host "请输入选项 (1-4) 并按回车"

        switch ($choice) {
            "1" { Install-Reminder; break }
            "2" { Uninstall-Reminder; break }
            "3" { Test-Reminder; break }
            "4" { exit }
        }
    }
}

function Install-Reminder {
    Write-Host "`n请选择提醒周期：" -ForegroundColor Cyan
    Write-Host "  1. 每天提醒 (包含周末)"
    Write-Host "  2. 仅工作日提醒 (周一至周五)"
    $scheduleChoice = Read-Host "请输入选项 (1-2) 并按回车"

    if ($scheduleChoice -ne "1" -and $scheduleChoice -ne "2") {
        Write-Host "输入无效，默认将为您设置为【仅工作日提醒】。" -ForegroundColor Yellow
        $scheduleChoice = "2"
    }

    Write-Host "`n正在配置运行环境..." -ForegroundColor Green

    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }

    # 1. 生成弹窗脚本 (使用 UTF8 保证中文不乱码)
    $ps1Path = "$targetDir\reminder.ps1"
    $ps1Content = @"
`$xml = @'
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>⏰ 久坐提醒</text>
            <text>工作辛苦了！您已经坐了很久，为了健康，请站起来活动几分钟、喝杯水吧！</text>
        </binding>
    </visual>
</toast>
'@
`$AppId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
`$doc = New-Object Windows.Data.Xml.Dom.XmlDocument
`$doc.LoadXml(`$xml)
`$toast = [Windows.UI.Notifications.ToastNotification]::new(`$doc)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(`$AppId).Show(`$toast)
"@
    [System.IO.File]::WriteAllText($ps1Path, $ps1Content, [System.Text.Encoding]::UTF8)

    # 2. 生成 VBS 隐藏启动脚本 (强制使用 ASCII 写入，兼容老引擎)
    $vbsPath = "$targetDir\run.vbs"
    $vbsContent = @"
Set objShell = CreateObject("WScript.Shell")
strAppData = objShell.ExpandEnvironmentStrings("%APPDATA%")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strAppData & "\StandUpReminder\reminder.ps1""", 0, False
"@
    [System.IO.File]::WriteAllText($vbsPath, $vbsContent, [System.Text.Encoding]::ASCII)

    # 3. 注册系统计划任务
    Write-Host "正在注册定时任务..." -ForegroundColor Green
    try {
        $TaskName = "StandUpReminderTask"
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""
        $times = @('10:00', '11:30', '13:30', '15:00', '16:30')
        
        # 根据用户选择配置触发器
        if ($scheduleChoice -eq "1") {
            $Triggers = $times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
            $descSchedule = "每天"
        } else {
            # 核心修改：指定为每周触发，且仅限周一至周五
            $Triggers = $times | ForEach-Object { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $_ }
            $descSchedule = "工作日(周一至周五)"
        }

        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        $Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

        Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Triggers -Settings $Settings -Principal $Principal -Description "久坐提醒定时任务($descSchedule)" -Force | Out-Null

        Write-Host "`n✅ 安装完成！以后 [$descSchedule] 的指定时间将自动在右下角提醒。" -ForegroundColor Yellow
        Write-Host "（你可以按 3 测试一下弹窗效果）" -ForegroundColor Gray
    } catch {
        Write-Host "`n❌ 注册计划任务失败，错误信息：" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    Write-Host "`n按任意键返回菜单..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Uninstall-Reminder {
    Write-Host "`n正在移除计划任务..." -ForegroundColor Green
    Unregister-ScheduledTask -TaskName "StandUpReminderTask" -Confirm:$false -ErrorAction SilentlyContinue

    Write-Host "正在清理系统文件..." -ForegroundColor Green
    if (Test-Path $targetDir) {
        Remove-Item -Path $targetDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host "`n🗑️ 卸载完成，已清理所有相关配置！" -ForegroundColor Yellow
    Write-Host "`n按任意键返回菜单..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Test-Reminder {
    if (Test-Path "$targetDir\run.vbs") {
        Write-Host "`n正在发送测试弹窗，请留意屏幕右下角..." -ForegroundColor Green
        Start-Process "wscript.exe" -ArgumentList "`"$targetDir\run.vbs`"" -NoNewWindow
    } else {
        Write-Host "`n⚠️ 请先输入 1 进行安装，然后才能测试！" -ForegroundColor Red
    }
    Write-Host "`n按任意键返回菜单..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

$host.UI.RawUI.WindowTitle = "久坐提醒工具"
Show-Menu