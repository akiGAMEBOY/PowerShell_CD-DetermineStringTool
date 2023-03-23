#################################################################################
# 処理名　｜CD-DetermineStringTool（メイン処理）
# 機能　　｜CDドライブ内のファイルをチェックするツール
#--------------------------------------------------------------------------------
# 戻り値　｜下記の通り。
# 　　　　｜   0：正常終了
# 　　　　｜-101：設定ファイル読込みエラー
# 　　　　｜-201：CDトレイのオープン（くり返し前の初回時）エラー
# 　　　　｜-301：CDの挿入（セット）エラー
# 　　　　｜-401：CDドライブの検証エラー
# 　　　　｜-402：ファイル名の検証エラー
# 　　　　｜-403：ファイル詳細の検証エラー
# 　　　　｜-501：CDトレイのオープン（くり返し中の終了時）エラー
# 引数　　｜-
#################################################################################
# 定数
[System.String]$c_config_file = "setup.ini"
[System.String]$c_cdlabel = "集計データ"
[System.Int32]$c_retry_count = 3
[System.Int32]$c_interval_sec = 10
[System.Int32]$c_wait_sec = 3
[System.Int32]$c_for_count = 50

# Function
#################################################################################
# 処理名　｜ExpandString
# 機能　　｜文字列を展開（先頭桁と最終桁にあるダブルクォーテーションを削除）
#--------------------------------------------------------------------------------
# 戻り値　｜String（展開後の文字列）
# 引数　　｜target_str：対象文字列
#################################################################################
Function ExpandString($target_str) {
    [System.String]$expand_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq "`"") -and
            ($target_str.Substring($target_str.Length - 1, 1) -eq "`"")) {
            $expand_str = $target_str.Substring(1, $target_str.Length - 2)
           }
    }

    return $expand_str
}

#################################################################################
# 処理名　｜OpenCdtray
# 機能　　｜CDトレイを自動でオープン
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：異常終了）
# 引数　　｜drive_full：対象ドライブ
#################################################################################
Function OpenCdtray($drive_full) {
    [System.Boolean]$return = $false
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

    # CDトレイのオープン
    [System.String]$prompt_message = ''
    try {
        (New-Object -com Shell.Application).Namespace(17).ParseName("${drive_full}").InvokeVerb("Eject")
        $return = $true
    }
    catch {
        $sbtemp=New-Object System.Text.StringBuilder
        $prompt_message = ''
        @("エラー　　：CDトレイ オープン処理`r`n",`
          "　　　　　　処理が失敗しました。`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }

    return $return
}

#################################################################################
# 処理名　｜InsertCD
# 機能　　｜CDトレイにCDメディアを手動で挿入（セット）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：処理中断）
# 引数　　｜label：対象CDのラベル名
#################################################################################
Function InsertCD($label) {
    [System.Boolean]$return = $false
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder
    [System.String]$prompt_message = ''

    @("捜査依頼　：CDのセット`r`n",`
      "　　　　　　処理を一時停止中。CDトレイに [ ${label} ] をセットしてください。`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    Write-Host $prompt_message -ForegroundColor DarkYellow
    
    $sbtemp=New-Object System.Text.StringBuilder
    $prompt_message = ''
    @("確認　　　：処理続行の確認`r`n",`
      "　　　　　　CDをトレイに入れた後に応答し再開してください。処理を再開しますか？`r`n",`
      "　　　　　　[ y: はい、n: いいえ ]`r`n",`
      "`r`n",`
      "入力　　　")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()

    # YesNo入力
    $return = ConfirmYesno $prompt_message

    return $return
}

#################################################################################
# 処理名　｜ConfirmYesno
# 機能　　｜YesNo入力
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：処理中断）
# 引数　　｜prompt_message：入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno($prompt_message) {
    [System.Boolean]$return = $false
    [System.String]$value = $null
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

    for($i=1; $i -le $c_retry_count; $i++) {
        # 入力受付
        try {
            [ValidateSet("y","Y","n","N")]$value = Read-Host $prompt_message
        }
        catch {
            $value = $null
        }
        Write-Host ''

        # 入力値チェック
        if ($value.ToLower() -eq "y") {
            $return = $true
            break
        }
        elseif ($value.ToLower() -eq "n") {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　：いいえを選択`r`n", `
              "　　　　　　処理を中断します。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
            break
        }
        elseif ($i -eq $c_retry_count) {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　：リトライ回数を超過`r`n", `
              "　　　　　　リトライ回数（", `
              [System.String]$c_retry_count, `
              "回）を超過した為、処理を中断します。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
        }
    }

    return $return
}

#################################################################################
# 処理名　｜ValidateDrive
# 機能　　｜CDドライブの検証
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：異常終了）
# 引数　　｜drive：対象ドライブ（ドライブレターのみ）, drive_full：対象ドライブ
#################################################################################
Function ValidateDrive($drive) {
    [System.Boolean]$return = $false
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder
    [System.Boolean]$is_exists = $false

    [System.String]$prompt_message = ''
    [System.Int32]$now = 1
    [System.Int32]$max = $c_wait_sec * $c_interval_sec
    [System.Management.Automation.PSDriveInfo]$psdrive = $null
    for($i=1; $i -le $c_interval_sec; $i++) {
        $psdrive = Get-PSDrive $drive 2>$null
        if ($null -ne $psdrive) {
            [Object[]]$itemlist = Get-ChildItem "${drive_full}" | Sort-Object -Descending {$_.Name}
            # CD内のファイル件数をカウント
            if ($itemlist.Count -ge 1) {
                $return = $true
                $sbtemp=New-Object System.Text.StringBuilder
                $prompt_message = ''
                @("通知　　　：CDドライブの検証`r`n",`
                  "　　　　　　正常にCDドライブを認識しました。`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $prompt_message = $sbtemp.ToString()
                Write-Host $prompt_message
                break
            }
            else {
                $is_exists = $true
                $sbtemp=New-Object System.Text.StringBuilder
                $prompt_message = ''
                @("エラー　　：CDドライブの検証`r`n",`
                  "　　　　　　CDドライブ内のデータがありませんでした。`r`n",`
                  "　　　　　　処理を中断します。`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $prompt_message = $sbtemp.ToString()
                Write-Host $prompt_message -ForegroundColor DarkRed
                break
            }
        }
        # スリープで待ち合わせ（読み込みに時間がかかった場合、後続処理が動いてしまう為）
        Start-Sleep $c_wait_sec
        $now = $c_wait_sec * $i
        $sbtemp=New-Object System.Text.StringBuilder
        $prompt_message = ''
        @("通知　　　：CDドライブの検証`r`n",`
          "　　　　　　チェック中。　経過時間 / 待ち時間 [ ${now} / ${max} sec ]`r`n",`
          "　　　　　　CDを認識するまで少々、お待ちください。`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message
    }
    # 待ち合わせたが、認識できなかった場合
    if ((-not $return) -And (-not $is_exists)) {
        $sbtemp=New-Object System.Text.StringBuilder
        $prompt_message = ''
        @("エラー　　：CDドライブの検証`r`n",`
          "　　　　　　CDを認識できませんでした。`r`n",`
          "　　　　　　処理を中断します。`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }

  return $return
}

#################################################################################
# 処理名　｜ValidateFileformat
# 機能　　｜ファイル形式（ファイル名と拡張子）の検証
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：異常終了）
# 引数　　｜drive_full：対象ドライブ
#################################################################################
Function ValidateFileformat($drive_full) {
    [System.Boolean]$return = $false

    # ファイルの命名規則をチェック
    $return = $true
    [Object[]]$itemlist = Get-ChildItem "${drive_full}" | Sort-Object {$_.Name}
    foreach($item in $itemlist) {
        # ファイル名が下記の通りである事を検証
        #   1～4桁目　　：アルファベット（大文字・小文字を区別しない）
        #   ファイル種類：csvファイル or テキストファイル
        if (-not(($item.Name -match '^[A-z][A-z][A-z][A-z]') -And `
            (($item.Name.ToLower() -match '\.csv$') -Or `
             ($item.Name.ToLower() -match '\.txt$')))) {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            $prompt_message = ''
            @("エラー　　：ファイル名の検証`r`n",`
              "　　　　　　既定のファイル名ではありません。`r`n",`
              "　　　　　　対象ファイル：[$($item.FullName)]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
            break
        }
    }

  return $return
}

#################################################################################
# 処理名　｜ValidateFiledetail
# 機能　　｜ファイル内の検証（文字列の有無を判定）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：正常終了, False：異常終了）
# 引数　　｜drive_full：対象ドライブ, findrange：検索範囲, findstring：検索文字列
#################################################################################
Function ValidateFiledetail($drive_full, $findrange, $findstring) {
    [System.Boolean]$return = $false
    
    [System.String]$without_ext = ''
    [System.String]$prompt_message = ''
    [System.Text.RegularExpressions.MatchCollection]$compared = $null

    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

    # CDドライブ内のファイル名の検証
    [Object[]]$itemlist = Get-ChildItem "${drive_full}" | Sort-Object {$_.Name}
    foreach($item in $itemlist) {
    
        # 最終行から10行分の文字列を検索
        $compared = [regex]::Matches((Get-Content $item.FullName -last $findrange),"${findstring}")
        # ファイル名から拡張子を除外した文字列
        $without_ext = [System.IO.Path]::GetFileNameWithoutExtension($item.FullName);
        if ($compared.Count -ge 1) {
            $return = $true
            $sbtemp=New-Object System.Text.StringBuilder
            $prompt_message = ''
            @("通知　　　：ファイルの詳細を検証`r`n",`
              "　　　　　　検証結果：成功（$(($without_ext))）`r`n",`
              "　　　　　　検索の文字列「$(($findstring))」が見つかりました。`r`n",`
              "　　　　　　対象：[$($item.FullName)]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor Blue
        }
        else {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            $prompt_message = ''
            @("エラー　　：ファイルの詳細を検証`r`n",`
              "　　　　　　検証結果：失敗（$(($without_ext))）`r`n",`
              "　　　　　　検索の文字列「$(($findstring))」が見つかりませんでした。`r`n",`
              "　　　　　　対象：[$($item.FullName)]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
            break
        }
    }
    return $return
}

#################################################################################
# 処理名　｜ConfirmLoop
# 機能　　｜くり返し確認
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True：繰り返す, False：繰り返し終了）
# 引数　　｜-
#################################################################################
Function ConfirmLoop() {
    [System.Boolean]$return = $false
    [System.String]$prompt_message = ''
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

    $return = $true

    # くり返し有無を確認
    $prompt_message = ''
    @("確認　　　：くり返し有無の確認`r`n",`
      "　　　　　　処理が終了しました。続けて処理しますか？`r`n",`
      "　　　　　　[ y: はい、n: いいえ ]`r`n",`
      "`r`n",`
      "入力　")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()

    # YesNo入力
    $return = ConfirmYesno $prompt_message

    return $return
}

#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
# 変数
[System.Int32]$result = 0
[System.Boolean]$return = $false
[System.String]$prompt_message = ''
[System.String]$result_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 設定ファイ読み込み
[System.String]$config_path = @(Split-Path $script:MyInvocation.MyCommand.path -parent).Trim()
$sbtemp=New-Object System.Text.StringBuilder
@("$config_path",`
  "\",`
  "$c_config_file")|
ForEach-Object{[void]$sbtemp.Append($_)}
$config_fullpath = $sbtemp.ToString()

try {
    [System.Collections.Hashtable]$param = Get-Content $config_fullpath -Raw | ConvertFrom-StringData
    # ドライブパス作成
    [System.String]$drive = $param.DriveLatter
    [System.String]$drive_full = ''
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder
    @("${drive}",":\")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $drive_full = $sbtemp.ToString()
    # 検索範囲
    [System.Int32]$findrange = $param.FindRange
    # 検索文字列
    # [System.String]$findstring = $param.FindString
    [System.String]$findstring = ExpandString($param.FindString)

    $sbtemp=New-Object System.Text.StringBuilder
    @("通知　　　：設定ファイル読み込み`r`n",`
      "　　　　　　設定ファイルの読み込みが正常終了しました。`r`n",`
      "　　　　　　対象：[${config_fullpath}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    Write-Host $prompt_message
}
catch {
    $result = -101
    $sbtemp=New-Object System.Text.StringBuilder
    @("エラー　　：設定ファイル読み込み`r`n",`
      "　　　　　　設定ファイルの読み込みが異常終了しました。`r`n",`
      "　　　　　　エラー内容：[${config_fullpath}",`
    "$($_.Exception.Message)]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
}

# CDトレイを自動でオープン（くり返し開始前）
if ($result -eq 0) {
    $return = $false
    $return = OpenCdtray($drive_full)
    if (-not $return) {
        $result = -201
    }
}

# くり返し開始
[System.Int32]$count = 0
for ($count = 1; $count -le $c_for_count; $count++) {
    # CDトレイにCDメディアを手動で挿入（セット）
    if ($result -eq 0) {
        $return = $false
        $return = InsertCD $c_cdlabel

        if (-not $return) {
            $result = -301
        }
    }

    # CDドライブの検証
    if ($result -eq 0) {
        $return = $false
        $return = ValidateDrive $drive $drive_full

        if (-not $return) {
            $result = -401
        }
    }

    # ファイル形式（ファイル名と拡張子）の検証
    if ($result -eq 0) {
        $return = $false
        $return = ValidateFileformat $drive_full

        if (-not $return) {
            $result = -402
        }
    }

    # ファイル内の検証（文字列の有無を判定）
    if ($result -eq 0) {
        $return = $false
        $return = ValidateFiledetail $drive_full $findrange $findstring
        if (-not $return) {
            $result = -403
        }
    }
    
    # CDトレイを自動でオープン（くり返し終了前）
    if ($result -ne -201) {
        $return = OpenCdtray $drive_full
        if (-not $return) {
            $result = -501
        }
    }

    # 処理結果の表示
    $sbtemp=New-Object System.Text.StringBuilder
    if ($result -eq 0) {
        @("処理結果　：正常終了`r`n",`
          "　　　　　　メッセージコード：[${result}]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $result_message = $sbtemp.ToString()
        Write-Host $result_message
    }
    else {
        @("処理結果　：異常終了`r`n",`
          "　　　　　　メッセージコード：[${result}]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $result_message = $sbtemp.ToString()
        Write-Host $result_message -ForegroundColor DarkRed
    }

    # くり返し確認
    $return = $false
    # 正常終了時に繰り返すか確認
    if ($result -eq 0) {
        $return = ConfirmLoop
    }
    # 異常終了、または処理中断した場合はくり返し終了
    if (-not $return) {
        break
    }
}
