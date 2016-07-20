Param(
    [switch]$GeneratePassword=$false,
    [switch]$SkipHeader=$false,
    [int]$PasswordLength=10,
    [Parameter(Mandatory)]$InputFile,
    [Parameter(Mandatory)]$OutputFile
) 
# Create local windows users from tsv file.
#
# Copyright (c) 2016 Masatsugu Mizuno
#
# This software is released under the MIT License.
# https://opensource.org/licenses/MIT


$ErrorActionPreference = "Stop"
$random = New-Object System.Random

function generatePassword ($passwordLength) {
    $characterTable = ($larges + $smalls + $digits + $symbols).ToCharArray()

    $candidatePassword = ""
    for($i = 1; $i -le $passwordLength; $i++) {
        $candidatePassword += $characterTable[$random.next(0, $characterTable.Length)]

    }
    return $candidatePassword
}


# パスワードに使用可能な文字列の定義
$larges = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
$smalls = "abcdefghijklmnopqrstuvwxyz"
$digits = "1234567890"
$symbols = " !`"#$%&'()*+,-./:;<=>?@[\]^_``{|}~"


# userFlagsの定義
# https://support.microsoft.com/ja-jp/kb/2812553=
$ADS_UF_ACCOUNTDISABLE = 2
$ADS_UF_PASSWD_CANT_CHANGE = 64
$ADS_UF_DONT_EXPIRE_PASSWD = 65536
# これでは設定できない
# $ADS_UF_PASSWORD_EXPIRED = [System.UInt32]8388608 # 0x800000

# パスワード再生成の最大値
$retryLimit = 10


$computerName = $env:COMPUTERNAME
$ou = [ADSI]"WinNT://$computerName"


$groupList=@{}

# ファイルの読み込み
# 先頭行はヘッダとして読み飛ばす
# TODO: Inport-CSVを用いるかどうかは検討
# フォーマット(タブ区切り)
    # name: ユーザ名
    # full name: フルネーム
    # group: 所属グループ(カンマ区切りで複数指定可能)
    # password: パスワード(フィールドを'RANDOM'にすると、乱数から自動生成
    # user must change password at next logon: ユーザーは次回ログオン時にパスワードの変更が必要(0, 無効, false, 1, 有効, trueのいずれか)
    # user cannot change password: ユーザーはパスワードを変更できない(0, 無効, false, 1, 有効, trueのいずれか)
    # password never expires: パスワードを無期限にする(0, 無効, false, 1, 有効, trueのいずれか)
    # account is disable: アカウントを無効にする(0, 無効, false, 1, 有効, trueのいずれか)
$userLists = Get-Content $InputFile `
    | Select-Object -Skip ([int]$SkipHeader.ToBool()) `
    | ConvertFrom-Csv -Delimiter "`t" -Header "name","fullName","groups","password","mustChangePasswordAtNextLogon","cantChangePassowrd","isNeverExpires","isDisable"


 ForEach($_ in $userLists) {
    Write-Host ("Adding user '" + $_.name + "'")

    $user = $ou.Create("User", $_.name)
    $userFlags = 0

    # ユーザへのオプション作成
    if($_.mustChangePasswordAtNextLogon -match "1|true|有効"){
        # ユーザーは次回ログオン時にパスワードの変更が必要
        $user.put("PasswordExpired", 1)

    } else {
        if($_.cantChangePassowrd -match "1|true|有効") {
            # ユーザーはパスワードを変更できない
            $userFlags = $userFlags -bor $ADS_UF_PASSWD_CANT_CHANGE
        }

        if($_.isNeverExpires -match "1|true|有効") {
            # パスワードを無期限にする
            $userFlags = $userFlags -bor $ADS_UF_DONT_EXPIRE_PASSWD
        }
    }

    if($_.isDisable -match "1|true|有効") {
        # アカウントを無効にする
        $userFlags = $userFlags -bor $ADS_UF_ACCOUNTDISABLE
    }

    $user.put("FullName", $_.fullName)
    $user.put("description", "")
    $user.put("UserFlags", $userFlags)

    # パスワードが'RANDOM'か、GeneratePasswordオプションが指定された場合はパスワードを生成
    if($_.password -eq "RANDOM" -or $GeneratePassword) {
        for($i = 0; $i -lt $retryLimit;$i++) {
            try {
                $_.password = generatePassword($PasswordLength)
                $user.setpassword($_.password)
                $user.setInfo()
                break

            } catch [Exception] {
                # 既に存在する
                # -2147022672

                # 複雑性の要件を満たしていない場合のみリトライ
                # -2147022651
                # https://www.manageengine.jp/products/ADSelfService_Plus/ADSSP_help_J/misc/troubleshooting_tips.html#error_800708c5
                if($error[0].Exception.InnerException.ErrorCode -ne -2147022651) {
                    throw $error[0]
                }
            }
        }
    } else {
        $user.setpassword($_.password)
        $user.setInfo()
    }

    if($i -eq $retryLimit) {
        throw "Reached retry limit"
    }


    $joinGroups = $_.groups -split " *, *"
    foreach($groupName in $joinGroups) {
        # グループが存在するかどうかチェック
        if($groupList[$groupName] -eq $null) {
            if(($ou.PSBase.Children | Where-Object { $_.psBase.schemaClassName -eq "Group" -and $_.Name -eq $groupName}) -eq $null) {
                Write-Host ("Group '" + $groupName + "' dose not exist. Create group.")

                # グループの作成
                $group = $ou.Create("Group", $groupName)
                $group.setInfo()
            }

            # チェック済としてマーク
            $groupList[$groupName] = [ADSI]"WinNT://$computerName/$groupName"
        }

        # グループにユーザを追加
        $groupList[$groupName].add("WinNT://$computerName/" + $_.name) 
        $groupList[$groupName].SetInfo() 
    }

    Write-Output ($_.name + "`t" + $_.password) | Add-Content $OutputFile -Encoding Default
}
