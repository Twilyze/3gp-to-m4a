# 3gpファイルをm4aに変換するスクリプト v1.3.0
using namespace System.IO
Add-Type -AssemblyName Microsoft.VisualBasic
chcp 65001
Write-Output '3gpファイルをm4aに変換するスクリプト v1.3.0'


# 同じ名前のm4aファイルが既にある場合、変換をスキップするか
$SkipDuplicateFile = $false

$picName = 'cover'

$fileAll   = '_all-3gp.tsv'
$fileMeta  = '_meta.tsv'
$file3gpPL = '_3gp.m3u8'
$fileM4aPL = '_converted-m4a.m3u8'
$fileNoPicPL      = '_no-picture.m3u8'
$fileNoPicAlbum   = '_no-picture-album.tsv'
$fileUnexpectedPL = '_unexpected.m3u8'
$fileUnexpected   = '_unexpected.tsv'

$RequiredExeFiles = @(
  'exiftool.exe',
  'MediaInfo.exe',
  'ffmpeg.exe'
)


#----------------------------
# 関数
#----------------------------
# カバーアート抽出
function ExtractCoverArt {
  param ([string]$3gpPath, [string]$picName, [string]$parentDirectory, [ref]$CoverArtPath)

  try {
    $picPath = "$parentDirectory\$picName"
    if (Test-Path -LiteralPath "$picPath.jpg" -PathType leaf) {
      $CoverArtPath.Value = "$picPath.jpg"
    } elseif (Test-Path -LiteralPath "$picPath.png" -PathType leaf) {
      $CoverArtPath.Value = "$picPath.png"
    } else {
      if (Test-Path -LiteralPath $picPath -PathType leaf) {
        # Write-Output "'$picName' already exists: '$parentDirectory'"
      } else {
        cmd /c "exiftool `"$3gpPath`" -b -CoverArt -charset filename=`"`" > `"$picPath`""

        # 「ソ」や特定の漢字など制御記号として認識される文字が含まれているとエラー
        if ($LASTEXITCODE -ne 0) {
          # 出力された空のファイルを削除
          Remove-Item -LiteralPath $picPath
          Write-Output "ERROR: Contains characters that cannot be used: '$3gpPath'"
          $CoverArtPath.Value = ''
          return
        }
      }

      # カバーアートに拡張子を設定
      if (Test-Path -LiteralPath $picPath -PathType leaf) {
        $ext = mediainfo --Inform="General;%Format%" $picPath
        if ($ext -eq 'JPEG') {
          Rename-Item -LiteralPath $picPath "$picName.jpg"
          $CoverArtPath.Value = "$picPath.jpg"
        } elseif ($ext -eq 'PNG') {
          Rename-Item -LiteralPath $picPath "$picName.png"
          $CoverArtPath.Value = "$picPath.png"
        } elseif ($ext -eq 'BMP') {
          Rename-Item -LiteralPath $picPath "$picName.bmp"
          $CoverArtPath.Value = "$picPath.bmp"
        } else {
          Write-Output "ERROR: Unexpected image format: '$picPath' '$ext'"
          $CoverArtPath.Value = ''
        }
      } else {
        Write-Output "ERROR: Cannot find '$picPath'"
        $CoverArtPath.Value = ''
      }
    }
  }
  catch {
    Write-Host $_ -ForegroundColor red
    $CoverArtPath.Value = ''
  }
}

# ファイルをゴミ箱へ移動
function FilesToRecycleBin([array]$paths, [bool]$showLog = $false) {
  $length = $paths.count
  for ($i = 0; $i -lt $length; $i++) {
    try {
      if (Test-Path -LiteralPath $paths[$i] -PathType Leaf) {
        $fullpath = Convert-Path -LiteralPath $paths[$i]
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($fullpath, 'OnlyErrorDialogs', 'SendToRecycleBin')
        if ($showLog) {
          Write-Output "SendToRecycle: '$($paths[$i])'"
        }
      } else {
        Write-Output "ERROR: '$($paths[$i])' is not file or not found."
      }
    }
    catch {
      Write-Host $_ -ForegroundColor red
    }
    Write-Progress -Activity 'SendToRecycle' -Status "$($i + 1)/$length" -PercentComplete (($i + 1) / $length * 100)
  }
}

# cmdに渡す引数用のエスケープ
function EscapeCMD([string]$arg) {
  # example
  #   arg: %path% & | < > ( ) % ^ \"
  #   ret: ^"^%path^% ^& ^| ^< ^> ^( ^) ^% ^^ \\\^"^"
  # 特殊文字(^ & | < > ( ) " %)を全て ^ でエスケープ
  # " の前にある連続する \ を同数の \ でエスケープ
  # 全ての " を \ でエスケープ
  # 全体を ^" 〜 ^" で囲む
  return '^"' + (
  $arg `
  -replace '\^', '^^' `
  -replace '&',  '^&' `
  -replace '\|', '^|' `
  -replace '<',  '^<' `
  -replace '>',  '^>' `
  -replace '\(', '^(' `
  -replace '\)', '^)' `
  -replace '"',  '^"' `
  -replace '%',  '^%' `
  -replace '\\(?=\\*?\^")', '\\' `
  -replace '\^"', '\^"'
  ) + '^"'
}


#--------------------
# ポップアップメッセージ用
$wshell = New-Object -ComObject wscript.shell
$vbOKOnly = 0
$vbYesNo  = 4
$vbCritical = 16
$vbExclamation = 48
$vbYes = 6

$result = 0
$popupSkip = $true


#----------------------------
# 必要なファイルがあるか確認
#----------------------------
$missingExeFiles = @()
foreach ($file in $RequiredExeFiles) {
  try {
    Get-Command $file -ErrorAction Stop > $null
  }
  catch {
    Write-Host "$file が見つかりません" -ForegroundColor red
    $missingExeFiles += $file
  }
}
if ($missingExeFiles.count -ne 0) {
  $wshell.popup("$($missingExeFiles -join ', ')が見つかりません`nスクリプトを終了します", 0, 'エラー', $vbOKOnly + $vbCritical)
  exit
}


#----------------------------
# ExifToolで抽出した情報からファイル(m3u8,img,tsv)を作成
#----------------------------
if (Test-Path -LiteralPath $fileM4aPL -PathType leaf) {
  $result = $wshell.popup("既に'$fileM4aPL'が存在します。`nもう一度ファイルリストを作成しますか？", 0, '確認', $vbYesNo)
  $popupSkip = $false
}
if ($popupSkip -or ($result -eq $vbYes)) {
  $NoPicCount = 0
  $UnexpectedCount = 0
  try {
    Write-Output "`n--- ExifToolで3gpファイルのメタデータを抽出 ---"
    cmd /c "exiftool -r . -ext 3gp -csv -csvDelim `"\t`" -FileType -AudioFormat -Title -Performer -Album -AlbumArtist -Year -Genre -CoverArt -progress: > `"$fileAll`""

    Write-Output "`n--- ExifToolで抽出した情報からファイル(m3u8,img,tsv)を作成 ---"
    $items = Import-Csv -LiteralPath $fileAll -Delimiter `t -Encoding UTF8

    $swMeta  = New-Object StreamWriter($fileMeta)
    $sw3gpPL = New-Object StreamWriter($file3gpPL)
    $swM4aPL = New-Object StreamWriter($fileM4aPL)
    $swNoPicPL      = New-Object StreamWriter($fileNoPicPL)
    $swNoPicAlbum   = New-Object StreamWriter($fileNoPicAlbum)
    $swUnexpectedPL = New-Object StreamWriter($fileUnexpectedPL)
    $swUnexpected   = New-Object StreamWriter($fileUnexpected)

    # tsvのヘッダー書き込み
    $swMeta.WriteLine(@('Title', 'Performer', 'Album', 'AlbumArtist', 'Year', 'Track', 'DiscNumber', 'Genre', 'CoverArtPath') -join "`t")
    $swNoPicAlbum.WriteLine(@('Album', 'AlbumArtist', 'Directory') -join "`t")
    $swUnexpected.WriteLine(@('Path', 'FileType', 'AudioFormat') -join "`t")

    $preParent = ''
    $CoverArtPath = ''
    foreach ($item in $items) {
      $parentDirectory = $item.SourceFile | Split-Path

      # スラッシュをバックスラッシュに置換
      $item.SourceFile = $item.SourceFile.Replace('/', '\')


      #--------------------
      # パスからプロパティ追加

      # 2階層上のフォルダ名（アルバムアーティスト名）を設定
      if (($null -eq $item.AlbumArtist) -or ('' -eq $item.AlbumArtist)) {
        $item | Add-Member AlbumArtist ($parentDirectory | Split-Path | Split-Path -leaf) -force
        # フォルダ名に使用できないよく使われる文字列を変換
        $item.AlbumArtist = $item.AlbumArtist -replace ' (F)eat_ ', ' $1eat. '
      }

      # 親フォルダ名（アルバム名）にディスクナンバーがあれば設定
      if (($parentDirectory | Split-Path -leaf) -match '[\[\(]disc (\d+)[\]\)]$') {
        $item | Add-Member DiscNumber ([int]$Matches.1)
      }

      # ファイル名にトラックナンバーがあれば設定
      if (($item.SourceFile | Split-Path -leaf) -match '^(\d+?)-') {
        $item | Add-Member Track ([int]$Matches.1)
      }


      #--------------------
      # ファイル書き込み
      $typeAndFormat = $item.FileType + $item.AudioFormat
      if (($typeAndFormat -eq '3GPmp4a') -or ($typeAndFormat -eq 'MP4mp4a')) {
        # 各ディレクトリ1回だけカバーアート抽出を実行
        if ($parentDirectory -ne $preParent) {
          $preParent = $parentDirectory
          $CoverArtPath = ''
          if ($item.CoverArt -ne '') {
            ExtractCoverArt $item.SourceFile $picName $parentDirectory ([ref]$CoverArtPath)
          }
          if ($CoverArtPath -eq '') {
            # カバーアートが設定できなかったアルバムのリスト(tsv)
            $swNoPicAlbum.WriteLine(@($item.Album, $item.AlbumArtist, $parentDirectory) -join "`t")
            $NoPicCount++
          }
        }
        # カバーアートが設定できなかった曲のリスト(m3u8)
        if ($CoverArtPath -eq '') {
          $swNoPicPL.WriteLine($item.SourceFile)
        }

        # 3gpのメタデータのリスト(tsv)と変換前と変換後のリスト(m3u8)
        $swMeta.WriteLine(@($item.Title, $item.Performer, $item.Album, $item.AlbumArtist, $item.Year, $item.Track, $item.DiscNumber, $item.Genre, $CoverArtPath) -join "`t")
        $sw3gpPL.WriteLine($item.SourceFile)
        $swM4aPL.WriteLine(($item.SourceFile -replace '\.3gp$', '.m4a'))
      }
      else {
        # 予期しない形式のリスト(m3u8,tsv)
        $swUnexpectedPL.WriteLine($item.SourceFile)
        $swUnexpected.WriteLine(@($item.SourceFile, $item.FileType, $item.AudioFormat) -join "`t")
        $UnexpectedCount++
      }
    }
  }
  catch {
    Write-Host $_ -ForegroundColor red
  }
  finally {
    if ($null -ne $swMeta)         { $swMeta.Close() }
    if ($null -ne $sw3gpPL)        { $sw3gpPL.Close() }
    if ($null -ne $swM4aPL)        { $swM4aPL.Close() }
    if ($null -ne $swNoPicPL)      { $swNoPicPL.Close() }
    if ($null -ne $swNoPicAlbum)   { $swNoPicAlbum.Close() }
    if ($null -ne $swUnexpectedPL) { $swUnexpectedPL.Close() }
    if ($null -ne $swUnexpected)   { $swUnexpected.Close() }
    Write-Output " > $fileMeta"
    Write-Output " > $file3gpPL"
    Write-Output " > $fileM4aPL"

    if ($NoPicCount -eq 0) {
      Remove-Item -LiteralPath $fileNoPicPL
      Remove-Item -LiteralPath $fileNoPicAlbum
    } else {
      Write-Output " > $fileNoPicPL"
      Write-Output " > $fileNoPicAlbum"
      Write-Output "カバーアートが設定できなかったアルバム件数：$NoPicCount"
      Start-Process -FilePath notepad.exe -ArgumentList $fileNoPicAlbum
    }

    if ($UnexpectedCount -eq 0) {
      Remove-Item -LiteralPath $fileUnexpectedPL
      Remove-Item -LiteralPath $fileUnexpected
    } else {
      Write-Output " > $fileUnexpectedPL"
      Write-Output " > $fileUnexpected"
      Write-Output "予期しない形式のファイル件数：$UnexpectedCount"
      Start-Process -FilePath notepad.exe -ArgumentList $fileUnexpected
    }
  }
} else {
  Write-Output '-- skip --'
}


#----------------------------
# 3gp → m4a 変換
#----------------------------
Write-Output "`n--- 3gp → m4a 変換 ---"
if (!$popupSkip) {
  $result = $wshell.popup("3gpファイルをm4aに変換しますか？", 0, '3gp → m4a 変換', $vbYesNo)
}
if ($popupSkip -or ($result -eq $vbYes)) {
  $inputPaths  = @(Get-Content -LiteralPath $file3gpPL -Encoding UTF8)
  $outputPaths = @(Get-Content -LiteralPath $fileM4aPL -Encoding UTF8)
  $items = Import-Csv -LiteralPath $fileMeta -Delimiter `t -Encoding UTF8
  $length = $inputPaths.count

  if ($length -eq 0) {
    $wshell.popup("変換できる3gpファイルがありません`nスクリプトを終了します", 0, 'エラー', $vbOKOnly + $vbCritical)
    exit
  }

  if (!$SkipDuplicateFile -and (Test-Path -LiteralPath $outputPaths[0] -PathType Leaf)) {
    $result = $wshell.popup("同じ名前のm4aファイルが既にある場合、変換をスキップしますか？", 0, '3gp → m4a 変換', $vbYesNo)
    if ($result -eq $vbYes) {
      $SkipDuplicateFile = $true
    }
  }

  Write-Output "変換中  件数：$($length)"
  for ($i = 0; $i -lt $length; $i++) {
    try {
      if ($SkipDuplicateFile -and (Test-Path -LiteralPath $outputPaths[$i] -PathType Leaf)) {
        # 同じ名前のファイルがある場合スキップ
      } elseif (Test-Path -LiteralPath $inputPaths[$i] -PathType Leaf) {
        $item = $items[$i]

        # ffmpegに渡す用のエスケープ
        $Title       = EscapeCMD $item.Title
        $Performer   = EscapeCMD $item.Performer
        $Album       = EscapeCMD $item.Album
        $AlbumArtist = EscapeCMD $item.AlbumArtist
        $Year        = EscapeCMD $item.Year
        $Genre       = EscapeCMD $item.Genre

        if ($item.CoverArtPath -ne '') {
          cmd /c "ffmpeg -hide_banner -loglevel error -y -i `"$($inputPaths[$i])`" -i `"$($item.CoverArtPath)`" -map 0:a -map 1:v -metadata `"title`"=$Title -metadata `"artist`"=$Performer -metadata `"album`"=$Album -metadata `"album_artist`"=$AlbumArtist -metadata `"date`"=$Year -metadata `"track`"=`"$($item.Track)`" -metadata `"disc`"=`"$($item.DiscNumber)`" -metadata `"genre`"=$Genre -c copy -disposition:1 attached_pic `"$($outputPaths[$i])`""
        } else {
          cmd /c "ffmpeg -hide_banner -loglevel error -y -i `"$($inputPaths[$i])`" -map 0:a -metadata `"title`"=$Title -metadata `"artist`"=$Performer -metadata `"album`"=$Album -metadata `"album_artist`"=$AlbumArtist -metadata `"date`"=$Year -metadata `"track`"=`"$($item.Track)`" -metadata `"disc`"=`"$($item.DiscNumber)`" -metadata `"genre`"=$Genre -c copy `"$($outputPaths[$i])`""
        }

        # タイムスタンプ引き継ぎ
        $inputItem = Get-Item -LiteralPath $inputPaths[$i]
        $outputItem = Get-Item -LiteralPath $outputPaths[$i]
        $outputItem.CreationTime     = $inputItem.CreationTime
        $outputItem.CreationTimeUtc  = $inputItem.CreationTimeUtc
        $outputItem.LastWriteTime    = $inputItem.LastWriteTime
        $outputItem.LastWriteTimeUtc = $inputItem.LastWriteTimeUtc
      } else {
        Write-Output "ERROR: '$($inputPaths[$i])' is not file or not found."
      }

      Write-Progress -Activity '3gp -> m4a 変換' -Status "$($i + 1)/$length" -PercentComplete (($i + 1) / $length * 100)
    }
    catch {
      Write-Host $_ -ForegroundColor red
    }
  }
} else {
  Write-Output '-- skip --'
}

#--------------------
# 3gpとm4aの数が違う場合は警告
$files = Get-ChildItem -File -Recurse -Include *.3gp, *.m4a | Group-Object Extension -NoElement
Write-Output $files
if ($files[0].count -ne $files[1].count) {
  $result = $wshell.popup("3gpファイルとm4aファイルの数が違います。`n処理を続行しますか？", 0, '警告', $vbYesNo + $vbExclamation)
  if ($result -ne $vbYes) {
    Read-Host "`n`n中断  エンターキーかウィンドウを閉じて終了"
    exit
  }
}


#----------------------------
# 変換前の3gpファイルをゴミ箱へ
#----------------------------
Write-Output "`n--- 変換前の3gpファイルをゴミ箱へ ---"
$result = $wshell.popup("変換前の3gpファイルをゴミ箱へ移動しますか？`n※3gp → m4aの変換前に削除しないように注意", 0, 'ファイル削除', $vbYesNo)
if ($result -eq $vbYes) {
  $paths = @(Get-Content -LiteralPath $file3gpPL -Encoding UTF8)
  Write-Output "削除件数：$($paths.count)"
  FilesToRecycleBin $paths

  #--------------------
  $paths = @($fileAll, $fileMeta, $file3gpPL)
  Write-Output "変換中に使用したファイルを削除：$($paths.count)"
  FilesToRecycleBin $paths $true
} else {
  Write-Output '-- skip --'
}


Read-Host "`n`n完了  エンターキーかウィンドウを閉じて終了"
exit
