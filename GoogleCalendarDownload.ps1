# iCal形式の非公開URL
$iCalURL = "(ここにURLを設定します)";

# 作業フォルダの作成と移動
$Current = Get-Location;
$WorkDir = New-TemporaryFile | ForEach-Object {
    Remove-Item $_
    mkdir $_
};
Set-Location -Path $WorkDir;

# ファイルのダウンロード準備
$wc = New-Object System.Net.WebClient;
if ($null -eq $wc) {
    Write-Warning "[Error 01] New-Object() failed. cannot create object of System.Net.WebClient.";
    exit 1;
}

# iCalダウンロード
$calUri = New-Object System.Uri($iCalURL);
$calFileName = "calendar.ics";
$calSavePath = (Join-Path $WorkDir $calFileName);
$wc.DownloadFile($calUri, $calSavePath);
Write-Information "Downloaded calendar. Path: ${calSavePath}";

# 検索
$outputFileName = (Get-Date).ToString("yyyyMMdd_hhmmss") + "_export_calendar.csv";
$calArray = @(Get-Content -Path $calSavePath -Encoding UTF8 | Select-String -Pattern "DTSTART:", "DTSTART;VALUE=DATE:", "DTEND:", "DTEND;VALUE=DATE:", "SUMMARY:", "DESCRIPTION:");

$result = "Start Date, End Date, Summary, Description`r`n";
$dtStart;
$dtEnd;
$summary;
$description;

for( $i = 0; $i -lt $calArray.Count; $i++ ) {

    $tmp = $calArray[$i].ToString();
    try {
        if ($tmp.Contains("DTSTART:")) {
            $tmp        =   $tmp.Replace( "DTSTART:", "" ).Replace( "T", " " );
            $dtStart    =   [DateTime]::ParseExact( $tmp, "yyyyMMdd HHmmssZ", $null );
        } elseif ($tmp.Contains("DTSTART;VALUE=DATE:")) {
            $tmp        =   $tmp.Replace( "DTSTART;VALUE=DATE:", "" );
            $dtStart    =   [DateTime]::ParseExact( $tmp, "yyyyMMdd", $null );
        } else {
            $i--;
        }
    } catch {
        Write-Output "DTSTART Error.";
        Write-Output $tmp
    }

    $i++;
    $tmp = $calArray[$i].ToString();
    try {
        if ($tmp.Contains("DTEND:")) {
            $tmp        =   $tmp.Replace( "DTEND:", "" ).Replace( "T", " " );
            $dtEnd      =   [DateTime]::ParseExact( $tmp.Replace( "DTEND:", "" ), "yyyyMMdd HHmmssZ", $null );
        } elseif ( $tmp.Contains("DTEND;VALUE=DATE:") ) {
            $tmp        =   $tmp.Replace( "DTEND;VALUE=DATE:", "" );
            $dtEnd      =   [DateTime]::ParseExact( $tmp, "yyyyMMdd", $null );
        } elseif ($tmp.Contains("DESCRIPTION:") -or $tmp.Contains("SUMMARY:")) {
            $i--;
        } else {
            $dtEnd      =   $dtStart;
            $i--;
        }
    } catch {
        Write-Output "DTEND Error.";
        Write-Output $tmp
    }

    $i++;
    $tmp = $calArray[$i].ToString();
    if ($tmp.Contains("DESCRIPTION:")) {
        $description    =   $tmp.Replace( "DESCRIPTION:", "" ).Replace( "<br>", ", " ).Replace( "\n", ", " ).Replace( "`\,", ", " );
        while ($description.Contains(", ,")) {
            $description = $description.Replace( ", ,", ", " );
        }
    } else {
        $i--;
    }

    $i++;
    $tmp = $calArray[$i].ToString();
    $summary        =   $tmp.Replace( "SUMMARY:", "" );
    while ($summary.Contains(", ,")) {
        $summary    = $summary.Replace( ", ,", ", " );
    }

    $result         +=  $dtStart.ToString("yyyy/MM/dd") + ", " + $dtEnd.ToString("yyyy/MM/dd") + ", " + $summary + ", " + $description + "`r`n";
    while ($result.Contains(", ,")) {
        $result    = $result.Replace( ", ,", ", " );
    }

}

$outputSavePath = (Join-Path $Current $outputFileName);
$result | Out-File -FilePath $outputSavePath;

# 作業フォルダの後片付け
Set-Location -Path $Current;
Remove-Item -Path $WorkDir -Recurse -Force;
