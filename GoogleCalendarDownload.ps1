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
#(Get-Content -Path $calSavePath -Encoding UTF8 | Select-String -Pattern { "DTSTART:", "DTEND:", "DESCRIPTION:" }) -replace "DESCRIPTION:","" | Out-File $outputSavePath;
$calArray = @(Get-Content -Path $calSavePath -Encoding UTF8 | Select-String -Pattern "DTSTART:", "DTSTART;VALUE=DATE:", "DTEND:", "DTEND;VALUE=DATE:", "SUMMARY:", "DESCRIPTION:");

$result = "";
$dtStart;
$dtEnd;
$summary;
$description;

for( $i = 0; $i -lt $calArray.Count; $i++ ) {

    Write-Host "++++++++++++++++++++++++++++++++++++++++++++++++++++++";

    $tmp = $calArray[$i].ToString();
    if ($tmp.Contains("DTSTART:")) {
        Write-Host "1-1) `$dtStart";
        $tmp        =   $tmp.Replace( "DTSTART:", "" ).Replace( "T", " " );
        Write-Host $tmp;
        $dtStart    =   [DateTime]::ParseExact( $tmp, "yyyyMMdd HHmmssZ", $null );
        Write-Host $dtStart;
    } else {
        Write-Host "1-2) `$dtStart";
        $tmp        =   $tmp.Replace( "DTSTART;VALUE=DATE:", "" );
        Write-Host $tmp;
        $dtStart    =   [DateTime]::ParseExact( $tmp, "yyyyMMdd", $null );
        Write-Host $dtStart;
    }

    $i++;
    $tmp = $calArray[$i].ToString();
    if ($tmp.Contains("DTEND:")) {
        Write-Host "2-1) `$dtEnd";
        $tmp        =   $tmp.Replace( "DTEND:", "" ).Replace( "T", " " );
        Write-Host $tmp;
        $dtEnd      =   [DateTime]::ParseExact( $tmp.Replace( "DTEND:", "" ), "yyyyMMdd HHmmssZ", $null );
        Write-Host $dtEnd;
    } elseif ( $tmp.Contains("DTEND;VALUE=DATE:") ) {
        Write-Host "2-2) `$dtEnd";
        $tmp        =   $tmp.Replace( "DTEND;VALUE=DATE:", "" );
        Write-Host $tmp;
        $dtEnd      =   [DateTime]::ParseExact( $tmp, "yyyyMMdd", $null );
        Write-Host $dtEnd;
    } else {
        $dtEnd      =   $dtStart;
        $i--;
    }

    $i++;
    Write-Host "3) `$DESCRIPTION";
    $tmp = $calArray[$i].ToString();
    Write-Host $tmp;
    $description    =   $tmp.Replace( "DESCRIPTION:", "" ).Replace( "<br>", ", " ).Replace( "\n", ", " );
    Write-Host $description;

    $i++;
    Write-Host "4) `$SUMMARY";
    $tmp = $calArray[$i].ToString();
    Write-Host $tmp;
    $summary        =   $tmp.Replace( "SUMMARY:", "" );
    Write-Host $summary;

    $result         +=  $dtStart.ToString("yyyy/MM/dd") + ", " + $dtEnd.ToString("yyyy/MM/dd") + ", " + $summary + ", " + $description + "`r`n";

}

$outputSavePath = (Join-Path $Current $outputFileName);
$result | Out-File -FilePath $outputSavePath;

# 作業フォルダの後片付け
Set-Location -Path $Current;
Remove-Item -Path $WorkDir -Recurse -Force;