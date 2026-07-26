# 생활인구 행정동(OA-14991) 나머지 seq 다운로드->파싱->삭제 (seq별, 디스크 1개만 보유)
# 2219(2020H1)은 이미 처리됨. 백그라운드 실행용. 로그 = livingpop_build.log
$ErrorActionPreference = "Continue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$dir = "C:\Users\wh850\AppData\Local\Temp\livingpop"
# ASCII 경로 사본 — PowerShell이 native python에 Korean 경로 인자를 넘기면 [Errno 22]로 깨짐 (실측 2026-06-21)
$py  = "C:\Users\wh850\AppData\Local\Temp\livingpop\parser.py"
$url = "https://datafile.seoul.go.kr/bigfile/iot/inf/nio_download.do?useCache=false"
$log = Join-Path $dir "livingpop_build.log"
function Log($m) { $t = Get-Date -Format "HH:mm:ss"; "$t  $m" | Tee-Object -FilePath $log -Append }

$seqs = 2220,2221,2222,2223,2224,2301,2302,2303,2304,2305,2306,2307,2308,2309,2310,2311,2312
Log "START livingpop download (17 seq)"
foreach ($seq in $seqs) {
    $zip = Join-Path $dir "seq$seq.bin"
    $partial = Join-Path $dir "partial_$seq.csv"
    if (Test-Path $partial) { Log "skip $seq (partial exists)"; continue }
    $body = "infId=OA-14991&seq=$seq&infSeq=3"
    try {
        Log "downloading seq $seq ..."
        Invoke-WebRequest -Uri $url -Method POST -Body $body -ContentType "application/x-www-form-urlencoded" -OutFile $zip -TimeoutSec 580
    } catch {
        Log "  ERROR download seq $seq : $($_.Exception.Message)"
        continue
    }
    $mb = (Get-Item $zip).Length / 1MB
    if ($mb -lt 1) {
        Log "  WARN seq $seq tiny ($([math]::Round($mb,2))MB) -> likely error page, kept for inspection, skip parse"
        continue
    }
    Log ("  got {0:N1}MB, parsing..." -f $mb)
    $out = & python3 $py $zip $partial 2>&1
    Log "  $out"
    Remove-Item $zip -Force
    Log "  done seq $seq, zip deleted"
}
Log "ALL DONE"
