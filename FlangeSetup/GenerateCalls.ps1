$csv = Import-Csv "$PSScriptRoot\dims.csv"

$configs = @(
  [pscustomobject]@{cfg="Pn2_5_Dn1200"; pn="Pn_2_5"; dn=1200},
  [pscustomobject]@{cfg="Pn6_Dn10";     pn="Pn_6";   dn=10},
  [pscustomobject]@{cfg="Pn6_Dn15";     pn="Pn_6";   dn=15},
  [pscustomobject]@{cfg="Pn6_Dn20";     pn="Pn_6";   dn=20},
  [pscustomobject]@{cfg="Pn6_Dn25";     pn="Pn_6";   dn=25},
  [pscustomobject]@{cfg="Pn6_Dn32";     pn="Pn_6";   dn=32},
  [pscustomobject]@{cfg="Pn6_Dn40";     pn="Pn_6";   dn=40},
  [pscustomobject]@{cfg="Pn6_Dn50";     pn="Pn_6";   dn=50},
  [pscustomobject]@{cfg="Pn6_Dn65";     pn="Pn_6";   dn=65},
  [pscustomobject]@{cfg="Pn6_Dn80";     pn="Pn_6";   dn=80},
  [pscustomobject]@{cfg="Pn6_Dn100";    pn="Pn_6";   dn=100},
  [pscustomobject]@{cfg="Pn6_Dn125";    pn="Pn_6";   dn=125},
  [pscustomobject]@{cfg="Pn6_Dn150";    pn="Pn_6";   dn=150},
  [pscustomobject]@{cfg="Pn6_Dn200";    pn="Pn_6";   dn=200},
  [pscustomobject]@{cfg="Pn6_Dn250";    pn="Pn_6";   dn=250},
  [pscustomobject]@{cfg="Pn6_Dn300";    pn="Pn_6";   dn=300},
  [pscustomobject]@{cfg="Pn6_Dn350";    pn="Pn_6";   dn=350},
  [pscustomobject]@{cfg="Pn6_Dn400";    pn="Pn_6";   dn=400},
  [pscustomobject]@{cfg="Pn6_Dn450";    pn="Pn_6";   dn=450},
  [pscustomobject]@{cfg="Pn6_Dn500";    pn="Pn_6";   dn=500},
  [pscustomobject]@{cfg="Pn6_Dn600";    pn="Pn_6";   dn=600},
  [pscustomobject]@{cfg="Pn6_Dn700";    pn="Pn_6";   dn=700},
  [pscustomobject]@{cfg="Pn6_Dn800";    pn="Pn_6";   dn=800},
  [pscustomobject]@{cfg="Pn6_Dn900";    pn="Pn_6";   dn=900},
  [pscustomobject]@{cfg="Pn6_Dn1000";   pn="Pn_6";   dn=1000},
  [pscustomobject]@{cfg="Pn6_Dn1200";   pn="Pn_6";   dn=1200},
  [pscustomobject]@{cfg="Pn6_Dn1400";   pn="Pn_6";   dn=1400},
  [pscustomobject]@{cfg="Pn6_Dn1600";   pn="Pn_6";   dn=1600},
  [pscustomobject]@{cfg="Pn6_Dn1800";   pn="Pn_6";   dn=1800},
  [pscustomobject]@{cfg="Pn6_Dn2000";   pn="Pn_6";   dn=2000},
  [pscustomobject]@{cfg="Pn10_Dn200";   pn="Pn_10";  dn=200},
  [pscustomobject]@{cfg="Pn10_Dn250";   pn="Pn_10";  dn=250},
  [pscustomobject]@{cfg="Pn10_Dn300";   pn="Pn_10";  dn=300},
  [pscustomobject]@{cfg="Pn10_Dn350";   pn="Pn_10";  dn=350},
  [pscustomobject]@{cfg="Pn10_Dn400";   pn="Pn_10";  dn=400},
  [pscustomobject]@{cfg="Pn10_Dn450";   pn="Pn_10";  dn=450},
  [pscustomobject]@{cfg="Pn10_Dn500";   pn="Pn_10";  dn=500},
  [pscustomobject]@{cfg="Pn10_Dn600";   pn="Pn_10";  dn=600},
  [pscustomobject]@{cfg="Pn10_Dn700";   pn="Pn_10";  dn=700},
  [pscustomobject]@{cfg="Pn10_Dn800";   pn="Pn_10";  dn=800},
  [pscustomobject]@{cfg="Pn10_Dn900";   pn="Pn_10";  dn=900},
  [pscustomobject]@{cfg="Pn10_Dn1000";  pn="Pn_10";  dn=1000},
  [pscustomobject]@{cfg="Pn10_Dn1200";  pn="Pn_10";  dn=1200},
  [pscustomobject]@{cfg="Pn16_Dn50";    pn="Pn_16";  dn=50},
  [pscustomobject]@{cfg="Pn16_Dn65";    pn="Pn_16";  dn=65},
  [pscustomobject]@{cfg="Pn16_Dn80";    pn="Pn_16";  dn=80},
  [pscustomobject]@{cfg="Pn16_Dn100";   pn="Pn_16";  dn=100},
  [pscustomobject]@{cfg="Pn16_Dn125";   pn="Pn_16";  dn=125},
  [pscustomobject]@{cfg="Pn16_Dn150";   pn="Pn_16";  dn=150},
  [pscustomobject]@{cfg="Pn16_Dn200";   pn="Pn_16";  dn=200},
  [pscustomobject]@{cfg="Pn16_Dn250";   pn="Pn_16";  dn=250},
  [pscustomobject]@{cfg="Pn16_Dn300";   pn="Pn_16";  dn=300},
  [pscustomobject]@{cfg="Pn16_Dn350";   pn="Pn_16";  dn=350},
  [pscustomobject]@{cfg="Pn16_Dn400";   pn="Pn_16";  dn=400},
  [pscustomobject]@{cfg="Pn16_Dn450";   pn="Pn_16";  dn=450},
  [pscustomobject]@{cfg="Pn16_Dn500";   pn="Pn_16";  dn=500},
  [pscustomobject]@{cfg="Pn16_Dn600";   pn="Pn_16";  dn=600},
  [pscustomobject]@{cfg="Pn16_Dn700";   pn="Pn_16";  dn=700},
  [pscustomobject]@{cfg="Pn16_Dn800";   pn="Pn_16";  dn=800},
  [pscustomobject]@{cfg="Pn16_Dn900";   pn="Pn_16";  dn=900},
  [pscustomobject]@{cfg="Pn16_Dn1000";  pn="Pn_16";  dn=1000},
  [pscustomobject]@{cfg="Pn25_Dn200";   pn="Pn_25";  dn=200},
  [pscustomobject]@{cfg="Pn25_Dn250";   pn="Pn_25";  dn=250},
  [pscustomobject]@{cfg="Pn25_Dn300";   pn="Pn_25";  dn=300},
  [pscustomobject]@{cfg="Pn25_Dn350";   pn="Pn_25";  dn=350},
  [pscustomobject]@{cfg="Pn25_Dn400";   pn="Pn_25";  dn=400},
  [pscustomobject]@{cfg="Pn25_Dn450";   pn="Pn_25";  dn=450},
  [pscustomobject]@{cfg="Pn25_Dn500";   pn="Pn_25";  dn=500},
  [pscustomobject]@{cfg="Pn25_Dn600";   pn="Pn_25";  dn=600},
  [pscustomobject]@{cfg="Pn25_Dn700";   pn="Pn_25";  dn=700},
  [pscustomobject]@{cfg="Pn25_Dn800";   pn="Pn_25";  dn=800},
  [pscustomobject]@{cfg="Pn40_Dn10";    pn="Pn_40";  dn=10},
  [pscustomobject]@{cfg="Pn40_Dn15";    pn="Pn_40";  dn=15},
  [pscustomobject]@{cfg="Pn40_Dn20";    pn="Pn_40";  dn=20},
  [pscustomobject]@{cfg="Pn40_Dn25";    pn="Pn_40";  dn=25},
  [pscustomobject]@{cfg="Pn40_Dn32";    pn="Pn_40";  dn=32},
  [pscustomobject]@{cfg="Pn40_Dn40";    pn="Pn_40";  dn=40},
  [pscustomobject]@{cfg="Pn40_Dn50";    pn="Pn_40";  dn=50},
  [pscustomobject]@{cfg="Pn40_Dn65";    pn="Pn_40";  dn=65},
  [pscustomobject]@{cfg="Pn40_Dn80";    pn="Pn_40";  dn=80},
  [pscustomobject]@{cfg="Pn40_Dn100";   pn="Pn_40";  dn=100},
  [pscustomobject]@{cfg="Pn40_Dn125";   pn="Pn_40";  dn=125},
  [pscustomobject]@{cfg="Pn40_Dn150";   pn="Pn_40";  dn=150},
  [pscustomobject]@{cfg="Pn40_Dn200";   pn="Pn_40";  dn=200},
  [pscustomobject]@{cfg="Pn40_Dn250";   pn="Pn_40";  dn=250},
  [pscustomobject]@{cfg="Pn40_Dn300";   pn="Pn_40";  dn=300},
  [pscustomobject]@{cfg="Pn40_Dn350";   pn="Pn_40";  dn=350},
  [pscustomobject]@{cfg="Pn40_Dn400";   pn="Pn_40";  dn=400},
  [pscustomobject]@{cfg="Pn63_Dn50";    pn="Pn_63";  dn=50},
  [pscustomobject]@{cfg="Pn63_Dn65";    pn="Pn_63";  dn=65},
  [pscustomobject]@{cfg="Pn63_Dn80";    pn="Pn_63";  dn=80},
  [pscustomobject]@{cfg="Pn63_Dn100";   pn="Pn_63";  dn=100},
  [pscustomobject]@{cfg="Pn63_Dn125";   pn="Pn_63";  dn=125},
  [pscustomobject]@{cfg="Pn63_Dn150";   pn="Pn_63";  dn=150},
  [pscustomobject]@{cfg="Pn63_Dn200";   pn="Pn_63";  dn=200},
  [pscustomobject]@{cfg="Pn63_Dn250";   pn="Pn_63";  dn=250},
  [pscustomobject]@{cfg="Pn63_Dn300";   pn="Pn_63";  dn=300},
  [pscustomobject]@{cfg="Pn63_Dn350";   pn="Pn_63";  dn=350},
  [pscustomobject]@{cfg="Pn63_Dn400";   pn="Pn_63";  dn=400},
  [pscustomobject]@{cfg="Pn100_Dn10";   pn="Pn_40";  dn=10},
  [pscustomobject]@{cfg="Pn100_Dn15";   pn="Pn_40";  dn=15},
  [pscustomobject]@{cfg="Pn100_Dn20";   pn="Pn_40";  dn=20},
  [pscustomobject]@{cfg="Pn100_Dn25";   pn="Pn_40";  dn=25},
  [pscustomobject]@{cfg="Pn100_Dn32";   pn="Pn_40";  dn=32},
  [pscustomobject]@{cfg="Pn100_Dn40";   pn="Pn_40";  dn=40},
  [pscustomobject]@{cfg="Pn100_Dn50";   pn="Pn_40";  dn=50},
  [pscustomobject]@{cfg="Pn100_Dn65";   pn="Pn_40";  dn=65},
  [pscustomobject]@{cfg="Pn100_Dn80";   pn="Pn_40";  dn=80},
  [pscustomobject]@{cfg="Pn100_Dn100";  pn="Pn_40";  dn=100},
  [pscustomobject]@{cfg="Pn100_Dn125";  pn="Pn_40";  dn=125},
  [pscustomobject]@{cfg="Pn100_Dn150";  pn="Pn_40";  dn=150},
  [pscustomobject]@{cfg="Pn100_Dn200";  pn="Pn_40";  dn=200},
  [pscustomobject]@{cfg="Pn100_Dn250";  pn="Pn_40";  dn=250},
  [pscustomobject]@{cfg="Pn100_Dn300";  pn="Pn_40";  dn=300},
  [pscustomobject]@{cfg="Pn100_Dn350";  pn="Pn_40";  dn=350}
)

# Track last known non-zero value per PN+key per field
$lastKnown = @{}

$out = @()
$prefix = "            "
$helper = "SetConfig(doc, flangeSketch, boreCutSketch, boltHoleSketch, boltHolesPatternFeature,"
$prevPnLabel = ""

foreach ($c in $configs) {
    $pnLabel = $c.cfg -replace "_Dn\d+", ""
    $pnKey = "$($c.pn)"

    if ($pnLabel -ne $prevPnLabel) {
        $out += ""
        $out += "${prefix}// ── $pnLabel ──────────────────────────────────────────────────────"
        $prevPnLabel = $pnLabel
        if (-not $lastKnown.ContainsKey($pnKey)) { $lastKnown[$pnKey] = @{} }
    }

    $row = $csv | Where-Object { $_.PN -eq $c.pn -and [int]$_.DN -eq [int]$c.dn } | Select-Object -First 1
    if (-not $row) { $out += "${prefix}// MISSING DATA: $($c.cfg)"; continue }

    function ResolveVal($pnKey2, $field, $raw) {
        if ($raw -and $raw -ne "" -and $raw -ne "0" -and $raw -ne "0.0") {
            $script:lastKnown[$pnKey2][$field] = $raw
            return $raw
        }
        if ($script:lastKnown[$pnKey2].ContainsKey($field)) { return $script:lastKnown[$pnKey2][$field] }
        return "0.0"
    }

    $D  = ResolveVal $pnKey "D"  $row.D
    $C1 = ResolveVal $pnKey "C1" $row.C1
    $B1 = ResolveVal $pnKey "B1" $row.B1
    $K  = ResolveVal $pnKey "K"  $row.K
    $L  = ResolveVal $pnKey "L"  $row.L
    $bc = $row.BoltCount

    $out += "${prefix}$helper `"$($c.cfg)`", D:$D, C1:$C1, B1:$B1, K:$K, L:$L, boltCount:$bc);"
}

$out | Set-Content "$PSScriptRoot\generated_calls.txt" -Encoding UTF8
Write-Host "Done - $($out.Count) lines written"
