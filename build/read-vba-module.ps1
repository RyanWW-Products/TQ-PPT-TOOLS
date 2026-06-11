# Reads the TRUE source of a VBA module out of a .ppam/.pptm by parsing the
# vbaProject.bin compound file and MS-OVBA-decompressing the module stream.
# Ignores the _SRP_* performance-cache streams (which can hold stale symbols),
# so this reflects the source that was actually imported - unlike a raw byte grep.
param(
    [Parameter(Mandatory=$true)][string]$Ppam,
    [Parameter(Mandatory=$true)][string]$Module,
    [string[]]$Contains = @()
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-VbaProjectBytes([string]$path) {
    if ($path -like '*.bin') { return [IO.File]::ReadAllBytes($path) }
    $zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $path))
    try {
        $e = $zip.Entries | Where-Object { $_.FullName -ieq 'ppt/vbaProject.bin' } | Select-Object -First 1
        if (-not $e) { throw "ppt/vbaProject.bin not found in $path" }
        $ms = New-Object IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Dispose()
        return $ms.ToArray()
    } finally { $zip.Dispose() }
}

function Read-CfbStream([byte[]]$d, [string]$name) {
    $secShift = [BitConverter]::ToUInt16($d,30); $secSize = 1 -shl $secShift
    $miniShift = [BitConverter]::ToUInt16($d,32); $miniSize = 1 -shl $miniShift
    $firstDir = [BitConverter]::ToUInt32($d,48)
    $miniCutoff = [BitConverter]::ToUInt32($d,56)
    $firstMiniFat = [BitConverter]::ToUInt32($d,60)
    $firstDifat = [BitConverter]::ToUInt32($d,68); $numDifat = [BitConverter]::ToUInt32($d,72)
    function Off($s){ 512 + [int64]$s * $secSize }
    $ENDC=[uint32]4294967294; $FREE=[uint32]4294967295

    $difat = New-Object Collections.Generic.List[uint32]
    for ($i=0;$i -lt 109;$i++){ $v=[BitConverter]::ToUInt32($d,76+$i*4); if($v -ne $FREE){ $difat.Add($v) } }
    $sec=$firstDifat; $n=$numDifat
    while ($n -gt 0 -and $sec -ne $ENDC -and $sec -ne $FREE) {
        $b=Off $sec
        for ($i=0;$i -lt ($secSize/4 - 1);$i++){ $v=[BitConverter]::ToUInt32($d,$b+$i*4); if($v -ne $FREE){ $difat.Add($v) } }
        $sec=[BitConverter]::ToUInt32($d,$b+$secSize-4); $n--
    }
    $fat = New-Object Collections.Generic.List[uint32]
    foreach ($fs in $difat){ $b=Off $fs; for ($i=0;$i -lt ($secSize/4);$i++){ $fat.Add([BitConverter]::ToUInt32($d,$b+$i*4)) } }

    function Chain([uint32]$start){
        $o=New-Object IO.MemoryStream; $s=$start
        while ($s -ne $ENDC -and $s -ne $FREE){ $o.Write($d,[int](Off $s),$secSize); if($s -ge $fat.Count){break}; $s=$fat[[int]$s] }
        return $o.ToArray()
    }
    $dir = Chain $firstDir
    $entries=@()
    for ($o=0;$o -lt $dir.Length;$o+=128){
        $nl=[BitConverter]::ToUInt16($dir,$o+64); if($nl -le 0){continue}
        $nm=[Text.Encoding]::Unicode.GetString($dir,$o,$nl-2)
        $entries += [pscustomobject]@{ Name=$nm; Type=$dir[$o+66]; Start=[BitConverter]::ToUInt32($dir,$o+116); Size=[BitConverter]::ToUInt32($dir,$o+120) }
    }
    $tgt=$entries | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if (-not $tgt){ return $null }
    if ($tgt.Size -ge $miniCutoff){
        $raw=Chain $tgt.Start; return $raw[0..($tgt.Size-1)]
    } else {
        $root=$entries | Where-Object { $_.Type -eq 5 } | Select-Object -First 1
        $cont=Chain $root.Start
        $mfat=New-Object Collections.Generic.List[uint32]; $s=$firstMiniFat
        while ($s -ne $ENDC -and $s -ne $FREE){ $b=Off $s; for($i=0;$i -lt ($secSize/4);$i++){ $mfat.Add([BitConverter]::ToUInt32($d,$b+$i*4)) }; $s=$fat[[int]$s] }
        $o=New-Object IO.MemoryStream; $m=$tgt.Start
        while ($m -ne $ENDC -and $m -ne $FREE){ $o.Write($cont,[int]($m*$miniSize),$miniSize); $m=$mfat[[int]$m] }
        $arr=$o.ToArray(); return $arr[0..($tgt.Size-1)]
    }
}

function Decompress-Ovba([byte[]]$buf,[int]$start,[int]$cap){
    if ($buf[$start] -ne 1){ return $null }
    $pos=$start+1; $out=New-Object Collections.Generic.List[byte]
    while ($pos+1 -lt $buf.Length){
        if ($cap -gt 0 -and $out.Count -ge $cap){ break }
        $hdr=[BitConverter]::ToUInt16($buf,$pos); $pos+=2
        $size=($hdr -band 0x0FFF)+3; $flag=($hdr -band 0x8000) -ne 0
        $decStart=$out.Count
        if (-not $flag){
            for ($i=0;$i -lt 4096 -and $pos -lt $buf.Length;$i++){ $out.Add($buf[$pos]); $pos++ }
        } else {
            $end=$pos+($size-2)
            while ($pos -lt $end -and $pos -lt $buf.Length){
                $fb=$buf[$pos]; $pos++
                for ($bit=0;$bit -lt 8 -and $pos -lt $end -and $pos -lt $buf.Length;$bit++){
                    if (($fb -band (1 -shl $bit)) -eq 0){ $out.Add($buf[$pos]); $pos++ }
                    else {
                        $tok=[BitConverter]::ToUInt16($buf,$pos); $pos+=2
                        $diff=$out.Count-$decStart
                        $bc=[Math]::Max([Math]::Ceiling([Math]::Log($diff,2)),4)
                        $lenMask=0xFFFF -shr [int]$bc; $offMask=((-bnot $lenMask) -band 0xFFFF)
                        $len=($tok -band $lenMask)+3
                        $offset=(($tok -band $offMask) -shr (16-[int]$bc))+1
                        $src=$out.Count-$offset
                        for ($k=0;$k -lt $len;$k++){ $out.Add($out[$src+$k]) }
                    }
                }
            }
        }
    }
    return ,$out.ToArray()
}

$vba = Get-VbaProjectBytes $Ppam
$stream = Read-CfbStream $vba $Module
if (-not $stream){ Write-Output "MODULE NOT FOUND: $Module"; exit 2 }

# scan for the compressed-source container (decompresses to 'Attribute VB_Name')
$source=$null
for ($i=0;$i -lt $stream.Length;$i++){
    if ($stream[$i] -ne 1){ continue }
    try {
        $probe=Decompress-Ovba $stream $i 64
        if ($probe -and $probe.Length -ge 16){
            $head=[Text.Encoding]::ASCII.GetString($probe,0,[Math]::Min(32,$probe.Length))
            if ($head -like 'Attribute VB_Name*'){
                $full=Decompress-Ovba $stream $i 0
                $source=[Text.Encoding]::ASCII.GetString($full)
                break
            }
        }
    } catch { }
}
if (-not $source){ Write-Output "COULD NOT DECOMPRESS SOURCE for $Module"; exit 3 }

$lines = ($source -split "`r?`n").Count
Write-Output ("MODULE {0}: {1} source lines decompressed OK" -f $Module, $lines)
foreach ($needle in $Contains){
    $present = $source -match [regex]::Escape($needle)
    Write-Output ("  [{0}] contains: {1}" -f ($(if($present){'YES'}else{'no '}), $needle))
}
