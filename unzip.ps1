Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    
    param(  [Parameter(ValueFromPipeline=$true)][string]$zip,
            [Parameter()][string]$target,
            [Parameter()][switch]$force,
            #ZIP uses a default codepage of IBM437.
            [Parameter()][System.Text.Encoding]$encoding=[System.Text.Encoding]::GetEncoding(437),
            [Parameter()][switch]$v)

    begin{
        [System.Diagnostics.Stopwatch] $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()
        if($v){
            Write-Output "verbose -> $v"
            Write-Output "target -> $target"
            Write-Output "force -> $force"
        }

        if($force -and [System.IO.Directory]::Exists($target)){
            Remove-Item -LiteralPath $target -Force -Recurse
            if($v){ Write-Output "removed directory: $target" }
        }
        if (![System.IO.Directory]::Exists($target)){
            $dir = [System.IO.Directory]::CreateDirectory($target)
            if($v){ Write-Output "created target directory: $dir `n" }
        }
    }

    process{
        if($v) {Write-Output "zip -> $zip" }
        $zip = Resolve-Path $zip
        if($v) {Write-Output "zip fullname -> $zip" }

        $entries = [System.IO.Compression.ZipFile]::Open($zip, [System.IO.Compression.ZipArchiveMode]::Read, $encoding).Entries
        foreach ($entry in $entries){
            if($v){ Write-Output "entry -> $entry" }

            $fullname = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine($target, $entry.FullName) )
            $dir = [System.IO.Path]::GetDirectoryName($fullname)
            if (![System.IO.Directory]::Exists($dir)){
                $dir = [System.IO.Directory]::CreateDirectory($dir)
                if($v) {Write-Output "created directory -> $dir" }
            }
            if($v){ Write-Output "extractToFile -> $fullname" }
            if(-Not $entry.Name -eq ""){
                extractToFile $entry $fullname $force
            }
        }
    }
    
    end{
        $stopwatch.Stop()
        $ms = $stopwatch.ElapsedMilliseconds
        Write-Output "unzipping $zip to $target finished in $ms ms"
    }
}

function extractToFile {
    param (
        [Parameter()][System.IO.Compression.ZipArchiveEntry]$source,
        [Parameter()][string]$destinationFileName,
        [Parameter()][bool]$overwrite=$false
    )
    $mode = [System.IO.FileMode]::CreateNew
    if($overwrite){ $mode = [System.IO.FileMode]::Create}

    try {
        [System.IO.Stream] $fs = [System.IO.FileStream]::new($destinationFileName, $mode, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        [System.IO.Stream] $es = $source.Open()
        $es.CopyTo($fs)
    }
    catch {
        Write-Output "error while extracting $source to $destinationFileName overwrite:$overwrite"
        Write-Output $Error[0]
    }
    finally {
        if($es){$es.Dispose()}
        if($fs){$fs.Dispose()}
    }

    [System.IO.File]::SetLastWriteTime($destinationFileName, $source.LastWriteTime.DateTime)
}


# Unzip with a custom name entry encoding
$encoding = [System.Text.Encoding]::GetEncoding(437)

#using piped input
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\gist\unzipped" -f -encoding $encoding -v
#using absolute input
Unzip -zip "C:\Users\renato\code\gist\a.zip" -target "C:\Users\renato\code\gist\a" -f -encoding $encoding 
#using relative input
Unzip -zip ".\b.zip" -target "C:\Users\renato\code\gist\b" -f -encoding $encoding -v