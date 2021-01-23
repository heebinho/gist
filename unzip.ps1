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
        if($v){
            Write-Output "verbose -> $v"
            Write-Output "target -> $target"
            Write-Output "force -> $force"
        }

        if($force -and [System.IO.Directory]::Exists($target)){
            Remove-Item -LiteralPath $target -Force -Recurse
            if($v){ Write-Host "removed directory: $target" }
        }
        if (![System.IO.Directory]::Exists($target)){
            $dir = [System.IO.Directory]::CreateDirectory($target)
            if($v){ Write-Host "created target directory: $dir `n" }
        }
    }

    process{
        if($v) {Write-Output "zip -> $zip" }
        $zip = Resolve-Path $zip
        if($v) {Write-Output "zip fullname -> $zip" }
        
        $entries = [System.IO.Compression.ZipFile]::Open($zip, [System.IO.Compression.ZipArchiveMode]::Read, $encoding).Entries
        foreach ($entry in $entries){

            $fullname = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine($target, $entry.FullName) )
            $dir = [System.IO.Path]::GetDirectoryName($fullname)
            if (![System.IO.Directory]::Exists($dir)){
                $dir = [System.IO.Directory]::CreateDirectory($dir)
                if($v) {Write-Host "created directory -> $dir" }
                continue
            }
            if($v){ write-host "extract file -> $fullname" }
            extractToFile $entry $fullname $force
        }
    }
    
    end{ if($v){Write-Output "unzipping finished"} }
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
        Write-Host "error while extracting $source to $destinationFileName overwrite:$overwrite"
        Write-Host $Error[0]
    }
    finally {
        if($es){$es.Dispose()}
        if($fs){$fs.Dispose()}
    }

    [System.IO.File]::SetLastWriteTime($destinationFileName, $source.LastWriteTime.DateTime)
}

#[System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $target)

# Unzip
$encoding = [System.Text.Encoding]::GetEncoding(437)
#Write-Output $encoding
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\ps\unzipped" -f -encoding $encoding -v


Unzip -zip "C:\Users\renato\code\ps\a.zip" -target "C:\Users\renato\code\ps\asdf" -f -encoding $encoding -v

Unzip -zip ".\b.zip" -target "C:\Users\renato\code\ps\asdf" -f -encoding $encoding -v