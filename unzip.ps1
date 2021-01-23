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
            #Write-Output $encoding
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
        if($v) {Write-Output "extracting -> $zip" }

        #[System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $target)

        $entries = [System.IO.Compression.ZipFile]::Open($zip, [System.IO.Compression.ZipArchiveMode]::Read, $encoding).Entries
        foreach ($entry in $entries) {

            $fullname = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine($target, $entry.FullName) )
            $dir = [System.IO.Path]::GetDirectoryName($fullname)
            
            if (![System.IO.Directory]::Exists($dir)){
                $dir = [System.IO.Directory]::CreateDirectory($dir)
                if($v) {Write-Host "created directory -> $dir" }
                continue
            }
            if($v){ write-host "extract file -> $fullname" }
            [System.IO.Compression.ZipArchiveEntryExtensions]::ExtractToFile($entry, $fullname, $force)           
        }
        
    }
    
    end{
        #Write-Output "teardown"
    }
     
}

$source = @"
namespace System.IO.Compression
{
    public static class ZipArchiveEntryExtensions
    {
        public static void ExtractToFile(ZipArchiveEntry source, string destinationFileName, bool overwrite)
        {
            FileMode fMode = overwrite ? FileMode.Create : FileMode.CreateNew;

            using (Stream fs = new FileStream(destinationFileName, fMode, FileAccess.Write, FileShare.None, bufferSize: 0x1000, useAsync: false))
            {
                using (Stream es = source.Open())
                    es.CopyTo(fs);
            }

            File.SetLastWriteTime(destinationFileName, source.LastWriteTime.DateTime);
        }
    }
}
"@

Add-Type -TypeDefinition $source


# Unzip
 
$encoding = [System.Text.Encoding]::GetEncoding(437)
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\ps\unzipped" -f -encoding $encoding -v
