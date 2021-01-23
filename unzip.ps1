Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    
    param(  [Parameter(ValueFromPipeline=$true)][string]$zip,
            [Parameter()][string]$target,
            [Parameter()][bool]$force=$true,
            [Parameter()][System.Text.Encoding]$encoding=[System.Text.Encoding]::GetEncoding(437) )
    
    begin{
        Write-Output "setup"
        Write-Output "target -> $target"
        Write-Output $force
        #Write-Output $encoding
        if([System.IO.Directory]::Exists($target)){
            Remove-Item -LiteralPath $target -Force -Recurse
            Write-Host "removed directory: $target"
        }
        if (![System.IO.Directory]::Exists($target)){
            [System.IO.Directory]::CreateDirectory($target)
            Write-Host "created directory: $target"
        }
        
    }

    process{
        Write-Output $zip
        # Write-Output $encoding
        #$encoding = [System.Text.Encoding]::GetEncoding($codepage)
        #[System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $target)

        $entries = [System.IO.Compression.ZipFile]::Open($zip, [System.IO.Compression.ZipArchiveMode]::Read, $encoding).Entries
        foreach ($entry in $entries) {
            Write-Host $entry
            #[System.IO.Compression.ZipFile]::ExtractToDirectory()
            $fullname = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine($target, $entry.FullName) )
            $dir = [System.IO.Path]::GetDirectoryName($fullname)
            Write-Host "dir -> $dir"
            if (![System.IO.Directory]::Exists($dir)){
                [System.IO.Directory]::CreateDirectory($dir)
                continue
            }
            write-host "fullname -> $fullname"
            #$entry.ExtractToFile($fullname, $force)
            [System.IO.Compression.ZipArchiveEntryExtensions]::ExtractToFile($entry, $fullname, $force)
        }
        
    }
    
    end{
        Write-Output "teardown"
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
# Get-ChildItem -Path ".\*.zip" "c" | Unzip
 
$encoding = [System.Text.Encoding]::GetEncoding(437)
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\ps\unzipped" -force $true -encoding $encoding
