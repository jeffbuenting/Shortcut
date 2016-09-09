#-------------------------------------------------------------------------------------------
# Modules Shortuct.psm1
#-------------------------------------------------------------------------------------------

Function New-Shortcut {

<#
    .Synopsis
        Creates a shortcut.

    .Description
        Creates a shortcut in the specified location.

    .Parameter Link
        Application or Webpage's link.

    .Parameter Path
        Location to save the shortcut.  Defaults to current path.

    .Parameter Name
        Name of the Shortcut
        

    .Link
        http://stackoverflow.com/questions/34164110/powershell-create-new-shortcut-named-after-website
#>

    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$True)]
        [String[]]$Link,

        [String]$Path = (Get-Location),
        
        [String]$Name
    )
    
    Begin {
        Write-Verbose "Saving Shortcut at path $Path"
    }

    Process {
        Foreach ( $L in $Link ) {
            write-verbose $L.Length
                        
            if ( $L -ne $Null -and $L.Length -ne 0 ) {
                    if ( -Not $Name ) { 
                        $L -match 'http:\/\/([^\/]*)' | out-Null
                        $Name = $Matches[1]
                    }


                    # ----- Check if file exist and increment number if it does
                    if ( Test-Path -Path "$Path\$Name.url" ) {
                        $I = 1
                        while ( Test-Path -Path "$Path\$Name$I.url" ) { $I++ }
                        $Name = "$Name$I"
                    }

                    Write-Verbose "Creating Shortcut for $L"
                    $WshShell = New-Object -comObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut("$Path\$Name.url")
                    $Shortcut.TargetPath = $L
                    $Shortcut.Save()
                }
                else {
                    Write-Verbose "Link was Empty or Null"
            }
        }
    }  
}

#-------------------------------------------------------------------------------------------

function Resolve-ShortcutFile {
     
<#
    .Synopsis
        converts a desktop shortcut to its url

    .Description
        Retrieves the web page address (URL) from a shortcut.

    .Parameter FileName
        Full Path to the shortcut

    .Parameter LikeString
        String to filter.  If the wep page address is 'Like' this string then it will be returned.  Otherwise it will be skipped.

    .Example
        get-childitem -path $RootPath | Resolve-ShortcutFile -LikeString 'fanhole'

    .LINK
        http://blogs.msdn.com/b/powershell/archive/2008/12/24/resolve-shortcutfile.aspx

    .Note
        Author : Jeff Buenting
#>          
           
    [CmdLetBinding()]
    param(
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position = 0)]
        [Alias("FullName")]
        [string]
        $fileName,

        [String]$LikeString = $Null
    )
    
    process {
        
        Write-Verbose "Processing $_"
        if ( $LikeString -ne $Null ) {
                if ($fileName -like "*.url") {
                    Write-Verbose "Filtering Urls"
                    $ShortCut = Get-Content $fileName | Where-Object { $_ -like "url=*" -and $_ -like "*$LikeString*" } 
                    if ( $ShortCut -ne $Null ) {
                        Write-Output ( New-Object -Type PSObject -Property @{'FileName'= $FileName; 'Url' = $ShortCut.Substring($ShortCut.IndexOf("=") + 1 ) } )
                    }
                }
            }
            Else {
                if ($fileName -like "*.url") {
                    Write-Verbose "Returning all Urls"
                    $ShortCut = Get-Content $fileName | Where-Object { $_ -like "url=*" } 
                    Write-Output ( New-Object -Type PSObject -Property @{'FileName'= $FileName; 'Url' = $ShortCut.Substring($ShortCut.IndexOf("=") + 1 ) } )

                }  
       }          
    
}

}  