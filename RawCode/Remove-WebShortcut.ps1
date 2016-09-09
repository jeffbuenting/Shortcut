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





$RootPath = 'p:\links\2009DEC'


get-childitem -path $RootPath | Resolve-ShortcutFile -LikeString 'fanhole' | Select-Object @{n='Path'; e={$_.FileName}} # | remove-item