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
     #   [Parameter(Mandatory=$True)]
      #  [String]$Link

    #    [String]$Path = 'c:\temp',
        
     #   [Parameter(Mandatory=$True)]
      #  [String]$Name
    )
    
    #if ( -Not $Path ) { $Path = Get-Location }

    #$Path
        
    Write-Verbose "Creating Shortcut for $L"
    #$WshShell = New-Object -comObject WScript.Shell
    #$Shortcut = $WshShell.CreateShortcut("$Path\$Name.url")
    #$Shortcut.TargetPath = $L
    #$Shortcut.Save()
     
}

New-Shortcut -Link 'http://stackoverflow.com/questions/34164110/powershell-create-new-shortcut-named-after-website' -Path c:\temp -Name 'Shortcut'