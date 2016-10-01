# ------------------------------------------------------------------------
# Move-shortcutsfromdesktop
#
# moves the shortcuts on the desktop to the correct dated folder
#-------------------------------------------------------------------------

Param (
    [String]$UserDesktop
)

#-------------------------------------------------------------------------
# Main
#-------------------------------------------------------------------------

# ----- Check if the P: drive exists.  If not then do not process...
if (  (Test-Path -Path "p:\links")  -eq $False ) { exit }

# ----- Get year and date
$Date = (Get-Date -Format "yyyy - MMM").ToUpper()

# ----- Check if folder exists
if ( Get-ChildItem -Path "p:\links" -Filter $Date ) {
		# ----- Do Nothing
	}
	else {
		md "p:\links\$Date"
}

# ----- Move .lnk files
Get-Item -path "$UserDesktop\*.url" | foreach {
	Write-Host "p:\links\$Date\$($_.Name)"
	if ( Test-Path "p:\links\$Date\$($_.Name)" ) {
#			Write-Host $_.name -ForegroundColor red
			$_ | Rename-Item -NewName "$($_.BaseName)($(Get-Random -Maximum 100))$($_.Extension)" -PassThru | Move-Item -Destination "p:\links\$Date"
		}
		else {
#			Write-Host $_.Name -ForegroundColor green
			$_ | Move-Item -Destination "p:\links\$Date"
	}
}