<# v3
- Added option to run again
- New version using .NET methods
- Update ForEach-Object to foreach
#>

<# Try this for faster comparison:
https://technet239.rssing.com/chan-4753999/article28253.html
#>

function printDate {
	Write-Host "===================" -ForegroundColor Red -NoNewline; Write-Host ([char]0xA0)
    # Write-Host " Date: $(Get-Date -UFormat "%d %b %Y") " -ForegroundColor DarkYellow -NoNewline; Write-Host ([char]0xA0)
    Write-Host " Time: $(Get-Date -f "HH:mm:ss") " -ForegroundColor DarkYellow -NoNewline; Write-Host ([char]0xA0)
	Write-Host "===================" -ForegroundColor Red -NoNewline; Write-Host ([char]0xA0)
}

# $cleanCard = {
function cleanCard {
	printDate

	# Add file extensions to this list as required:
	$extensions = ('.cr2','.cr3','.raf','.jpg','.orf','.mov','.mp4')

	Write-Host "`n Clean up Memory Card " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host ([char]0xA0) -NoNewline
	Write-Host "`n"

	# Default locations. Set a location or set as $null
	# Card:
	$uDefault = "M:"
	$promptC = If ($uDefault -ne $null) {"($uDefault)"} Else {"(not set)"}
	# Photos:
	$mDefault = "E:\Photos-01\2023"
	$promptM = If ($mDefault -ne $null) {"($mDefault)"} Else {"(not set)"}

	# Locate directories
	$u = Read-Host -Prompt "Enter path / drive letter for the Memory Card or leave blank to use default $promptC"
	If (-not ($u)){
		$u = $uDefault
	}
	If ($u -match '^([a-zA-Z])(?!\s:)$'){
		$u = $u + ':'
	}
	$usbDir = Get-Item -Path $u

	$m = Read-Host -Prompt "Enter path / drive letter for the Photos / Media folder on the computer or leave blank to use`ndefault $promptM"
	If (-not ($m)){
		$m = $mDefault
	}
	If ($m -match '^([a-zA-Z])(?!\s:)$'){
		$m = $m + ':'
	}
	$mediaDir = Get-Item -Path $m

	Write-Host "`n Scanning files... " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host ([char]0xA0)

	# Index files from folders, based on file extensions
	# Check folders are not empty (end script if either is empty)
	

	# Original version using cmdlets:
	# Measure-Command -Expression {
	# $usb = Get-ChildItem -Path $usbDir -Recurse | Where-Object {$_.Extension -in $extensions}
	# $media = Get-ChildItem -Path $mediaDir -Recurse | Where-Object {$_.Extension -in $extensions}
	# }
	
	# New version using .NET methods
	# Measure-Command -Expression {
	$a = foreach ($extension in $extensions){[System.IO.Directory]::EnumerateFiles($usbDir,"*$extension","AllDirectories")}
	$usb = foreach ($usbfile in $a){[System.IO.FileInfo]$usbfile}
	$b = foreach ($extension in $extensions){[System.IO.Directory]::EnumerateFiles($mediaDir,"*$extension","AllDirectories")}
	$media = foreach ($mediafile in $b){[System.IO.FileInfo]$mediafile}
	# }

	If (!$usb){
		Write-Host "No files found at specified USB location. Press enter to exit." -ForegroundColor Red
		Read-Host
		exit
	}
	If (!$media){
		Write-Host "No files found at specified local media location. Press enter to exit." -ForegroundColor Red
		Read-Host
		exit
	}

	# Compare indexed directories
	Write-Host "`n Comparing files... " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host ([char]0xA0)
	
	$notcopied = Compare-Object -ReferenceObject $usb -DifferenceObject $media -Property Name, Length -PassThru | Where-Object {$_.SideIndicator -eq "<="}
	$copied = $usb | Where-Object {$notcopied -notcontains $_}

	# Display comparison results
	# Calculate matches for already copied files:
	Write-Host "`n Locating matched files. This may take a while... " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host ([char]0xA0)
	printDate
	$start = Get-Date
	
	# Debug
	# Read-Host
	
	# Original method:
	# $copied | ForEach-Object {
	# 	$name = $_.Name -replace $_.Extension
	# 	$length = $_.Length
	# 	$localfile = ($media | Where-Object {($_.Name -like "*$name*") -and ($_.Length -match $length)}).FullName
	# 	Add-Member -InputObject $_ -MemberType NoteProperty -Name "MatchedFile" -Value $localfile
	# }

	<# Speed for above original:
	Minutes           : 13
	Seconds           : 22
	Milliseconds      : 562
	Ticks             : 8025620577
	TotalMinutes      : 13.376034295
	TotalSeconds      : 802.5620577
	TotalMilliseconds : 802562.0577
	#>

	# Rewrite using foreach instead of ForEach-Object:
	$i = 0
	foreach ($cfile in $copied) {
		$i++
		$j = [math]::round(100*($i / $copied.Count))
		Write-Progress -Activity "Working..." -Status " $j% complete" -PercentComplete $j
		$name = $cfile.Name -replace $cfile.Extension
		$length = $cfile.Length
		$MatchedFile = foreach ($mfile in $media) {If ($mfile.Name -match $name -and $mfile.Length -match $length){$mfile.FullName}}
		Add-Member -InputObject $cfile -MemberType NoteProperty -Name "MatchedFile" -Value $MatchedFile
	}
	Start-Sleep -Seconds 1
	<# Speed for above rewrite:
	Minutes           : 1
	Seconds           : 22
	Milliseconds      : 639
	Ticks             : 826393711
	TotalMinutes      : 1.37732285166667
	TotalSeconds      : 82.6393711
	TotalMilliseconds : 82639.3711
	#>

	
	# Create a custom size property to format file size in MB
	$size = @{label="Size(MB)";expression={"{0:n2} MB" -f ($_.length/1MB)}}

	# Print outputs
	# Write-Host "`n Files remaining to be copied to computer: " -ForegroundColor Black -BackgroundColor DarkYellow -NoNewline; Write-Host ([char]0xA0)
	# $notcopied | Select-Object -Property DirectoryName,Name,$size,LastWriteTime | Format-Table -AutoSize | Out-Host
	# Write-Host "`n Files to be deleted from the SD Card: " -ForegroundColor Black -BackgroundColor Red -NoNewline; Write-Host ([char]0xA0); Write-Host ([char]0xA0)
	# $copied | Select-Object -Property DirectoryName,Name,$size,MatchedFile | Format-Table -AutoSize | Out-Host
	
	
	# Print operation time:
	# Start-Sleep 3s
	$end = Get-Date
	$ts = New-TimeSpan -Start $start -End $end
	
	# Debugging:
	# $start = Get-Date
	# Start-Sleep -Milliseconds 5
	# $end = Get-Date
	# $ts = New-TimeSpan -Start $start -End $end

	switch ($ts)
	{
		{$_.Days -gt 0} {
			$days = $($_.Days.ToString()) + " days"
		}
		{$_.Hours -gt 0} {
			$hours = $($_.Hours.ToString()) + " hours"
		}
		{$_.Minutes -gt 0} {
			$mins = $($_.Minutes.ToString()) + " minutes"
		}
		{$_.Seconds -ge 1} {
			$secs = $($_.Seconds.ToString()) + " seconds"
		}
		{$_.TotalSeconds -lt 1} {
			$ltsecs = "less than a second"
		}
	}
	# Put in array and filter out empty items
	$duration = ($days, $hours, $mins, $secs, $ltsecs) | Where-Object {$_ -ne $null}
	Write-Output "End. Operation took $($duration -join ', ')."
	
	# Summary and confirmation to delete
	Write-Host "`n`tSummary" -ForegroundColor Green
	Write-Host "`t-------" -ForegroundColor Green
	Write-Host "`tNot yet copied:",$notcopied.Count
	Write-Host "`tAlready copied:",$copied.Count
	
	If ($notcopied.Count -gt 0){
		$viewRemain = Read-Host "`nEnter [v] to view $($notcopied.Count) files remaining to be copied to computer:"
		If ($viewRemain -eq "v") {
			$notcopied | Select-Object -Property DirectoryName,Name,$size,LastWriteTime | Out-GridView
		}
	}
	Write-Host ([char]0xA0)
	If ($copied.Count -gt 0){
		$viewDelete = Read-Host "Enter [v] to view $($copied.Count) files to be deleted from the SD Card:"
		If ($viewDelete -eq "v") {
			$copied | Select-Object -Property DirectoryName,Name,$size,MatchedFile | Out-GridView
		}
	}

	If ($copied.Count -gt 0){
		Write-Host "`nRemove $($copied.Count) files from the SD card that have already been copied to the computer? [Y/n]" -ForegroundColor Red
		$yn = Read-Host
		While ("Y","n" -cnotcontains $yn) {
			$yn = Read-Host "Try again. Type 'Y' or 'n' (case sensitive)"
		}

		# Check response and action
		If ($yn -clike 'Y'){
			try {
				Write-Host "`nDeleting..."
				$copied | Remove-Item -ErrorAction Stop
				Write-Host "$($copied.Count) files have been deleted from the memory card."
			}
			catch {
				Write-Host "`nThe following error occurred during file deletion. Check files manually."
				Write-Host $_
				Read-Host -Prompt "Press enter to close"
			}
		}
	} Else {
		Write-Host "`nNo files will be deleted."
		# Read-Host -Prompt "`t"
	}

	# Ask to copy new files to dated folders:
	If ($notcopied.Count -gt 0){
		Write-Host "`nCopy new files to dated folders? [Y/n]" -ForegroundColor Red
		$yn = Read-Host
		While ("Y","n" -cnotcontains $yn) {
			$yn = Read-Host "Try again. Type 'Y' or 'n' (case sensitive)"
		}
		
		If ($yn -clike 'Y'){
			$i = 0
			foreach ($file in $notcopied){
				$i++
				$j = [math]::round(100*($i / $notcopied.Count))
				Write-Progress -Activity "Copying files." -Status " $j% Complete." -PercentComplete $j
				# Create directory structure
				# $base = (Get-Item "E:\Photos-01").FullName
				$base = $mediaDir.FullName
				# $year = $file.CreationTime | Get-Date -Format "yyyy"
				$folder = $file.CreationTime | Get-Date -Format "yyyy-MM-dd -"
				# $copylocation = $base + "\" + $year + "\" + $folder + "\" + "Unsorted\"
				$copylocation = $base + "\" + $folder + "\" + "Unsorted\"
				
				## Debug:
				# $copylocation
				# Read-Host -Prompt 'Continue?'
				
				# Check directory exists and create if not
				If (!(Test-Path $copylocation)){
					$null = New-Item -ItemType Directory -Path $copylocation
				}

				# Copy to folder
				# Copy-Item $file -Destination $copylocation
				xcopy $($file.FullName) $copylocation | Out-Null
			}
			Start-Sleep -Seconds 1
		}
	}
}

Do {
	# &$cleanCard
	cleanCard
	$runAgain = Read-Host -Prompt "`nType [a] to run again, CTRL-C to quit"
} While ($runAgain -like "a")
