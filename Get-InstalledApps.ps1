function Get-Programs {
	param(
		[string]$Sortby
	)
	function Convert-Bytes {
		[CmdletBinding()]
		param(
			[Parameter(Mandatory)]
			[string]$bytes
		)
		switch ($bytes.length) {
			{$_ -le 3} {""}
			{($_ -ge 4) -and ($_ -le 6)} {[string]("{0:g3}" -f ($bytes/[Math]::Pow(1024,1))) + " KB"}
			{($_ -ge 7) -and ($_ -le 9)} {[string]("{0:g3}" -f ($bytes/[Math]::Pow(1024,2))) + " MB"}
			{($_ -ge 10) -and ($_ -le 12)} {[string]("{0:g3}" -f ($bytes/[Math]::Pow(1024,3))) + " GB"}
			{($_ -ge 13) -and ($_ -le 15)} {[string]("{0:g3}" -f ($bytes/[Math]::Pow(1024,4))) + " TB"}
			{$_ -ge 16} {" >=1 PiB!!"}
		}
	}

	# Installed programs as reported by Windows itself.
	$ShellPrograms = ((New-Object -ComObject Shell.Application).NameSpace("::{26EE0668-A00A-44D7-9371-BEB064C98683}\8\"+"::{7B81BE6A-CE2B-4676-A29E-EB907A5126C5}")).Items()

	$ProgramList = foreach ($Program in $ShellPrograms) {
		$Program | Select-Object @{
			name='Name'
			expr={$_.ExtendedProperty("Name")}
		},@{
			name='Publisher'
			expr={$_.ExtendedProperty("System.Software.Publisher")}
		},@{
			name='Installed On'
			expr={($_.ExtendedProperty("System.Software.DateInstalled")).ToShortDateString()}
		},@{
			name='Size'
			expr={Convert-Bytes $_.ExtendedProperty("Size")}
		},@{
			name='SizeInBytes'
			expr={$_.ExtendedProperty("Size")}
		},@{
			name='Version'
			expr={$_.ExtendedProperty("System.Software.ProductVersion")}
		}
	}
	if ($SortBy -eq 2) {
		$ProgramList | Sort-Object {$_.'Installed On' -as [datetime]} -Descending | Select-Object -Property * -ExcludeProperty SizeInBytes | Format-Table
	}
	elseif ($SortBy -eq 3) {
		$ProgramList | Sort-Object SizeInBytes -Descending | Select-Object -Property * -ExcludeProperty SizeInBytes | Format-Table
	}
	else { # If input is 1 or was mistyped as something else entirely
		$ProgramList | Sort-Object Name | Select-Object -Property * -ExcludeProperty SizeInBytes | Format-Table
	}
}

$Computer = Read-Host "Press Enter see programs on this computer, or enter a computer name to remotely check its installed programs"

if (($Computer) -and ($Computer -ne $env:COMPUTERNAME)){
	if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $Computer -Quiet) {
		$Remote = "Online"
	}
}
else {
	$Local = $True
}

if ($Local -or $Remote) {
	$SortBy = Read-Host "Sort programs? 1=Name 2=Install Date 3=Size (please type a number)"

	# Maximize the powershell window so you can (hopefully) see all the output.
	(Add-Type -MemberDefinition '[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru)::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 3) | Out-Null

	"Working...`n"

	if ($Local) {
		Get-Programs $SortBy
	}
	else {
		# Running commands on remote machines will cache your user account on those machines (if it doesn't already exist). 
		# If you want to avoid this, run the script as a user account that is already on the target machine, or prompt for the correct credentials and add -Credential to Invoke-Command.
		Invoke-Command -ComputerName $Computer -ScriptBlock ${function:Get-Programs} -ArgumentList $SortBy
	}
}
else {
	"Attempted ping to $Computer failed"
}

# Offer to restart the script.
# https://github.com/JedHarb/Restart-Powershell-Script/blob/main/Restart-PSScript.ps1
if ((Read-Host "`nEnter Y to restart this script") -eq "Y") {
	# Reset most of the local automatic variables that started with powershell back to their initial values (some are read-only).
	try {
		((& powershell "Get-Variable") | Select-Object -Skip 3 | ConvertFrom-String -PropertyNames Name, Value).ForEach({
			Set-Variable -Name $_.Name -Value $_.Value -ErrorAction SilentlyContinue
		})
	}
	catch {}

	# Remove all additional variables created in this session.
	try {
		Remove-Variable -Name (Compare-Object (Get-Variable) ((& powershell "Get-Variable") | ConvertFrom-String -PropertyNames Name) -Property Name | Where-Object SideIndicator -eq "<=").Name -ErrorAction SilentlyContinue
	}
	catch {}

	# Reset the last few stragglers
	# $Error.Clear() # Every once in awhile, this throws a "Method invocation failed because [System.String] does not contain a method named 'Clear'." and I haven't been able to pin down why.
	$$, $StackTrace = ""

	# The automatic variable $^ can't be manually removed, reset, or changed in any way (at least in all of my testing).
	# It will become equal to the literal text 'try' at this point, and change with each command run from here (as usual).
	
	# Start the script from the beginning.
	.$PSCommandPath
}
