$Computer = Read-Host "Enter 1 to see all programs on this computer, or enter a computer name to remotely check its installed programs"

if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $Computer -Quiet) {
	$SortBy = Read-Host "Sort programs? 1=Name 2=Install Date 3=Size (please type a number)"
	
	# Maximize the powershell window so you can (hopefully) see all the output.
	(Add-Type -MemberDefinition '[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru)::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 3) | Out-Null

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
	
	# I'd normally make this check first, but I only want to prompt for $SortBy once at the beggining of the script, but only if the remote computer is reachable.
	if (($Computer -eq $env:COMPUTERNAME) -or ($Computer -eq 1)) { 
		Get-Programs $SortBy
	}
	else {
		Invoke-Command -ComputerName $Computer -ScriptBlock ${function:Get-Programs} -ArgumentList $SortBy
	}
}
else {
	"Attempted ping to $Computer failed"
}

if ((Read-Host "`nEnter Y to go again") -eq "Y") {
	.$PSCommandPath # Start the script from the beginning.
}