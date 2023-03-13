# Get-InstalledApps
Retrieve all installed programs on a local or remote Windows computer.

The output is sent back to the host and laid out in the same old-style layout as the view in Control Panel > Programs and Features.

This script gets the same installed programs as reported by that view, as it checks the same NameSpace. This means certain malware or other "invisibly installed" programs that are not seen in Windows' own installed programs list won't be reported by this script either.

You might notice that the reported Size is rounded differently from the Windows view. In Windows, it's always rounded down. In this view, it's rounded up or down, whichever is closer. This is intentional, as I feel it is more accurate.
