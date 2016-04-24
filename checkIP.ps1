Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension
}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName

#define "i" for progress bar
$i = 0
$data = @()

$ErrorActionPreference= 'silentlycontinue'


foreach($server in $serverList){
	#empty variables in case previous loop errored out
	$IP = ''
	$Error = ''
	
	#determine if it is online
	$pingIt = Test-Connection -ComputerName $server -quiet -count 1
	#get the nbtstat info for the server, using the IP
	try {
	#get the IP
	
	$IP = Test-Connection -ComputerName $server -count 1 | select ipv4address
	$IP = $IP.IPV4Address.IPAddressToString
		if($pingIt -eq "True"){
			$returnedName = nbtstat -a $IP | ?{$_ -match '\<00\>  UNIQUE'} | %{$_.SubString(4,14)}
			$returnedName = $returnedName.trim()
			if($server -eq $returnedName){
				$props = [ordered]@{
				'Server' = $server
				'IP' = $IP
				'Status' = 'DNS Name Matches'
				'Details' = ''
				}
			}
			else {
				$props = [ordered]@{
				'Server' = $server
				'IP' = $IP
				'Status' = 'DNS Name does not match'
				'Details' = 'DNS returned ' + $returnedName
				}	
			}
		}
		Else{
			$props = [ordered]@{
			'Server' = $server
			'IP' = $IP
			'Status' = 'Ping failed'
			'Details' = ''	
			}
		}

	}
	Catch {
		$ErrorMessage = $_.Exception.Message
		if($ErrorMessage -eq "You cannot call a method on a null-valued expression."){
			$props = [ordered]@{
			'Server' = $server
			'IP' = $IP
			'Status' = 'Error'
			'Details' = 'No DNS for this IP'
			}
		}
		else{
			$props = [ordered]@{
			'Server' = $server
			'IP' = $IP
			'Status' = 'Error'
			'Details' = $ErrorMessage
			}
		}
		
	}
	$obj = New-Object -TypeName PSObject -Property $props
	$data += $obj
	#$data | Where-Object {$_} | Export-Csv $filename -noTypeInformation -append

	$i++
	Write-Progress -activity "Validating server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)
}
$data | Where-Object {$_} | Export-Csv $filename -noTypeInformation -append
