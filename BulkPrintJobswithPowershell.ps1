
Write-Host '   ------------Bulk Printjob with Powershell------------'
Write-Host
Write-Host 'Please select the file to be printed once the Open file dialog opens..' -ForegroundColor Yellow

$FileFullPath         = $null
$PrinterSharePath     = '\\printserver\FollowMePrint'
$PrinterName          = ($PrinterSharePath -split '\\')[-1]
$CriticalErrorMessage = 'It seems we are facing a critical error, we are exiting the script, kindly inform service desk.'
 
sleep 5
Function Get-FileName($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = 'All files (*.*)| *.*'
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function Get-FileName

$FileFullPath = Get-FileName

function Get-UserToRetrySelectFile{
	$wshell = New-Object -ComObject Wscript.Shell
    $answer = $wshell.Popup('It seems there has been a problem selecting the file, Do you want to retry ?',0,'Retry',64+4)
	If($answer -eq 6){
		$FileFullPath = Get-FileName
		If($FileFullPath.length -eq 0){
            $answer = $wshell.Popup($CriticalErrorMessage)
            $FileFullPath = 'Error'
        }
	}
    Else{
        $FileFullPath = 'NoFileSelected'
    }
    return $FileFullPath
}

If($FileFullPath.length -eq 0){
    $FileFullPath = Get-UserToRetrySelectFile
    If($FileFullPath -eq 'Error'){
		sleep 10
        exit 1
    }
    ElseIf($FileFullPath -eq 'NoFileSelected'){
		Sleep 10
        exit 0
    }
    Else{
        # Everything is All right
    }
}

$wshell = New-Object -ComObject Wscript.Shell
$wshell.AppActivate('BulkPrintJobswithPowershell')|Out-Null

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$title = 'RE-PRINT COUNT'
$msg   = 'Enter the number of times you want to re-print the selected document'
$count = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

If($count.Length -eq 0){
	$answer = $wshell.Popup($CriticalErrorMessage)
	Sleep 10
	exit 1
}
Else{

	Try{
		Write-Host "We are trying to connect to `" $PrinterName `" printer, kindly wait for few seconds." -ForegroundColor Yellow
		$printers = Get-CimInstance -Class Win32_Printer |Where-Object -FilterScript {$_.Name -eq $PrinterSharePath} 
		If($null -eq $printers){
			$WSHNewPrinterCreation=(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($PrinterSharePath)
			sleep 3
		}
		$printers = Get-CimInstance -Class Win32_Printer |Where-Object -FilterScript {$_.Name -eq $PrinterSharePath} 
		If($null -eq $printers){
			Write-Host "Failed to connect to `"$($PrinterName)`" printer" -ForegroundColor Red
			Write-Host $CriticalErrorMessage -ForegroundColor Red
			Sleep 10
			exit 1
		}
		Else{
			Write-Host "`"FollowMePrint`" printer found `, Sending a test print to printer.." -ForegroundColor White
			$wshell.AppActivate('Windows PowerShell')|Out-Null
			Get-item -LiteralPath $FileFullPath | Out-Printer -Name $PrinterSharePath
		}
	}
	Catch{
		Write-Host "Failed to connect to `"$($PrinterName)`" printer, ErrorMessage : $($_.Exception.Message)" -ForegroundColor Red
		Write-Host 'Failed to execute ' -ForegroundColor Red -nonewLine
		Write-Host "$($_.InvocationInfo.Line) "
		Write-Host $CriticalErrorMessage -ForegroundColor Red
		sleep 10
		exit 1
	}
	Finally{
		
	}

	Write-Host ' '
	Write-Host -nonewLine 'Everything seems ok, we are going ahead with printing the selected document [ '
	Write-Host (Get-Item -literalPath $FileFullPath).Name -nonewLine -f yellow 
	Write-Host -nonewLine ' ] '
	Write-Host "$count" -nonewLine -f Green 
	Write-Host ' times! ' 
	Write-Host 'Press Enter to proceed' -f Red -nonewLine
	Read-Host ' '
	Write-Host 
	Try{
		for ($i = 1; $i -le $count; $i++ )
		{
			$Progress = [Int](($i/$count)*100)
			Write-Progress -Activity 'Printing in Progress' -Status $Progress'% Complete:' -PercentComplete $Progress
			$StartTime = Get-Date
			Get-item -LiteralPath $FileFullPath | Out-Printer -Name $PrinterSharePath
			$EndTime = Get-Date
			Start-Sleep -seconds ($EndTime - $StartTime).TotalSeconds
		}
		Write-Host 'Successfully printed document,exiting the script in few moments.'
		Sleep 10
		exit 1
	}
	Catch{
		Write-Host "Failed to print, ErrorMessage : $($_.Exception.Message)" -ForegroundColor Red
		Write-Host 'Failed to execute ' -ForegroundColor Red -nonewLine
		Write-Host "$($_.InvocationInfo.Line) "
		Write-Host $CriticalErrorMessage -ForegroundColor Red
		sleep 100
		exit 1
	}
}
