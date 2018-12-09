# Backend Functions 

Function New-MTPrinter {
	<#
		.SYNOPSIS
		Function for adding local or remote printers

		.DESCRIPTION
		Function for adding local or remote printers by utilizing multithreading and combining existing functions.

		.PARAMETER

		.EXAMPLE

		.NOTES
		This function utilizes a static value for multithreading when mapping the printer. 
	#>

	Param (
		[Parameter (
			Mandatory = $False,
			ParameterSetName = 'List'
		)]
		[System.String[]]$ComputerName = $env:COMPUTERNAME ,

		[Parameter (
			Mandatory = $False,
			ParameterSetName = 'Session'
		)]
		[System.Management.Automation.Runspaces.PSSession[]]$Session = $null,

		[Parameter (
			Mandatory = $True,
			HelpMessage = "Enter your administrator credentials.",
			ParameterSetName = 'List'
		)]
		[System.Management.Automation.CredentialAttribute()]$Credential,

		[Parameter (
			Mandatory = $True,
			HelpMessage = "Enter the name you would like for the new printer." 
		)]
		[System.String]$PrinterName,

		[Parameter (
			Mandatory = $True,
			HelpMessage = "Enter the path to the driver inf file."
		)]
		[System.String]$infPath,

		[Parameter (
			Mandatory = $True,
			HelpMessage = "Enter the IP Address of the remote printer"
		)]
		[System.String]$IPAddress 
	)

	BEGIN {
		$localErrors = [System.Collections.ArrayList]::New()
		$CompletedData = [System.Collections.ArrayList]::New()

		$SessionScript = {
			$Session = [System.Management.Automation.Runspaces.PSSession[]]::New() 
			Write-Debug -Message "Attempting to add Computername list into PSSession"
			$ComputerName | Foreach {
				TRY {
					$Session += New-PSSession -ComputerName $_ -Credential $Credential -ErrorAction Stop 
				} CATCH {
					[void]$localErrors.Add($_)
				}
			}
			Write-Output $Session 
		}

		IF ($Session -eq $null) {
			IF ($ComputerName.Count -eq 1) {
				IF ($ComputerName -eq $env:ComputerName) {
					Start-Sleep -Milliseconds 10 
				} ELSE {
					$Session = $SessionScript.Invoke() 
				} 
			} ELSE {
				$Session = $SessionScript.Invoke() 
			}
		}
	}

	PROCESS {
		Write-Debug -Message "Attempting to add printer."
		TRY {
			$DriverName = (Get-WindowsDriver -Driver $infPath -Online)[0].HardwareDescription
			$PortName = "IP_" + $IPAddress 
			$Script = {
				$RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 2)
				$RunspacePool.Open() 
				$Threads = [System.Collections.ArrayList]::New()

				$ScriptErrors = [System.Collections.ArrayList]::New()
				$ScriptData = [System.Collections.ArrayList]::New()

				$DriverPS = [powershell]::Create() 
				$DriverPS.RunspacePool = $RunspacePool
				[void]$DriverPS.AddScript({
					Param (
						[System.String]$infPath,
						[System.String]$DriverName 
					)
					$ScriptData = [System.Collections.ArrayList]::New()
					$ScriptErrors = [System.Collections.ArrayList]::New() 
					TRY {
						Write-Debug -Message "Attempting to map driver"
						#$DriverName = (Get-WindowsDriver -Driver $infPath -Online)[0].HardwareDescription 
						pnputil.exe -a $infPath
						$DriverMap = (Add-PrinterDriver -Name $DriverName -ErrorAction Stop)
						[void]$ScriptData.Add($DriverMap)
					} CATCH {
						Write-Verbose -Message "Failed adding driver"
						[void]$ScriptErrors.add($DriverMap)
					}
					Write-Output $ScriptData
					Write-Output $ScriptErrors 
				})
				$DriverParameters = @{
					infPath = $infPath
					DriverName = $DriverName
				}
				[void]$DriverPS.AddParameters($DriverParameters)
				$DriverHandle = $DriverPS.BeginInvoke()
				$DriverContainer = [System.String]::Empty
				$DriverContainer | Add-Member -MemberType NoteProperty -Name Powershell -Value $null
				$DriverContainer | Add-Member -MemberType NoteProperty -Name Handle -Value $null
				$DriverContainer = $DriverContainer | Select-Object -Property Powershell, Handle
				$DriverContainer.Powershell = $DriverPS
				$DriverContainer.Handle = $DriverHandle 
				[void]$Threads.Add($DriverContainer) 

				$PortPS = [powershell]::Create() 
				$PortPS.RunspacePool = $RunspacePool 
				[void]$PortPS.AddScript({
					Param (
						[System.String]$IPAddress,
						[System.String]$PortName 
					)
					$ScriptData = [System.Collections.ArrayList]::New()
					$ScriptErrors = [System.Collections.ArrayList]::New() 
					TRY {
						Write-Debug -Message "Attempting to create printer port"
						#$PortName = "IP_" + $IPAddress
						$PortMap = (Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -ErrorAction Stop)
						[void]$ScriptData.Add($PortMap)
					} CATCH {
						Write-Verbose -Message "Failed mapping printer port."
						[void]$ScriptErrors.Add($PortMap)
					}
					Write-Output $ScriptData
					Write-Output $ScriptErrors 
				})
				$PortParameters = @{
					IPAddress = $IPAddress
					PortName = $PortName
				}
				[void]$PortPS.AddParameters($PortParameters)
				$PortHandle = $PortPS.BeginInvoke() 
				$PortContainer = [System.String]::Empty
				$PortContainer | Add-Member -MemberType NoteProperty -Name Powershell -Value $null
				$PortContainer | Add-Member -MemberType NoteProperty -Name Handle -Value $null 
				$PortContainer = $PortContainer | Select-Object -Property Powershell, Handle
				$PortContainer.Powershell = $PortPS
				$PortContainer.Handle = $PortHandle 
				[void]$Threads.Add($PortContainer) 

				DO {
					$Threads | Foreach {
						IF ($Threads.Handle.IsCompleted -eq $True) {
							$PSID = $_.Powershell.InstanceID.Guid 
							[void]$ScriptData.Add($_.Powershell.EndInvoke($_.Handle))
							$_.Powershell.Dispose()
							$Threads = $Threads | Where-Object {$_.Powershell.InstanceID.Guid -ne $PSID} 
						}
					}
				} until ($Threads -eq $null) 
				$RunspacePool.Close()
				$RunspacePool.Dispose() 

				TRY {
					Write-Debug -Message "Attempting to setup remote printer."
					$Printer = (Add-Printer -DriverName $DriverName -Name $PrinterName -PortName $PortName -ErrorAction Stop)
					[void]$ScriptData.add($Printer)
				} CATCH {
					Write-Verbose -Message "Failed adding printer"
				}

				Write-Output $ScriptData
				Write-Output $ScriptErrors 
			}

			IF ($Session -eq $null) {
				IF ($ComputerName.Count -eq 1) {
					IF ($ComputerName -eq $env:COMPUTERNAME) {
						TRY {
							Write-Debug -Message "Attempting to add printer to local computer"
							$RunScript = $Script.Invoke()
							[void]$CompletedData.Add($RunScript) 
						} CATCH {
							Write-Verbose -Message "Unable to run script locally" 
							[void]$localErrors.Add($RunScript) 
						}
					} 
				} 
			} ELSE {
				TRY {
					Write-Debug -Message "Attempting to Invoke Commands remotely"
					$InvokeData = Invoke-Command -Session $Session -ScriptBlock {
						$Script = $using:Script 
						$PrinterName = $using:PrinterName
						$infPath = $using:infPath
						$IPAddress = $using:IPAddress
						$RemoteData = [System.Collections.ArrayList]::New()
						$RemoteErrors = [System.Collections.ArrayList]::New() 

						TRY {
							$RunScript = $Script.Invoke() 
							[void]$RemoteData.Add($RunScript) 
						} CATCH {
							Write-Verbose -Message "Could not run script remotely"
							[void]$RemoteErrors.Add($RunScript) 
						}

						Write-Output $RemoteData
						Write-Output $RemoteErrors 
					} -ErrorAction Stop 
					[void]$CompletedData.Add($InvokeData)
				} CATCH {
					Write-Verbose -Message "Could not invoke commands" 
					[void]$localErrors.Add($InvokeData) 
				}
			}
		} CATCH {
			Write-Verbose -Message "Could not add printer." 
		}
	}

	END {
		Write-Output $CompletedData
		Write-Output $localErrors
	}
}
