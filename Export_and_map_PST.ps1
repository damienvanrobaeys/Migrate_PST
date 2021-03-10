[CmdletBinding()]
Param(
  [Parameter(Mandatory=$true)] 
  [string]$PST_New_Path 
  )  
 
Function Close_Outlook
	{
		$Outlook_Process = Get-Process outlook -ErrorAction SilentlyContinue
		$Outlook_Process_ID = $Outlook_Process.id
		If ($Outlook_Process) 
			{
				Try
					{						
						$Outlook_Process | Stop-Process -Force
						$Outlook_Process.WaitForExit()
						write-host "Closing Outlook: OK"
					}
				Catch
					{
						write-host "Closing Outlook: KO"
						Break
					}
			} 
	}
  
  
Function Export_PST
	{ 
		If(!(test-path $PST_New_Path))
			{
				write-warning "The PST folder you have typed does not exist. Please select an existing folder"
				Break
			}

		$PST_Folders = $PST_New_Path
		Close_Outlook

		# List PST files
		write-host "Step 1 - Listing PST files: OK"
		$outlook = New-Object -comobject Outlook.Application 
		$User_All_PST_Files = $outlook.Session.Stores | where {($_.FilePath -like '*.PST')} 
		If($User_All_PST_Files -eq $null)
			{
				write-warning "No PST files opened in Outlook"
				break
			}

		# Export PST properties in an XML file
		$List_PST_To_Copy_Path = "$PST_Folders\pst_files.xml"
		$User_All_PST_Files | select displayname, filepath | export-clixml $List_PST_To_Copy_Path
		write-host "Step 2 - Exporting PST files name: OK"

		# Disconnect PST files
		write-host "Step 3 - Trying to disconnect PST files"
		$List_All_PST = import-clixml $List_PST_To_Copy_Path
		$namespace = $outlook.getnamespace("MAPI")
		ForEach ($pst in $List_All_PST)
			{
				Try
					{
						$PSTPath = $pst.filepath
						$PSTDisplayName = $pst.displayname
						$namespace.AddStore($PSTPath) 
						$PST = $namespace.Stores | ? {$_.FilePath -eq $PSTPath} 
						$PSTRoot = $PST.GetRootFolder() 
						$PSTFolder = $Namespace.Folders.Item($PSTDisplayName) 
						$Namespace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$Namespace,($PSTFolder))

						write-host "Step 4 - PST $PSTDisplayName has been disconnected"
					}
				Catch 
					{
						write-warning "Step 4 - PST $PSTDisplayName can not be disconnected"
					}
			}

		Close_Outlook

		# Copy PST files in the new location
		write-host "Step 5 - Copying PST files to the new location"
		ForEach ($pst in $List_All_PST)
			{
				Try
					{
						$PSTPath = $pst.filepath
						copy-item $PSTPath $PST_Folders -force
						write-host "Step 6 - $PSTPath has been copied to $PST_Folders"
					}
				Catch
					{
						write-warning "Step 6 - $PSTPath has not been copied"
					}
			}

		# Remapp PST files from the new location
		Close_Outlook
		write-host "Step 7 - PST files mapping from the new location"
		$outlook = New-Object -comObject Outlook.Application 
		$namespace = $outlook.getnamespace("MAPI")
		$Get_Temp_Outlook_Folder = Get-Childitem $PST_Folders | Where-Object {$_.Extension -eq ".pst"}
		ForEach ($New_PST in $Get_Temp_Outlook_Folder)
			{ 
				Try
					{
						$PSTPath = "$PST_Folders\$New_PST"
						$namespace.AddStore($PSTPath)
						write-host "Step 8 - $PSTPath has been mapped to Outlook"
					}
				Catch
					{
						write-warning "Step 8 - $PSTPath has not been mapped to Outlook"
					}

			}

			Close_Outlook 
			Remove-item $List_PST_To_Copy_Path -Force
	}
  
Export_PST -PST_Folders $PST_New_Path