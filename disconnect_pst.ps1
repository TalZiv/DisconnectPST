Clear
Write-Host "PST Disconnect - v0.1 - Tal Ziv"
Sleep -Seconds 2
try{
	$Test = Add-Type -AssemblyName microsoft.office.interop.outlook
	$outlook = New-Object -ComObject outlook.application
	$namespace = $outlook.GetNamespace('MAPI')
}catch{
	Write-Host "There was a problem Connecting to Outlook..."
	sleep -Seconds 5
	exit
	}
 
try {
     foreach ($Store in $namespace.Stores){
           $CheckMe = $Store.FilePath
           $PSTRootFolder = $store.GetRootFolder()
           if ($CheckMe -like "*pst" ){
                Write-Host "Disconnecting: $CheckMe"
                $PSTFolder = $namespace.Folders.item($PSTRootFolder.name)
                $namespace.GetType().InvokeMember('Removestore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$namespace,($PSTFolder))
           }
     }
}Catch{
}
Write-Host "Closing Outlook..."
$outlook.Quit()
Write-Host "Starting Outlook in 10 seconds..."
sleep -Seconds 10
start outlook.exe