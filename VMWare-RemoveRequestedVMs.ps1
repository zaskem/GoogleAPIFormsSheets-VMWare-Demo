# Google API scopes and variables:
$invocationPath = (Get-Item -Path ".\" -Verbose).FullName + "\"
$certFile = "local\path\to.p12"
$certPath = $invocationPath + $certFile
$iss = "your-api-service-name@somethingorother.iam.gserviceaccount.com"
$certPswd = "notasecret"
$scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive"
$vmwareSheetName = "VMWare-Actions"
$vmwareSheetID = ""
$vmwareVMSheetName = "RemoveVM"
$vmwareVMAvailableSheetName = "VMsAvailable"

# VMWare Infrastructure variables
$vmwareServer = "host.com"
$vmwareRootContainer = "Safe Sandbox"
$vmwareSrvAcct = "svcaccount-example"
$vmwareSrvCredPath = $invocationPath + "local\path\to\" + $vmwareSrvAcct + ".txt"
$vmwareSrvPwrdESSS = Get-Content $vmwareSrvCredPath | ConvertTo-SecureString

# Create the Google access token and get the ID of the action sheet
$token = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPswd -iss $iss
$vmwareSheetID = Get-GSheetSpreadSheetID -fileName $vmwareSheetName -accessToken $token
# Obtain list of the action items in scope (grabbing all data)
Try {
    $dataToProcess = $true
    $vmData = Get-GSheetData -spreadsheetID $vmwareSheetID -accessToken $token -sheetName $vmwareVMSheetName -cell 'AllData'
}
Catch {
    # If no records beyond the header were found, the API call will throw an error
    $dataToProcess = $false
}
if ($dataToProcess -eq $true) {
    # Look for unprocessed records and stage for action (This should be wrapped into a foreach for N>1 rows)
    $vmToUse = $vmData.'VM Name'

    # Connect to the VMWare infrastructure:
    Connect-VIServer $vmwareServer -User $vmwareSrvAcct -Password ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($vmwareSrvPwrdESSS)))

    # Set some variables based on Known Things (from above) now we're connected to VMWare
    $VM = Get-VM -Name $vmToUse
    # Remove a VM template from the source VM, then spin up a clone from the template
    #   Should really add a sanity check to make sure the submitted template name still exists (cross-check with available list)
    # Power off VM before removing -- THIS IS REQUIRED FOR REMOVAL TO SUCCEED
    if ($VM.PowerState -eq "PoweredOn") {
        Stop-VM -VM $VM -Confirm:$false
    }
    Remove-VM -VM $VM -DeletePermanently:$true -Confirm:$false
    # List all Templates Available in VMWare
    $vms = Get-VM -Location $vmwareRootContainer

    # Gracefully exit VMWare
    Disconnect-VIServer -Confirm:$false

    # Process the list of Templates
    $vmProperties = ('Name', 'PowerState')
    $vmArray = New-Object -TypeName 'System.Collections.ArrayList'
    $vmArray.Add(@($vmProperties))
    $vms | ForEach-Object {
        $data = $_
        $array = $vmProperties | ForEach-Object {$data.$_}
        $vmArray.Add(@($array))
    }

    # Write These to the VM Sheet (overwrite, after clear since rowcount--), then reset the input sheet appropriately
    Clear-GSheetSheet -spreadSheetID $vmwareSheetID -accessToken $token -sheetName $vmwareVMAvailableSheetName
    Set-GSheetData -spreadsheetID $vmwareSheetID -accessToken $token -sheetName $vmwareVMAvailableSheetName -rangeA1 'A1' -values $vmArray
    Clear-GSheetSheet -spreadSheetID $vmwareSheetID -accessToken $token -sheetName $vmwareVMSheetName
    # Write a new header row (and "null" first data row)
    Set-GSheetData -spreadSheetID $vmwareSheetID -accessToken $token -sheetName $vmwareVMSheetName -rangeA1 'A1' -values @(@("Timestamp", "VM Name"), @())
}