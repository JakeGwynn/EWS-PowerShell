# MyModule.psm1
# This file is used to import all .ps1 files in the directory and export all functions to the module


# Get the directory of the current script
$FunctionPath = $PSScriptRoot + "\functions\"

write-host $FunctionPath 

# Get all .ps1 files in the directory
$scriptFiles = Get-ChildItem -Path $FunctionPath -Filter *.ps1

# Dot source each file
foreach ($file in $scriptFiles) {
    . $file.FullName
}

$DllPath = $PSScriptRoot + "\Microsoft.Exchange.WebServices.dll"
Add-Type -Path $DllPath

# Export all functions to the module
Export-ModuleMember -Function *