function Log-Message {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [string]$Path = "$EmailDirectory\Copy-MailFolder_$ScriptStartTime.log",

        [Parameter(Mandatory=$false, HelpMessage="Outputs the message to the console.")]
        [switch]$Output,

        [Parameter(Mandatory=$false, HelpMessage="Sets the color of the message.")]
        [ValidateSet("Information", "Success", "Warning", "Error")]
        [string]$MessageType = "Information",

        [Parameter(Mandatory=$false, HelpMessage="Logs the message to a file.")]
        [switch]$LogToFile
    )
    $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Message = "$Date >> $Message"
    if ($DebugMode -or $Output) {
        $ForegroundColor = switch ($MessageType) {
            "Success" { "Green" }
            "Warning" { "Yellow" }
            "Error" { "Red" }
            "Information" { "White" }
        }
        Write-Host $Message -ForegroundColor $ForegroundColor
    } 
    if ($LogToFile) {
        Add-Content -Path $Path -Value $Message
    }
}