# Define your regex pattern (case-insensitive)
$pattern = "pattern"  # Change this as needed

# Define the mailbox to search — can be an email or display name
$mailboxName = "example@example.com"  # ← CHANGE THIS

# Start Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Try to resolve the mailbox name
$recipient = $Namespace.CreateRecipient($mailboxName)
$recipient.Resolve()

if (-not $recipient.Resolved) {
    Write-Error "Mailbox '$mailboxName' could not be resolved. Check the name or email."
    exit
}

# Get root folder of the shared mailbox
try {
    $RootFolder = $Namespace.GetSharedDefaultFolder($recipient, [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Parent
    Write-Host "Resolved and accessing mailbox: '$mailboxName'"
} catch {
    Write-Error "Failed to access shared mailbox root for '$mailboxName': $_"
    exit
}

Write-Host "Searching ALL folders under '$mailboxName' for pattern: '$pattern' in subject or body (case-insensitive)...`n"

# Recursive search function
function Search-FolderRecursively {
    param (
        $Folder
    )

    try {
        $items = $Folder.Items
    } catch {
        Write-Warning "Cannot access items in folder: $($Folder.Name). Skipping."
        return
    }

    foreach ($item in $items) {
        if ($item -is [__ComObject] -and $item.MessageClass -like "IPM.Note*") {
            $subject = $item.Subject
            $body = $item.Body

            if (
                ($null -ne $subject -and $subject -match "(?i)$pattern") -or
                ($null -ne $body -and $body -match "(?i)$pattern")
            ) {
                Write-Host "Match found:"
                Write-Host "  Folder  : $($Folder.FolderPath)"
                Write-Host "  Subject : $($item.Subject)"
                Write-Host "  Received: $($item.ReceivedTime)"
                Write-Host "  Sender  : $($item.SenderName)`n"
            }
        }
    }

    foreach ($subFolder in $Folder.Folders) {
        Search-FolderRecursively -Folder $subFolder
    }
}

# Kick it off
Search-FolderRecursively -Folder $RootFolder
