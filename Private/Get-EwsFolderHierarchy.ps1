function Get-EwsFolderHierarchy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Service,

        [Parameter(Mandatory = $true)]
        [string]$PrimarySmtpAddress,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 1000)]
        [int]$PageSize,

        [ValidateRange(0, 10)]
        [int]$RetryCount = 3,

        [switch]$IncludeSubfolders,

        [switch]$IncludeArchive,

        [switch]$SearchDumpster
    )

    $roots = New-Object 'System.Collections.Generic.List[object]'
    $mailbox = New-Object -TypeName Microsoft.Exchange.WebServices.Data.Mailbox -ArgumentList $PrimarySmtpAddress
    $primaryRoot = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId -ArgumentList @(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,
        $mailbox
    )
    $roots.Add([pscustomobject]@{ FolderId = $primaryRoot; Path = ''; Source = 'Primary' })

    if ($IncludeArchive) {
        try {
            $archiveName = [Enum]::Parse([Microsoft.Exchange.WebServices.Data.WellKnownFolderName], 'ArchiveMsgFolderRoot', $true)
            $archiveRoot = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId -ArgumentList @($archiveName, $mailbox)
            $roots.Add([pscustomobject]@{ FolderId = $archiveRoot; Path = 'Archive'; Source = 'Archive' })
        }
        catch {
            throw 'The loaded EWS Managed API does not support the ArchiveMsgFolderRoot folder. Verify the Exchange/EWS version before requesting archive search.'
        }
    }

    if ($SearchDumpster) {
        try {
            $dumpsterName = [Enum]::Parse([Microsoft.Exchange.WebServices.Data.WellKnownFolderName], 'RecoverableItemsRoot', $true)
            $dumpsterRoot = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId -ArgumentList @($dumpsterName, $mailbox)
            $roots.Add([pscustomobject]@{ FolderId = $dumpsterRoot; Path = 'Recoverable Items'; Source = 'Dumpster' })
        }
        catch {
            throw 'The loaded EWS Managed API does not support the RecoverableItemsRoot folder. Verify the Exchange/EWS version before requesting dumpster search.'
        }
    }

    $queue = New-Object 'System.Collections.Generic.Queue[object]'
    foreach ($root in $roots) {
        try {
            $rootFolder = Invoke-EwsRequestWithRetry -ScriptBlock {
                [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $root.FolderId)
            } -RetryCount $RetryCount
            $queue.Enqueue([pscustomobject]@{ FolderId = $root.FolderId; Path = $root.Path; Source = $root.Source; IsRoot = $true; Folder = $rootFolder })
        }
        catch {
            [pscustomobject]@{
                PSTypeName   = 'ExchangeNamedProperties.SearchError'
                FolderPath   = $root.Path
                ErrorCode    = $_.Exception.GetType().Name
                ErrorMessage = $_.Exception.Message
            }
        }
    }

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        if (-not $current.IsRoot) {
            if ($current.Folder -is [Microsoft.Exchange.WebServices.Data.SearchFolder]) {
                continue
            }
            [pscustomobject]@{
                Folder     = $current.Folder
                FolderId   = $current.Folder.Id
                FolderPath = $current.Path
                Source     = $current.Source
            }
        }

        if (-not $IncludeSubfolders -and -not $current.IsRoot) {
            continue
        }

        $offset = 0
        do {
            $view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList $PageSize
            $view.Offset = $offset
            $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
            $view.PropertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            $view.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)

            try {
                $folderPage = Invoke-EwsRequestWithRetry -ScriptBlock {
                    Write-Output -NoEnumerate ($Service.FindFolders($current.FolderId, $view))
                } -RetryCount $RetryCount
            }
            catch {
                [pscustomobject]@{
                    PSTypeName   = 'ExchangeNamedProperties.SearchError'
                    FolderPath   = $current.Path
                    ErrorCode    = $_.Exception.GetType().Name
                    ErrorMessage = $_.Exception.Message
                }
                break
            }

            foreach ($folder in $folderPage.Folders) {
                if ($folder -is [Microsoft.Exchange.WebServices.Data.SearchFolder]) {
                    continue
                }

                $folderName = [string]$folder.DisplayName
                $folderPath = if ([string]::IsNullOrEmpty($current.Path)) { $folderName } else { '{0}\{1}' -f $current.Path, $folderName }
                $queue.Enqueue([pscustomobject]@{ FolderId = $folder.Id; Path = $folderPath; Source = $current.Source; IsRoot = $false; Folder = $folder })
            }

            $offset += $folderPage.Folders.Count
            $moreFolders = $folderPage.MoreAvailable
            if ($folderPage.Folders.Count -eq 0) {
                $moreFolders = $false
            }
        } while ($moreFolders)
    }
}