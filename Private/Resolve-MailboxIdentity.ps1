function Resolve-MailboxIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    $getMailboxCommand = Get-Command -Name 'Get-Mailbox' -ErrorAction SilentlyContinue
    if ($getMailboxCommand) {
        try {
            $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        }
        catch {
            $exception = New-Object System.InvalidOperationException("Identity '$Identity' could not be resolved by Get-Mailbox: $($_.Exception.Message)", $_.Exception)
            throw $exception
        }

        $primarySmtpAddress = [string]$mailbox.PrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($primarySmtpAddress)) {
            throw "Get-Mailbox resolved '$Identity' but did not return PrimarySmtpAddress."
        }

        $mailboxGuid = $null
        if ($mailbox.PSObject.Properties['ExchangeGuid'] -and $mailbox.ExchangeGuid -and
            [guid]$mailbox.ExchangeGuid -ne [guid]::Empty) {
            $mailboxGuid = ([guid]$mailbox.ExchangeGuid).ToString('D')
        }

        return [pscustomobject]@{
            PSTypeName = 'ExchangeNamedProperties.MailboxIdentity'
            Identity = $Identity
            PrimarySmtpAddress = $primarySmtpAddress
            MailboxGuid = $mailboxGuid
        }
    }

    $parsedAddress = $null
    if (-not [System.Net.Mail.MailAddress]::TryCreate($Identity, [ref]$parsedAddress) -or
        $parsedAddress.Address -ne $Identity) {
        throw "Identity '$Identity' is not an SMTP address and Get-Mailbox is unavailable. Import the Exchange Management Shell tools or use the mailbox's SMTP address."
    }

    return [pscustomobject]@{
        PSTypeName = 'ExchangeNamedProperties.MailboxIdentity'
        Identity = $Identity
        PrimarySmtpAddress = $parsedAddress.Address
        MailboxGuid = $null
    }
}