function Import-EwsManagedApi {
    [CmdletBinding()]
    param(
        [string]$Path
    )

    $assemblyName = 'Microsoft.Exchange.WebServices'
    $loadedAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
    Where-Object { $_.GetName().Name -eq $assemblyName } |
    Select-Object -First 1

    if ($loadedAssembly) {
        if ($Path) {
            $requestedPath = if (Test-Path -LiteralPath $Path -PathType Container) {
                Join-Path $Path 'Microsoft.Exchange.WebServices.dll'
            }
            else {
                $Path
            }

            if ((Test-Path -LiteralPath $requestedPath -PathType Leaf) -and
                -not [string]::Equals([IO.Path]::GetFullPath($requestedPath), $loadedAssembly.Location, [StringComparison]::OrdinalIgnoreCase)) {
                throw "The EWS Managed API is already loaded from '$($loadedAssembly.Location)' and cannot be replaced in this PowerShell session. Start a new session to load '$requestedPath'."
            }
        }

        return $loadedAssembly.Location
    }

    $candidates = New-Object 'System.Collections.Generic.List[string]'
    if ($Path) {
        if (Test-Path -LiteralPath $Path -PathType Container) {
            $candidates.Add((Join-Path $Path 'Microsoft.Exchange.WebServices.dll'))
        }
        else {
            $candidates.Add($Path)
        }
    }

    $moduleRoot = Split-Path -Parent $PSScriptRoot
    $currentDirectory = (Get-Location).Path
    if ($currentDirectory) {
        $candidates.Add((Join-Path $currentDirectory 'Microsoft.Exchange.WebServices.dll'))
    }
    $candidates.Add((Join-Path $moduleRoot 'Microsoft.Exchange.WebServices.dll'))
    $candidates.Add((Join-Path $moduleRoot 'lib\Microsoft.Exchange.WebServices.dll'))

    foreach ($programFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($programFiles) {
            foreach ($versionDirectory in @('2.2', '2.1', '2.0')) {
                $candidates.Add((Join-Path $programFiles "Microsoft\Exchange\Web Services\$versionDirectory\Microsoft.Exchange.WebServices.dll"))
            }
        }
    }

    $assemblyPath = $candidates |
    Select-Object -Unique |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

    if (-not $assemblyPath) {
        throw 'Microsoft.Exchange.WebServices.dll was not found. Supply EwsManagedApiPath, place the DLL in the current working directory, module directory, or module lib directory, or install the EWS Managed API in a standard location.'
    }

    try {
        Add-Type -Path $assemblyPath -ErrorAction Stop
    }
    catch {
        throw "Microsoft.Exchange.WebServices.dll was found at '$assemblyPath' but could not be loaded. Verify that it is a compatible EWS Managed API assembly: $($_.Exception.Message)"
    }

    return [IO.Path]::GetFullPath($assemblyPath)
}