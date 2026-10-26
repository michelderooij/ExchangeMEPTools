# EWS Managed API

The Microsoft Exchange Web Services Managed API assembly is not bundled with this module. Verify the assembly's applicable redistribution terms before distributing it.

EWS Managed API is required only for item-level search and removal, not Store inventory. Install a compatible assembly and pass the DLL path to `Connect-EWSSession -EwsManagedApiPath`, place the DLL in the current working directory, module directory, or this directory as `Microsoft.Exchange.WebServices.dll`, or install it in a standard EWS Managed API location.

Do not commit the assembly unless its redistribution terms have been verified. The module uses the current Windows security context by default; connection settings are kept only for the current module session.
