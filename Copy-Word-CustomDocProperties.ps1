<#
.SYNOPSIS
Copies Word Custom Document Properties from one document to another

.DESCRIPTION
Given two word documents (a source and a target), this script will copy all
custom document properties from the source document to the target document.

Document properties that already exist in the target document will remain 
untouched.

.EXAMPLE
.\Copy-WordDocProperties.ps1 -SourceDocumentPath Source.docx -TargetDocumentPath Target.docx

.LINK
https://flamingkeys.com/BLOG-POST-SLUG

.NOTES
Written by: Chris Brown

Find me on:
* My blog: https://flamingkeys.com/
* Github: https://github.com/chrisbrownie

#>
Param(
    [Parameter(Mandatory)][string]$SourceDocumentPath,
    [Parameter(Mandatory)][string]$TargetDocumentPath
)

[CmdletBinding()]
[string]$ResolvedSourceDocumentPath = Resolve-Path $SourceDocumentPath
Write-Verbose "Resolved '-SourceDocumentPath' to '$ResolvedSourceDocumentPath'"

[string]$ResolvedTargetDocumentPath = Resolve-Path $TargetDocumentPath
Write-Verbose "Resolved '-TargetDocumentPath' to '$ResolvedTargetDocumentPath'"

# Make sure the source and target files exist
if (-not [system.io.file]::Exists($ResolvedSourceDocumentPath)) { throw [System.IO.FileNotFoundException]}
if (-not [system.io.file]::Exists($ResolvedTargetDocumentPath)) { throw [System.IO.FileNotFoundException]}

Write-Verbose "Opening Microsoft Word"
$wordApplication = New-Object -ComObject Word.Application -Verbose:$false

Write-Verbose "Hiding Word"
$wordApplication.Visible = $false
$wordApplication.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

Write-Verbose "Opening Word document (source)"
$sourceDocument = $wordApplication.Documents.Open($ResolvedSourceDocumentPath)
Write-Verbose "Opening Word document (target)"
$targetDocument = $wordApplication.Documents.Open($ResolvedTargetDocumentPath)

Write-Verbose "Retrieving custom document properties from source document"
$customDocProperties = $sourceDocument.CustomDocumentProperties
Write-Verbose "Retrieving custom document properties from target document"
$targetDocProperties = $targetDocument.CustomDocumentProperties
$DocPropertiesForTransfer = @{}

foreach ($customDocProperty in $customDocProperties) {
    # Put each docproperty into the hashtable
    $docPropertyName = [System.__ComObject].InvokeMember("name", [System.Reflection.BindingFlags]::GetProperty, $null, $customDocProperty, $null)
    $docPropertyValue = [System.__ComObject].InvokeMember("value", [System.Reflection.BindingFlags]::GetProperty, $null, $customDocProperty, $null)
    Write-Verbose "Retrieved property: $docPropertyName='$docPropertyValue'"
    $DocPropertiesForTransfer.$docPropertyName = $docPropertyValue
}
Write-Verbose "Processing document properties."
foreach ($customDocProperty in $DocPropertiesForTransfer.GetEnumerator()) {
    Write-Verbose "- Processing '$($customDocProperty.Name)'"
    $ThisDocPropertyArgs = $customDocProperty.Name, $false, 4, $customdocProperty.Value
    try {
        # Try to add the doc property
        [System.__ComObject].InvokeMember(
            "add", 
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $targetDocProperties,
            $ThisDocPropertyArgs
        ) | Out-Null
        Write-Verbose "-- Added doc property."
    } catch {
        Write-Verbose "-- Failed to add doc property. It probably already exists. Deleting it first."
        # If adding the doc property failed, it probably already exists. so try to delete it
        $propertyObject = [System.__ComObject].InvokeMember(
            "Item", 
            [System.Reflection.BindingFlags]::GetProperty,
            $null,
            $targetDocProperties,
            $customDocProperty.Name
        )

        [System.__ComObject].InvokeMember(
            "Delete",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $propertyObject,
            $null
        )

        # Try to add the doc property
        [System.__ComObject].InvokeMember(
            "add", 
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $targetDocProperties,
            $ThisDocPropertyArgs
        ) | Out-Null
    }
}
Write-Verbose "All document properties processed. Wrapping up."

Write-Verbose "Saving target document"
$targetDocument.Saved = $false
$targetDocument.Save()

Write-Verbose "Closing Word"
$WordApplication.Documents.Close()
$wordApplication.quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApplication)
Write-Verbose "Garbage collecting"
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
