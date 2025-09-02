
Param(
    [Parameter(Mandatory = $true )][string] $SiteUrl,
    [Parameter(Mandatory = $false)][string[]] $BuildLists,
    [Parameter(Mandatory = $false)][Switch] $LoadSecondaryData,
    [Parameter(Mandatory = $false)][Switch] $Help
)
$env:ENTRAID_APP_ID = '39a54693-963c-4ebb-aad0-3e2bc2c9a6b4'
Import-Module Microsoft.PowerShell.Utility;
Import-Module PnP.PowerShell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
#Import-Module BHF.PnP.Addons


If ($null -eq ${Function New-Guid}) {
    Function New-Guid {
        return [guid]::NewGuid()
    }
}
#Config Variables1
#$SiteURL = "https://bhfonline.sharepoint.com/teams-and-projects/platform-dev/ITBAUFrontDoor"  
$CsvPath = "$PSScriptRoot\CSVs";
$listNameAuthorisations = "FrontDoor"
$listUrlAuthorisations = "Lists/FrontDoor"

$listNameComments = "NatureOfRequest"
$listUrlComments = "Lists/NatureOfRequest"

$listNameDesiredOutcome = "DesiredOutcome"
$listUrlDesiredOutcome = "Lists/DesiredOutcome"

$listNameDistributionList  = "DistributionList"
$listUrlDistributionList  = "Lists/DistributionList"

$listNameDivision  = "Division"
$listUrlDivision = "Lists/Division"

$listNameAreaOfImpact  = "AreaOfImpact"
$listUrlAreaOfImpact = "Lists/AreaOfImpact"


$listNameUrgency  = "Urgency"
$listUrlUrgency = "Lists/Urgency"



$listNameLimitValues="LimitValues"
$listUrlLimitValues  = "Lists/LimitValues"



#$listNameStatusValues = "StatusValues"
#$listUrlStatusValues = "Lists/$listNameStatusValues"



$Template = "GenericList"
#$Script:AllLists = @("FrontDoor")
$Script:AllLists = @("FrontDoor", "NatureOfRequest", "DesiredOutcome","DistributionList", "Division", "Urgency", "LimitValues","AreaOfImpact")

#Get Credentials to connect
#$Cred = Get-Credential

Function Set-DefaultElement(
    [xml] $xml,
    $DefaultValue
)
{
    try {
        $defaultElt = $xml.CreateElement("Default")
        $defaultElt.InnerText = $DefaultValue
        $xml.DocumentElement.AppendChild($defaultElt) | Out-null
    }
    catch {
        #$e = $_.Exception;
        Write-Host -ForegroundColor Red "Error Setting Default: ${$e.Message}"
        throw;
    }
    return $xml;
}

<#
.SYNOPSIS
Add-LookupField: adds a primary Lookup Field to a SharePoint List

.DESCRIPTION

Lookup fields create a _relationship_ between one List (the _source list_) and another
list (the _target list_).  These relationships let you join information from the two lists
and keep it consistent as users of the list data add, edit and delete list items.

This is similar to the use of *foreign keys* in a database.

.PARAMETER ListName

Identifier for the List in which the field is to be placed

.PARAMETER InternalName
Internal Name of the Lookup Field

.PARAMETER DisplayName
Display Name of the Lookup Field (defaults to InternalName)

.PARAMETER Description
Field description

.PARAMETER Indexed

Is this field to be indexed?  Note that if the `EnforceUniqueValues` Switch is present, the field **shall**
be indexed and its Required property set.

.PARAMETER Required

Must a value be set for this field


.PARAMETER EnforceUniqueValues
Enforces unique values across Rows in SharePoint List

.PARAMETER RelationshipDeleteBehavior

How should SharePoint behave when the row referenced (lookup record) in the lookup list is deleted?

# Cascade Delete: If the lookup record is deleted, then the list record (containing the value referred to in the lookup list) is removed.

# Restrict Delete: If the lookup record is attempted to be deleted, Sharepoint will raise an error and
stop the operation

# None: No action is taken if the lookup record is deleted; the field will point to a non-existent lookup record


.PARAMETER LookupListName
Name of Lookup List

.PARAMETER LookupField
Name of lookup field.

.EXAMPLE

Establishes a link between a Header List "STDInvoice" and a Detail List "STDCollections".
The Lookup Field (on STDInvoice) is "ID"; the field on the Detail List is RequisitionID
```
  Add-LookupField -ListName STDCollections -InternalName 'RequisitionID' `
  -Indexed -Required `
  -RelationShipDeleteBehaviour "Cascade" `
  -LookupListName "STDInvoice" `
  -LookupField "ID"
- 
```

.NOTES
General notes
#>
Function Add-LookupField(
    [Parameter(Mandatory = $true)] [string]  $ListName, # List to host lookup field
    [Parameter(Mandatory = $true)] [string]  $InternalName, # Name of field
    [Parameter(Mandatory = $false)] [string] $DisplayName = $InternalName, # DisplayName
    [Parameter(Mandatory = $false)] [string] $Description = [string]::Empty, # Description of field
    [Parameter(Mandatory = $false)] [Switch] $Indexed, # Is lookup field indexed?
    [Parameter(Mandatory = $false)] [Switch] $Required, # Is lookup field Required?
    [Parameter(Mandatory = $false)] [Switch] $EnforceUniqueValues = $false, # Enforce Unique values
    [Parameter(Mandatory = $false)] [string] $RelationshipDeleteBehavior = "None", # What happens when child record removed
    [Parameter(Mandatory = $false)] [Switch] $AddToDefaultView,
    [Parameter(Mandatory = $true)] [string] $LookupListName, # Name of List consulted for field's value
    [Parameter(Mandatory = $true)] [string] $LookupField                           # Name of Reference field on Lookup List.
) {
    $FieldID = New-Guid;

    try {
#        $local:context = Get-PnPContext;
        $List = Get-PnPList -Identity $ListName;
        $LookupList = Get-PnPList -Identity $LookupListName;
        $LookupListId = $LookupList.Id
      
        $FieldSchema = @"
<Field 
    Type="Lookup" 
    ID="{$FieldID}" 
    Name="$InternalName"
    StaticName="$InternalName" 
    DisplayName="$DisplayName"
    Description="$Description"
    List="{$LookupListId}" 
    ShowField="$LookupField" 
/>
"@;
        # Use PowerShell's XML subsystem to add optional child elements/attributes
        [xml] $fieldXml = $FieldSchema;

        if ($Required.IsPresent) {
            $fieldXml.DocumentElement.SetAttribute("Required", "TRUE");
        }
        If ($Indexed.IsPresent -or $EnforceUniqueValues.IsPresent) {
            $fieldXml.DocumentElement.SetAttribute("Indexed", "TRUE");
        }
        if ($EnforceUniqueValues.IsPresent) {
            $fieldXml.DocumentElement.SetAttribute("EnforceUniqueValues", "TRUE");
        }
        If ($RelationshipDeleteBehavior -ne "None" -or $IsIndexed -eq $true) {
            $fieldXml.Field.SetAttribute("Indexed", "TRUE")
        }
        if ($AddToDefaultView.IsPresent) {
            $fieldXml.DocumentElement.SetAttribute("AddToDefaultView", "TRUE");
        }
        If ($null -ne $RelationshipDeleteBehavior) {
            $fieldXml.Field.SetAttribute("RelationshipDeleteBehavior", $RelationshipDeleteBehavior);
        }
        Add-PnPFieldFromXml -List $List -FieldXml $fieldXml.OuterXml
    }
    catch {
        Write-Host -ForegroundColor Red $_.Exception.Message
        Write-Host -ForegroundColor Red $_.ScriptStackTrace
        throw $_.Exception
    }
    finally {
    
    }
}
Function Add-BHFCalculatedField(
    [Parameter(Mandatory = $true)] $List,
    [Parameter(Mandatory = $true)] $InternalName,
    [Parameter(Mandatory = $false)] $DisplayName = $InternalName,
    [Parameter(Mandatory = $true)] $Formula,
    [Parameter(Mandatory = $false)] $FieldReferenceList,
    [Parameter(Mandatory = $false)] $resultType,
    [Parameter(Mandatory = $false)] $decimals,
    [switch] $ShowAsPercent,
    [switch] $AddToDefaultView
) {
    try {
        if (-not $Formula.StartsWith("=")) {
            Write-Warning "Formula *MUST* start with equals (just like Excel)"
            $Formula = "=$($Formula)"
        }
        $schemaCalculatedField = @"
        <Field 
            ID="{$([GUID]::NewGuid())}" 
            Name="$InternalName"
            StaticName="$InternalName" 
            DisplayName="$DisplayName" 
            Type="Calculated" 
            ResultType="$resultType" 
            ReadOnly="TRUE">
          <Formula>$formula</Formula>
        </Field>
"@;
        [xml] $xmlCalculatedField = $schemaCalculatedField;
        If ($null -ne $resultType) {
            if ($resultType.ToLower() -notin ("number", "currency") -and $null -ne $decimals) {
                throw "`"Decimals`" invalid for result type `"$resultType`"";
            }
            elseif ($null -ne $decimals) {
                $xmlCalculatedField.DocumentElement.SetAttribute("Decimals", $decimals);
            }
            if ($resultType.ToLower() -eq "currency") {
                $xmlCalculatedField.DocumentElement.SetAttribute("LCID", "2057") # en-GB
            }
        }
        if ($ShowAsPercent.IsPresent) {
            $xmlCalculatedField.DocumentElement.SetAttribute("Percentage", "TRUE");
        }
        if ($AddToDefaultView.IsPresent) {
            $xmlCalculatedField.DocumentElement.SetAttribute("AddToDefaultView", "TRUE");
        }
        If($null -ne $FieldReferenceList -and $FieldReferenceList.Count -gt 0) {
            $fieldRefs = $xmlCalculatedField.CreateElement("FieldRefs");
            ForEach ($field in $FieldReferenceList) {
                $fieldRef = $xmlCalculatedField.CreateElement("FieldRef");
                $fieldRef.SetAttribute("Name", $field);
                $fieldRefs.AppendChild($fieldRef);
            }
            $xmlCalculatedField.Field.AppendChild($fieldRefs);
        }
        return Add-PnPFieldFromXml -List $List  -FieldXml $xmlCalculatedField.OuterXml;
    }
    catch {
        Write-Host -ForegroundColor Red $_.Exception.Message;
        Write-Host -ForegroundColor Red $_.ScriptStackTrace
    }
}

<#
.SYNOPSIS

Add-BHFChoiceField: Adds a Choice field to a SharePoint List

.PARAMETER List

A list identifier, typically the URL of the list e.g. "Lists/TargetList" or the Name of a List "TargetList".

.PARAMETER InternalName

The Internal or Static Name of the field

.PARAMETER DisplayName

The display name of the field; this will be used by default by PowerApps.  This is an OPTIONAL parameter; the Displayname is set to the value of the InternalName parameter by default

.PARAMETER Choices

A list of strings representing the choices that the user of the SharePoint form (or PowerApp) may choose from.

.PARAMETER Format

By default, shows Choices as a dropdown/combobox; other styles are also available.

.PARAMETER Default

Default value

.PARAMETER Required

Should this SharePoint field always be set?

.PARAMETER Indexed

Should this SharePoint field be indexed?

#>
Function Add-BHFChoiceField(
    [Parameter(Mandatory=$true)]  $List,
    [Parameter(Mandatory=$true)][string]  $InternalName,
    [Parameter(Mandatory=$false)][string] $DisplayName = $InternalName,
    [Parameter(Mandatory=$true)][string[]]  $Choices,
    [Parameter(Mandatory=$false)] $Format="Dropdown",
    [Parameter(Mandatory=$false)] $Default = $null,
    [Parameter(Mandatory=$false)] [Switch] $Required,
    [Parameter(Mandatory=$false)] [Switch] $Indexed,
    [Parameter(Mandatory=$false)] [Switch] $AddToDefaultView
 )
{
    [string] $local:fieldId = New-Guid
    [string] $local:fieldSchemaXml = @"
    <Field 
    Type="Choice"
    Name="$InternalName" 
    StaticName="$InternalName" 
    DisplayName="$displayName"
    ID="{$local:fieldId}" 
/>
"@
    [xml] $local:fieldXmlDoc = $fieldSchemaXml
    if ($null -ne $Default) {
        Set-DefaultElement -xml $fieldXmlDoc -DefaultValue $Default
    }
    If ($null -ne $Format) {
        $fieldXmlDoc.Field.SetAttribute("Format", $Format)
    }
    If ($Required.IsPresent) {
        $fieldXmlDoc.Field.SetAttribute("Required", "TRUE");
    }
    If ($Indexed.IsPresent) {
        $fieldXmlDoc.Field.SetAttribute("Indexed", "TRUE")
    }
    If ($AddToDefaultView.IsPresent) {
        $fieldXmlDoc.Field.SetAttribute("AddToDefaultView", "TRUE");
    }
    $local:ChoicesElt = $fieldXmlDoc.CreateElement("CHOICES");
    ForEach($local:Choice in $Choices) {
        $local:ChoiceChildElt = $fieldXmlDoc.CreateElement("CHOICE");
        $local:ChoiceChildElt.InnerText = $local:Choice
        $ChoicesElt.AppendChild($local:ChoiceChildElt)
    }
    $fieldXmlDoc.Field.AppendChild($local:ChoicesElt)
    Add-PnPFieldFromXml -List $List -FieldXml $local:fieldXmlDoc.OuterXml
    
}

Function Add-BHFNumberField(
    [Parameter(Mandatory = $true)] $List,
    [Parameter(Mandatory = $true)] $InternalName,
    [Parameter(Mandatory = $false)] $DisplayName = $InternalName,
    [Parameter(Mandatory = $false)] $Decimals = 0,
    [Parameter(Mandatory = $false)][Switch] $Required,
    [Parameter(Mandatory = $false)] $minimumValue = $null,
    [Parameter(Mandatory = $false)] $maximumValue = $null,
    [Parameter(Mandatory = $false)] $DefaultValue = $null,
    [Parameter(Mandatory = $false)] [Switch] $Percentage,
    [Parameter(Mandatory = $false)] [switch] $AddToDefaultView
) {
    try {
        if ($DisplayName -eq [string]::Empty) {
            $DisplayName = $InternalName
        }
        $local:fieldId = (New-GUID).Guid;
        $local:fieldSpec = @"
    <Field 
        Type="Number" 
        Name="$InternalName" 
        DisplayName="$DisplayName"
        ID="{$local:fieldId}" 
        StaticName="$InternalName" 
        AddToDefaultView="$($AddToDefaultView.IsPresent)" 
    />
"@;
        [xml] $local:fieldXml = $local:fieldSpec;
        If($Required.IsPresent) {
            $local:fieldXml.Field.SetAttribute("Required", "TRUE");
        }
        If ($null -ne $minimumValue) {
            $local:fieldXml.Field.SetAttribute("Min", $minimumValue);
        }
        If ($null -ne $maximumValue) {
            $local:fieldXml.Field.SetAttribute("Max", $maximumValue);
        }
        if ($null -ne $decimals) {
            $local:fieldXml.Field.SetAttribute("Decimals", $decimals)
        }
        If ($Percentage.IsPresent) {
            $local:fieldXml.Field.SetAttribute("Percentage", "TRUE")
        }
        if ($null -ne $DefaultValue) {
            Set-DefaultElement -xml $local:fieldXml -DefaultValue DefaultValue
        }
        Add-PnPFieldFromXml -Connection $pnpConnection -List $List -FieldXml $local:fieldXml.OuterXml
    }
    catch {
        Write-Host -ForegroundColor Red $_.Exception.Message;
        Write-Host -ForeGroundColor Red $_.ScriptStackTrace
        throw
    }
}
Function Add-BHFDateTimeField(
    [Parameter(Mandatory = $true)] $List,
    [Parameter(Mandatory = $true)] $InternalName,
    [Parameter(Mandatory = $false)] $displayName = $InternalName,
    [Parameter(Mandatory = $false)][Switch] $AddToDefaultView,
    [Parameter(Mandatory = $false)][Switch] $Required,
    [Parameter(Mandatory = $false)] $DateFormat = "DateOnly",
    [Parameter(Mandatory = $false)] $DefaultValue = $null,
    [Parameter(Mandatory = $false)] $FriendlyDisplayFormat = [string]::Empty

)
{
    $local:fieldId = (New-GUID).Guid;
    $local:fieldSpec = @"
<Field 
    Type="DateTime" 
    Name="$InternalName" 
    DisplayName="$displayName"
    ID="{$local:fieldId}" 
    Format="$DateFormat"
/>
"@;
    try{
        [xml] $local:fieldXml = $local:fieldSpec
        If ($Required.IsPresent) {
            $local:fieldXml.Field.SetAttribute("Required", "TRUE");
        }
        If ($AddToDefaultView.IsPresent) {
            $local:fieldXml.Field.SetAttribute("AddToDefaultView", "TRUE");
        }
        If ($null -ne $DefaultValue) {
            Set-DefaultElement -xml $local:fieldXml -DefaultValue $DefaultValue
        }
        If ($null -ne $FriendlyDisplayFormat) {
            $local:fieldXml.Field.SetAttribute("FriendlyDisplayFormat", $FriendlyDisplayFormat)
        }

        Add-PnPFieldFromXml -List $List -FieldXml $local:fieldXml.OuterXml;
    }
    catch {
        Write-Host -ForegroundColor Red $_.Exception.Message;
        Write-Host -ForeGroundColor Red $_.ScriptStackTrace
        throw
    }
}

Function Hide-TitleField (
    $List,
    $TitleDisplayName = "_Old Title"
) {
    $f = Get-PnPField -List $List -Identity "Title" -ErrorAction SilentlyContinue
    if ($null -eq $f) {
        Write-Warning "Failed to locate ""Title"" field in $List"
        return
    }
    $local:ctx = $Connection.Context;
    Set-PnPField -List $List -Identity $f.Id -Values @{Title = $TitleDisplayName; Hidden = $true; Required = $false } -Connection $Connection
    $local:ctx.Load($f);
    $local:ctx.ExecuteQuery();
    $local:view = Get-PnPView -List $List -Identity "All Items"
    $retainedFields = $($view.ViewFields | ForEach-Object { if ($_ -ne "LinkTitle" -and $_ -ne "Title") {return $_}});
    
    if($null -ne $retainedFields -and $retainedFields.Count -gt 0) {
        Set-PnPView -List $list -Identity "All Items" -Fields $retainedFields
    }
    Invoke-PnPQuery
}

Function Set-View (
    [Parameter(Mandatory = $true)] $List,
    [Parameter(Mandatory = $true)] [string[]] $Fields,
    [Parameter(Mandatory = $false)] [Switch] $IsDefault = $false,
    [Parameter(Mandatory = $false)] $ViewName = "All Items",
    [Parameter(Mandatory = $false)] $ViewIdentity = "AllItems",
    [Parameter(Mandatory = $false)][string] $ViewQuery = [string]::Empty
) {
    #    $ViewQuery = [string]::Empty;
    try {
        Remove-PnPView -Connection $connection -List $List -Identity $ViewName -Force # AllItems doesn't cut it!
        $view = Add-PnpView -Connection $connection -List $List -Title $viewIdentity -ViewType None `
            -Fields $Fields -RowLimit 30 -Paged
        $view = $view.TypedObject;
        Set-PnPView -List $List -Connection $connection -Identity $view.Id -Values @{"Title" = $viewName }

        If ($IsDefault) {
            Set-PnPView -List $List -Identity $view.Id -Values @{ "DefaultView" = $true }
        }
        If ($ViewQuery -ne [string]::Empty) {
            Set-PnPView -List $List -Identity $view.Id -Values @{ "ViewQuery" = $ViewQuery }
        }
		
    }
    catch {
        Write-Host -ForegroundColor Red $_.ScriptStackTrace;
        Write-Host -ForegroundColor Red $_.Exception.Message;
        throw;
    }
}

<#
.SYNOPSIS
Adds a DateTime Field to a List
#>
Function Add-BHFDateTimeField(
    [Parameter(Mandatory = $true)] $List,
    [Parameter(Mandatory = $true)] $InternalName,
    [Parameter(Mandatory = $false)] $displayName = $InternalName,
    [Parameter(Mandatory = $false)][Switch] $AddToDefaultView,
    [Parameter(Mandatory = $false)][Switch] $Required,
    [Parameter(Mandatory = $false)] $DateFormat = "DateOnly",
    [Parameter(Mandatory = $false)] $DefaultValue = $null,
    [Parameter(Mandatory = $false)] $FriendlyDisplayFormat = [string]::Empty

)
{
    $local:fieldId = (New-GUID).Guid;
    $local:fieldSpec = @"
<Field 
    Type="DateTime" 
    Name="$InternalName" 
    DisplayName="$displayName"
    ID="{$local:fieldId}" 
    Format="$DateFormat"
/>
"@;
    [xml] $local:fieldXml = $local:fieldSpec
    If ($Required.IsPresent) {
        $local:fieldXml.Field.SetAttribute("Required", "TRUE");
    }
    If ($AddToDefaultView.IsPresent) {
        $local:fieldXml.Field.SetAttribute("AddToDefaultView", "TRUE");
    }
    If ($null -ne $DefaultValue) {
        Set-DefaultElement -xml $local:fieldXml -DefaultValue $DefaultValue
    }
    If ($null -ne $FriendlyDisplayFormat) {
        $local:fieldXml.Field.SetAttribute("FriendlyDisplayFormat", $FriendlyDisplayFormat)
    }

    Add-PnPFieldFromXml -List $List -FieldXml $local:fieldXml.OuterXml;
}

#Custom function to add column to list
Function Add-CalculatedColumnToList() {
    param
    (
        [Parameter(Mandatory = $true)] [string] $ListName,
        [Parameter(Mandatory = $true)] [string] $Name,
        [Parameter(Mandatory = $true)] [string] $DisplayName,
        [Parameter(Mandatory = $false)] [string] $Description = [string]::Empty,
        [Parameter(Mandatory = $true)] [string] $Formula,
        [Parameter(Mandatory = $true)] [string] $ResultType,
        [Parameter(Mandatory = $true)] [string[]] $FieldsReferenced
    )
 
    #Generate new GUID for Field ID
    $FieldID = New-Guid
    $local:ctx = Get-PnPContext;
    Try {
 
        Write-host "List name = " $ListName -f Yellow
         
        #Get the List
        $List = Get-PnPList -Identity $ListName
 
        #Check if the column exists in list already
        $Fields = $List.Fields
        $local:ctx.Load($Fields)
        $local:ctx.executeQuery()
        $NewField = $Fields | Where-Object { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
        if ( $NULL -ne $NewField) {
            Write-host "Column $Name already exists in the List!" -f Yellow
        }
        else {
            #Frame FieldRef Field
            $FieldRefXML = [string]::Empty
            #            $FieldRefs = $FieldsReferenced.Split(",")
            foreach ($Ref in $FieldsReferenced) {
                $FieldRefXML = $FieldRefXML + "<FieldRef Name='$Ref' />"
            }
 
            #Create Column in the list
            $FieldSchema = @"
<Field 
    Type='Calculated'
    ID='{$FieldID}' 
    DisplayName='$DisplayName'
    Name='$Name' 
    Description='$Description' 
    ResultType='$ResultType' 
    ReadOnly='TRUE'>
    <Formula>$Formula</Formula>
    <FieldRefs>$FieldRefXML</FieldRefs>
</Field>
"@
            $FieldSchema
            #$NewField = $List.Fields.AddFieldAsXml($FieldSchema,$True,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $NewField = Add-PnPFieldFromXml -List $List -FieldXml $FieldSchema
             
            Write-host "New Column Added to the List Successfully!" -ForegroundColor Green 
        }
    }
    Catch {
        #write-host -f Red "Error Adding Column to List!" $_.Exception.Message
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Host -ForegroundColor Red "caught exception: $e at $line"
        Write-Host -ForegroundColor Red $_.ScriptStackTrace
    }
}

Function IsListBuildRequired($ListName) {
    If ($ListName -in $Script:listsToBuild) {
        return $true
    }
    Else {
        Write-Warning "$ListName not required - not rebuilding"
        return $false
    }
}


Function New-FrontDoorList($ListName, $ListUrl, $Connection)
{
    #$ListName = "FrontDoor"
    #$ListUrl = "List/FrontDoor"
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }    
    If ($null -ne (Get-PnPField -Identity ChangeAuthorisationID -List "NatureOfRequest" -ErrorAction SilentlyContinue)) {
        Remove-PnPField -Identity ChangeAuthorisationID -List "NatureOfRequest" -Force
        Invoke-PnPQuery
    }
    If ($null -ne (Get-PnPField -Identity ChangeAuthorisationID -List "ChangeAuthorisationChanges" -ErrorAction SilentlyContinue)) {
        Remove-PnPField -Identity ChangeAuthorisationID -List "ChangeAuthorisationChanges" -Force
        Invoke-PnPQuery
    }
    If($null -ne (Get-PnpList -Identity $ListName -ErrorAction SilentlyContinue)) {
        Remove-PnPList -Identity $ListName -Force
    }
    New-PnPList -Title $ListName             -Url $ListUrl             -Template $Template -ErrorAction Stop
    #NewStarters - Employee details
    #cannot create a field with Internal name Title as field already exists.
    #Set-PnPField -List $ListName -Identity Title -Values @{Required = $false; "Title" = "Old_Title" }second
<#
    $changeTypeList = 
    "Change To Job Title", "Change To Job Level", "Change Of Line Manager", "Contract Extension (Of Existing Contract)"                                         `
    , "Change Of Salary", "Change Of Name", "Secondment (Don't appear for 'Only Shops/Stores')", "End of Secondment (Don't appear for 'Only Shops/Stores')"     `
    , "End Of Temporary Change", "Permanent Promotion", "Additional Responsibility Allowance (Don't appear for 'Only Shops/Stores')"                            `
    , "Compensation Change/Benchmarking", "Change to Cost Centre", "Change to Working Style and/or Office Location"                                             `
    , "Temporary Promotion (Only appear for 'Only Shops/Stores')", "End of Temporary Promotion (Only appear for 'Only Shops/Stores')";
#>


    #
    # Change Type Fields

   


    #Add-BHFChoiceField -List $ListName -InternalName 'WorkerTypeChoice' -DisplayName 'WorkerTypeChoice'     `
    #-Required -AddToDefaultView                                                                         `
    #-Choices "Fixed Term Employee", "Volunteer", "Permanent Employee", "Contractor/Contingent Worker"   `
    
    

    # The Data Subject
    $field = Add-PnPField -List $ListName -InternalName 'RequestorName' -DisplayName 'RequestorName' -Type User -Required
    Set-PnPField -List $ListName -Identity $field.Id -Values @{"SelectionMode" = 0}  #People only.  Not groups.
    
        Add-PnPField -List $ListName -InternalName 'Directorate' -DisplayName 'Directorate' -Type Text -AddToDefaultView -Required
	    Add-PnPField -List $ListName -InternalName 'AreaOfImpact' -DisplayName 'AreaOfImpact' -Type Text -AddToDefaultView -Required
  	    Add-PnPField -List $ListName -InternalName 'NatureOfRequest' -DisplayName 'NatureOfRequest' -Type Text -AddToDefaultView -Required   
	    Add-PnPField -List $ListName -InternalName 'RequestOutline' -DisplayName 'RequestOutline' -Type Text -AddToDefaultView -Required   
        Add-PnPField -List $ListName -InternalName  'RequestDescription' -DisplayName 'RequestDescription' -Type Note -AddToDefaultView  -Required
Add-PnPField -List $ListName -InternalName 'TestingSignoff' -DisplayName 'TestingSignoff' -Type User -Required

    Add-PnPField -List $ListName -Type Boolean -AddToDefaultView -InternalName 'ThisARegulatoryRequirement' -DisplayName 'ThisARegulatoryRequirement'  -Required
	
	Add-PnPField -List $ListName -InternalName 'DesiredOutcome' -DisplayName 'DesiredOutcome' -Type Text -AddToDefaultView -Required
Add-PnPField -List $ListName -InternalName 'Urgency' -DisplayName 'Urgency' -Type Text -AddToDefaultView -Required
   Add-PnPField -List $ListName -InternalName  'Risks' -DisplayName 'Risks' -Type Note -AddToDefaultView -Required 
    


    Add-BHFDateTimeField -List $ListName -InternalName 'Deadline' -DisplayName 'Deadline' -DateFormat "DateOnly" -FriendlyDisplayFormat $null -AddToDefaultView 
    #Add-BHFDateTimeField -List $ListName -InternalName 'EndDate' -DisplayName 'EndDate' -DateFormat "DateOnly" -FriendlyDisplayFormat $null -AddToDefaultView


    


 
    

	#Add-BHFDateTimeField -List $ListName -InternalName 'StatusChangeDate' -DisplayName 'StatusChangeDate' -DateFormat "DateTime" -FriendlyDisplayFormat $null -AddToDefaultView
    Hide-TitleField -List $listUrl

    Invoke-PnPQuery
}


<#
Function New-StatusValuesList($ListName, $ListUrl, $Connection)
{
    $local:dataStatusValuesList = "$($CsvPath)/$($ListName).csv"
    If(-not (IsListBuildRequired $ListName)) {
        return;
    }
    If ($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)) {
        Remove-PnPList -Identity $ListName -Force
    }

    New-PnPList -Title $ListName -Url $ListUrl -Connection $Connection -Template GenericList
    Invoke-PnPQuery

    Add-PnPField -List $ListName -InternalName 'Status' -DisplayName 'Status' -Type Text -AddToDefaultView
	Add-PnPField -List $ListName -InternalName 'Visibility' -DisplayName 'Visibility' -Type Text -AddToDefaultView
    Hide-TitleField -List $ListName
    If(Test-Path $dataStatusValuesList) {
        $statusValueItems = Import-Csv $dataStatusValuesList
        Write-Host -NoNewline "Loading $ListName"
        ForEach ($item in $statusValueItems) {
            Add-PnPListItem -List $ListName -Values @{"Status" = $item.Status;"Visibility" = $item.Visibility;} | Out-Null
            Write-Host -NoNewLine "."
        }
        Write-Host "`n"        
    }
}#>


Function New-LimitValuesList($ListName, $ListUrl, $Connection)
{
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }
	If($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)){
        Remove-PnPList -Identity $ListName -Force
	}
   New-PnPList -Title $ListName -Url $listUrl -Template GenericList
	
    #Set-PnPField -List $ListName -Identity Title -Values @{Required = $false; "Title" = "Old_Title" }
    Add-PnPField -List $ListName -InternalName  'NameOfValue' -DisplayName 'NameOfValue' -Type Text -AddToDefaultView 
    Add-PnPField -List $ListName -InternalName  'TypeOfLimitationRecord' -DisplayName 'TypeOfLimitationRecord' -Type Choice -Choices 'Minimum', 'Maximum', 'Absolute' -AddToDefaultView
   
    Add-PnPField -List $ListName -InternalName  'NumericValue' -DisplayName 'NumericValue' -Type Number -AddToDefaultView
    Add-PnPField -List $ListName -InternalName  'CurrencyValue' -DisplayName 'CurrencyValue' -Type Currency -AddToDefaultView
    Add-PnPField -List $ListName -InternalName  'EmailFinance' -DisplayName 'EmailFinance' -Type Text -AddToDefaultView
    Add-PnPField -List $ListName -InternalName  'TextValue' -DisplayName 'TextValue' -Type Text -AddToDefaultView
    Add-PnPField -List $ListName -InternalName  'Fascia' -DisplayName 'Fascia' -Type Choice -Choices 'HR', 'FE' -AddToDefaultView
    #
    # Load 2 LimitValues entries via CSV - really??
    Add-PnPListItem -List $ListName -Values @{"Title"="OperatingEnvironment";"NameOfValue"="OperatingEnvironment";"TypeOfLimitationRecord"="Minimum";"Fascia"="HR";"TextValue"="Dev"}
    Add-PnPListItem -List $ListName -Values @{"Title"="Version";"NameOfValue"="Version";"NumericValue"="0.1"}
}

Function New-DesiredOutcomeList($ListName, $ListUrl, $Connection) {
    #$ListName = "DesiredOutcome"
    #$ListUrl = "Lists/DesiredOutcome"
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }
	If($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)){
       Remove-PnPList -Force -Identity $ListName -Connection $connection 
	}
    New-PnPList -Title $ListName -Url $ListUrl -Template GenericList
    Set-PnPField -List $ListName -Identity Title -Values @{Required = $false; "Title" = "Old_Title" }
    Add-PnPField -List $ListName -InternalName  'DesiredOutcome' -DisplayName 'DesiredOutcome' -Type Text -AddToDefaultView -Required
    Add-PnPField -List $ListName -InternalName 'Active' -DisplayName 'Active' -Type Boolean -AddToDefaultView -Required
    Set-View -List $ListName -Fields DesiredOutcome,Active  -IsDefault
    #If ($LoadSecondaryData.IsPresent) {
        If($null -ne ($dataDesiredOutcomeCsv = Get-Item -Path "$CsvPath/DesiredOutcome.csv" -ErrorAction SilentlyContinue)) {
            $DesiredOutcomeItems = Import-Csv $dataDesiredOutcomeCsv;
            Write-Host -NoNewLine "Loading $ListName "
            ForEach ($DesiredOutcome in $DesiredOutcomeItems) {
                $local:isActive = Set-BooleanField $DesiredOutcome.Active
                Add-PnPListItem -List $ListName -Values @{"DesiredOutcome" = $DesiredOutcome.DesiredOutcome} | Out-Null
                Write-Host -NoNewLine "."
            }
            Write-Host "`n"
        }
        else {
            Write-Warning "Cannot find CSV file ""$CsvPath/DesiredOutcome.csv"""
        }
    #}
}



Function New-DivisionList ($ListName, $ListUrl, $Connection) 
{
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }
#    $ListName = "Division"
#    $ListUrl = "Lists/Division"
    $csvFilePath = "$csvPath/Division.csv"
	If($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)){
    Remove-PnPList -Identity $ListName -Connection $connection -Force
	}
    New-PnPList -Title $ListName -Url $ListUrl -Template GenericList

    Add-PnPField -List $ListName -InternalName  'Division' -DisplayName 'Division' -Type Text -AddToDefaultView -Required
    Set-View -List $ListName -Fields Division -IsDefault
    #If($LoadSecondaryData.IsPresent) {
        If($null -ne ($csvLocation = Get-Item -Path "$csvFilePath" -ErrorAction SilentlyContinue)) {
            Write-Host -NoNewline "Loading $ListName "
            $locationData = Import-csv -Path $csvLocation
            ForEach ($location in $locationData) {
                Add-PnPListItem -List $ListName -Values @{"Division" = $location.Division} | Out-Null
                Write-Host -NoNewline "."
            }
            Write-Host "`n"
        }
        else {
            Write-Warning "Can't locate CSV file ""$csvFilePath""."
        }
    #}
}
Function New-AreaOfImpactList ($ListName, $ListUrl, $Connection) 
{
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }

    $csvFilePath = "$csvPath/AreaOfImpact.csv"
	If($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)){
    Remove-PnPList -Identity $ListName -Connection $connection -Force
	}
    New-PnPList -Title $ListName -Url $ListUrl -Template GenericList

    Add-PnPField -List $ListName -InternalName  'AreaOfImpact' -DisplayName 'AreaOfImpact' -Type Text -AddToDefaultView -Required
    Set-View -List $ListName -Fields AreaOfImpact -IsDefault
    #If($LoadSecondaryData.IsPresent) {
        If($null -ne ($csvLocation = Get-Item -Path "$csvFilePath" -ErrorAction SilentlyContinue)) {
            Write-Host -NoNewline "Loading $ListName "
            $locationData = Import-csv -Path $csvLocation
            ForEach ($location in $locationData) {
                Add-PnPListItem -List $ListName -Values @{"AreaOfImpact" = $location.AreaOfImpact} | Out-Null
                Write-Host -NoNewline "."
            }
            Write-Host "`n"
        }
        else {
            Write-Warning "Can't locate CSV file ""$csvFilePath""."
        }
    #}
}

Function New-UrgencyValues($ListName, $ListUrl, $Connection)
{
#    $listNameUrgency = "UrgencyValues"
#    $listUrlUrgency  = "Lists/$listNameUrgency"
    If (-not (IsListBuildRequired $ListName)) {
        return;
    }
    If($null -ne (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)){
        Remove-PnPList -Force -Identity $ListName -Connection $Connection
    }
    New-PnPList -Title $ListName -Url $ListUrl -Template GenericList -Connection $connection
    Set-PnPField -List $ListUrl -Identity Title -Values @{ Title = "Old_Title"; Required = $false;}
    Add-PnPField -List $ListName -InternalName 'UrgencyValue' -DisplayName 'UrgencyValue' -Type Text
    #Set-View -List $ListName -Fields UrgencyValue -IsDefault
    Invoke-PnPQuery -Connection $Connection
    #If($LoadSecondaryData.IsPresent) {
        If ($null -ne ($dataUrgencyValues = Get-Item -Path "$CsvPath/UrgencyValues.csv" -ErrorAction SilentlyContinue)) {
            $UrgencyValues = Import-Csv -Path $dataUrgencyValues
            Write-Host -NoNewLine "Loading $ListName "
            ForEach($UrgencyValue in $UrgencyValues) {
                Add-PnPListItem -List $ListName -Values @{"UrgencyValue" = $UrgencyValue.UrgencyValue} | Out-Null
                Write-Host -NoNewline "."
            }
            Write-Host "`n"
        }
        else {
            Write-Warning "CSV file can't be loaded for $ListName"
        }
    #}
}



Function Set-BooleanField($value) {
    if ($null -eq $value) {
        return $false;
    }
    if ($value || ! $value) {
        return $value;
    }
    if ($value.ToLower() -in ("false", "no", "n")) {
        return $false
    }
    if ($value.ToLower() -in ("true", "yes", "y")) {
        return $true
    }
    return $false
}



Function Write-Help ($message = "") 
{
    $($ScriptName)
    Write-Output @"
$($ScriptName): $($message)

Creates Change Authorisation Lists in a SharePoint Site

Usage: $($ScriptName) -SiteUrl https://samplesite.sharepoint.com/teams-and-projects/site// `
                              -BuildLists <list of SharePoint Lists to build>

I know how to build the following lists:-
List
====================================
"@
    $Script:AllLists | ForEach-Object {
        Write-Output "$_"
    }   

}

$ScriptInvocation = (Get-Variable MyInvocation -Scope Script).Value
$ScriptName = $ScriptInvocation.MyCommand.Name
#Read more: https://www.sharepointdiary.com/2017/03/add-yes-no-check-box-field-to-sharepoint-list-using-powershell.html#ixzz7PfnvrnL0
Try {
    if ($Help.IsPresent) {
        Write-Help
        return;
    }
    $script:listsToBuild = @();
    if ($null -eq $BuildLists -or $BuildLists.Count -eq 0 -or "all" -in $BuildLists) {
        $script:listsToBuild = $Script:AllLists
    }
    else {
        $BuildLists | ForEach-Object {
            $local:listName = $_.ToLower();
            if ($local:listName -in $Script:AllLists) {
                $script:listsToBuild += $local:listName
            }
            else {
                Write-Warning "Cannot build list `"$local:listName`" "
            }
        }
    }
    if ($null -eq $script:listsToBuild -or $script:listsToBuild.Count -eq 0 ) {
        Write-Help "No lists found to build"
        return
    }
    #Connect to PnP Online
    #Connect-PnPOnline -Url $SiteURL -Credentials $Cred 
    #$connection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection
    Connect-PnPOnline -Url $SiteUrl -Interactive 
    $Connection= Get-PnpConnection
    # Remove Projected fields first, then the lookup field, then the containing list.
    # Each of the New... functions in this block takes care of loading
	<#New-ChangeAuthEmailNotificationEmailAddressesList -ListName $listNameChangeAuthEmailNotificationEmailAddresses  -ListUrl $listUrlChangeAuthEmailNotificationEmailAddresses -Connection $connection
	New-CheckPermissionsList -ListName $listNameCheckPermissions  -ListUrl $listUrlCheckPermissions -Connection $connection
	New-LimitValuesList -ListName $listNameLimitValues  -ListUrl $listUrlLimitValues -Connection $connection
	New-SecondmentReasons -ListName $listNameSecondmentReasons  -ListUrl $listUrlSecondmentReasons -Connection $connection
	New-ShopBandsList -ListName $listNameShopBands -ListUrl $listUrlShopBands -Connection $connection
#>
New-LimitValuesList -ListName $listNameLimitValues  -ListUrl $listUrlLimitValues -Connection $connection
    New-FrontDoorList -ListName $listNameAuthorisations -ListUrl $listUrlAuthorisations -Connection $connection
 New-DesiredOutcomeList -ListName $listNameDesiredOutcome -ListUrl $listUrlDesiredOutcome -Connection $connection
    New-AreaOfImpactList -ListName $listNameAreaOfImpact -ListUrl $listUrlAreaOfImpact -Connection $connection
    New-DivisionList -ListName $listNameDivision -ListUrl $listUrlDivision -Connection $connection
    New-UrgencyValues -ListName $listNameUrgency -ListUrl $listUrlUrgency -Connection $connection
#
  <#  New-FullOrPartTimeList -ListName $listNameFullOrPartTime -ListUrl $listUrlFullOrPartTime -Connection $connection

     New-DesiredOutcomeList -ListName $listNameDesiredOutcome -ListUrl $listUrlDesiredOutcome -Connection $connection
    New-DistributionList -ListName $listNameDistributionList -ListUrl $listUrlDistributionList -Connection $connection
    New-DivisionList -ListName $listNameDivision -ListUrl $listUrlDivision -Connection $connection
    New-UrgencyValues -ListName $listNameUrgency -ListUrl $listUrlUrgency -Connection $connection
    New-OfficeLocationValues -ListName $listNameOfficeLocation -ListUrl $listUrlOfficeLocation -Connection $connection
    New-DualHomeLocation -ListName $listNameDualHomeLocationValues  -ListUrl $listUrlDualHomeLocationValues -Connection $connection
    #New-ApprovalDestination -ListName $listNameApprovalDestination -ListUrl $listUrlApprovalDestination -Connection $connection
    New-JobLevel -ListName $listNameJobLevel  -ListUrl $listUrlJobLevel -Connection $connection
    New-YesNoList -ListName $listNameYesNo -ListUrl $listUrlYesNo -Connection $Connection
    New-StatusValuesList -ListName $listNameStatusValues -ListUrl $listUrlStatusValues -Connection $Connection
    New-HolidaysAndNonWorkDays -ListName $listNameHolidaysAndNonWorkDays -ListUrl $listUrlHolidaysAndNonWorkDays -Connection $Connection
	New-AuthorisationMatrixList -ListName $listNameAuthorisationMatrix -ListUrl $listUrlAuthorisationMatrix -Connection $Connection
    New-ChangeAuthorisationChangesList -ListName $listNameChangeAuthorisationChanges -ListUrl $listUrlChangeAuthorisationChanges -LookupListName $listNameAuthorisations -LookupField "ID"  -Connection $Connection
    New-CommentList -ListName $listNameComments -ListUrl $listUrlComments -LookupListName $listNameAuthorisations -LookupField "ID" -Connection $Connection #>
}
catch {
    $e =    $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg =  $e.Message 
    Write-Host -ForegroundColor Red "caught exception: $msg at $line"
    write-host "Error adding field: $($_.Exception.Message)" -foregroundcolor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)"  -ForegroundColor Red
    Exit
}
