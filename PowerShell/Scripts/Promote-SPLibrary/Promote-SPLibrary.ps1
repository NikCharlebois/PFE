Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue

function Promote-SPLibrary
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory="true")]
        [string] $Url,

        
        [Parameter(Mandatory="true")]
        [string] $Name,

        [Parameter(Mandatory="true")]
        [string] $DestinationSiteCollectionUrl
    )

    #region Prerequisites
    Write-Host "Validating the library against the predefined set of prerequisites..." -NoNewLine
    # Ensure site exists
    $web = Get-SPWeb $Url -ErrorAction SilentlyContinue
    if ($null -eq $web)
    {
        throw "The specified URL could not be found."
    }

    # Ensure library exists
    $library = $web.Lists.TryGetList($Name)
    if ($null -eq $library)
    {
        throw "The specified library doesn't exist."
    }

    # Query for Checked-Out files with no versions
    if ($library.CheckedOutFiles.Count -gt 0)
    {
        Write-Host "The specified library has files that are checked-out with no versions. Please ask the client to make sure they check-in the following files:" -ForegroundColor Red
        foreach ($file in $library.CheckedOutFiles)
        {
            Write-Host $file.Url -ForegroundColor Red
        }
        throw "Checked out files with no versions were found."
    }

    # Query for all checked-out files
    $CheckedOutCAML = "<Where><IsNotNull><FieldRef Name='CheckoutUser' /></IsNotNull></Where>"
    $camlQuery = New-Object Microsoft.SharePoint.SPQuery
    $camlQuery.Query = $CheckedOutCAML
    $checkedOutFiles = $library.GetItems($camlQuery)

    if ($checkedOutFiles.Count -gt 0)
    {
        Write-Host "The specified library has files that are checked-out. Please ask the client to make sure they check-in the following files:" -ForegroundColor Red
        foreach ($file in $checkedOutFiles)
        {
            Write-Host $file.Url -ForegroundColor Red
        }
        throw "Checked out files were found."
    }

    # Check for active workflows
    $filesWithActiveWorkflows
    foreach ($file in $library.Items)
    {
        foreach ($workflow in $file.Workflows)
        {
            if ($workflow.InternalState -eq "Running")
            {
                $filesWithActiveWorkflows += $file.Url
            }
        }
    }

    if ($filesWithActiveWorkflows.Count -gt 0)
    {
        Write-Host "It was detected that active workflows are currently running. Please make sure you either cancel or complete the workflows on the following files:" -ForegroundColor Red
        foreach ($entry in $filesWithActiveWorkflows)
        {
            Write-Host $entry -ForegroundColor Red
        }
        throw "Active Workflows were found."
    }

    Write-Host "Done" -ForegroundColor Green
    #endregion

    #region Save the source library as template (without content)
    $guidPart = (New-Guid).ToString().Split('-')[0]
    $templateName = "promote-" + $Name.Replace(" ", "") + "-" + $guidPart + ".stp"
    $library.SaveAsTemplate($templateName, $templateName.Split('.')[0], "Template created by the Promote-SPLibrary cmdlet", $false)
    $client = New-Object System.Net.WebClient
    $path = $env:TEMP + $templateName
    $client.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    $client.DownloadFile($web.Site.RootWeb.Url + "/_catalogs/lt/" + $templateName, $path)
    #endregion

    #region Create Destination Site Collection
    Write-Host "Creating the destination site collection..." -NoNewline
    $NewSite = Get-SPSite $DestinationSiteCollectionUrl -EA SilentlyContinue
    if ($null -ne $NewSite)
    {
        Write-Host ""
        $answer = Read-Host "A site already exists at $($DestinationSiteCollectionUrl). Do you wish to delete it first? (y/n)"
        if ("y" -eq $answer.ToLower() -and !($web.Url.ToLower().Contains($DestinationSiteCollectionUrl.ToLower())))
        {
            Remove-SPSite $DestinationSiteCollectionUrl -Confirm:$false
            $NewSite = New-SPSite -Url $DestinationSiteCollectionUrl -OwnerAlias ($env:USERDOMAIN + "\" + $env:USERNAME) -Template STS#0
        }
    }
    Write-Host "Done" -ForegroundColor Green
    #endregion

    #region Upload List Template to new site
    Write-Host "Uploading the list template to the new site collection..." -NoNewline
    $file = (Get-Item $path)
    $spfolder = $NewSite.RootWeb.getfolder("_catalogs\lt")
    $spfileCollection = $spfolder.Files
    $catch = $spfileCollection.Add($templateName, $file.OpenRead(), $true)
    Write-Host "Done" -ForegroundColor Green
    #endregion

    #region Create Destination Library
    Write-Host "Creating the destination library..." -NoNewLine
    $newList = $NewSite.RootWeb.Lists[$Name] 
    
    if ($null -ne $newList)
    {
        Write-Host ""
        $answer = Read-Host "A list named $($Name) already exists on the destination. Do you wish to delete it first? (y/n)"
        if ("y" -eq $answer.ToLower())
        {
            $newList.Delete()            
            $ListTemplate = $NewSite.GetCustomListTemplates($NewSite.RootWeb)[$templateName.SPlit('.')[0]]
            $newList = $NewSite.RootWeb.Lists.Add($Name, "Created with Promote-SPLibrary", $ListTemplate)
        }
    }
    Write-Host "Done" -ForegroundColor Green
    #endregion

    #region Move Files
    Write-Host "Moving files to the destination..." -NoNewLine
    $count = 1
    foreach ($file in $library.Items)
    {
        Write-Progress -Activity "Copying Files" -Status "Copying Files" -PercentComplete ($count / $library.Items.Count * 100)
        Write-Host $file["ows_EncodedAbsUrl"].ToString()
        Copy-SPFileWithVersions -SourceFileUrl $file["ows_EncodedAbsUrl"].ToString() -DestinationLibraryUrl $($NewSite.Url + "/" + $Name)
        $count ++
    }
    Write-Host "Done" -ForegroundColor Green
    #endregion

    Write-Host "Completed on $($NewSite.Url)"
}

function Add-SPWebPart
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Url,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]
        $Web,

        [Parameter(Mandatory = $true)]
        [string]
        $ZoneName,

        [Parameter(Mandatory = $true)]
        [string]
        $ZoneIndex,

        [Parameter(Mandatory = $true)]
        [string]
        $WebPartXML
    )

    $spWpManager = $web.GetLimitedWebPartManager($Url, [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
    $SR = New-Object System.IO.StringReader($WebPartXml)
    $XTR = New-Object System.Xml.XmlTextReader($SR)
    $Err = $null
    $WebPartDefinition = $spWpManager.ImportWebPart($XTR, [ref] $Err)
    $spWpManager.AddWebPart($WebPartDefinition, $ZoneName, $ZoneIndex)
}

function Copy-SPFileWithVersions
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $SourceFileUrl,
        
        [Parameter(Mandatory = $true)]
        [System.String]
        $DestinationLibraryUrl
    )
    Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue

    #region Source
    # Getting Source lists and files
    $SourceSite = New-Object Microsoft.SharePoint.SPSite $SourceFileUrl
    $SourceWeb = $SourceSite.OpenWeb()
    $SourceFile = $SourceWeb.GetFile($SourceFileUrl)
    $SourceSite.Dispose()
    $SourceWeb.Dispose()
    #endregion

    #region Destination
    # Getting Destination lists and files
    if(!$DestinationLibraryUrl.EndsWith("/"))
    {
        $DestinationLibraryUrl += "/"
    }
    $DestinationSite = New-Object Microsoft.SharePoint.SPSite $DestinationLibraryUrl
    $DestinationWeb = $DestinationSite.OpenWeb()
    $DestinationLibrary = $DestinationWeb.GetList($DestinationLibraryUrl)
    $DestinationFileUrl = $DestinationWeb.Url + "/" + $SourceFile.Url
    #endregion

    if (!$SourceFile.Item.COntentTypeID.ToString().StartsWith("0x0120D520"))
    {
        $userCreatedBy = $SourceFile.Author
        $dateCreatedOn = $SourceFile.TimeCreated.ToLocalTime()

        for ($i = 0; $i -le $SourceFile.Versions.Count; $i++)
        {
            $hashSourceProperties = $null
            $streamFile= $null
            $userModifiedBy= $null
            $dateModifiedOn= $null
            $strVersionComment = ""
            $majorVersion = $false
            $contact = $null
            $GoodToProceedWithCopy = $false
            $IsVersion = $false

            # If the version is not the last one (not the Published Version);
            if ($i -lt $SourceFile.Versions.Count) 
            {
                $SourceFileVersion = $SourceFile.Versions[$i + $SourceFile.Versions.Count - 1]
                if($null -ne $SourceFileVersion)
                {
                    #region Previous Versions
                    $IsVersion = $true
                    $hashSourceProperties = $SourceFileVersion.Properties

                    if ($i -eq 0) 
                    {
                        $userModifiedBy = $userCreatedBy
                    }
                    else
                    {
                        $userModifiedBy = $SourceFileVersion.CreatedBy
                    }

                    $dateModifiedOn = $SourceFileVersion.Created.ToLocalTime()
                    $strVersionComment = $SourceFileVersion.CheckInComment
            
                    if($SourceFileVersion.VersionLabel.EndsWith("0"))
                    {
                        $majorVersion = $true
                    }

                    if($hashSourceProperties["PublishingStartDate"] -ne $null)
                    {
                        try
                        {
                            $dateModifiedOn = [System.TimezoneInfo]::ConvertTimeToUtc(($hashSourceProperties["PublishingStartDate"].ToString()))
                        }
                        catch
                        {
                            $str = "Error changing the publishingstartdate for file: " + $SourceFile.Url
                            Write-Host $str
                            $dateModifiedOn = $SourceFileVersion.Created.ToLocalTime()
                        }
                    }  

                    $streamFile = $SourceFileVersion.OpenBinaryStream()
                    $GoodToProceedWithCopy = $true
                    #endregion
                }
            }
            else
            {
                #region Current Version
                if ($SourceFile.Item["Editor"] -ne $null)
                {
                    try
                    {
                        $userBy = $SourceFile.Item["Editor"]

                        $userValue = New-Object Microsoft.SharePoint.SPFieldUserValue($SourceWeb, $userBy)

                        if($userValue.User -ne $null -and $userValue.User -ne "")
                        {
                            $userModifiedBy = $DestinationWeb.EnsureUser($userValue.User)
                        }
                        else
                        {
                            $userModifiedBy = $DestinationWeb.EnsureUser($userValue.LookupValue)
                        }
                    }
                    catch
                    {
                        $str = "Error while obtaining the Edited by User for file: " + $SourceFile.Url
                        $userModifiedBy = $null
                    }
                }
                $dateModifiedOn = $SourceFile.Item["Modified"]

                if ($SourceFile.MinorVersion -eq 0)
                {
                    $majorVersion = $true
                }
                else
                {
                    $majorVersion = $false
                }

                $hashSourceProperties = $SourceFile.Properties
                $strVersionComment = $SourceFile.CheckInComment
                $streamFile = $SourceFile.OpenBinaryStream()
                $GoodToProceedWithCopy = $true
                #endregion
            }

            try
            {
                if ($hashSourceProperties.'display_urn:schemas-microsoft-com:office:office#PublishingContact' -ne $null)
                {
                    $contact = ($DestinationWeb.EnsureUser($hashSourceProperties.'display_urn:schemas-microsoft-com:office:office#PublishingContact')).ID
                }
            }
            catch
            {
                $str = "Error adding the user: " + $hashSourceProperties.'display_urn:schemas-microsoft-com:office:office#PublishingContact'
                Write-Host $str
            }

            try
            {
                if ($null -ne $userCreatedBy)
                {
                    $userCreatedByDest = $DestinationWeb.EnsureUser($userCreatedBy);
                }
            }
            catch
            {
                if ($null -ne $DefaultUser)
                {
                    $userCreatedByDest = $DestinationWeb.EnsureUser($DefaultUser);
                }
            }

            try
            {
                if ($null -ne $userModifiedBy)
                {
                    $userModifiedByDest = $DestinationWeb.EnsureUser($userModifiedBy);
                }
            }
            catch
            {
                if ($null -ne $DefaultUser)
                {
                    $userModifiedByDest = $DestinationWeb.EnsureUser($DefaultUser);
                }
            }
            finally
            {
                # Variable $GoodToProceedWithCopy is used as a defense mechanism to ensure that everything 
                # is ready to initiate the copy operation of a previous or current version.
                if ($GoodToProceedWithCopy)
                {
                    $DestinationFile = $DestinationLibrary.RootFolder.Files.Add(
                                            $DestinationFileUrl,
                                            $streamFile,
                                            $hashSourceProperties, 
                                            $userCreatedByDest, 
                                            $userModifiedByDest,
                                            $dateCreatedOn,
                                            $dateModifiedOn,
                                            $strVersionComment,
                                            $true)

                    if ($majorVersion)
                    {                    
                        $DestinationFile.Item.UpdateOverwriteVersion()

                        if($DestinationLibrary.EnableMinorVersions)
                        {
                            $DestinationFile.Publish($strVersionComment)
                        }

                        if($DestinationLibrary.EnableModeration)
                        {
                            $DestinationFile.Approve("Approved by Copy-SPFileWithVersions.")
                        }
                    }
                    else
                    {
                        $DestinationFile.Item["Created"] = $dateCreatedOn
                        $DestinationFile.Item["Modified"] = $dateModifiedOn
                        $DestinationFile.Item.UpdateOverwriteVersion()
                    }
                }
            }
        }
    }
    else
    { 
        #region Copy the Document Set
        $modifiedByColumnName = "Editor" 
        $modifiedDateColumnName = "Modified" 
        $createdByColumnName = "Author" 
        $createdDateColumnName = "Created" 
 
        #checkDocSetToCopy $SourceFile.Item.ParentList $DestinationLibrary $SourceFile.Name $DestinationWeb
        $docsetID = [Microsoft.SharePoint.SPBuiltInContentTypeId]::DocumentSet 
        [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]$docsetToMove = $null  
 
        $docSetListItem = $SourceFile.Item.ParentList.Items | where {$_['Title'] -eq $SourceFile.Name} 
        $docSetToMove =  [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::GetDocumentSet($docSetListItem.Folder) 
        $docsetContentTypeId = $docSetListItem.ContentType.Parent.Id 
 
        if($docsetToMove -ne $null -and $docsetContentTypeId -ne $null -and $docsetContentTypeId.ToString().StartsWith($docsetID.ToString()))  
        {
            Copy-SPDocumentSet $docsetToMove $DestinationLibrary $SourceFile.Item.ContentType.Name $DestinationWeb
        }
        #endregion

        #region Fix Web Part on the Document Set Home Page
        $pageUrl = $DestinationLibrary.RootFolder.Url + "/forms/" + $SourceFile.Item.ContentType.Name + "/docsethomepage.aspx"
        $ImageWebPartXml = '<WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">' +
          '<Title>Image</Title>' +
          '<FrameType>Default</FrameType>' +
          '<Description>Use to display pictures and photos.</Description>' +
          '<IsIncluded>true</IsIncluded>' +
          '<ZoneID>WebPartZone_TopLeft</ZoneID>' +
          '<PartOrder>0</PartOrder>' +
          '<FrameState>Normal</FrameState>' +
          '<Height />' +
          '<Width />' +
          '<AllowRemove>true</AllowRemove>' +
          '<AllowZoneChange>true</AllowZoneChange>' +
          '<AllowMinimize>true</AllowMinimize>' +
          '<AllowConnect>true</AllowConnect>' +
          '<AllowEdit>true</AllowEdit>' +
          '<AllowHide>true</AllowHide>' +
          '<IsVisible>true</IsVisible>' +
          '<DetailLink />' +
          '<HelpLink />' +
          '<HelpMode>Modeless</HelpMode>' +
          '<Dir>Default</Dir>' +
          '<PartImageSmall />' +
          '<MissingAssembly>Cannot import this Web Part.</MissingAssembly>' +
          '<PartImageLarge />' +
          '<IsIncludedFilter />' +
          '<Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>' +
          '<TypeName>Microsoft.SharePoint.WebPartPages.ImageWebPart</TypeName>' +
          '<ImageLink xmlns="http://schemas.microsoft.com/WebPart/v2/Image">/_layouts/images/docset_welcomepage_big.png</ImageLink>' +
          '<AlternativeText xmlns="http://schemas.microsoft.com/WebPart/v2/Image" />' +
          '<VerticalAlignment xmlns="http://schemas.microsoft.com/WebPart/v2/Image">Middle</VerticalAlignment>' +
          '<HorizontalAlignment xmlns="http://schemas.microsoft.com/WebPart/v2/Image">Center</HorizontalAlignment>' +
          '<BackgroundColor xmlns="http://schemas.microsoft.com/WebPart/v2/Image">transparent</BackgroundColor>' +
        '</WebPart>';

        Add-SPWebPart -Url $pageUrl -Web $NewSite.RootWeb -ZoneName "WebPartZone_TopLeft" -ZoneIndex "0" -WebPartXML $ImageWebPartXml

        $DoSetPropertiesWebPartXml = '<WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">' +
          '<Title>Document Set Properties</Title>' +
          '<FrameType>Default</FrameType>' +
          '<Description>Displays the properties of the Document Set.</Description>' +
          '<IsIncluded>true</IsIncluded>' +
          '<ZoneID>WebPartZone_Top</ZoneID>' +
          '<PartOrder>0</PartOrder>' +
          '<FrameState>Normal</FrameState>' +
          '<Height />' +
          '<Width />' +
          '<AllowRemove>true</AllowRemove>' +
          '<AllowZoneChange>true</AllowZoneChange>' +
          '<AllowMinimize>true</AllowMinimize>' +
          '<AllowConnect>true</AllowConnect>' +
          '<AllowEdit>true</AllowEdit>' +
          '<AllowHide>true</AllowHide>' +
          '<IsVisible>true</IsVisible>' +
          '<DetailLink />' +
          '<HelpLink />' +
          '<HelpMode>Modeless</HelpMode>' +
          '<Dir>Default</Dir>' +
          '<PartImageSmall />' +
          '<MissingAssembly>Cannot import this Web Part.</MissingAssembly>' +
          '<PartImageLarge>/_layouts/15/images/msimagel.gif</PartImageLarge>' +
          '<IsIncludedFilter />' +
          '<DisplayText>' +
          '</DisplayText>' +
          '<Assembly>Microsoft.Office.DocumentManagement, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>' +
          '<TypeName>Microsoft.Office.Server.WebControls.DocumentSetPropertiesWebPart</TypeName>' +
        '</WebPart>';

        Add-SPWebPart -Url $pageUrl -Web $NewSite.RootWeb -ZoneName "WebPartZone_Top" -ZoneIndex "0" -WebPartXML $DoSetPropertiesWebPartXml

        $DocSetContentwebPartXml = '<WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">' +
          '<Title>Document Set Contents</Title>' +
          '<FrameType>Default</FrameType>' +
          '<Description>Displays the contents of the Document Set.</Description>' +
          '<IsIncluded>true</IsIncluded>' +
          '<ZoneID>WebPartZone_CenterMain</ZoneID>' +
          '<PartOrder>0</PartOrder>' +
          '<FrameState>Normal</FrameState>' +
          '<Height />' +
          '<Width />' +
          '<AllowRemove>true</AllowRemove>' +
          '<AllowZoneChange>true</AllowZoneChange>' +
          '<AllowMinimize>true</AllowMinimize>' +
          '<AllowConnect>true</AllowConnect>' +
          '<AllowEdit>true</AllowEdit>' +
          '<AllowHide>true</AllowHide>' +
          '<IsVisible>true</IsVisible>' +
          '<DetailLink />' +
          '<HelpLink />' +
          '<HelpMode>Modeless</HelpMode>' +
          '<Dir>Default</Dir>' +
          '<PartImageSmall />' +
          '<MissingAssembly>Cannot import this Web Part.</MissingAssembly>' +
          '<PartImageLarge>/_layouts/15/images/msimagel.gif</PartImageLarge>' +
          '<IsIncludedFilter />' +
          '<DisplayText>' +
          '</DisplayText>' +
          '<Assembly>Microsoft.Office.DocumentManagement, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>' +
          '<TypeName>Microsoft.Office.Server.WebControls.DocumentSetContentsWebPart</TypeName>' +
        '</WebPart>';

        Add-SPWebPart -Url $pageUrl -Web $NewSite.RootWeb -ZoneName "WebPartZone_CenterMai" -ZoneIndex "0" -WebPartXML $DocSetContentwebPartXml
        #endregion
    }
}

function Copy-SPDocumentSet
{ 
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]
        $docset,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPList]
        $moveToList,

        [Parameter(Mandatory = $true)]
        [string]
        $docsetContentType,

        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]
        $destinationWeb
    )
    [Microsoft.SharePoint.SPFolder]$targetFolder = $destinationWeb.GetFolder($moveToList.RootFolder.ServerRelativeUrl + "/" + $docSet.Item.Name)    
    if (!$targetFolder.Exists) 
    {          
        $rootFolder = $moveToList.RootFolder 
        $cType = $moveToList.ContentTypes[$docsetContentType]  
        [Byte[]]$compressedFile = $docset.Export() 
        $user = $moveToList.ParentWeb.EnsureUser($env:UserDomain + "\" + $env:UserName) 
         
        $properties = new-object System.Collections.Hashtable 
         
        $AllFields = $docset.Item.Fields | ? {!($_.sealed)}  
  
        foreach ($Field in $AllFields)  
        {  
            if (!$Field.ReadOnlyField) 
            { 
                $propValue = $docset.Item[$Field.Title] 
                 
                if ($docset.Item.Properties[$Field.Title] -eq $null) 
                { 
                    $propValue = "" 
                } 
                  
                if (!$properties.ContainsKey($Field.Title))  
                { 
                    $properties.Add($Field.Title,$propValue) 
                }
                else 
                { 
                    $properties[$Field.Title] = $propValue  
                }  
            } 
        }  
        
        $newDocumentSet = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Import($compressedFile, $docset.Item.Name, $rootFolder, $cType.Id, $properties, $user)
        Update-SPDocumentSet $docset $moveToList $newDocumentSet.Item.ID 
    } 
 
} 

function Update-SPDocumentSet
{ 
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]
        $docset,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPList]
        $moveToList,

        [Parameter(Mandatory = $true)]
        [string]
        $docsetId
    )
    $docSetListItem = $moveToList.Items | where {$_['ID'] -eq $docsetId} 
    $newdocset =  [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::GetDocumentSet($docSetListItem.Folder) 
     
    $item = $docset.Item 
    $newItem = $newdocset.Item 
 
    if ($item.Folder -ne $null) 
    { 
        foreach ($spFile in $item.Folder.Files) 
        { 
            if ($newItem.Folder -ne $null) 
            { 
                foreach ($newspFile in $newItem.Folder.Files) 
                { 
                    if ($spFile.Name -ne $newspFile.Name) 
                    { 
                        continue 
                    } 
                    try 
                    {  
                        $newspFile.Item[$modifiedDateColumnName] = $spFile.Item[$modifiedDateColumnName] 
                        $newspFile.Item[$modifiedByColumnName] = $spFile.Item[$modifiedByColumnName] 
                        $newspFile.Item[$createdDateColumnName] = $spFile.Item[$createdDateColumnName] 
                        $newspFile.Item[$createdByColumnName] = $spFile.Item[$createdByColumnName] 
                          
                        $newspFile.Item.Update() 
                        break 
                    } 
                    catch 
                    { 
                        Write-Host "Caught the exception : " $_.Exception.Message 
                        Write-Host $newItem.Name "|" $spFile.Name  
                    } 
                } 
            }    
        } 
    } 
     
    try 
    { 
        $newItem[$modifiedDateColumnName] = $item[$modifiedDateColumnName] 
        $newItem[$modifiedByColumnName] = $item[$modifiedByColumnName] 
        $newItem[$createdDateColumnName] = $item[$createdDateColumnName] 
        $newItem[$createdByColumnName] = $item[$createdByColumnName] 
         
        $newItem.Update() 
    } 
    catch 
    { 
        Write-Host "Caught the exception : " $_.Exception.Message 
        Write-Host $newItem.Name 
    } 
} 

$guid = New-Guid
$SourceWebUrl = "http://osfi-sharepoint.canadaeast.cloudapp.azure.com/sites/pension/B002/"
$SourceLibraryName = "Risk Assessment and Intervention"
$DestinationWebUrl = "http://osfi-sharepoint.canadaeast.cloudapp.azure.com/"#sites/$($guid.ToString().Split('-')[0])"
Promote-SPLibrary -Url $SourceWebUrl -Name $SourceLibraryName -DestinationSiteCollectionUrl $DestinationWebUrl