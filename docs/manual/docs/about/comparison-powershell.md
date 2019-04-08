# Comparison to SharePoint and Office 365 PowerShell

Following table lists the different Office 365 CLI commands and how they map to PowerShell cmdlets for SharePoint and Office 365.

PowerShell Cmdlet|Source|Office 365 CLI command
-----------------|------|----------------------
Add-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Add-SPOHubSiteAssociation|Microsoft.Online.SharePoint.PowerShell|[spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md)
Add-SPOSiteCollectionAppCatalog|Microsoft.Online.SharePoint.PowerShell|[spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md)
Add-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md)
Add-SPOSiteDesignTask|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Add-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript add](../cmd/spo/sitescript/sitescript-add.md)
Add-SPOTenantCdnOrigin|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md)
Add-SPOTenantCentralAssetRepositoryLibrary|Microsoft.Online.SharePoint.PowerShell|
Add-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme set](../cmd/spo/theme/theme-set.md)
Add-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Approve-SPOTenantServicePrincipalPermissionGrant|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant add](../cmd/spo/serviceprincipal/serviceprincipal-grant-add.md)
Approve-SPOTenantServicePrincipalPermissionRequest|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)
Connect-SPOService|Microsoft.Online.SharePoint.PowerShell|[spo login](../cmd/spo/login.md)
ConvertTo-SPOMigrationEncryptedPackage|Microsoft.Online.SharePoint.PowerShell|
ConvertTo-SPOMigrationTargetedPackage|Microsoft.Online.SharePoint.PowerShell|
Deny-SPOTenantServicePrincipalPermissionRequest|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)
Disable-SPOTenantServicePrincipal|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Disconnect-SPOService|Microsoft.Online.SharePoint.PowerShell|[spo logout](../cmd/spo/logout.md)
Enable-SPOTenantServicePrincipal|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Export-SPOUserInfo|Microsoft.Online.SharePoint.PowerShell|
Export-SPOUserProfile|Microsoft.Online.SharePoint.PowerShell|
Get-SPOAppErrors|Microsoft.Online.SharePoint.PowerShell|
Get-SPOAppInfo|Microsoft.Online.SharePoint.PowerShell|
Get-SPOBrowserIdleSignOut|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoMovedUsers|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoMoveReport|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoUsers|Microsoft.Online.SharePoint.PowerShell|
Get-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Get-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|
Get-SPOExternalUser|Microsoft.Online.SharePoint.PowerShell|[spo externaluser list](../cmd/spo/externaluser/externaluser-list.md)
Get-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Get-SPOGeoMoveCrossCompatibilityStatus|Microsoft.Online.SharePoint.PowerShell|
Get-SPOGeoStorageQuota|Microsoft.Online.SharePoint.PowerShell|
Get-SPOHideDefaultThemes|Microsoft.Online.SharePoint.PowerShell|[spo hidedefaultthemes get](../cmd/spo/hidedefaultthemes/hidedefaultthemes-get.md)
Get-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite get](../cmd/spo/hubsite/hubsite-get.md), [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md)
Get-SPOMigrationJobProgress|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMigrationJobStatus|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMultiGeoExperience|Microsoft.Online.SharePoint.PowerShell|
Get-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Get-SPOPublicCdnOrigins|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site classic list](../cmd/spo/site/site-classic-list.md)
Get-SPOSiteContentMoveState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteDataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md), [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md)
Get-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md)
Get-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteRenameState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript get](../cmd/spo/sitescript/sitescript-get.md), [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md)
Get-SPOSiteScriptFromList|Microsoft.Online.SharePoint.PowerShell|[spo list sitescript get](../cmd/spo/list/list-sitescript-get.md)
Get-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity get](../cmd/spo/storageentity/storageentity-get.md), [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md)
Get-SPOTenant|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantCdnEnabled|Microsoft.Online.SharePoint.PowerShell|[spo cdn get](../cmd/spo/cdn/cdn-get.md)
Get-SPOTenantCdnOrigins|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin list](../cmd/spo/cdn/cdn-origin-list.md)
Get-SPOTenantCdnPolicies|Microsoft.Online.SharePoint.PowerShell|[spo cdn policy list](../cmd/spo/cdn/cdn-policy-list.md)
Get-SPOTenantCentralAssetRepository|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantContentTypeReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantLogEntry|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantLogLastAvailableTimeInUtc|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantServicePrincipalPermissionGrants|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md)
Get-SPOTenantServicePrincipalPermissionRequests|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md)
Get-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantTaxonomyReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme list](../cmd/spo/theme/theme-list.md)
Get-SPOUnifiedGroup|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUserAndContentMoveState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUserOneDriveLocation|Microsoft.Online.SharePoint.PowerShell|
Get-SPOWebTemplate|Microsoft.Online.SharePoint.PowerShell|
Grant-SPOHubSiteRights|Microsoft.Online.SharePoint.PowerShell|[spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md)
Grant-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md)
Invoke-SPOMigrationEncryptUploadSubmit|Microsoft.Online.SharePoint.PowerShell|
Invoke-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
New-SPOMigrationEncryptionParameters|Microsoft.Online.SharePoint.PowerShell|
New-SPOMigrationPackage|Microsoft.Online.SharePoint.PowerShell|
New-SPOPublicCdnOrigin|Microsoft.Online.SharePoint.PowerShell|
New-SPOSdnProvider|Microsoft.Online.SharePoint.PowerShell|
New-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site classic add](../cmd/spo/site/site-classic-add.md)
New-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Register-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Register-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite register](../cmd/spo/hubsite/hubsite-register.md)
Remove-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOExternalUser|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOHubSiteAssociation|Microsoft.Online.SharePoint.PowerShell|[spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md)
Remove-SPOMigrationJob|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOPublicCdnOrigin|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSdnProvider|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site classic remove](../cmd/spo/site/site-classic-remove.md)
Remove-SPOSiteCollectionAppCatalog|Microsoft.Online.SharePoint.PowerShell|[spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md)
Remove-SPOSiteCollectionAppCatalogById|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md)
Remove-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md)
Remove-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md)
Remove-SPOTenantCdnOrigin|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin remove](../cmd/spo/cdn/cdn-origin-remove.md)
Remove-SPOTenantCentralAssetRepositoryLibrary|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme remove](../cmd/spo/theme/theme-remove.md)
Remove-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOUserInfo|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOUserProfile|Microsoft.Online.SharePoint.PowerShell|
Repair-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Request-SPOPersonalSite|Microsoft.Online.SharePoint.PowerShell|
Request-SPOUpgradeEvaluationSite|Microsoft.Online.SharePoint.PowerShell|
Restore-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Restore-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|
Revoke-SPOHubSiteRights|Microsoft.Online.SharePoint.PowerShell|[spo hubsite rights revoke](../cmd/spo/hubsite/hubsite-rights-revoke.md)
Revoke-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md)
Revoke-SPOTenantServicePrincipalPermission|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)
Revoke-SPOUserSession|Microsoft.Online.SharePoint.PowerShell|
Set-SPOBrowserIdleSignOut|Microsoft.Online.SharePoint.PowerShell|
Set-SPOGeoStorageQuota|Microsoft.Online.SharePoint.PowerShell|
Set-SPOHideDefaultThemes|Microsoft.Online.SharePoint.PowerShell|[spo hidedefaultthemes set](../cmd/spo/hidedefaultthemes/hidedefaultthemes-set.md)
Set-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite set](../cmd/spo/hubsite/hubsite-set.md)
Set-SPOMigrationPackageAzureSource|Microsoft.Online.SharePoint.PowerShell|
Set-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Set-SPOMultiGeoExperience|Microsoft.Online.SharePoint.PowerShell|
Set-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md)
Set-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSiteOffice365Group|Microsoft.Online.SharePoint.PowerShell|[spo site o365group set](../cmd/spo/site/site-o365group-set.md)
Set-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|
Set-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity set](../cmd/spo/storageentity/storageentity-set.md)
Set-SPOTenant|Microsoft.Online.SharePoint.PowerShell|[spo tenant settings set](../cmd/spo/tenant/tenant-settings-set.md)
Set-SPOTenantCdnEnabled|Microsoft.Online.SharePoint.PowerShell|[spo cdn set](../cmd/spo/cdn/cdn-set.md)
Set-SPOTenantCdnPolicy|Microsoft.Online.SharePoint.PowerShell|[spo cdn policy set](../cmd/spo/cdn/cdn-policy-set.md)
Set-SPOTenantContentTypeReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Set-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Set-SPOTenantTaxonomyReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Set-SPOUnifiedGroup|Microsoft.Online.SharePoint.PowerShell|
Set-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Set-SPOWebTheme|Microsoft.Online.SharePoint.PowerShell|
Start-SPOSiteContentMove|Microsoft.Online.SharePoint.PowerShell|
Start-SPOUserAndContentMove|Microsoft.Online.SharePoint.PowerShell|
Stop-SPOSiteContentMove|Microsoft.Online.SharePoint.PowerShell|
Stop-SPOUserAndContentMove|Microsoft.Online.SharePoint.PowerShell|
Submit-SPOMigrationJob|Microsoft.Online.SharePoint.PowerShell|
Test-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Unregister-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md)
Update-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Update-UserType|Microsoft.Online.SharePoint.PowerShell|
Upgrade-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Add-PnPApp|SharePointPnPPowerShellOnline|[spo app add](../cmd/spo/app/app-add.md)
Add-PnPTenantCdnOrigin|SharePointPnPPowerShellOnline|[spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md)
Add-PnPClientSidePage|SharePointPnPPowerShellOnline|[spo page add](../cmd/spo/page/page-add.md)
Add-PnPClientSidePageSection|SharePointPnPPowerShellOnline|[spo page section add](../cmd/spo/page/page-section-add.md)
Add-PnPClientSideText|SharePointPnPPowerShellOnline|
Add-PnPClientSideWebPart|SharePointPnPPowerShellOnline|[spo page clientsidewebpart add](../cmd/spo/page/page-clientsidewebpart-add.md)
Add-PnPContentType|SharePointPnPPowerShellOnline|[spo contenttype add](../cmd/spo/contenttype/contenttype-add.md)
Add-PnPContentTypeToDocumentSet|SharePointPnPPowerShellOnline|
Add-PnPContentTypeToList|SharePointPnPPowerShellOnline|[spo list contenttype add](../cmd/spo/list/list-contenttype-add.md)
Add-PnPCustomAction|SharePointPnPPowerShellOnline|[spo customaction add](../cmd/spo/customaction/customaction-add.md)
Add-PnPDataRowsToProvisioningTemplate|SharePointPnPPowerShellOnline|
Add-PnPDocumentSet|SharePointPnPPowerShellOnline|
Add-PnPEventReceiver|SharePointPnPPowerShellOnline|
Add-PnPField|SharePointPnPPowerShellOnline|
Add-PnPFieldFromXml|SharePointPnPPowerShellOnline|[spo field add](../cmd/spo/field/field-add.md)
Add-PnPFieldToContentType|SharePointPnPPowerShellOnline|[spo contenttype field set](../cmd/spo/contenttype/contenttype-field-set.md)
Add-PnPFile|SharePointPnPPowerShellOnline|[spo file add](../cmd/spo/file/file-add.md)
Add-PnPFileToProvisioningTemplate|SharePointPnPPowerShellOnline|
Add-PnPFolder|SharePointPnPPowerShellOnline|[spo folder add](../cmd/spo/folder/folder-add.md)
Add-PnPHtmlPublishingPageLayout|SharePointPnPPowerShellOnline|
Add-PnPIndexedProperty|SharePointPnPPowerShellOnline|
Add-PnPJavaScriptBlock|SharePointPnPPowerShellOnline|
Add-PnPJavaScriptLink|SharePointPnPPowerShellOnline|
Add-PnPListFoldersToProvisioningTemplate|SharePointPnPPowerShellOnline|
Add-PnPListItem|SharePointPnPPowerShellOnline|[spo listitem add](../cmd/spo/listitem/listitem-add.md)
Add-PnPMasterPage|SharePointPnPPowerShellOnline|
Add-PnPNavigationNode|SharePointPnPPowerShellOnline|[spo navigation node add](../cmd/spo/navigation/navigation-node-add.md)
Add-PnPOffice365GroupToSite|SharePointPnPPowerShellOnline|
Add-PnPProvisioningSequence|SharePointPnPPowerShellOnline|
Add-PnPProvisioningSite|SharePointPnPPowerShellOnline|
Add-PnPProvisioningSubSite|SharePointPnPPowerShellOnline|
Add-PnPPublishingImageRendition|SharePointPnPPowerShellOnline|
Add-PnPPublishingPage|SharePointPnPPowerShellOnline|
Add-PnPPublishingPageLayout|SharePointPnPPowerShellOnline|
Add-PnPRoleDefinition|SharePointPnPPowerShellOnline|
Add-PnPSiteClassification|SharePointPnPPowerShellOnline|
Add-PnPSiteCollectionAdmin|SharePointPnPPowerShellOnline|
Add-PnPSiteCollectionAppCatalog|SharePointPnPPowerShellOnline|[spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md)
Add-PnPSiteDesign|SharePointPnPPowerShellOnline|[spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md)
Add-PnPSiteDesignTask|SharePointPnPPowerShellOnline|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Add-PnPSiteScript|SharePointPnPPowerShellOnline|[spo sitescript add](../cmd/spo/sitescript/sitescript-add.md)
Add-PnPStoredCredential|SharePointPnPPowerShellOnline|
Add-PnPTenantTheme|SharePointPnPPowerShellOnline|[spo theme set](../cmd/spo/theme/theme-set.md)
Add-PnPTaxonomyField|SharePointPnPPowerShellOnline|
Add-PnPUserToGroup|SharePointPnPPowerShellOnline|
Add-PnPView|SharePointPnPPowerShellOnline|
Add-PnPWebhookSubscription|SharePointPnPPowerShellOnline|[spo list webhook add](../cmd/spo/list/list-webhook-add.md)
Add-PnPWebPartToWebPartPage|SharePointPnPPowerShellOnline|
Add-PnPWebPartToWikiPage|SharePointPnPPowerShellOnline|
Add-PnPWikiPage|SharePointPnPPowerShellOnline|
Add-PnPWorkflowDefinition|SharePointPnPPowerShellOnline|
Add-PnPWorkflowSubscription|SharePointPnPPowerShellOnline|
Apply-PnPProvisioningHierarchy|SharePointPnPPowerShellOnline|
Apply-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
Approve-PnPTenantServicePrincipalPermissionRequest|SharePointPnPPowerShellOnline|[spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)
Clear-PnPListItemAsRecord|SharePointPnPPowerShellOnline|[spo listitem record undeclare](../cmd/spo/listitem/listitem-record-undeclare.md)
Clear-PnPRecycleBinItem|SharePointPnPPowerShellOnline|
Clear-PnPTenantRecycleBinItem|SharePointPnPPowerShellOnline|
Connect-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md)
Connect-PnPMicrosoftGraph|SharePointPnPPowerShellOnline|
Connect-PnPOnline|SharePointPnPPowerShellOnline|[spo login](../cmd/spo/login.md)
Convert-PnPFolderToProvisioningTemplate|SharePointPnPPowerShellOnline|
Convert-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
ConvertTo-PnPClientSidePage|SharePointPnPPowerShellOnline|
Copy-PnPFile|SharePointPnPPowerShellOnline|[spo file copy](../cmd/spo/file/file-copy.md), [spo folder copy](../cmd/spo/folder/folder-copy.md)
Copy-PnPItemProxy|SharePointPnPPowerShellOnline|
Deny-PnPTenantServicePrincipalPermissionRequest|SharePointPnPPowerShellOnline|[spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)
Disable-PnPFeature|SharePointPnPPowerShellOnline|
Disable-PnPInPlaceRecordsManagementForSite|SharePointPnPPowerShellOnline|[spo site inplacerecordsmanagement set](../cmd/spo/site/site-inplacerecordsmanagement-set.md)
Disable-PnPResponsiveUI|SharePointPnPPowerShellOnline|
Disable-PnPSiteClassification|SharePointPnPPowerShellOnline|[graph siteclassification disable](../cmd/graph/siteclassification/siteclassification-disable.md)
Disable-PnPTenantServicePrincipal|SharePointPnPPowerShellOnline|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Disconnect-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md)
Disconnect-PnPOnline|SharePointPnPPowerShellOnline|[spo logout](../cmd/spo/logout.md)
Enable-PnPFeature|SharePointPnPPowerShellOnline|
Enable-PnPInPlaceRecordsManagementForSite|SharePointPnPPowerShellOnline|
Enable-PnPResponsiveUI|SharePointPnPPowerShellOnline|
Enable-PnPSiteClassification|SharePointPnPPowerShellOnline|[graph siteclassification enable](../cmd/graph/siteclassification/siteclassification-enable.md)
Enable-PnPTenantServicePrincipal|SharePointPnPPowerShellOnline|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Ensure-PnPFolder|SharePointPnPPowerShellOnline|
Execute-PnPQuery|SharePointPnPPowerShellOnline|
Export-PnPTaxonomy|SharePointPnPPowerShellOnline|
Export-PnPTermGroupToXml|SharePointPnPPowerShellOnline|
Find-PnPFile|SharePointPnPPowerShellOnline|
Get-PnPAccessToken|SharePointPnPPowerShellOnline|
Get-PnPApp|SharePointPnPPowerShellOnline|[spo app get](../cmd/spo/app/app-get.md), [spo app list](../cmd/spo/app/app-list.md)
Get-PnPAppAuthAccessToken|SharePointPnPPowerShellOnline|
Get-PnPAppInstance|SharePointPnPPowerShellOnline|
Get-PnPAuditing|SharePointPnPPowerShellOnline|
Get-PnPAuthenticationRealm|SharePointPnPPowerShellOnline|
Get-PnPAvailableClientSideComponents|SharePointPnPPowerShellOnline|
Get-PnPAzureADManifestKeyCredentials|SharePointPnPPowerShellOnline|
Get-PnPAzureCertificate|SharePointPnPPowerShellOnline|
Get-PnPClientSideComponent|SharePointPnPPowerShellOnline|
Get-PnPClientSidePage|SharePointPnPPowerShellOnline|[spo page get](../cmd/spo/page/page-get.md), [spo page control list](../cmd/spo/page/page-control-list.md), [spo page control get](../cmd/spo/page/page-control-get.md), [spo page section get](../cmd/spo/page/page-section-get.md), [spo page section list](../cmd/spo/page/page-section-list.md), [spo page column get](../cmd/spo/page/page-column-get.md), [spo page column list](../cmd/spo/page/page-column-list.md), [spo page text add](../cmd/spo/page/page-text-add.md)
Get-PnPConnection|SharePointPnPPowerShellOnline|
Get-PnPContentType|SharePointPnPPowerShellOnline|[spo contenttype get](../cmd/spo/contenttype/contenttype-get.md), [spo list contenttype list](../cmd/spo/list/list-contenttype-list.md)
Get-PnPContentTypePublishingHubUrl|SharePointPnPPowerShellOnline|
Get-PnPContext|SharePointPnPPowerShellOnline|
Get-PnPCustomAction|SharePointPnPPowerShellOnline|[spo customaction get](../cmd/spo/customaction/customaction-get.md), [spo customaction list](../cmd/spo/customaction/customaction-list.md)
Get-PnPDefaultColumnValues|SharePointPnPPowerShellOnline|
Get-PnPDocumentSetTemplate|SharePointPnPPowerShellOnline|
Get-PnPEventReceiver|SharePointPnPPowerShellOnline|
Get-PnPException|SharePointPnPPowerShellOnline|
Get-PnPFeature|SharePointPnPPowerShellOnline|[spo feature list](../cmd/spo/feature/feature-list.md)
Get-PnPField|SharePointPnPPowerShellOnline|[spo field get](../cmd/spo/field/field-get.md)
Get-PnPFile|SharePointPnPPowerShellOnline|[spo file get](../cmd/spo/file/file-get.md), [spo file list](../cmd/spo/file/file-list.md)
Get-PnPFolder|SharePointPnPPowerShellOnline|[spo folder get](../cmd/spo/folder/folder-get.md), [spo folder list](../cmd/spo/folder/folder-list.md)
Get-PnPFolderItem|SharePointPnPPowerShellOnline|
Get-PnPGroup|SharePointPnPPowerShellOnline|
Get-PnPGroupMembers|SharePointPnPPowerShellOnline|
Get-PnPGroupPermissions|SharePointPnPPowerShellOnline|
Get-PnPHealthScore|SharePointPnPPowerShellOnline|
Get-PnPHideDefaultThemes|SharePointPnPPowerShellOnline|[spo hidedefaultthemes get](../cmd/spo/hidedefaultthemes/hidedefaultthemes-get.md)
Get-PnPHomePage|SharePointPnPPowerShellOnline|
Get-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite get](../cmd/spo/hubsite/hubsite-get.md), [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md)
Get-PnPIndexedPropertyKeys|SharePointPnPPowerShellOnline|
Get-PnPInformationRightsManagement|SharePointPnPPowerShellOnline|
Get-PnPInPlaceRecordsManagement|SharePointPnPPowerShellOnline|
Get-PnPJavaScriptLink|SharePointPnPPowerShellOnline|
Get-PnPLabel|SharePointPnPPowerShellOnline|[spo list label get](../cmd/spo/list/list-label-get.md)
Get-PnPList|SharePointPnPPowerShellOnline|[spo list get](../cmd/spo/list/list-get.md), [spo list list](../cmd/spo/list/list-list.md)
Get-PnPListItem|SharePointPnPPowerShellOnline|[spo listitem get](../cmd/spo/listitem/listitem-get.md), [spo listitem list](../cmd/spo/listitem/listitem-list.md)
Get-PnPListRecordDeclaration|SharePointPnPPowerShellOnline|
Get-PnPMasterPage|SharePointPnPPowerShellOnline|
Get-PnPNavigationNode|SharePointPnPPowerShellOnline|[spo navigation node list](../cmd/spo/navigation/navigation-node-list.md)
Get-PnPProperty|SharePointPnPPowerShellOnline|
Get-PnPPropertyBag|SharePointPnPPowerShellOnline|[spo propertybag get](../cmd/spo/propertybag/propertybag-get.md), [spo propertybag list](../cmd/spo/propertybag/propertybag-list.md)
Get-PnPProvisioningSite|SharePointPnPPowerShellOnline|
Get-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
Get-PnPProvisioningTemplateFromGallery|SharePointPnPPowerShellOnline|
Get-PnPPublishingImageRendition|SharePointPnPPowerShellOnline|
Get-PnPRecycleBinItem|SharePointPnPPowerShellOnline|
Get-PnPRequestAccessEmails|SharePointPnPPowerShellOnline|
Get-PnPRoleDefinition|SharePointPnPPowerShellOnline|
Get-PnPSearchConfiguration|SharePointPnPPowerShellOnline|
Get-PnPSearchCrawlLog|SharePointPnPPowerShellOnline|
Get-PnPSite|SharePointPnPPowerShellOnline|[spo site get](../cmd/spo/site/site-get.md), [spo site list](../cmd/spo/site/site-list.md)
Get-PnPSiteClassification|SharePointPnPPowerShellOnline|[graph siteclassification get](../cmd/graph/siteclassification/siteclassification-get.md)
Get-PnPSiteClosure|SharePointPnPPowerShellOnline|
Get-PnPSiteCollectionAdmin|SharePointPnPPowerShellOnline|
Get-PnPSiteCollectionTermStore|SharePointPnPPowerShellOnline|
Get-PnPSiteDesign|SharePointPnPPowerShellOnline|[spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md), [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md)
Get-PnPSiteDesignRights|SharePointPnPPowerShellOnline|[spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md)
Get-PnPSiteDesignRun|SharePointPnPPowerShellOnline|[spo sitedesign run list](../cmd/spo/sitedesign/sitedesign-run-list.md)
Get-PnPSiteDesignRunStatus|SharePointPnPPowerShellOnline|[spo sitedesign run status get](../cmd/spo/sitedesign/sitedesign-run-status-get.md)
Get-PnPSiteDesignTask|SharePointPnPPowerShellOnline|[spo sitedesign task get](../cmd/spo/sitedesign/sitedesign-task-get.md), [spo sitedesign task list](../cmd/spo/sitedesign/sitedesign-task-list.md)
Get-PnPSitePolicy|SharePointPnPPowerShellOnline|
Get-PnPSiteScript|SharePointPnPPowerShellOnline|[spo sitescript get](../cmd/spo/sitescript/sitescript-get.md), [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md)
Get-PnPSiteSearchQueryResults|SharePointPnPPowerShellOnline|
Get-PnPStorageEntity|SharePointPnPPowerShellOnline|[spo storageentity get](../cmd/spo/storageentity/storageentity-get.md), [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md)
Get-PnPStoredCredential|SharePointPnPPowerShellOnline|
Get-PnPSubWebs|SharePointPnPPowerShellOnline|
Get-PnPTaxonomyItem|SharePointPnPPowerShellOnline|
Get-PnPTaxonomySession|SharePointPnPPowerShellOnline|
Get-PnPTenant|SharePointPnPPowerShellOnline|[spo tenant settings list](../cmd/spo/tenant/tenant-settings-list.md)
Get-PnPTenantAppCatalogUrl|SharePointPnPPowerShellOnline|[spo tenant appcatalogurl get](../cmd/spo/tenant/tenant-appcatalogurl-get.md)
Get-PnPTenantCdnEnabled|SharePointPnPPowerShellOnline|[spo cdn get](../cmd/spo/cdn/cdn-get.md)
Get-PnPTenantCdnOrigin|SharePointPnPPowerShellOnline|[spo cdn origin list](../cmd/spo/cdn/cdn-origin-list.md)
Get-PnPTenantCdnPolicies|SharePointPnPPowerShellOnline|[spo cdn policy list](../cmd/spo/cdn/cdn-policy-list.md)
Get-PnPTenantRecycleBinItem|SharePointPnPPowerShellOnline|
Get-PnPTenantServicePermissionGrants|SharePointPnPPowerShellOnline|[spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md)
Get-PnPTenantServicePermissionRequests|SharePointPnPPowerShellOnline|[spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md)
Get-PnPTenantServicePrincipal|SharePointPnPPowerShellOnline|
Get-PnPTenantSite|SharePointPnPPowerShellOnline|[spo site get](../cmd/spo/site/site-get.md), [spo site classic list](../cmd/spo/site/site-classic-list.md)
Get-PnPTenantTheme|SharePointPnPPowerShellOnline|[spo theme get](../cmd/spo/theme/theme-get.md), [spo theme list](../cmd/spo/theme/theme-list.md)
Get-PnPTerm|SharePointPnPPowerShellOnline|[spo term get](../cmd/spo/term/term-get.md), [spo term list](../cmd/spo/term/term-list.md)
Get-PnPTermGroup|SharePointPnPPowerShellOnline|[spo term group get](../cmd/spo/term/term-group-get.md), [spo term group list](../cmd/spo/term/term-group-list.md)
Get-PnPTermSet|SharePointPnPPowerShellOnline|[spo term set get](../cmd/spo/term/term-set-get.md), [spo term set list](../cmd/spo/term/term-set-list.md)
Get-PnPTheme|SharePointPnPPowerShellOnline|
Get-PnPTimeZoneId|SharePointPnPPowerShellOnline|
Get-PnPUnifiedGroup|SharePointPnPPowerShellOnline|[graph o365group get](../cmd/graph/o365group/o365group-get.md), [graph o365group list](../cmd/graph/o365group/o365group-list.md)
Get-PnPUnifiedGroupMembers|SharePointPnPPowerShellOnline|
Get-PnPUnifiedGroupOwners|SharePointPnPPowerShellOnline|
Get-PnPUPABulkImportStatus|SharePointPnPPowerShellOnline|
Get-PnPUser|SharePointPnPPowerShellOnline|
Get-PnPUserProfileProperty|SharePointPnPPowerShellOnline|
Get-PnPView|SharePointPnPPowerShellOnline|[spo list view get](../cmd/spo/list/list-view-get.md), [spo list view list](../cmd/spo/list/list-view-list.md)
Get-PnPWeb|SharePointPnPPowerShellOnline|[spo web get](../cmd/spo/web/web-get.md), [spo web list](../cmd/spo/web/web-list.md)
Get-PnPWebhookSubscriptions|SharePointPnPPowerShellOnline|[spo list webhook get](../cmd/spo/list/list-webhook-get.md), [spo list webhook list](../cmd/spo/list/list-webhook-list.md)
Get-PnPWebPart|SharePointPnPPowerShellOnline|
Get-PnPWebPartProperty|SharePointPnPPowerShellOnline|
Get-PnPWebPartXml|SharePointPnPPowerShellOnline|
Get-PnPWebTemplates|SharePointPnPPowerShellOnline|
Get-PnPWikiPageContent|SharePointPnPPowerShellOnline|
Get-PnPWorkflowDefinition|SharePointPnPPowerShellOnline|
Get-PnPWorkflowInstance|SharePointPnPPowerShellOnline|
Get-PnPWorkflowSubscription|SharePointPnPPowerShellOnline|
Grant-PnPHubSiteRights|SharePointPnPPowerShellOnline|[spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md)
Grant-PnPSiteDesignRights|SharePointPnPPowerShellOnline|[spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md)
Grant-PnPTenantServicePrincipalPermission|SharePointPnPPowerShellOnline|[aad oauth2grant add](../cmd/aad/oauth2grant/oauth2grant-add.md)
Import-PnPAppPackage|SharePointPnPPowerShellOnline|
Import-PnPTaxonomy|SharePointPnPPowerShellOnline|
Import-PnPTermGroupFromXml|SharePointPnPPowerShellOnline|
Import-PnPTermSet|SharePointPnPPowerShellOnline|
Install-PnPApp|SharePointPnPPowerShellOnline|[spo app install](../cmd/spo/app/app-install.md)
Install-PnPSolution|SharePointPnPPowerShellOnline|
Invoke-PnPQuery|SharePointPnPPowerShellOnline|
Invoke-PnPSiteDesign|SharePointPnPPowerShellOnline|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Invoke-PnPWebAction|SharePointPnPPowerShellOnline|
Load-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
Measure-PnPList|SharePointPnPPowerShellOnline|
Measure-PnPResponseTime|SharePointPnPPowerShellOnline|
Measure-PnPWeb|SharePointPnPPowerShellOnline|
Move-PnPClientSideComponent|SharePointPnPPowerShellOnline|
Move-PnPFile|SharePointPnPPowerShellOnline|[spo file move](../cmd/spo/file/file-copy.md)
Move-PnPFolder|SharePointPnPPowerShellOnline|[spo folder move](../cmd/spo/folder/folder-move.md)
Move-PnPItemProxy|SharePointPnPPowerShellOnline|
Move-PnPListItemToRecycleBin|SharePointPnPPowerShellOnline|
Move-PnPRecycleBinItem|SharePointPnPPowerShellOnline|
New-PnPAzureCertificate|SharePointPnPPowerShellOnline|
New-PnPExtensbilityHandlerObject|SharePointPnPPowerShellOnline|
New-PnPGroup|SharePointPnPPowerShellOnline|
New-PnPList|SharePointPnPPowerShellOnline|[spo list add](../cmd/spo/list/list-add.md)
New-PnPPersonalSite|SharePointPnPPowerShellOnline|
New-PnPProvisioningCommunicationSite|SharePointPnPPowerShellOnline|
New-PnPProvisioningHierarchy|SharePointPnPPowerShellOnline|
New-PnPProvisioningSequence|SharePointPnPPowerShellOnline|
New-PnPProvisioningTeamNoGroupSite|SharePointPnPPowerShellOnline|
New-PnPProvisioningTeamNoGroupSubSite|SharePointPnPPowerShellOnline|
New-PnPProvisioningTeamSite|SharePointPnPPowerShellOnline|
New-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
New-PnPProvisioningTemplateFromFolder|SharePointPnPPowerShellOnline|
New-PnPSite|SharePointPnPPowerShellOnline|[spo site add](../cmd/spo/site/site-add.md)
New-PnPTenantSite|SharePointPnPPowerShellOnline|[spo site classic add](../cmd/spo/site/site-classic-add.md)
New-PnPTerm|SharePointPnPPowerShellOnline|[spo term add](../cmd/spo/term/term-add.md)
New-PnPTermGroup|SharePointPnPPowerShellOnline|[spo term group add](../cmd/spo/term/term-group-add.md)
New-PnPTermSet|SharePointPnPPowerShellOnline|[spo term set add](../cmd/spo/term/term-set-add.md)
New-PnPUnifiedGroup|SharePointPnPPowerShellOnline|[graph o365group add](../cmd/graph/o365group/o365group-add.md)
New-PnPUPABulkImportJob|SharePointPnPPowerShellOnline|
New-PnPUser|SharePointPnPPowerShellOnline|
New-PnPWeb|SharePointPnPPowerShellOnline|[spo web add](../cmd/spo/web/web-add.md)
Publish-PnPApp|SharePointPnPPowerShellOnline|[spo app deploy](../cmd/spo/app/app-deploy.md)
Read-PnPProvisioningHierarchy|SharePointPnPPowerShellOnline|
Read-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
Register-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite register](../cmd/spo/hubsite/hubsite-register.md)
Remove-PnPApp|SharePointPnPPowerShellOnline|[spo app remove](../cmd/spo/app/app-remove.md)
Remove-PnPClientSideComponent|SharePointPnPPowerShellOnline|
Remove-PnPClientSidePage|SharePointPnPPowerShellOnline|[spo page remove](../cmd/spo/page/page-remove.md)
Remove-PnPContentType|SharePointPnPPowerShellOnline|
Remove-PnPContentTypeFromDocumentSet|SharePointPnPPowerShellOnline|
Remove-PnPContentTypeFromList|SharePointPnPPowerShellOnline|[spo list contenttype remove](../cmd/spo/list/list-contenttype-remove.md)
Remove-PnPCustomAction|SharePointPnPPowerShellOnline|[spo customaction remove](../cmd/spo/customaction/customaction-remove.md)
Remove-PnPEventReceiver|SharePointPnPPowerShellOnline|
Remove-PnPField|SharePointPnPPowerShellOnline|[spo field remove](../cmd/spo/field/field-remove.md)
Remove-PnPFieldFromContentType|SharePointPnPPowerShellOnline|
Remove-PnPFile|SharePointPnPPowerShellOnline|[spo file remove](../cmd/spo/file/file-remove.md)
Remove-PnPFileFromProvisioningTemplate|SharePointPnPPowerShellOnline|
Remove-PnPFolder|SharePointPnPPowerShellOnline|[spo folder remove](../cmd/spo/folder/folder-remove.md)
Remove-PnPGroup|SharePointPnPPowerShellOnline|
Remove-PnPIndexedProperty|SharePointPnPPowerShellOnline|
Remove-PnPJavaScriptLink|SharePointPnPPowerShellOnline|
Remove-PnPList|SharePointPnPPowerShellOnline|[spo list remove](../cmd/spo/list/list-remove.md)
Remove-PnPListItem|SharePointPnPPowerShellOnline|[spo listitem remove](../cmd/spo/listitem/listitem-remove.md)
Remove-PnPNavigationNode|SharePointPnPPowerShellOnline|[spo navigation node remove](../cmd/spo/navigation/navigation-node-remove.md)
Remove-PnPPropertyBagValue|SharePointPnPPowerShellOnline|[spo propertybag remove](../cmd/spo/propertybag/propertybag-remove.md)
Remove-PnPPublishingImageRendition|SharePointPnPPowerShellOnline|
Remove-PnPRoleDefinition|SharePointPnPPowerShellOnline|
Remove-PnPSiteClassification|SharePointPnPPowerShellOnline|
Remove-PnPSiteCollectionAdmin|SharePointPnPPowerShellOnline|
Remove-PnPSiteCollectionAppCatalog|SharePointPnPPowerShellOnline|[spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md)
Remove-PnPSiteDesign|SharePointPnPPowerShellOnline|[spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md)
Remove-PnPSiteScript|SharePointPnPPowerShellOnline|[spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md)
Remove-PnPStorageEntity|SharePointPnPPowerShellOnline|[spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md)
Remove-PnPStoredCredential|SharePointPnPPowerShellOnline|
Remove-PnPTaxonomyItem|SharePointPnPPowerShellOnline|
Remove-PnPTenantCdnOrigin|SharePointPnPPowerShellOnline|[spo cdn origin remove](../cmd/spo/cdn/cdn-origin-remove.md)
Remove-PnPTenantSite|SharePointPnPPowerShellOnline|
Remove-PnPTenantTheme|SharePointPnPPowerShellOnline|[spo theme remove](../cmd/spo/theme/theme-remove.md)
Remove-PnPTermGroup|SharePointPnPPowerShellOnline|
Remove-PnPUnifiedGroup|SharePointPnPPowerShellOnline|[graph o365group remove](../cmd/graph/o365group/o365group-remove.md)
Remove-PnPUser|SharePointPnPPowerShellOnline|
Remove-PnPUserFromGroup|SharePointPnPPowerShellOnline|
Remove-PnPView|SharePointPnPPowerShellOnline|[spo list view remove](../cmd/spo/list/list-view-remove.md)
Remove-PnPWeb|SharePointPnPPowerShellOnline|[spo web remove](../cmd/spo/web/web-remove.md)
Remove-PnPWebhookSubscription|SharePointPnPPowerShellOnline|[spo list webhook remove](../cmd/spo/list/list-webhook-remove.md)
Remove-PnPWebPart|SharePointPnPPowerShellOnline|
Remove-PnPWikiPage|SharePointPnPPowerShellOnline|
Remove-PnPWorkflowDefinition|SharePointPnPPowerShellOnline|
Remove-PnPWorkflowSubscription|SharePointPnPPowerShellOnline|
Rename-PnPFile|SharePointPnPPowerShellOnline|
Rename-PnPFolder|SharePointPnPPowerShellOnline|[spo folder rename](../cmd/spo/folder/folder-rename.md)
Request-PnPReIndexList|SharePointPnPPowerShellOnline|
Request-PnPReIndexWeb|SharePointPnPPowerShellOnline|[spo web reindex](../cmd/spo/web/web-reindex.md)
Restore-PnPRecycleBinItem|SharePointPnPPowerShellOnline|
Restore-PnPTenantRecycleBinItem|SharePointPnPPowerShellOnline|
Resolve-PnPFolder|SharePointPnPPowerShellOnline|
Resume-PnPWorkflowInstance|SharePointPnPPowerShellOnline|
Revoke-PnPSiteDesignRights|SharePointPnPPowerShellOnline|[spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md)
Revoke-PnPTenantServicePrincipalPermission|SharePointPnPPowerShellOnline|[spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)
Save-PnPProvisioningHierarchy|SharePointPnPPowerShellOnline|
Save-PnPProvisioningTemplate|SharePointPnPPowerShellOnline|
Send-PnPMail|SharePointPnPPowerShellOnline|
Set-PnPAppSideLoading|SharePointPnPPowerShellOnline|
Set-PnPAuditing|SharePointPnPPowerShellOnline|
Set-PnPAvailablePageLayouts|SharePointPnPPowerShellOnline|
Set-PnPClientSidePage|SharePointPnPPowerShellOnline|[spo page set](../cmd/spo/page/page-set.md), [spo page header set](../cmd/spo/page/page-header-set.md)
Set-PnPClientSideText|SharePointPnPPowerShellOnline|
Set-PnPClientSideWebPart|SharePointPnPPowerShellOnline|
Set-PnPContext|SharePointPnPPowerShellOnline|
Set-PnPDefaultColumnValues|SharePointPnPPowerShellOnline|
Set-PnPDefaultContentTypeToList|SharePointPnPPowerShellOnline|
Set-PnPDefaultPageLayout|SharePointPnPPowerShellOnline|
Set-PnPDocumentSetField|SharePointPnPPowerShellOnline|
Set-PnPField|SharePointPnPPowerShellOnline|[spo field set](../cmd/spo/field/field-set.md)
Set-PnPFileCheckedIn|SharePointPnPPowerShellOnline|[spo file checkin](../cmd/spo/file/file-checkin.md)
Set-PnPFileCheckedOut|SharePointPnPPowerShellOnline|[spo file checkout](../cmd/spo/file/file-checkout.md)
Set-PnPGroup|SharePointPnPPowerShellOnline|
Set-PnPGroupPermissions|SharePointPnPPowerShellOnline|
Set-PnPHideDefaultThemes|SharePointPnPPowerShellOnline|[spo hidedefaultthemes set](../cmd/spo/hidedefaultthemes/hidedefaultthemes-set.md)
Set-PnPHomePage|SharePointPnPPowerShellOnline|
Set-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite set](../cmd/spo/hubsite/hubsite-set.md)
Set-PnPIndexedProperties|SharePointPnPPowerShellOnline|
Set-PnPInformationRightsManagement|SharePointPnPPowerShellOnline|
Set-PnPInPlaceRecordsManagement|SharePointPnPPowerShellOnline|[spo site inplacerecordsmanagement set](../cmd/spo/site/site-inplacerecordsmanagement-set.md)
Set-PnPLabel|SharePointPnPPowerShellOnline|[spo list label set](../cmd/spo/list/list-label-set.md)
Set-PnPList|SharePointPnPPowerShellOnline|[spo list set](../cmd/spo/list/list-set.md)
Set-PnPListItem|SharePointPnPPowerShellOnline|[spo listitem set](../cmd/spo/listitem/listitem-set.md)
Set-PnPListItemAsRecord|SharePointPnPPowerShellOnline|[spo listitem record declare](../cmd/spo/listitem/listitem-record-declare.md)
Set-PnPListItemPermission|SharePointPnPPowerShellOnline|
Set-PnPListPermission|SharePointPnPPowerShellOnline|
Set-PnPListRecordDeclaration|SharePointPnPPowerShellOnline|
Set-PnPMasterPage|SharePointPnPPowerShellOnline|
Set-PnPMinimalDownloadStrategy|SharePointPnPPowerShellOnline|
Set-PnPPropertyBagValue|SharePointPnPPowerShellOnline|[spo propertybag set](../cmd/spo/propertybag/propertybag-set.md)
Set-PnPProvisioningTemplateMetadata|SharePointPnPPowerShellOnline|
Set-PnPRequestAccessEmails|SharePointPnPPowerShellOnline|
Set-PnPSearchConfiguration|SharePointPnPPowerShellOnline|
Set-PnPSite|SharePointPnPPowerShellOnline|[spo site set](../cmd/spo/site/site-set.md)
Set-PnPSiteClosure|SharePointPnPPowerShellOnline|
Set-PnPSiteDesign|SharePointPnPPowerShellOnline|[spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md)
Set-PnPSitePolicy|SharePointPnPPowerShellOnline|
Set-PnPSiteScript|SharePointPnPPowerShellOnline|[spo sitescript set](../cmd/spo/sitescript/sitescript-set.md)
Set-PnPStorageEntity|SharePointPnPPowerShellOnline|[spo storageentity set](../cmd/spo/storageentity/storageentity-set.md)
Set-PnPTaxonomyFieldValue|SharePointPnPPowerShellOnline|
Set-PnPTenant|SharePointPnPPowerShellOnline|[spo tenant settings set](../cmd/spo/tenant/tenant-settings-set.md)
Set-PnPTenantCdnEnabled|SharePointPnPPowerShellOnline|[spo cdn set](../cmd/spo/cdn/cdn-set.md)
Set-PnPTenantCdnPolicy|SharePointPnPPowerShellOnline|[spo cdn policy set](../cmd/spo/cdn/cdn-policy-set.md)
Set-PnPTenantSite|SharePointPnPPowerShellOnline|[spo site classic set](../cmd/spo/site/site-classic-set.md)
Set-PnPTheme|SharePointPnPPowerShellOnline|
Set-PnPTraceLog|SharePointPnPPowerShellOnline|
Set-PnPUnifiedGroup|SharePointPnPPowerShellOnline|[graph o365group set](../cmd/graph/o365group/o365group-set.md)
Set-PnPUserProfileProperty|SharePointPnPPowerShellOnline|
Set-PnPView|SharePointPnPPowerShellOnline|[spo list view set](../cmd/spo/list/list-view-set.md)
Set-PnPWeb|SharePointPnPPowerShellOnline|[spo web set](../cmd/spo/web/web-set.md)
Set-PnPWebhookSubscription|SharePointPnPPowerShellOnline|[spo list webhook set](../cmd/spo/list/list-webhook-set.md)
Set-PnPWebPartProperty|SharePointPnPPowerShellOnline|
Set-PnPWebPermission|SharePointPnPPowerShellOnline|
Set-PnPWebTheme|SharePointPnPPowerShellOnline|
Set-PnPWikiPageContent|SharePointPnPPowerShellOnline|
Start-PnPWorkflowInstance|SharePointPnPPowerShellOnline|
Stop-PnPWorkflowInstance|SharePointPnPPowerShellOnline|
Submit-PnPSearchQuery|SharePointPnPPowerShellOnline|[spo search](../cmd/spo/search/search.md)
Test-PnPListItemIsRecord|SharePointPnPPowerShellOnline|
Test-PnPOffice365GroupAliasIsUsed|SharePointPnPPowerShellOnline|
Test-PnPProvisioningHierarchy|SharePointPnPPowerShellOnline|
Uninstall-PnPApp|SharePointPnPPowerShellOnline|[spo app uninstall](../cmd/spo/app/app-uninstall.md)
Uninstall-PnPAppInstance|SharePointPnPPowerShellOnline|
Uninstall-PnPSolution|SharePointPnPPowerShellOnline|
Unpublish-PnPApp|SharePointPnPPowerShellOnline|[spo app retract](../cmd/spo/app/app-retract.md)
Unregister-PnPHubSite|SharePointPnPPowerShellOnline|[spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md)
Update-PnPApp|SharePointPnPPowerShellOnline|[spo app upgrade](../cmd/spo/app/app-upgrade.md)
Update-PnPSiteClassification|SharePointPnPPowerShellOnline|[graph siteclassification set](../cmd/graph/siteclassification/siteclassification-set.md)
Approve-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Deny-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Disable-Flow|Microsoft.PowerApps.PowerShell|
Enable-Flow|Microsoft.PowerApps.PowerShell|
Get-Flow|Microsoft.PowerApps.PowerShell|[azmgmt flow list](../cmd/azmgmt/flow/flow-list.md), [azmgmt flow get](../cmd/azmgmt/flow/flow-get.md)
Get-FlowApproval|Microsoft.PowerApps.PowerShell|
Get-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Get-FlowEnvironment|Microsoft.PowerApps.PowerShell|[azmgmt flow environment list](../cmd/azmgmt/flow/flow-environment-list.md), [azmgmt flow environment get](../cmd/azmgmt/flow/flow-environment-get.md)
Get-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Get-FlowRun|Microsoft.PowerApps.PowerShell|[azmgmt flow run list](../cmd/azmgmt/flow/flow-run-list.md), [azmgmt flow run get](../cmd/azmgmt/flow/flow-run-get.md)
Get-PowerApp|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnection|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnector|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppEnvironment|Microsoft.PowerApps.PowerShell|
Get-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppsNotification|Microsoft.PowerApps.PowerShell|
Get-PowerAppVersion|Microsoft.PowerApps.PowerShell|
Publish-PowerApp|Microsoft.PowerApps.PowerShell|
Remove-Flow|Microsoft.PowerApps.PowerShell|
Remove-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Remove-PowerApp|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnection|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnector|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Remove-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Restore-PowerAppVersion|Microsoft.PowerApps.PowerShell|
Set-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Set-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Set-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Set-PowerAppDisplayName|Microsoft.PowerApps.PowerShell|
Set-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Add-ConnectorToBusinessDataGroup|Microsoft.PowerApps.Administration.PowerShell|
Add-CustomConnectorToPolicy|Microsoft.PowerApps.Administration.PowerShell|
Add-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppApisToBypassConsent|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppAsFeatured|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppAsHero|Microsoft.PowerApps.Administration.PowerShell|
Disable-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Enable-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlowUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerApp|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppCdsDatabaseCurrencies|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppCdsDatabaseLanguages|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnection|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectionReferences|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnector|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironmentLocations|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppsUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Get-JwtToken|Microsoft.PowerApps.Administration.PowerShell|
Get-TenantDetailsFromGraph|Microsoft.PowerApps.Administration.PowerShell|
Get-UsersOrGroupsFromGraph|Microsoft.PowerApps.Administration.PowerShell|
InvokeApi|Microsoft.PowerApps.Administration.PowerShell|
New-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
New-AdminPowerAppCdsDatabase|Microsoft.PowerApps.Administration.PowerShell|
New-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowApprovals|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerApp|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnection|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnector|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-ConnectorFromBusinessDataGroup|Microsoft.PowerApps.Administration.PowerShell|
Remove-CustomConnectorFromPolicy|Microsoft.PowerApps.Administration.PowerShell|
Remove-LegacyCDSDatabase|Microsoft.PowerApps.Administration.PowerShell|
Remove-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
ReplaceMacro|Microsoft.PowerApps.Administration.PowerShell|
Select-CurrentEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppApisToBypassConsent|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppAsFeatured|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppAsHero|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppEnvironmentDisplayName|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppOwner|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Test-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
Add-TeamUser|MicrosoftTeams|[graph teams user add](../cmd/graph/o365group/o365group-user-add.md)
Connect-MicrosoftTeams|MicrosoftTeams|[graph login](../cmd/graph/login.md)
Disconnect-MicrosoftTeams|MicrosoftTeams|[graph logout](../cmd/graph/logout.md)
Get-Team|MicrosoftTeams|[graph teams list](../cmd/graph/teams/teams-list.md)
Get-TeamChannel|MicrosoftTeams|[graph teams channel list](../cmd/graph/teams/teams-channel-list.md)
Get-TeamFunSettings|MicrosoftTeams|[graph teams funsettings list](../cmd/graph/teams/teams-funsettings-list.md)
Get-TeamGuestSettings|MicrosoftTeams|[graph teams guestsettings list](../cmd/graph/teams/teams-guestsettings-list.md)
Get-TeamHelp|MicrosoftTeams|
Get-TeamMemberSettings|MicrosoftTeams|[graph teams membersettings list](../cmd/graph/teams/teams-membersettings-list.md)
Get-TeamMessagingSettings|MicrosoftTeams|[graph teams messagingsettings list](../cmd/graph/teams/teams-messagingsettings-list.md)
Get-TeamUser|MicrosoftTeams|[graph teams user list](../cmd/graph/teams/teams-user-list.md)
New-Team|MicrosoftTeams|[graph teams add](../cmd/graph/teams/teams-add.md)
New-TeamChannel|MicrosoftTeams|[graph teams channel add](../cmd/graph/teams/teams-channel-add.md)
Remove-Team|MicrosoftTeams|
Remove-TeamChannel|MicrosoftTeams|
Remove-TeamUser|MicrosoftTeams|[graph teams user remove](../cmd/graph/teams/teams-user-remove.md)
Set-Team|MicrosoftTeams|
Set-TeamChannel|MicrosoftTeams|[graph teams channel set](../cmd/graph/teams/teams-channel-set.md)
Set-TeamFunSettings|MicrosoftTeams|
Set-TeamGuestSettings|MicrosoftTeams|[graph teams guestsettings set](../cmd/graph/teams/teams-guestsettings-set.md)
Set-TeamMemberSettings|MicrosoftTeams|[graph teams membersettings set](../cmd/graph/teams/teams-membersettings-set.md)
Set-TeamMessagingSettings|MicrosoftTeams|
Set-TeamPicture|MicrosoftTeams|