#This is script will connect to OneDrive sites tenant with AppID and APP Secret
#global variables 
$Global:OneDriveSites=$null
function SPOnlineConnection()
{
    try
    {

Connect-PnPOnline -Url "https://domain-admin.sharepoint.com" -UseWebLogin

#Get All OneDrive Sites using filter
$global:OneDriveSites = Get-PNPTenantSite -IncludeOneDrivesites -Filter "Url -like '-my.sharepoint.com/personal/'"
             
#Exporting all One-Drive site collection details 
$global:OneDriveSites | Out-File $OnedriveSiteCollReport                          
    }
    catch [System.Exception] 
    {
         write-host -f Red "Error in OneDrive Connection Alerts!" $_.Exception.Message 
	
    }

 }

 function WhiteListDomain()
 {
    [cmdletbinding()]
    param(
		[Parameter(Mandatory=$true)]$OnedriveSiteCollReport           
    )
        
#Loop through each site and add site collection admin
          
            foreach($site in $OneDriveSites)  
            {  
	try
	{
                Write-host "Processing URL" $site.Url -ForegroundColor Green 
                Set-PNPTenantSite -url $site.url -SharingAllowedDomainList "newdomain.com" -SharingDomainRestrictionMode AllowList
                $site.url >> WhitelistedReport.csv
	}
 catch [System.Exception] 
    {
         write-host -f Red "Error in OneDrive Connection Alerts!" $_.Exception.Message 
	 $_.Exception.Message + $site.url | Out-File $OnedriveSiteCollReport
    }
			}
 }

#Defining variables for export reports
    $CurrentLocation =  $(get-location).Path;
    $OnedriveSiteCollReport =$($CurrentLocation +"\"+ "OnedriveSiteCollReport.csv")
 
#Establishing online connection
SPOnlineConnection

#Calling Onedrive operations
WhiteListDomain -OnedriveSiteCollReport $OnedriveSiteCollReport 