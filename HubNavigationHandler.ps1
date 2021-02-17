function Get-SPOHubSiteNavigation{
        <#
        .SYNOPSIS
         Exports the hub navigation for the SPO site provided to CSV.
        .DESCRIPTION
         This custom function gets the Hub site navigation links for the provided SPO site.
         It then iterates through each of the links and builds a collection to export to CSV.
         This collection can also be integrated using pipe functions.
        .PARAMETER Identity 
         Specifies the Url for the SPO site where the navigation should be exported from.
        .PARAMETER Export 
         Specifies whether the navigation should be exported or not. The export is saved to the current directory.
        .EXAMPLE
         PS C:\>Get-SPOHubNavigation -Identity https://[tenant].sharepoint.com/sites/[site] -Export:$true
        #>
        param(
                [Parameter(Mandatory=$true)]
                [string] $Identity,
                [Parameter(Mandatory=$true)]
                [boolean] $Export
        )
        begin{
                $exportNavCol = @()
                # This counter is used in order to maintain the order of the navigation. 
                # The navigation is returned in the order it appears.
                $counter = 1
                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Get-SPOHubNavigation function started."
        }
        process{
                # Connect to the hub site with the navigation to base all the other sites on
                $connection = Connect-PnPOnline -Url $identity -UseWebLogin -ReturnConnection
                $site = Get-PnPSite

                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Exporting navigation from $($site.Url)..."

                # Get the master navigation
                $navigationNodes = Get-PnPNavigationNode -Location TopNavigationBar -Connection $connection
                
                # Iterate through the navigation and capture all the nodes on all 3 levels
                foreach($navigationNode in $navigationNodes){
                        $parentNode = Get-PnPNavigationNode -id $navigationNode.Id
                        $navInfo = New-Object PSObject -property @{
                                Level = "Level 1"
                                Id = $navigationNode.Id
                                Title = $navigationNode.Title
                                Url = $navigationNode.Url
                                ParentId = "0"
                                ParentTitle = ""
                                Visible = $navigationNode.IsVisible
                                Order = $counter
                        }
                        # Add the navInfo collection to the collection we're going to export.
                        $exportNavCol += $navInfo
                        $counter++

                        # Get the second level navigation
                        $navigation = Get-PnPNavigationNode -Id $navigationNode.Id
                        $children = $navigation.Children

                        # If children exist proceed
                        if($children){
                                foreach($child in $children){
                                        # Get the node and further information about the link
                                        $childNode = Get-PnPNavigationNode -Id $child.Id
                                        $navInfo = New-Object PSObject -property @{
                                                Level = "Level 2"
                                                Id = $childNode.Id
                                                Title = $childNode.Title
                                                Url = $childNode.Url
                                                ParentId = $parentNode.Id
                                                ParentTitle = $parentNode.Title
                                                Visible = $childNode.IsVisible
                                                Order = $counter
                                        }

                                        # Add the navInfo collection to the collection we're going to export.
                                        $exportNavCol += $navInfo
                                        $counter++

                                        # Get the third level navigation
                                        $subChildren = $childNode.Children

                                        # if children exist proceed
                                        if($subChildren) {
                                                foreach($subChild in $subChildren) {
                                                        # Get the node and further information about the link
                                                        $subChildNode = Get-PnPNavigationNode -Id $subChild.Id
                                                        $navInfo = New-Object PSObject -property @{
                                                                Level = "Level 3"
                                                                Id = $subChildNode.Id
                                                                Title = $subChildNode.Title
                                                                Url = $subChildNode.Url
                                                                ParentId = $childNode.Id
                                                                ParentTitle = $childNode.Title
                                                                Visible = $childNode.IsVisible
                                                                Order = $counter
                                                        }
                                                        # Add the navInfo collection to the collection we're going to export.
                                                        $exportNavCol += $navInfo
                                                        $counter++
                                                }
                                        }       
                                }
                        }
                }
                Disconnect-PnPOnline -Connection $connection
        }
        end{                
                # Rebuild collection with sort
                $returnCol = @()
                $returnCol = $exportNavCol | Sort-Object Order

                # Export the navigation to a CSV file if the switch is enabled
                if($Export -eq $true){
                        #$exportFile = ".\Output-HubNavigation-$((Get-Date).ToString("yyyymmddhhss")).csv"
                        $exportFile = "C:\Users\tripaabh\OneDrive - TietoEVRY\Desktop\Output-HubNavigation.csv"
                        Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Navigation exported to '$($exportFile)'."
                        $returnCol | Export-Csv $exportFile -NoTypeInformation -Append:$false -Force:$true
                }

                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Get-SPOHubNavigation finished."
                
                return $returnCol
        }
}       

function Copy-SPOHubSiteNavigation {
        <#
        .SYNOPSIS
         Imports the hub navigation to the SPO site provided..
        .DESCRIPTION
         This custom function imports the Hub site navigation provided in the CSV file to the provided SPO site.
         It then iterates through each of the links and adds them using the Add-PnPNavigationNode.
        .PARAMETER Identity 
         Specifies the Url for the SPO site where the navigation should be imported to.
        .PARAMETER importCsv 
         Specifies the path to the CSV where the hub navigation was exported to.
        .EXAMPLE
         PS C:\>Copy-SPOHubNavigation -Identity https://[tenant].sharepoint.com/sites/[targetsite] -importCsv ".\Output-HubNavigation.csv"
        #>
        param (
                [Parameter(Mandatory=$true)]
                [string] $Identity,
                [Parameter(Mandatory=$true)]
                [string] $importCsv
        )
        begin{
                $newNavigationArray = @()
                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Copy-SPOHubNavigation function started."
        }
        process{
                # Connect to the target site with where the navigation will be added
                $connection = Connect-PnPOnline -Url  $identity -UseWebLogin -ReturnConnection
                $site = Get-PnPSite

                $currentLevel = 1
                $parentId = 0
            
                # Import navigation CSV
                $navigationCol = Import-Csv $importCsv

                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Removing existing navigation."
                
                # Remove existing hub navigation
                Remove-PnPNavigationNode -All -Force

                # Remove any remaining nodes. Experienced instances where the RemovePnpNavigationNode -All hasn't worked.
                $navNodes = Get-PnPNavigationNode -Location TopNavigationBar -Connection $connection
                foreach($navNode in $navNodes) {
                        Remove-PnPNavigationNode -Identity $navNode.Id -Force -Connection $connection         
                }

                Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Importing navigation from $($importCsv) to $($site.Url)..."
                            
                foreach($navigationObject in $navigationCol){

                        Write-Debug "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Adding navigation node: $($navigationObject.Title)"
                        #Write-Debug "$($navigationObject.ParentId) $($navigationObject.Title)"

                        # If the parent is 0 then it is a top-level navigation node
                        if($navigationObject.ParentId -eq 0) {
                                $parentId = 0
                                $level = 1                    
                                if($navigationObject.Url.Length -gt 0) {
                                        # Navigation node has a Url so create it as a top level menu item
                                        $addedNav = Add-PnPNavigationNode -Url $navigationObject.Url -Title $navigationObject.Title -Location TopNavigationBar -External
                                }else{
                                        # Navigation node hasn't a Url so create it as a top level header item
                                        $addedNav = Add-PnPNavigationNode -Title $navigationObject.Title -Location TopNavigationBar -External
                                }
                        }else{
                                # Get the parent for the child node from the master navigation array
                                $parentNav = $navigationCol | Where-Object{$_.Id -eq $navigationObject.ParentId}
                                if($parentNav.Url.Length -gt 0) {
                                        # Navigation node has a Url so get the newly created navigation node from the new array
                                        # via the Title and Url to get the new Id for the parent item and add it to the navigation
                                        $newNav = $newNavigationArray | Where-Object {$_.Title -eq $parentNav.Title -and $_.Url -eq $parentNav.Url}
                                        if($currentLevel -ne $navigationObject.Level){
                                                $parentId = $newNav.Id
                                                $level = $navigationObject.Level
                                        } 
                                        $addedNav = Add-PnPNavigationNode -Url $navigationObject.Url -Title $navigationObject.Title -Location TopNavigationBar -Parent $parentId -External
                                }else{
                                        # Navigation node does not have a Url so get the newly created navigation node from the new array
                                        # via the Title only to get the new Id for the parent item and add it to the navigation
                                        $newNav = $newNavigationArray | Where-Object{$_.Title -eq $parentNav.Title}
                                        if($currentLevel -ne $navigationObject.Level){
                                                $parentId = $newNav.Id
                                                $level = $navigationObject.Level
                                        } 
                                        $addedNav = Add-PnPNavigationNode -Title $navigationObject.Title -Location TopNavigationBar -Parent $parentId -External
                                }
                                $parentNav = $null
                        }

                        $currentLevel = $navigationObject.Level

                        # Capture the navigation node created to store their IDs
                        $navigationObject = New-Object PSObject -property @{
                                Id = $addedNav.Id
                                Url = $addedNav.Url
                                Title = $addedNav.Title
                                Level = $level
                                ParentId = $parentId
                        }
    
                        # Add the recently added navigation to the new navigation array
                        $newNavigationArray += $navigationObject
                }    
        }
        end{
                Disconnect-PnPOnline -Connection $connection
        }         
}

$DebugPreference = "Continue" #Continue; SilentlyContinue
$masterHubNavUrl = "https://keskogroup.sharepoint.com/sites/Intra-SE-Kesko"

$exportFile = "C:\Users\tripaabh\OneDrive - TietoEVRY\Desktop\Output-HubNavigation.csv"
$nav = Get-SPOHubSiteNavigation -identity $masterHubNavUrl -Export:$true

$nav | Format-Table

$siteUrl = "https://keskogroup.sharepoint.com/sites/Abhishek"
# Copy-SPOHubSiteNavigation -identity $siteUrl -importCsv $exportFile #

# Example how this could work with a collection of other sites.

$sites = @("https://keskogroup.sharepoint.com/sites/Abhishek")
foreach($site in $sites){
	Copy-SPOHubSiteNavigation -identity $siteUrl -importCsv $exportFile
} 
