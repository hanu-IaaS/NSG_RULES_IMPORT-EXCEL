# Clear all variables
#Login-AzureRmAccount
Remove-Variable * -ErrorAction SilentlyContinue; 
$progressPreference = 'silentlyContinue'
$WarningPreference  = 'silentlyContinue' 

# Check and install Import-Excel module
$module = Get-Module -Name ImportExcel
if ($module -eq $null) { Write-Host "Module ImportExcel is not installed. Installing.." -ForegroundColor Yellow;
Install-Module -Name ImportExcel -RequiredVersion 4.0.8 -Force -WarningAction SilentlyContinue }

$rg_name   = "nsg"
$location  = "west central us"
$ExcelFile = "https://raw.githubusercontent.com/hanu-IaaS/NSG_RULES_IMPORT-EXCEL/NSG-Rules-v0.2.xlsx"

# NSG name array
$nsg_name = @("adr-web-dmz-nsg01","adr-web-dmz-nsg02","adr-web-dmz-nsg03","adr-web-dmz-nsg04","adr-app-dmz-nsg01")

# List VNET
$vnet = Get-AzureRmVirtualNetwork -ResourceGroupName $rg_name

# List subnets
$subnets = Get-AzureRmVirtualNetwork -Name $vnet.Name -ResourceGroupName $rg_name | Get-AzureRmVirtualNetworkSubnetConfig

# List Address prefixes
$addrprefix = (Get-AzureRmVirtualNetworkSubnetConfig -VirtualNetwork $vnet).AddressPrefix

# List NSGs
$NSGs = Get-AzureRmNetworkSecurityGroup -ResourceGroupName $rg_name

# Create NSG in each subnet
Foreach($nsg in $nsg_name) {New-AzureRmNetworkSecurityGroup -ResourceGroupName $rg_name -Name $nsg -Location $location -Force  }

# Attach NSG to each subnet

for($i=0; $i -lt $NSGs.Length; $i++){
Set-AzureRmVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $subnets[$i].Name -AddressPrefix $addrprefix[$i] -NetworkSecurityGroup $NSGs[$i] | Set-AzureRmVirtualNetwork
Write-Host $NSGs[$i].Name "attached to" $subnets[$i].Name -ForegroundColor Green }

#Import the excel with the rules and apply on each NSG
for($n=0; $n -lt $nsg_name.Length; $n++) {
Write-Host "Importing NSG rules for" $nsg_name[$n].Name -ForegroundColor Green
$Rule = Import-Excel -Path $ExcelFile -WorksheetName $subnets[$n].Name}

for($r=0; $r -lt $Rule.Length; $r++) {
Write-Host "Adding rules to" $nsg_name[$n].Name "for Subnet" $subnets[$n].Name -ForegroundColor Green
Add-AzureRmNetworkSecurityRuleConfig -Name $Rule[$r].ruleName  -NetworkSecurityGroup $nsg_name[$n] -Protocol $Rule[$r].protocol -SourcePortRange $Rule[$r].sourcePort -DestinationPortRange $Rule[$r].destinationPort -SourceAddressPrefix $Rule[$r].sourcePrefix -DestinationAddressPrefix $Rule[$r].destinationPrefix -Access $Rule[$r].access -Priority $Rule[$r].priority -Direction $Rule[$r].direction -Description $Rule[$r].description | Set-AzureRmNetworkSecurityGroup 
}
