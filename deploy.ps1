# Clear all variables
Remove-Variable * -ErrorAction SilentlyContinue; 
$progressPreference = 'silentlyContinue'
$WarningPreference  = 'silentlyContinue' 

# Check and install Import-Excel module
$module = Get-Module -Name ImportExcel
if ($module -eq $null) { Write-Host "Module ImportExcel is not installed. Installing.." -ForegroundColor Yellow;
Install-Module -Name ImportExcel -RequiredVersion 4.0.8 -Force -WarningAction SilentlyContinue }

$rg_name   = "NSG"
$location  = "eastus"
$ExcelFile = "C:\tfa\my_NSG.xlsx"

# NSG name array
$nsg_name = @("NSG1","NSG2","NSG3")

# List VNET
$vnet = Get-AzureRmVirtualNetwork -ResourceGroupName $rg_name

# List subnets
$subnets = Get-AzureRmVirtualNetwork -Name $vnet.Name -ResourceGroupName $rg_name | Get-AzureRmVirtualNetworkSubnetConfig

# List Address prefixes
$addrprefix = (Get-AzureRmVirtualNetworkSubnetConfig -VirtualNetwork $vnet).AddressPrefix

# List NSGs
$NSGs = Get-AzureRmNetworkSecurityGroup -ResourceGroupName $rg_name

# Create NSG in each subnet
Foreach($nsg in $nsg_name) {New-AzureRmNetworkSecurityGroup -ResourceGroupName $rg_name -Name $nsg -Location $location}

# Attach NSG to each subnet
for($i=0; $i -lt $NSGs.Length; $i++){
Set-AzureRmVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $subnets[$i].Name -AddressPrefix $addrprefix[$i] -NetworkSecurityGroup $NSGs[$i] | Set-AzureRmVirtualNetwork
Write-Host $NSGs[$i].Name "attached to" $subnets[$i].Name -ForegroundColor Green }

#Import the excel with the rules and apply on each NSG
for($n=0; $n -lt $NSGs.Length; $n++) {
Write-Host "Importing NSG rules for" $NSGs[$n].Name -ForegroundColor Green
$Rule = Import-Excel -Path $ExcelFile -WorksheetName $subnets[$n].Name
Write-Host "Adding rules to" $NSGs[$n].Name "for Subnet" $subnets[$n].Name -ForegroundColor Green
for($r=0; $r -lt $Rule.Length; $r++) {
Add-AzureRmNetworkSecurityRuleConfig -Name $Rule[$r].ruleName  -NetworkSecurityGroup $NSGs[$n] -Protocol $Rule[$r].protocol -SourcePortRange $Rule[$r].sourcePort -DestinationPortRange $Rule[$r].destinationPort -SourceAddressPrefix $Rule[$r].sourcePrefix -DestinationAddressPrefix $Rule[$r].destinationPrefix -Access $Rule[$r].access -Priority $Rule[$r].priority -Direction $Rule[$r].direction | Set-AzureRmNetworkSecurityGroup 
}}
