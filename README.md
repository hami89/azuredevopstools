# azuredevopstools
Azure DevOps utilities

## Basic usage
    ipmo SprintWorkItems
    
    az login

    $config = New-DevOpsConfig -org "https://dev.azure.com/itron" -project "SoftwareProducts" -team "GDS" -areaPath "SoftwareProducts\Outcomes\Operations Management\Gas Distribution Safety"
 
    $sprints=Get-SprintSet -config $config

    $items = Get-SprintWorkItems -config $config -sprint $sprints.last -includeHistory

    Get-SprintReport -sprint $sprints.last -workItems $items -path .\test.csv

    .\test.csv
