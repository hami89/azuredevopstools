# azuredevopstools
Azure DevOps utilities

## Basic usage
    ipmo SprintWorkItems
    
    az login

    $config = New-DevOpsConfig -org "https://dev.azure.com/your-company" -project "SoftwareProject" -team "YourTeam" -areaPath "SoftwareProject\SomeProject"
 
    $sprints=Get-SprintSet -config $config

    $items = Get-SprintWorkItems -config $config -sprint $sprints.last -includeHistory

    Get-SprintReport -sprint $sprints.last -workItems $items -path .\test.csv

    .\test.csv
