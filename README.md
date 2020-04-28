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

## Members

    Get-Command -Module SprintWorkItems

    CommandType     Name                                               Version    Source
    -----------     ----                                               -------    ------
    Function        Check-ActivationAfter                              0.0        SprintWorkItems
    Function        Check-ActivationBefore                             0.0        SprintWorkItems
    Function        Check-Done                                         0.0        SprintWorkItems
    Function        Check-DoneAfter                                    0.0        SprintWorkItems
    Function        Check-DoneBefore                                   0.0        SprintWorkItems
    Function        Check-EverAssigned                                 0.0        SprintWorkItems
    Function        Check-Planned                                      0.0        SprintWorkItems
    Function        Check-Removed                                      0.0        SprintWorkItems
    Function        Check-SprintAt                                     0.0        SprintWorkItems
    Function        Check-Unplanned                                    0.0        SprintWorkItems
    Function        Get-DoneBy                                         0.0        SprintWorkItems
    Function        Get-DoneDate                                       0.0        SprintWorkItems
    Function        Get-FirstAssignedBy                                0.0        SprintWorkItems
    Function        Get-FirstAssignedDate                              0.0        SprintWorkItems
    Function        Get-NewRevisionsBetween                            0.0        SprintWorkItems
    Function        Get-RevisionAt                                     0.0        SprintWorkItems
    Function        Get-SprintReport                                   0.0        SprintWorkItems
    Function        Get-SprintSet                                      0.0        SprintWorkItems
    Function        Get-SprintWorkItems                                0.0        SprintWorkItems
    Function        Get-WorkItemHistory                                0.0        SprintWorkItems
    Function        Install-Az                                         0.0        SprintWorkItems
    Function        New-DevOpsConfig                                   0.0        SprintWorkItems
