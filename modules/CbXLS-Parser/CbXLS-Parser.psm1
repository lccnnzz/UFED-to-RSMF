#========================================================
# Modules
#========================================================
Import-Module -Name .\modules\CbXLS-ParsingLib -DisableNameChecking

#========================================================
# Functions
#========================================================
function Get-EndDate{
    param(
        [Parameter(Mandatory)] 
        [DateTime]$beginDate,

        [Parameter()]
        [ValidateSet('None', 'Day', 'Week', 'Month')]
        [String]$TimePeriod = 'None'
    )

    switch ($TimePeriod){
        'None'{
            return Get-Date -Hour 23 -Minute 59 -Second 59
        }
        'Week'{
            return Get-Date $beginDate.AddDays(7) -Hour 23 -Minute 59 -Second 59
        }
        'Month'{
            return (Get-Date $beginDate.AddMonths(1) -Day 1 -Hour 23 -Minute 59 -Second 59).AddDays(-1)
        }
        'Day'{
            return (Get-Date $beginDate -Hour 23 -Minute 59 -Second 59)
        }
    }
}

function Parse-Chat{
    param(
        [Parameter(Mandatory)] [System.Collections.ArrayList]$EventRows,
        [Parameter(Mandatory)] $FieldNameCols,
        [Parameter(Mandatory)] $AttachmentCols,
        [Parameter()]
        [ValidateSet('None', 'Week', 'Month')]
        [String]$GroupBy = 'None',
        [Parameter()] $CustodianID = '',
        [Parameter()] [int]$MaxMessagesPerChat = 9999,
        [Parameter()] [ref]$ProgressHelper
    )
    
    $Participants  = New-Object Collections.Generic.List[PSCustomObject]
    $Conversations = New-Object Collections.Generic.List[PSCustomObject]
    $Events        = New-Object Collections.Generic.List[PSCustomObject]
    $EventsGroups  = [System.Collections.ArrayList]@()
        
    #Get first row in the range
    $FirstRow = $($EventRows | Select-Object -First 1)

    # Get the first event object
    $FirstEvent = Get-Event -EventRow $FirstRow -FieldCols $FieldNameCols -AttachmentCols $AttachmentCols -CustodianID $CustodianID

    # Add the first event to Events list
    [void] $Events.Add($FirstEvent)
    
    # Update the progress bar
    # Add new task to helper
    $taskID   = [int]$FirstEvent.conversation + 100
    $parentID = 1
    $activity = "Chat # $($FirstEvent.conversation)"
    $items    = $EventRows.Count

    $Helper = $ProgressHelper.Value
    if($Helper){
        $Helper.Add($taskID, $parentID, "Chat # $($FirstEvent.conversation)", $items)
        $Helper.Show()
        $Helper.Update($taskID, 1)
        $Helper.Show()
    }
    
    # Get participant list from the first event
    $Participants = Get-Participants -EventRow $firstRow -ParticipantsCol $FieldNameCols["Participants"]

    # Get conversation info from the first event
    $conversation = Get-Conversation -EventRow $FirstRow -FieldCols $FieldNameCols -Participants $Participants
    [void] $Conversations.Add($conversation)

    $beginDate = [DateTime] $FirstEvent.timestamp
    $endDate   = [DateTime]$(Get-EndDate -beginDate $beginDate -Timeperiod $GroupBy)
    $messageCounter = 1
    $EventGroupN     = 1

    # Iterate event rows, skipping the first row
    foreach ($row in ($EventRows | Select-Object -Skip 1)){
        $Event = Get-Event -EventRow $row -FieldCols $FieldNameCols -AttachmentCols $AttachmentCols -CustodianID $CustodianID
        if (([datetime]$Event.timestamp -gt $endDate) -or ($messageCounter -gt $MaxMessagesPerChat) ){
            $EventGroup = [PSCustomObject]@{
                "groupNumber"   = $EventGroupN
                "participants"  = $Participants;
                "conversations" = $Conversations;
                "events"        = $Events;
            }
            [void] $EventsGroups.Add($EventGroup)

            #Create a new Events Group
            $Events = New-Object Collections.Generic.List[PSCustomObject]
            $EventGroupN++

            # Update startTime and endTime
            $beginDate  = [DateTime]$Event.timestamp
            $endDate    = [DateTime](Get-EndDate -beginDate $beginDate -TimePeriod 'Week')
            $messageCounter = 0
        }

        # Add current event to Events list
        [void] $Events.Add($Event)
        
        if($Helper){
            $Helper.Update($taskID, 1)
            $Helper.Show()
        }
        $messageCounter++
    }

    $EventGroup = [PSCustomObject]@{
        "groupNumber"   = $EventGroupN
        "participants"  = $Participants;
        "conversations" = $Conversations;
        "events"        = $Events;
    }
    [void] $EventsGroups.Add($EventGroup)
    return $EventsGroups
}

Export-ModuleMember -Function *