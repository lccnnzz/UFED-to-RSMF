function Get-HeaderRow{
    param(
        [Parameter(Mandatory)] $Worksheet,
        [Parameter()] [int]$IheaderRow = 1
    )
    $HeaderRow = $Worksheet.UsedRange.Rows[$IheaderRow]
    return $HeaderRow
}

function Get-Rows{
    param(
        [Parameter(Mandatory)] $Worksheet,
        [Parameter(Mandatory)] [int]$beginRow,
        [Parameter(Mandatory)] [int]$endRow
    )

    $RowList = [System.Collections.ArrayList]@()
    foreach ($rowN in $beginRow..$endRow){
        [void] $RowList.Add($Worksheet.UsedRange.Rows[$rowN])
    }
    return $RowList
}

function Get-FieldColumn{
    param(
        [Parameter(Mandatory)] $HeaderRow,
        [Parameter(Mandatory)] [String]$FieldName
    )

    foreach ($field in $HeaderRow.Columns.Cells){
        if ($field.Value2 -eq "$FieldName") {
            return $field.column
        }
    }
}

function Get-FieldMultiColumn{
    param(
        [Parameter(Mandatory)] $HeaderRow,
        [Parameter(Mandatory)] [String]$Pattern
    )
    $col_pattern   = [Regex]::new($Pattern)
    $FieldColumns = [System.Collections.ArrayList]@()

    foreach ($field in $HeaderRow.Columns.Cells){
        if ($col_pattern.Match($field.Value2).Success){
            [void] $FieldColumns.Add($field.Column)      
        }
    }
    return $fieldColumns
}

function Get-ChatEventsCount{
    param(
        [Parameter(Mandatory)] $Worksheet,
        [Parameter(Mandatory)] $ChatNCol,
        [Parameter(Mandatory)] $beginRow,
        [Parameter(Mandatory)] $endRow
    )
 
    $ChatEvents = @{}
    $beginRange = $Worksheet.Cells($beginRow, $ChatNCol).Address()
    $endRange =   $Worksheet.Cells($endRow, $ChatNCol).Address()
    
    $EventsRange = $Worksheet.Range($beginRange, $endRange)

    $EventsRange.Value2 | Group-Object -NoElement | ForEach-Object{
        [void] $ChatEvents.Add($_.Values[0], $_.Count)
    }
    
    return $ChatEvents 
}

function Get-Participants{
    param(
        [Parameter(Mandatory)] $eventRow,
        [Parameter(Mandatory)] [int]$ParticipantsCol,
        [Parameter()] [String]$pattern = "[0-9]+@s.whatsapp.net"
    )
    $id_pattern   = [Regex]::new($pattern)
    $participants = New-Object Collections.Generic.List[PSCustomObject]
    
    foreach ($p in $eventRow.Columns[$ParticipantsCol].Value2.split("`r`n")){
        if ($p | Select-String -Pattern $id_pattern){
            $id, $display   = $p.split(" ", 2)
            $participant    = [PSCustomObject]@{
                "id"         = $id;
                "display"    = ($display -eq "") ? "Unknown" : $display;
                "account_id" = $id
            }
            [void]$participants.Add($participant)
        }
    }
    return $participants
}

function Get-SenderID{
    param(
        [Parameter(Mandatory)] $From,
        [Parameter()] [String]$pattern = "[0-9]+@s.whatsapp.net"
    )

    $SenderIdRegex = [regex]::new($pattern)
    $result = $SenderIdRegex.Match($From)
    if($result.value -ne ""){
        return $result.value
    }else{return $from.split(" ", 2)[0]}
}

function Get-Attachments{
    param(
        [Parameter(Mandatory)] $EventRow,
        [Parameter(Mandatory)] [int[]]$FieldCols
    )
    $attachments = New-Object Collections.Generic.List[PSCustomObject]

    foreach ($column in $FieldCols){
        $id = $EventRow.Columns[$column].Value2
        if ($id -ne ""){
            $attachment = [PSCustomObject]@{
                "id"      = $id
                "display" = $id
                "size"    = 100
            }
        [void] $attachments.add($attachment)
        }
    }
    return $attachments
}

function Get-Conversation{
    param(
        [Parameter(Mandatory)] $EventRow,
        [Parameter(Mandatory)] $FieldCols,
        [Parameter(Mandatory)] $Participants
    )

    $conversation = [PSCustomObject]@{
        "id"           = $eventRow.Columns[$FieldCols["Chat #"]].Value2
        "display"      = $eventRow.Columns[$FieldCols["Name"]].Value2
        "platform"     = $eventRow.Columns[$FieldCols["Source"]].Value2
        "type"         = ($eventRow.Columns[$FieldCols["Name"]].Value2 -eq "") ? "direct" : "channel"
        "participants" = $Participants.ID
    }
    return $conversation
}

function Format-Timestamp{
    param(
        [Parameter(Mandatory)] $Timestamp
    )
    $regexUTC       = [Regex]::new("\(UTC[+|-][0-12]\)")
    $match          = [datetime]($timestamp.Replace($regexUTC.Match($timestamp).Value, ""))
    $ISOTimestamp   = Get-Date $match -Format "yyyy-MM-ddTHH:mm:ss"
    return $ISOTimestamp
}

function Get-Event{
    param(
        [Parameter(Mandatory)] $eventRow,
        [Parameter(Mandatory)] $FieldCols,
        [Parameter(Mandatory)] $AttachmentCols,
        [Parameter()] $CustodianID
    )
    $attachments = New-Object Collections.Generic.List[PSCustomObject]
    
    $chatN     = [string]($eventRow.Columns[$FieldCols["Chat #"]].Value2)
    $messageN  = [string]($eventRow.Columns[$FieldCols["Instant Message #"]].Value2)
    $id        = ($chatN.PadLeft(3,'0'), $messageN.PadLeft(5,'0')) -join "-"
    $parent    = ($chatN.PadLeft(3,'0'), '1'.PadLeft(5,'0')) -join "-"
    $body      = $eventRow.Columns[$FieldCols["Body"]].Value2;
    $from      = $eventRow.Columns[$FieldCols["From"]].Value2;
    $sender    = Get-SenderID -From $from;
    $direction = switch($CustodianID) {"" {""} $sender {"outgoing"} default {"incoming"}}
    $type      = ($sender -match "System") ? "disclaimer" : "message"
    $timestamp = Format-Timestamp($($eventRow.Columns[$FieldCols["Timestamp: Time"]].Value2))
    $deleted   = ($eventRow.Columns[$FieldCols["Deleted - Instant Message"]] -ne "") ? $false : $true
    
    #Note: uses foreach instead of assignment to always return the list (an never $null)
    foreach ($attachment in (Get-Attachments -EventRow $eventRow -FieldCols $AttachmentCols)){
        [void] $attachments.Add($attachment)
    } 
    
    $event = [PSCustomObject]@{

        "conversation" = $chatN;
        "id"           = $id;
        "parent"       = $parent;
        "body"         = $body
        "participant"  = $sender;
        "direction"    = $direction;
        "type"         = $type;
        "timestamp"    = $timestamp;
        "deleted"      = $deleted;
        "attachments"  = $attachments
    }
    return $event
}

Export-ModuleMember -Function *