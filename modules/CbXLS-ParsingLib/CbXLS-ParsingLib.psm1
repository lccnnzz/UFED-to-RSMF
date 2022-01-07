#========================================================
# Modules
#========================================================
Import-Module -Name .\modules\Utils -DisableNameChecking
$platforms = Import-PowerShellDataFile .\Platforms.psd1
#========================================================
# Functions
#========================================================
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

function Get-FieldNameCols{
    param(
        [Parameter(Mandatory)] $FieldNamelist,
        [Parameter(Mandatory)] $HeaderRow
    )   
    
    $FieldNameCol = @{}
    foreach ($fieldName in $FieldNameList){
        $fieldColumn = Get-FieldColumn -HeaderRow $HeaderRow -FieldName $fieldName
        [void] $FieldNameCol.Add($fieldName, $fieldColumn)
    }

    return $FieldNameCol
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

function Get-UserPattern{
    param(
        [Parameter(Mandatory)] $EventRow,
        [Parameter(Mandatory)] $SourceCol
    )
    $platform = $EventRow.Columns[$SourceCol].Value2
    if ($Platforms.containsKey($platform)){
        return $Platforms[$platform].accountPattern
    }
    else{
        return $Platforms['Generic'].accountPattern
    }
}

function Get-ChatIcon{
    param(
        [Parameter(Mandatory)] $EventRow,
        [Parameter(Mandatory)] $SourceCol
    )
    $platform = $EventRow.Columns[$SourceCol].Value2
    if ($Platforms.containsKey($platform)){
        return $Platforms[$platform].icon
    }
    else{
        return $Platforms['Generic'].icon
    }
}

function Get-Participants{
    param(
        [Parameter(Mandatory)] $eventRow,
        [Parameter(Mandatory)] [int]$ParticipantsCol,
        [Parameter(Mandatory)] $UserPattern
    )
    $id_pattern   = [Regex]::new($pattern)
    $participants = New-Object Collections.Generic.List[PSCustomObject]
    
    foreach ($p in $eventRow.Columns[$ParticipantsCol].Value2.split("`r`n")){
        if ($p | Select-String -Pattern $UserPattern){
            $id, $display   = $p.split(" ", 2)
            $custom = New-Object Collections.Generic.List[PSCustomObject]
            $hash = [PSCustomObject]@{
                "name" = "hash"
                "value" = $id #todo
            }
            [void] $custom.Add($hash)

            $participant    = [PSCustomObject]@{
                "id"         = $id;
                "display"    = ($display -eq "") ? "Unknown" : $display;
                "account_id" = $id
                "custom"     = $custom
            }
            [void]$participants.Add($participant)
        }
    }
    return $participants
}

function Get-SenderID{
    param(
        [Parameter(Mandatory)] $From,
        [Parameter(Mandatory)] [String]$pattern
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
        [Parameter()] $AttachmentDir,
        [Parameter(Mandatory)] [int[]]$FieldCols
    )
    $attachments = New-Object Collections.Generic.List[PSCustomObject]
    foreach ($column in $FieldCols){
        $cell = $EventRow.Columns[$column]
        if ($cell.Value2 -ne ""){
            foreach ($hlink in $cell.Hyperlinks){
                $address = Join-Path -Path $AttachmentDir -ChildPath $hlink.Address
                $attachment = [PSCustomObject]@{
                    "id"      = $hlink.Name
                    "display" = $hlink.Name
                    "size"    = (Get-Item $address).Length
                }
            [void] $attachments.add($attachment)
            }
        }
    }
    return $attachments
}

function Get-Conversation{
    param(
        [Parameter(Mandatory)] $EventRow,
        [Parameter(Mandatory)] $FieldCols,
        [Parameter(Mandatory)] $Participants,
        [Parameter(Mandatory)] $Icon
    )

    $source = $eventRow.Columns[$FieldCols["Source"]].Value2
    foreach ($key in $Platforms.keys){ 
        if ($source -match $key){
            $platform = $key
            break
        }
    }

    $conversation = [PSCustomObject]@{
        "id"           = $eventRow.Columns[$FieldCols["Chat #"]].Value2
        "display"      = $eventRow.Columns[$FieldCols["Name"]].Value2
        "platform"     = $platform
        "type"         = ($eventRow.Columns[$FieldCols["Name"]].Value2 -eq "") ? "direct" : "channel"
        "icon"         = $Icon
        "participants" = $Participants.ID

    }
    return $conversation
}

function Get-Event{
    param(
        [Parameter(Mandatory)] $eventRow,
        [Parameter(Mandatory)] $FieldCols,
        [Parameter(Mandatory)] $AttachmentCols,
        [Parameter(Mandatory)] $UserPattern,
        [Parameter()] $AttachmentDir,
        [Parameter()] $CustodianID = ''

    )
    $attachments = New-Object Collections.Generic.List[PSCustomObject]
    $custom      = New-Object Collections.Generic.List[PSCustomObject]

    $chatN          = [string]($eventRow.Columns[$FieldCols["Chat #"]].Value2)
    $messageN       = [string]($eventRow.Columns[$FieldCols["Instant Message #"]].Value2)
    $id             = ($chatN.PadLeft(3,'0'), $messageN.PadLeft(5,'0')) -join "-"
    $parent         = ($chatN.PadLeft(3,'0'), '1'.PadLeft(5,'0')) -join "-"
    $body           = $eventRow.Columns[$FieldCols["Body"]].Value2;
    $from           = $eventRow.Columns[$FieldCols["From"]].Value2;
    $sender         = Get-SenderID -From $from -Pattern $UserPattern;
    $status         = $eventRow.Columns[$FieldCols["Status"]].Value2
    $direction      = ($status -match "Sent") ? "outgoing" : "incoming"
    $starredMessage = $eventRow.Columns[$FieldCols["Starred Message"]].Value2
    $importance     = ($status -match "Yes") ? "high" : "normal"
    $type           = ($sender -match "System") ? "disclaimer" : "message"
    $timestamp      = Format-Timestamp($($eventRow.Columns[$FieldCols["Timestamp: Time"]].Value2))
    $deleted        = ($eventRow.Columns[$FieldCols["Deleted - Instant Message"]] -ne "") ? $false : $true
    
    #Note: uses foreach instead of assignment to always return the list (an never $null)
    foreach ($attachment in (Get-Attachments -EventRow $eventRow -FieldCols $AttachmentCols -AttachmentDir $AttachmentDir)){
        [void] $attachments.Add($attachment)
    }
    $eventhash = [PSCustomObject]@{
        "name" = "eventHash"
        "value" = ConcatHash -Values $sender, $timestamp, $body -Algorithm 'MD5'
    }
    [void] $custom.Add($eventhash)
    
    $event = [PSCustomObject]@{
        "conversation" = $chatN;
        "id"           = $id;
        "parent"       = $parent;
        "body"         = $body;
        "participant"  = $sender;
        "direction"    = $direction;
        "type"         = $type;
        "timestamp"    = $timestamp;
        "deleted"      = $deleted;
        "importance"   = $importance;
        "attachments"  = $attachments;
        "custom"       = $custom
    }
    return $event
}

Export-ModuleMember -Function *