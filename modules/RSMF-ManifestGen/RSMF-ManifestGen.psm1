function New-RSMFManifest{
    param(
        [Parameter(Mandatory)] $conversationList,
        [Parameter(Mandatory)] $participantList,
        [Parameter(Mandatory)] $eventList,
        [Parameter()]          $eventCollectionId = ''
    )
    $Manifest = @{
        "version"           = "2.0.0";
        "participants"      = @();
        "events"            = @();
        "conversations"     = @();
        "eventcollectionid" = ""
    }

    $manifest.conversations     = $conversationList
    $manifest.participants      = $participantList
    $manifest.events            = $eventList
    $manifest.eventCollectionId = $eventCollectionId

    
    return ($manifest | ConvertTo-JSON -Depth 4)
}

function New-OutFilePath{
    param(
        [Parameter(Mandatory)] $OutputRootDir,
        [Parameter(Mandatory)] [string]$InputFileName,
        [Parameter()] [int]$ConversationN,
        [Parameter()] [int]$eventGroupN
    )

    $outputDirName = $(Get-Item $InputFileName).BaseName

    if($conversationN){
        $outputDirName = ($outputDirName, $([string]$conversationN).PadLeft(3, '0')) -Join "_"
    }

    if($eventGroupN){
        $outputDirName = ($outputDirName, $([string]$eventGroupN).PadLeft(2, '0')) -Join "_"
    }

    $outputDirPath = New-Item -Path $OutputRootDir -Name $outputDirName -ItemType "directory"
    $outputFilePath = Join-Path -Path $outputDirPath -ChildPath  "rsmf_manifest.json"
    return $outputFilePath
}

function Save-RSMFManifest{
    param(
        [Parameter(Mandatory)] $ManifestJSON,
        [Parameter(Mandatory)] $OutputFilePath
    )
    $manifestJSON | Out-File $outputFilePath
}

Export-ModuleMember -Function *

