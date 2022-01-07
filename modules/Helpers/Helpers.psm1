class ProgressHelper{
    [System.Collections.ArrayList] $Tasks = @()

    [void]Add([int]$Id, [int]$ParentID, [String]$Activity, [int]$Items){
        $task = [PSCustomObject]@{
            "id"        = $Id;
            "parentid"  = $ParentID;
            "activity"  = $Activity;
            "items"     = $Items;
            "completed" = 0
        }
        $this.Tasks.Add($task)
    }

    [void]Update([int]$TaskId, [int]$ItemsToAdd){
        $task = $this.Tasks | Where-Object {$_.id -eq $TaskId}
        $task.completed += $ItemsToAdd
    }

    [void]Show(){
        foreach ($task in $this.Tasks){
            $percent = [math]::Round( ((100 / $task.items) * $task.completed),0)
            $status = "| $($task.completed) of $($task.items) items"
            if ($task.parentid -eq $task.id){
                Write-Progress -Id $task.id -Activity $task.activity -Status $status -PercentComplete $percent
            } else{
                Write-Progress -Id $task.id -ParentID $task.parentid -Activity $task.activity -Status $status -PercentComplete $percent
            } 
        }
    }
}

class StatsHelper{
    [System.Collections.ArrayList] $Counters = @()
    [void]Add([String]$Name){
        $counter = [PSCustomObject]@{
            "name"        = $Name;
            "count" = 0
        }
        $this.Counters.Add($counter)
    }

    [void]Update([String]$Counter, [int]$ItemsToAdd){
        $counterToUpdate = $this.Counters | Where-Object {$_.name -eq $Counter}
        $counterToUpdate.count += $ItemsToAdd
    }

    [void]Increment([String]$Counter){
        Update $Counter 1
    }

    [void]GetCount([String]$Counter){
        ($this.Counters | Where-Object {$_.name -eq $Counter}).count
    }
    [System.Collections.ArrayList]GetShats(){
        return $this.Counters 
    }

}