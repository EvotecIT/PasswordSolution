function Format-ReminderDays {
    <#
    .SYNOPSIS
    Formats an array of reminder days into a readable, concise format.

    .DESCRIPTION
    This function accepts an array of numbers (which may include nested arrays)
    and always sorts them in ascending order. It then groups contiguous sequences
    (where each subsequent number is either equal to or exactly 1 greater than the previous)
    and formats groups of three or more unique numbers as a range. For clarity, if any
    number in the range is negative the range is displayed using " to " (e.g. "-500 to 500")
    to avoid confusion with hyphenated negatives.

    .PARAMETER Days
    An array of integers (or nested arrays of integers) representing reminder days.

    .EXAMPLE
    Format-ReminderDays -Days @(1,2,3,7,30)
    # Returns: "1-3, 7, 30"

    .EXAMPLE
    $xxx = @(500..-500), 60, 59, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45, 600, -505
    Format-ReminderDays -Days $xxx
    # Returns: "-505, -500 to 500, 600"

    .EXAMPLE
    Format-ReminderDays -Days @(15..-500), 60, 59, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45, 600, -505
    # Returns: -505, -500 to 15, 30, 59, 60, 600
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Array]$Days
    )

    # Flatten the input array (to handle nested arrays like those produced by ranges)
    $flatDays = @()
    foreach ($item in $Days) {
        if ($item -is [Array]) {
            $flatDays += $item
        } else {
            $flatDays += $item
        }
    }

    # Convert all items to integers.
    $flatDays = $flatDays | ForEach-Object { [int]$_ }

    # Always sort in ascending order.
    $sortedDays = $flatDays | Sort-Object

    # Group contiguous numbers.
    $groups = @()
    $currentGroup = @($sortedDays[0])
    for ($i = 1; $i -lt $sortedDays.Count; $i++) {
        $current = $sortedDays[$i]
        $previous = $sortedDays[$i - 1]
        # Group if the current number is equal (duplicate) or exactly 1 greater than the previous.
        if ($current -eq $previous -or $current -eq $previous + 1) {
            $currentGroup += $current
        } else {
            $groups += , @($currentGroup)
            $currentGroup = @($current)
        }
    }
    $groups += , @($currentGroup)

    # Format each group into a string.
    $formattedGroups = foreach ($group in $groups) {
        # Count unique numbers in the group.
        $uniqueCount = ($group | Select-Object -Unique).Count
        if ($uniqueCount -ge 3) {
            # Use " to " when any value is negative (for clarity)
            if ($group[0] -lt 0 -or $group[-1] -lt 0) {
                "$($group[0]) to $($group[-1])"
            } else {
                "$($group[0])-$($group[-1])"
            }
        } else {
            # For groups with fewer than 3 unique numbers, list the numbers separated by commas.
            $group -join ", "
        }
    }

    return ($formattedGroups -join ", ")
}