$yaketysaxfile = 'c:\yaketysax.wav' # This MUST be the full path.

<#
.Synopsis
   Creates a timer with display
.DESCRIPTION
   Creates a timer that writes the progress of the current action and the estimated remaining time.

   Suppose you have $Items, an array of items to process.
   First, you create a timer initialized to the number of items in the array:
   $MyTimer = New-Timer -TotalItems $Items.count -Activity 'Processing items'

   Then, while processing the items, you just call the next() function of the timer:
   foreach ( $item in $Items ) {
     # do things with $item
     $MyTimer.next()
   }

   The progressbar is given a random Id, but the -Id parameter can be used to give all progressbars
   in the script the same Id.
.EXAMPLE
   $MyTimer = New-Timer -TotalItems 100
   Creates a timer.
.EXAMPLE
   $MyTimer.next()
   Increments the timer.
.INPUTS
    None
    You cannot pipe input to this cmdlet.
.OUTPUTS
    Object
    New-Timer returns the timer object that is created.
#>
function New-Timer
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Total number of items to process
        [Parameter(Mandatory=$true,Position=0)]
        [int]$TotalItems,

        # Activity name
        [Parameter(Position=1)]
        [string]$Activity = 'processing',

        # Progressbar Id
        [int]$Id = (Get-Random),

        # For very long processes
        [switch]$yaketysax
    )

    If ( $TotalItems -le 0 ) {
        Throw 'The number of items must be positive'
    }

    $Timer = New-Object -TypeName psobject -Property @{
        total = $TotalItems
        current = 0
        timer = [Diagnostics.Stopwatch]::StartNew()
        activity = $Activity
        id = $Id
        media = $null
    }

    if ( $YaketySax ) {
        $Timer.media = New-Object System.Media.SoundPlayer
        $Timer.media.SoundLocation = $yaketysaxfile
        $Timer.media.playlooping()
    }
    Add-Member -InputObject $Timer -MemberType ScriptProperty -Name percent -Value { ( 100 * $this.current / $this.total ) }
    Add-Member -InputObject $Timer -MemberType ScriptProperty -Name timeleft -Value { ( $this.timer.ElapsedMilliseconds / (1000*$this.percent) ) * ( 100 - $this.percent ) }
    Add-Member -InputObject $Timer -MemberType ScriptMethod -Name progress -Value { Write-Progress -Activity $this.activity -PercentComplete $this.percent -Id $this.id -SecondsRemaining $this.timeleft }
    Add-Member -InputObject $Timer -MemberType ScriptMethod -Name next -Value { $this.current++ ; if ( $this.media -and $this.current -ge $this.total ) { $this.media.stop() } ; $this.progress() }
    Add-Member -InputObject $Timer -MemberType ScriptMethod -Name done -Value { Write-Progress -Activity $this.activity -Id $this.id -Completed }
    
    Return $Timer
}

#region demo

help New-Timer -ShowWindow

# array creation
$Items = 1..100

# timer initialisation
$timer = New-Timer -TotalItems $Items.count -Activity 'testing' -YaketySax

# timer use
$Items | % {
    Write-Host test -ForegroundColor ( $_ % 16 ) # random treatment
    Start-Sleep -Milliseconds 100                # let's make things slower for a better view
    $timer.next()                                # increment timer.
}
#endregion
