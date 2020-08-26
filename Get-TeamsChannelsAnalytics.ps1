<#
.SYNOPSIS
    Export in CSV the Analytics data of all Channels of a given Team
.DESCRIPTION
    Export in CSV the Analytics data of all Channels of a given Team
.PARAMETER TeamDisplayName
    The Name as shown in your Teams Client
.PARAMETER UserPrincipalName
    Your User Principal Name (often time it's your email address)
.PARAMETER Bearer
    Copy it from your Browser Network trace (See Readme.md for more detail)
.PARAMETER Path
    Where you want your csv to be exported
.EXAMPLE
    Get-TeamChannelAnalytics -TeamDisplayName "My Teams" -UserPrincipalName user@contoso.com -Bearer "Bearer XYZ8978979879[...]sx" -Path "C:\Temp"
    This will export all the analytics of the channels under My Teams to C:\Temp\My Teams-Analytics-2020-08-26.csv
.NOTES
    Author: Simon Poirier
    Source: https://github.com/poiriersimon/TeamChannelAnalytics
    LinkedIn: https://www.linkedin.com/in/poiriersimon/?locale=en_US
    Twitter: @SimonSaysEhlo
#>
function Get-TeamChannelAnalytics
{
    [OutputType([Hashtable])]
    [cmdletbinding()]
	param(
    [Parameter(Mandatory = $True)]
      	[string]$TeamDisplayName,
    [Parameter(Mandatory = $True)]
          [string]$UserPrincipalName,
    [Parameter(Mandatory = $True)]
          [string]$Bearer,
    [Parameter(Mandatory = $False)]
      	[string]$Path = "."
    )

    #Build the Headers based on web request
    $headers = @{
        "Authorization" = $Bearer
        "Content-Type"  = "application/json"
        }


    Connect-MicrosoftTeams -AccountId $UserPrincipalName
    $Groupid = (Get-Team -User $UserPrincipalName |where{$_.displayname -eq $TeamDisplayName}).GroupID
    $Channels = Get-TeamChannel -GroupId $GroupID

    $ChannelsStats = @()

    foreach($Channel in $Channels){
        $URL = "https://teams.microsoft.com/tas/prod/v2/teams/$($Groupid)/channels/$($Channel.Id)/summary-timeseries?metrics=all&timeperiod=last-ninety-days&includeTimeseries=postMessages,replyMessages,mentions,reactions,meetingsOrganized"
        $Reply = Invoke-WebRequest -Method get -UseDefaultCredentials -Uri $URL -Headers $headers
        $Data = ConvertFrom-Json $Reply.Content
        $ChannelStat = New-Object PSObject
        $ChannelStat | Add-Member NoteProperty -Name "Channel" -Value $Channel.DisplayName
        $ChannelStat | Add-Member NoteProperty -Name "replyMessages" -Value $Data.channels.metrics.replyMessages.value
        $ChannelStat | Add-Member NoteProperty -Name "LastreplyMessages" -Value $Data.channels.metrics.replyMessages.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "postMessages" -Value $Data.channels.metrics.postMessages.value
        $ChannelStat | Add-Member NoteProperty -Name "LastpostMessages" -Value $Data.channels.metrics.postMessages.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "meetingsOrganized" -Value $Data.channels.metrics.meetingsOrganized.value
        $ChannelStat | Add-Member NoteProperty -Name "LastmeetingsOrganized" -Value $Data.channels.metrics.meetingsOrganized.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "activeUsers" -Value $Data.channels.metrics.activeUsers.value
        $ChannelStat | Add-Member NoteProperty -Name "LastactiveUsers" -Value $Data.channels.metrics.activeUsers.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "channelMessages" -Value $Data.channels.metrics.channelMessages.value
        $ChannelStat | Add-Member NoteProperty -Name "LastchannelMessages" -Value $Data.channels.metrics.channelMessages.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "reactions" -Value $Data.channels.metrics.reactions.value
        $ChannelStat | Add-Member NoteProperty -Name "Lastreactions" -Value $Data.channels.metrics.reactions.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelStat | Add-Member NoteProperty -Name "mentions" -Value $Data.channels.metrics.mentions.value
        $ChannelStat | Add-Member NoteProperty -Name "Lastmentions" -Value $Data.channels.metrics.mentions.timeSeries | sort -Descending date | where {$_.value -gt 0} |select -First 1
        $ChannelsStats += $ChannelStat
    }
    $ChannelsStats | Export-csv "$($Path)\$($TeamDisplayName)-Analytics-$(Get-date -Format yyyy-MM-dd).csv" -NoTypeInformation
}

