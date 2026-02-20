# Script to "mute" Slack's tray notification dot
# To achieve this, the script will override the Slack tray icons with the blank one

# Only change the following values if the names of the icons either change or new dotted ones appear
$blankIconFile = 'slack-taskbar-rest.ico'
$dottedIconFiles = 'slack-taskbar-highlight.ico', 'slack-taskbar-unread.ico'


function Main {

    "This script 'disables' Slack's tray notification dots"
    "(Please note that it won't work if you use the Windows 10 Universal App."
    " This is because its directory is not writable by regular (even admin) users."
    " Sorry for that :/)`n"

    $slackVersion = Read-Host -Prompt "Please type your slack version in the from '1.2.3':`n(You can find it via Help > About Slack)"
    $slackDirectory = "$env:LOCALAPPDATA\slack\app-$slackVersion"

    "`nGoing to mute slack in the following directory: $slackDirectory"

    if (ExecutionShouldContinue) {
        $error.clear()

        # Get default icons folder path
        $iconsFolder = "$slackDirectory\resources\app.asar.unpacked\dist\static"

        foreach ($dottedIcon in $dottedIconFiles ) {
            # Rename dotted icon (as backup)
            Rename-Item -Path "$iconsFolder\$dottedIcon" -NewName "$iconsFolder\$dottedIcon.bak.ico"

            # Copy blank icon using its name
            Copy-Item -Path "$iconsFolder\$blankIconFile" -Destination "$iconsFolder\$dottedIcon"
        }

        if (!$error) {
            ExitWithMessage "`nSuccessfully muted Slack"
        }
        else {
            ExitWithMessage "Seems like something went wrong...`n
Maybe you forgot to set the Slack app directory?
Or did you execute the script twice and the backup files already exist?`n
Check the error message above to get a hint..."
        }
    }
    else {
        ExitWithMessage "`nCancelled"
    }
}

function ExecutionShouldContinue {
    $title = ''
    $question = 'Continue?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    return ($decision -eq 0)
}

function Wait-For-Keypress {
    Write-Host -NoNewLine "`nPress any key to continue...`n"
    [Console]::ReadKey($true) | Out-Null
}

function ExitWithMessage([string]$messageToDisplay) {
    Write-Host $messageToDisplay
    Wait-For-Keypress
    Exit
}

Main

