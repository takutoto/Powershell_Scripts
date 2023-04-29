# Script Author: Brian Pringle
# Script Purpose: Finds COM GUIDs by CLSID, AppID, or Application. Supports exact and partial matches.

Function Pause($Message = "Press any key to continue . . .")
{
    # Only show a prompt if this is being run outside of ISE.
    If (!(Test-Path variable:psISE) -Or !$psISE)
    {
        Write-Host $Message
        [void][System.Console]::ReadKey($true)
    }
}

# Function Source: https://stackoverflow.com/questions/27689678/color-parts-of-string-in-powershell
function Write-HostColored()
{
    <#
        .SYNOPSIS
            A wrapper around Write-Host that supports selective coloring of
            substrings.

        .DESCRIPTION
            In addition to accepting a default foreground and background color,
            you can embed one or more color specifications in the string to write, 
            using the following syntax:
            #<fgcolor>[:<bgcolor>]#<text>#

            <fgcolor> and <bgcolor> must be valid [ConsoleColor] values, such as 'green' or 'white' (case does not matter).
            Everything following the color specification up to the next '#' or, impliclitly, the end of the string
            is written in that color.

            Note that nesting of color specifications is not supported.
            As a corollary, any token that immediately follows a color specification is treated
            as text to write, even if it happens to be a technically valid color spec too.
            This allows you to use, e.g., 'The next word is #green#green#.', without fear
            of having the second '#green' be interpreted as a color specification as well.

        .PARAMETER ForegroundColor
            Specifies the default text color for all text portions
            for which no embedded foreground color is specified.

        .PARAMETER BackgroundColor
            Specifies the default background color for all text portions
            for which no embedded background color is specified.

        .PARAMETER NoNewline
            Output the specified string withpout a trailing newline.

        .NOTES
            While this function is convenient, it will be slow with many embedded colors, because,
            behind the scenes, Write-Host must be called for every colored span.

        .EXAMPLE
            Write-HostColored "#green#Green foreground.# Default colors. #blue:white#Blue on white."

        .EXAMPLE
            '#black#Black on white (by default).#Blue# Blue on white.' | Write-HostColored -BackgroundColor White
    #>

    [CmdletBinding(ConfirmImpact = 'None', SupportsShouldProcess = $false, SupportsTransactions = $false)]
    param
    (
        [parameter(Position = 0, ValueFromPipeline = $true)]
        [string[]] $Text
        , [switch] $NoNewline
        , [ConsoleColor] $BackgroundColor = $host.UI.RawUI.BackgroundColor
        , [ConsoleColor] $ForegroundColor = $host.UI.RawUI.ForegroundColor
    )

    Begin
    {
        # If text was given as an operand, it'll be an array.
        # Like Write-Host, we flatten the array into a single string
        # using simple string interpolation (which defaults to separating elements with a space,
        # which can be changed by setting $OFS).
        if ($Text -ne $null)
        {
            $Text = "$Text"
        }
    }

    Process
    {
        if ($Text)
        {
            # Start with the foreground and background color specified via
            # -ForegroundColor / -BackgroundColor, or the current defaults.
            $curFgColor = $ForegroundColor
            $curBgColor = $BackgroundColor

            # Split message into tokens by '#'.
            # A token between to '#' instances is either the name of a color or text to write (in the color set by the previous token).
            $tokens = $Text.split("#")

            # Iterate over tokens.            
            $prevWasColorSpec = $false
            foreach($token in $tokens)
            {
                if (-not $prevWasColorSpec -and $token -match '^([a-z]+)(:([a-z]+))?$')
                {
                    # A potential color spec.
                    # If a token is a color spec, set the color for the next token to write.
                    # Color spec can be a foreground color only (e.g., 'green'), or a foreground-background color pair (e.g., 'green:white')
                    try
                    {
                        $curFgColor = [ConsoleColor]$matches[1]
                        $prevWasColorSpec = $true
                    }
                    catch {}

                    if ($matches[3])
                    {
                        try
                        {
                            $curBgColor = [ConsoleColor]$matches[3]
                            $prevWasColorSpec = $true
                        }
                        catch {}
                    }

                    if ($prevWasColorSpec)
                    {
                        continue                    
                    }
                }

                $prevWasColorSpec = $false

                if ($token)
                {
                    <#
                        A text token: write with (with no trailing line break).
                        !! In the ISE - as opposed to a regular PowerShell console window,
                        !! $host.UI.RawUI.ForegroundColor and $host.UI.RawUI.ForegroundColor inexcplicably 
                        !! report value -1, which causes an error when passed to Write-Host.
                        !! Thus, we only specify the -ForegroundColor and -BackgroundColor parameters
                        !! for values other than -1.
                    #>
                    $argsHash = @{}
                    if ([int]$curFgColor -ne -1)
                    {
                        $argsHash += @{ 'ForegroundColor' = $curFgColor }
                    }

                    if ([int]$curBgColor -ne -1)
                    {
                        $argsHash += @{ 'BackgroundColor' = $curBgColor }
                    }

                    Write-Host -NoNewline @argsHash $token
                }

                # Revert to default colors.
                $curFgColor = $ForegroundColor
                $curBgColor = $BackgroundColor
            }
        }
        # Terminate with a newline, unless suppressed
        if (-not $NoNewLine)
        {
            Write-Host
        }
    }
}

$search_id = Read-Host "Enter a full or partial ID for a CLSID, AppID, or Application"

$clsid_path = 'HKLM:\SOFTWARE\Classes\CLSID'
$appid_path = 'HKLM:\SOFTWARE\Classes\AppID'

$valid_clsid_path = ""
$valid_appid_path = ""

$match_color = "green"

<#
    Check for a full GUID by testing its path in the registry.
    The path requires the GUID to be enclosed by {}, so encase it if necessary.
#>
if (Test-Path $clsid_path\$search_id)
{
    $valid_clsid_path = "$clsid_path\$search_id"
}
elseif (Test-Path "$clsid_path\{$search_id}")
{
    $valid_clsid_path = "$clsid_path\{$search_id}"
}
elseif (Test-Path $appid_path\$search_id)
{
    $valid_appid_path = "$appid_path\$search_id"
}
elseif (Test-Path "$appid_path\{$search_id}")
{
    $valid_appid_path = "$appid_path\{$search_id}"
}

<#
    Get item properties based on path validation:
    - Valid CLSID Path: Find the properties using the CLSID's exact path in the registry.
    - Valid AppID Path: Find the properties by searching for the CLSIDs using the AppID.
    - No Valid Path: Find properties by searching for CLSIDs where the GUID partially matches the CLSID or AppID.
#>
if ($valid_clsid_path)
{
    $results = Get-ItemProperty -Path $valid_clsid_path |
    Select PSChildName, AppID, `(default`)
}
elseif ($valid_appid_path)
{
    Write-Host "`nSearching for AppID by exact GUID match. Please wait."

    $results = Get-ItemProperty -Path $clsid_path\* |
    Select PSChildName, AppID, `(default`) |
    Where-Object AppID -Match $search_id
}
else #No Valid Path
{
    Write-Host "`nCould not find CLSID or AppID by exact GUID match.`nAttempting partial ID search. Please wait."

    $results = Get-ItemProperty -Path $clsid_path\* |
    Select PSChildName, AppID, `(default`) |
    Where-Object{$_.PSChildName -like "*$search_id*" -or $_.AppID -like "*$search_id*" -or $_.{(default)} -like "*$search_id*"}
}

if ($results)
{
    <#
        Formats the results as an output table.
        Renames the following headers and highlights ID:
            1. (default)   -> Application
            2. PSChildName -> CLSID
            3. AppID       -> AppID
    #>
    $results = $results |
    Sort -Property `(default`) |
    # TODO: Need to fix issue where the case is changed when we replace text to add color tags.
    Format-Table -Wrap @{L = "Application"; E = {$_.{(default)} -replace $search_id, "#$match_color#$search_id#"}},
        @{L = "CLSID"; E = {$_.PSChildName -replace $search_id, "#$match_color#$search_id#"}},
        @{L = "AppID"; E = {$_.AppID -replace $search_id, "#$match_color#$search_id#"}} |
    Out-String

    # TODO: Need to fix issue where headers become misaligned with the text when rows are search_id is colored.
    Write-HostColored $results
}
else
{
    Write-HostColored "`nUnable to find CLSID, AppID, or Application using ID filter #$match_color#$search_id#."
}

Pause