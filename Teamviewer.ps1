function Get-TVIncomingLog_byDate {
    <#
    .SYNOPSIS
        Parses the connections_incoming.txt file and returns data before, after, or between a specific date
    
    .PARAMETER File
        Location to the connections_incoming.txt file
    
    .PARAMETER BeforeDate
        Returns data before the specified date
    
    .PARAMETER AfterDate
        Returns data after the specified date
    
    .EXAMPLE
        Get-TVIncomingLog_byDate -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -before "12/25/2020"

        Returns data before December 25, 2020
    
    .EXAMPLE
        Get-TVIncomingLog_byDate -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -after "12/25/2020"

        Returns data after December 25, 2020

    .EXAMPLE
        Get-TVIncomingLog_byDate -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -before "3/1/2021" -after "12/25/2020"

        Returns data after March 1, 2021 before December 25, 2020
    #>

    [CmdletBinding()]
    param(
        [string]$File,
        [datetime]$BeforeDate,
        [datetime]$AfterDate
    )

    # Read the file content
    $logs = Get-Content $File

    # Initialize an empty array to store log objects
    $obj = @()

    # Process each log entry
    $logs = $logs -replace (' ', '_')

    $obj = foreach ($log in $logs) {
        $dur = ''
        $logParts = $log -split "\s+"
        $logParts = $logParts -replace ('_', ' ')

        # Initialize start and end date
        $startDate = $null
        $endDate = $null

        # Parse dates with error handling
        try { 
            if ($logParts[2]) { $startDate = [datetime]::ParseExact($logParts[2], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        try { 
            if ($logParts[3]) { $endDate = [datetime]::ParseExact($logParts[3], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        # Calculate duration if both dates are valid
        if ($startDate -and $endDate) {
            $dur = New-TimeSpan -Start $startDate -End $endDate
        }

        # Create the custom object for each log entry
        [PSCustomObject]@{
            IncomingID     = $logParts[0]
            DisplayName    = $logParts[1]
            StartDate      = $startDate
            EndDate        = $endDate
            Duration       = if ($dur) { $dur.ToString("dd'd.'hh'h:'mm'm:'ss's'") } else { "Invalid Duration" }
            LoggedOnUser   = $logParts[4]
            ConnectionType = $logParts[5]
            ConnectionID   = $logParts[6]
        }
    }

    # Filter logs based on the date range provided
    if ($AfterDate -and $BeforeDate) {
        return $obj | Where-Object { $_.StartDate -gt $AfterDate -and $_.StartDate -lt $BeforeDate }
    }
    elseif ($AfterDate) {
        return $obj | Where-Object { $_.StartDate -gt $AfterDate }
    }
    elseif ($BeforeDate) {
        return $obj | Where-Object { $_.StartDate -lt $BeforeDate }
    }
    else {
        return $obj
    }
}

function Get-TVIncomingLog_Top10Duration {
    <#
    .SYNOPSIS
        Parses the connections_incoming.txt file and returns the duration for the top 10 longest or shortest incoming connections

    .PARAMETER File
        Location to the connections_incoming.txt file

    .PARAMETER shortest
        Used to select the shortest connections

    .PARAMETER longest
        Used to select the longest connections

    .EXAMPLE
        Get-TVIncomingLog_Top10Duration -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt"

        Returns the duration for all incoming connections

    .EXAMPLE
        Get-TVIncomingLog_Top10Duration -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -shortest

        Returns the top 10 shortest durations for all incoming connections

    .EXAMPLE
        Get-TVIncomingLog_Top10Duration -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -longest

        Returns the top 10 longest durations for all incoming connections
    #>
    [CmdletBinding()]
    param(
        [string]$File,
        [switch]$shortest,
        [switch]$longest
    )

    # Read the file content
    $logs = Get-Content $File

    # Initialize an empty array to store log objects
    $obj = @()

    # Process each log entry
    $logs = $logs -replace (' ', '_')

    $obj = foreach ($log in $logs) {
        $dur = ''
        $logParts = $log -split "\s+"
        $logParts = $logParts -replace ('_', ' ')

        # Initialize start and end date
        $startDate = $null
        $endDate = $null

        # Parse dates with error handling
        try {
            if ($logParts[2]) { $startDate = [datetime]::ParseExact($logParts[2], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        try {
            if ($logParts[3]) { $endDate = [datetime]::ParseExact($logParts[3], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        # Calculate duration if both dates are valid
        if ($startDate -and $endDate) {
            $dur = New-TimeSpan -Start $startDate -End $endDate
        }

        # Create the custom object for each log entry
        [PSCustomObject]@{
            IncomingID     = $logParts[0]
            DisplayName    = $logParts[1]
            StartDate      = $startDate
            EndDate        = $endDate
            Duration       = if ($dur) { $dur.ToString("dd'd.'hh'h:'mm'm:'ss's'") } else { "Invalid Duration" }
            LoggedOnUser   = $logParts[4]
            ConnectionType = $logParts[5]
            ConnectionID   = $logParts[6]
        }
    }

    # Sort by duration and select top 10 based on shortest or longest
    if ($shortest) {
        return $obj | Sort-Object Duration | Select-Object -First 10 | Format-Table
    }
    elseif ($longest) {
        return $obj | Sort-Object Duration -Descending | Select-Object -First 10 | Format-Table
    }
    else {
        return $obj
    }
}

function Get-TVIncomingLog_Unique {
    <#
    .SYNOPSIS
        Parses the connections_incoming.txt and returns the unique incoming IDs, display names, or logged on users

    .PARAMETER File
        Location to the connections_incoming.txt file

    .PARAMETER IncomingID
        Used to select the returning of unique incoming IDs

    .PARAMETER DisplayName
        Used to select the returning of unique display names

    .PARAMETER LoggedOnUser
        Used to select the returning of unique logged on users

    .EXAMPLE
        Get-TVIncomingLog_Unique -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -incomingid

        Returns the entries containing unique incoming IDs

    .EXAMPLE
        Get-TVIncomingLog_Unique -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -displayname

        Returns the entries containing unique display names

    .EXAMPLE
        Get-TVIncomingLog_Unique -file "C:\Program Files (x86)\TeamViewer\connections_incoming.txt" -loggedonuser

        Returns the entries containing unique logged on user
    #>

    [CmdletBinding()]
    param(
        [string]$File,
        [switch]$IncomingID,
        [switch]$DisplayName,
        [switch]$LoggedOnUser
    )

    # Read the file content
    $logs = Get-Content $File
    Write-Host "got file: $File"

    # Initialize an empty array to store log objects
    $obj = @()

    # Process each log entry
    $logs = $logs -replace (' ', '_')

    $obj = foreach ($log in $logs) {
        $dur = ''
        $logParts = $log -split "\s+"
        $logParts = $logParts -replace ('_', ' ')

        # Initialize start and end date
        $startDate = $null
        $endDate = $null

        # Parse dates with error handling
        try {
            if ($logParts[2]) { $startDate = [datetime]::ParseExact($logParts[2], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        try {
            if ($logParts[3]) { $endDate = [datetime]::ParseExact($logParts[3], 'dd-MM-yyyy HH:mm:ss', $null) }
        } catch {}

        # Calculate duration if both dates are valid
        if ($startDate -and $endDate) {
            $dur = New-TimeSpan -Start $startDate -End $endDate
        }

        # Create the custom object for each log entry
        [PSCustomObject]@{
            IncomingID     = $logParts[0]
            DisplayName    = $logParts[1]
            StartDate      = $startDate
            EndDate        = $endDate
            Duration       = if ($dur) { $dur.ToString("dd'd.'hh'h:'mm'm:'ss's'") } else { "Invalid Duration" }
            LoggedOnUser   = $logParts[4]
            ConnectionType = $logParts[5]
            ConnectionID   = $logParts[6]
        }
    }

    # Return unique values based on the specified criteria
    if ($IncomingID) {
        return $obj | Sort-Object -Property IncomingID -Unique | Select-Object IncomingID
    }
    elseif ($DisplayName) {
        return $obj | Sort-Object -Property DisplayName -Unique | Select-Object DisplayName
    }
    elseif ($LoggedOnUser) {
        return $obj | Sort-Object -Property LoggedOnUser -Unique | Select-Object LoggedOnUser
    }
}

function Get-TVLogFile_RunTimes {
    <#
    .SYNOPSIS
        Parses the Teamviewer15_logfile.log and Teamviewer15_logfile_OLD.log for the run time of the Teamviewer program

    .PARAMETER directory
        Used to specify the directory containing the log files

    .EXAMPLE
        Get-TVLogFile_RunTimes -directory "C:\Program Files (x86)\TeamViewer"

        Will search the specified directory and parse the log files, returning the run time of the Teamviewer program
    #>
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Get all relevant log files in the directory
    $logs = Get-ChildItem -Path ($directory + "\TeamViewer15_Logfile*.log")
    Write-Host "Logs found: $($logs.Count)"

    # Initialize an empty array for storing the results
    $obj = @()

    # Iterate through each log file
    foreach ($line in $logs.FullName) {
        # Read and search for relevant patterns in the log file
        $logfile = Get-Content $line | Select-String -Pattern '(2]::processconnected:) | (Closing TeamViewer)'

        # Process each matched item
        foreach ($item in $logfile) {
            $data = $shutdownData = ''

            # Check if the line indicates program start (process connected)
            if ($item -like "*2]::processconnected: *") {
                $data = $item.Line -split ' '
                $data = $data[0] + ' ' + $data[1]

                # Safely convert to DateTime
                try { 
                    $data = [datetime]$data
                } catch {
                    $data = $null
                }

                # Check for the shutdown event after the process start
                $index = ($logfile.IndexOf($item) + 1)
                if ($logfile[$index] -match "Closing TeamViewer") {
                    $shutdownData = $logfile[$index].Line -split ' '
                    $shutdownData = $shutdownData[0] + ' ' + $shutdownData[1]

                    # Safely convert to DateTime
                    try {
                        $shutdownData = [datetime]$shutdownData
                    } catch {
                        $shutdownData = $null
                    }
                }

                # Create and add the custom object to the result array
                if ($data -and $shutdownData) {
                    $obj += [pscustomobject]@{
                        ProgramStart = $data.ToString("MM/dd/yyyy HH:mm:ss")
                        ProgramEnd   = $shutdownData.ToString("MM/dd/yyyy HH:mm:ss")
                    }
                }
            }
        }
    }

    # Return the results
    return $obj
}

function Get-TVLogFile_AccountLogons {
    <#
    .SYNOPSIS
        Parses the Teamviewer15_logfile.log and Teamviewer15_logfile_OLD.log for the account names

    .PARAMETER directory
        Used to specify the directory containing the log files

    .EXAMPLE
        Get-TVLogFile_AccountLogons -directory "C:\Program Files (x86)\TeamViewer"

        Returns the account names
    #>
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Get all relevant log files in the directory
    $logs = Get-ChildItem -Path ($directory + "\teamviewer15_Logfile*.log")

    # Initialize an empty array for storing the results
    $obj = @()

    # Iterate through each log file
    foreach ($line in $logs.FullName) {
        # Read and search for relevant patterns in the log file
        $logfile = Get-Content $line | Select-String -Pattern "(HandleLoginFinished|HandleLoginFinishedWithOld): Authentication successful", "Account::Logout: Account session terminated successfully"

        # Process each matched item
        foreach ($item in $logfile) {
            $data = $logoutData = '--'

            # Check for successful authentication (login)
            if ($item.Matches.Value -like "*authentication*") {
                $data = $item.Line -split ' '
                $data = $data[0] + ' ' + $data[1]

                # Safely convert to DateTime
                try { 
                    $data = [datetime]$data
                } catch {
                    $data = '--'
                }

                $data = $data.ToString("MM/dd/yyyy HH:mm:ss")
            }

            # Check for account logout event
            if ($item.Matches.Value -like "*terminated*") {
                $logoutData = $item.Line -split ' '
                $logoutData = $logoutData[0] + ' ' + $logoutData[1]

                # Safely convert to DateTime
                try { 
                    $logoutData = [datetime]$logoutData
                } catch {
                    $logoutData = '--'
                }

                $logoutData = $logoutData.ToString("MM/dd/yyyy HH:mm:ss")
            }

            # Create and add the custom object to the result array
            $obj += [pscustomobject]@{
                AccountLogon  = $data
                AccountLogout = $logoutData
            }
        }
    }

    # Return the results
    return $obj
}

function Get-TVLogFile_IPs {
<#
.SYNOPSIS
    Parses the Teamviewer15_logfile.log and Teamviewer15_logfile_OLD.log for a list of IPs used for incoming connections

.PARAMETER directory
    Used to specify the directory containing the log files

.EXAMPLE
    Get-TVLogFile_IPs -directory "C:\Program Files (x86)\TeamViewer"

    Returns the IPs used during the incoming connections
#>
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Get all matching log files
    $logs = Get-ChildItem -Path ($directory + "\teamviewer15_Logfile*.log") -ErrorAction SilentlyContinue
    if (-not $logs) {
        Write-Warning "No log files found in: $directory"
        return
    }

    $pattern = 'punch received a=(\d+\.\d+\.\d+\.\d+):(\d+):'
    $obj = @()

    foreach ($logfile in $logs.FullName) {
        $matches = Get-Content $logfile | Select-String -Pattern $pattern

        foreach ($line in $matches) {
            try {
                # Extract timestamp (first 23 characters)
                $dateStr = $line.Line.Substring(0, 23)
                $date = [datetime]::ParseExact($dateStr, "yyyy/MM/dd HH:mm:ss.fff", $null)

                if ($line.Line -match $pattern) {
                    $ip = $matches[1]
                    $port = $matches[2]

                    $obj += [pscustomobject]@{
                        Date     = $date.ToString("MM/dd/yyyy HH:mm:ss")
                        IP       = $ip
                        SrcPort  = $port
                    }
                }
            } catch {
                Write-Warning "Error processing line: $($line.Line)"
            }
        }
    }

    return $obj
}




function Get-TVLogFile_PIDs {
    <#
    .SYNOPSIS
        Parses the Teamviewer15_logfile.log and Teamviewer15_logfile_OLD.log and returns the PID for the connection

    .PARAMETER directory
        Used to specify the directory containing the log files

    .EXAMPLE
        Get-TVLogFile_PIDs -directory "C:\Program Files (x86)\TeamViewer"

        Returns the PIDs associated with the incoming connection
    #>
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Check if directory exists
    if (-not (Test-Path $directory)) {
        Write-Error "Directory does not exist: $directory"
        return
    }

    # Get all relevant log files in the directory
    $logs = Get-ChildItem (Join-Path $directory "teamviewer15_Logfile*.log")

    $obj = foreach ($log in $logs.FullName) {
        $temp = Get-Content $log | Select-String "Start Desktop process"
        foreach ($line in $temp) {
            $split = $line -Split ' '
            try {
                $data = $split[0] + ' ' + $split[1]
                $data = [datetime]$data
            } catch {
                Write-Warning "Date parsing failed for line: $line"
                continue
            }

            if ($split.Count -gt 1) {
                $processID = $split[-1]  # Renamed from PID to processID to avoid conflict
            } else {
                Write-Warning "PID not found in line: $line"
                continue
            }

            [pscustomobject]@{
                Date = $data.ToString("MM/dd/yyyy HH:mm")
                PID  = $processID   # Updated the reference to processID
            }
        }
    }

    # Return the results
    Write-Output $obj
}



Function Get-TVLogFile_Outgoing {
    <#
    .SYNOPSIS
        Parses the TeamViewer log files and returns outgoing connections (successes and failures)

    .PARAMETER directory
        The directory containing the log files

    .EXAMPLE
        Get-TVLogFile_Outgoing -directory "C:\Program Files (x86)\TeamViewer"
        Returns the outgoing connections (successes and failures)
    #>
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Get all relevant log files in the directory
    $logs = Get-ChildItem -Path (Join-Path $directory "teamviewer15_Logfile*.log")
    
    # Initialize an empty array for storing the results
    $obj = @()

    # Iterate through each log file
    foreach ($log in $logs.FullName) {
        # Read the log file
        $logContent = Get-Content $log

        # Select relevant lines from the log content
        $temp = $logContent | Select-String -Pattern "trying connection to", "LoginOutgoing: ConnectFinished - error: KeepAliveLost"

        foreach ($line in $temp) {
            # Initialize variables
            $tvID = $null
            $success = $null
            $data = $null

            # Split the line by spaces
            $split = $line.Line -split ' '

            # Ensure the line contains sufficient parts
            if ($split.Length -ge 5) {
                # If the line matches the expected pattern for "mode = 1"
                if ($line.Line -like "*mode = 1*") {
                    $tvID = $split[-4].Trim(',')  # Extract TV ID
                    $index = $logContent.IndexOf($line.Line) + 1

                    # Check for "KeepAliveLost" in the next line, if available
                    if ($index -lt $logContent.Count -and $logContent[$index] -match "KeepAliveLost") {
                        $success = "No"
                    } else {
                        $success = "Yes"
                    }

                    # Extract date
                    $data = $split[0] + ' ' + $split[1]
                    try {
                        $data = [datetime]$data
                        $data = $data.ToString("MM/dd/yyyy HH:mm:ss")
                    } catch {
                        Write-Warning "Failed to parse date for line: $($line.Line)"
                        continue
                    }

                    # Add the result as a custom object
                    $obj += [pscustomobject]@{
                        Date       = $data
                        ID         = $tvID
                        Successful = $success
                    }
                }
            } else {
                Write-Warning "Line does not contain expected number of parts: $($line.Line)"
            }
        }
    }

    # Output results or message
    if ($obj.Count -eq 0) {
        Write-Output "No outgoing connections found."
    } else {
        return $obj
    }
}



Function Get-TVLogFile_KeyboardLayout {
    <#
    .SYNOPSIS
        Parses the Teamviewer15_logfile.log and Teamviewer15_logfile_OLD.log and returns the keyboard layout associated with the incoming connection

    .PARAMETER directory
        Used to specify the directory containing the log files

    .EXAMPLE
        Get-TVLogFile_KeyboardLayout -directory "C:\Program Files (x86)\TeamViewer"

        Returns the keyboard layout associated with the incoming connection
    #>    
    [CmdletBinding()]
    param(
        [string]$directory
    )

    # Get all relevant log files in the specified directory
    $logs = Get-ChildItem -Path ($directory + "\teamviewer15_Logfile*.log")

    # Initialize an empty array for storing the results
    $obj = @()

    # Iterate through each log file
    foreach ($log in $logs.FullName) {
        # Search for lines related to changing keyboard layout
        $temp = Get-Content $log | Select-String -Pattern "changing keyboard layout to"

        # Process each matched line
        foreach ($line in $temp) {
            $keyboard = $date = '--'

            try {
                # Split the line into parts
                $split = $line -split ' '

                # Parse the date and convert it to DateTime
                $data = $split[0] + ' ' + $split[1]
                $date = ([datetime]$data).ToString("MM/dd/yyyy HH:mm")

                # Extract the keyboard layout from the last part of the line
                $keyboard = $split[-1]
            } catch {
                # In case of any error, set default values for date and keyboard
                $date = '--'
                $keyboard = '--'
            }

            # Create and add the custom object to the result array
            $obj += [pscustomobject]@{
                Date     = $date
                Keyboard = $keyboard
            }
        }
    }

    # Return the result array containing the custom objects
    return $obj
}
function Get-TVConnectionsLog_byDate {
    <#
    .SYNOPSIS
        Parses the connections.txt and returns data before, after, or between a specific date

    .DESCRIPTION
        Location to the connections.txt file

    .PARAMETER BeforeDate
        Returns data before the specified date

    .PARAMETER AfterDate
        Returns data after the specified date

    .EXAMPLE
        Get-TVConnectionsLog_byDate -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" -before "12/25/2020"
       
        Returns data before December 25, 2020

    .EXAMPLE
        Get-TVConnectionsLog_byDate -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" -after "12/25/2020"
        
        Returns data after December 25, 2020

    .EXAMPLE
        Get-TVConnectionsLog_byDate -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" -before "3/1/2021" -after "12/25/2020"
        
        Returns data after March 1, 2021 before December 25, 2020
    #>
    [CmdletBinding()]
    param(
        [string]$File,
        [datetime]$BeforeDate,
        [datetime]$AfterDate
    )

    # Read the content from the specified file
    $logs = Get-Content $File

    # Initialize an empty array for the results
    $obj = @()

    # Iterate through each line in the log file
    foreach ($log in $logs) {
        $dur = ''
        $log = $log -split '\s+'

        # Extract start and end dates from the log
        $dataStart = $log[1] + ' ' + $log[2]
        $dataEnd = $log[3] + ' ' + $log[4]
        $startDate = $null
        $endDate = $null

        # Attempt to parse the start and end dates, handling errors gracefully
        try {
            $startDate = [datetime]::ParseExact($dataStart, 'dd-MM-yyyy HH:mm:ss', $null)
        } catch {
            $startDate = $null
        }

        try {
            $endDate = [datetime]::ParseExact($dataEnd, 'dd-MM-yyyy HH:mm:ss', $null)
        } catch {
            $endDate = $null
        }

        # Calculate duration if both dates were successfully parsed
        if ($startDate -and $endDate) {
            $dur = New-TimeSpan -Start $startDate -End $endDate -ErrorAction SilentlyContinue
        }

        # Create an object with the parsed data
        $obj += [PSCustomObject]@{
            IncomingID     = $log[0]
            StartDate      = $startDate
            EndDate        = $endDate
            Duration       = if ($dur) { $dur.ToString("dd'd.'hh'h:'mm'm:'ss's'") } else { '--' }
            LoggedOnUser   = $log[5]
            ConnectionType = $log[6]
            ConnectionID   = $log[7]
        }
    }

    # Filter results based on date conditions
    if ($AfterDate -and $BeforeDate) {
        $obj | Where-Object { $_.StartDate -gt $AfterDate -and $_.StartDate -lt $BeforeDate }
    }
    elseif ($AfterDate) {
        $obj | Where-Object { $_.StartDate -gt $AfterDate }
    }
    elseif ($BeforeDate) {
        $obj | Where-Object { $_.StartDate -lt $BeforeDate }
    }
}

function Get-TVConnectionsLog_Top10Duration {
    <#
    .SYNOPSIS
        Parses the connections.txt file and returns the duration for the top 10 longest or shortest outgoing connections

    .PARAMETER File
        Location to the connections.txt file

    .PARAMETER shortest
        Used to select the shortest connections

    .PARAMETER longest
        Used to select the longest connections

    .EXAMPLE
        Get-TVConnectionsLog_Top10Duration -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" 

        Returns the duration for all outgoing connections

    .EXAMPLE
        Get-TVConnectionsLog_Top10Duration -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" -shortest

        Returns the top 10 shortest durations for all outgoing connections

    .EXAMPLE
        Get-TVConnectionsLog_Top10Duration -file "C:\Users\<user>\AppData\Roaming\TeamViewer\connections.txt" -longest

        Returns the top 10 longest durations for all outgoing connections
    #>
    [CmdletBinding()]
    param(
        [string]$File,
        [switch]$shortest,
        [switch]$longest
    )

    # Read the content from the specified file
    $logs = Get-Content $File

    # Initialize an empty array for the results
    $obj = @()

    # Iterate through each line in the log file
    foreach ($log in $logs) {
        $dur = ''
        $log = $log -split '\s+'

        # Extract start and end dates from the log
        $dataStart = $log[1] + ' ' + $log[2]
        $dataEnd = $log[3] + ' ' + $log[4]
        $startDate = $null
        $endDate = $null

        # Attempt to parse the start and end dates, handling errors gracefully
        try {
            $startDate = [datetime]::ParseExact($dataStart, 'dd-MM-yyyy HH:mm:ss', $null)
        } catch {
            $startDate = $null
        }

        try {
            $endDate = [datetime]::ParseExact($dataEnd, 'dd-MM-yyyy HH:mm:ss', $null)
        } catch {
            $endDate = $null
        }

        # Calculate duration if both dates are successfully parsed
        if ($startDate -and $endDate) {
            $dur = New-TimeSpan -Start $startDate -End $endDate -ErrorAction SilentlyContinue
        }

        # Create an object with the parsed data
        $obj += [PSCustomObject]@{
            IncomingID     = $log[0]
            StartDate      = $startDate
            EndDate        = $endDate
            Duration       = if ($dur) { $dur.ToString("dd'd.'hh'h:'mm'm:'ss's'") } else { '--' }
            LoggedOnUser   = $log[5]
            ConnectionType = $log[6]
            ConnectionID   = $log[7]
        }
    }

    # Sort and display the top 10 shortest or longest durations
    if ($shortest) {
        $obj | Sort-Object Duration | Select-Object -First 10 | Format-Table
    }

    if ($longest) {
        $obj | Sort-Object Duration -Descending | Select-Object -First 10 | Format-Table
    }
}
