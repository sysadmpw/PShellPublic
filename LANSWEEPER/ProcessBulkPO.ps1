<#
.SYNOPSIS
    LANSWEEPER bulk field update tool.
.DESCRIPTION
    ProcessBulkPO is a bulk update tool to insert PO numbers for assets directly into the LANSWEEPER database instead of manually editing and entering the data individually. This is the none GUI version, 
    simply place a well formed CSV file using the template in the LANSWEEPER / processing share and let the scheduled task take care of it for you. Once processed, a log file is created.
.EXAMPLE
    1. Create a CSV file that is MS-DOS text based in the format of:
    hostname (which is the serial number),POnumber
    2. Wait for the scheduled task to start and process your CSV file
    3. Log file is created  for you to check in the same processing folder under Logs.
    
    PO's are added against Assets based on the primary key AssetID
.INPUTS
    N/A
.OUTPUTS
    On completing the processing, write log file to Log folder inside the Processing folder.
.NOTES
    N/A
#>


Clear-Host

#Global Variables

#Key Value Pairs from from CSV File
$global:assetHashTbl = @{ }
#Key Value Pairs prepared for SQL update
$global:sqlUpdateTbl = @{ }

# This is a simple user/pass connection string.
# Feel free to substitute "Integrated Security=True" for system logins.
$global:sqlConnStr = "Data Source=LOCALHOST;Database=lansweeperdb;User ID=XX;Password=XXXX"
$global:conn = $null
#CSV Target file to process
$global:CSVFile = $null
$global:LogFileLocation = $null



Function Open-DatabaseConnection {

    $connState = $null
    try {
        #For each Asset, get the AssetID from SQL and inject into new hash table
        $conn = New-Object System.Data.SqlClient.SqlConnection $sqlConnStr
        #Attempt to open the connection
        $conn.Open()
        if ($conn.State -eq "Open") {
            $connState = $true
            $currentTime = Get-LogTime
            Write-host "`n[$currentTime] Connected to DB." -ForegroundColor Green
            return $connState
        }
        #$conn.Close()
        # We could not connect here
        # Notify connection was not in the "open" state
        $connState = $false
        return $connState
    }
    catch {
        # We could not connect here
        # Notify there was an error connecting to the database
        $currentTime = Get-LogTime
        Write-host "`n[$currentTime] Not connected to DB." -ForegroundColor Green
        $connState = $false
        return $connState
    }
}

Function Get-LogTime {
    $currentTime = Get-Date -Format " ddMMyyyy-HHmm"
    return $currentTime
}

Function Get-CSVFile {
    $fileExists = Test-Path -Path $CSVFile

    if($fileExists){
        $currentTime = Get-LogTime
        Write-host "`n[$currentTime] Reading CSV File." -ForegroundColor Yellow
        $data = Import-Csv -Path $CSVFile -Delimiter ',' -Header "Asset", "PO"
        foreach ($line in $data) {
            $tmpAsset = $line.Asset
            $tmpPO = $line.PO
            $assetHashTbl[$tmpAsset] = $tmpPO
            $currentTime = Get-LogTime
            Write-host "[$currentTime] Read: $tmpAsset, $tmpPO."
        }
    }else{
        $currentTime = Get-LogTime
        Write-host "`n[$currentTime] Unable to find $CSVFile CSV File." -ForegroundColor Red
    }
}


Function Get-AssetNames {
    try {
        #For each Asset, get the AssetID from SQL and inject into new hash table
        $conn = New-Object System.Data.SqlClient.SqlConnection $sqlConnStr

        #Attempt to open the connection
        $conn.Open()
        if ($conn.State -eq "Open") {
            # We have a successful connection here
            # Notify of successful connection
            $currentTime = Get-LogTime
            Write-Host "`n[$currentTime] Retrieving matching AssetNames." -ForegroundColor Yellow
            #Loop through hashtable, grab Asset name, query based on Asset name - return list of Assets
            foreach ($key in $assetHashTbl.GetEnumerator()) {
                $tmpAssetName = $key.Name
                $queryAssetID = "SELECT AssetName, count(*) as NUM from tblAssets GROUP BY AssetName HAVING AssetName = '$tmpAssetName'"
                $returnAssetName = Invoke-Sqlcmd -ServerInstance LOCALHOST -Database lansweeperdb -Query $queryAssetID -ErrorAction Stop

                $duplicate = $returnAssetName.NUM
                if($duplicate -gt 1){
                    $currentTime = Get-LogTime
                    Write-Host "[$currentTime] Asset $tmpAssetName appears twice in LANSWEEPER, this will be ignored. Please fix manually." -ForegroundColor Yellow
                }else{
                    $currentTime = Get-LogTime
                    Write-Host "[$currentTime] Asset $tmpAssetName PO set to be applied." -ForegroundColor Green
                    #Add Asset ID and Asset Name pairs to new hash table, make sure your update query is based on ASSETID as thats the primary KEY
                    $queryAssetID = "SELECT AssetID, AssetName from dbo.tblAssets WHERE dbo.tblAssets.AssetName='$tmpAssetName'"
                    $returnAssetName = Invoke-Sqlcmd -ServerInstance LOCALHOST -Database lansweeperdb -Query $queryAssetID -ErrorAction Stop
                    $tmpAssetID = $returnAssetName.AssetID
                    $tmpAssetName = $returnAssetName.AssetName
                    $sqlUpdateTbl[$tmpAssetID] = $tmpAssetName
                }
            }
            $conn.Close()
        }
        # We could not connect here
        # Notify connection was not in the "open" state
    }
    catch {
        # We could not connect here
        # Notify there was an error connecting to the database
        $currentTime = Get-LogTime
        Write-Host "`n[$currentTime] Unable to connect to database." -ForegroundColor Red

    }
}


Function Set-AssetPO {
    #Make sure your UPDATE Query is on ASSETID, as there are duplicate hostnames
    foreach ($key in $sqlUpdateTbl.GetEnumerator()) {
        foreach ($keyTmp in $assetHashTbl.GetEnumerator()) {
            $sqlHostName = $key.Value
            $csvHostName = $keyTmp.Name
            
            if ($sqlHostName -eq $csvHostName) {

                $tmpAssetID = $key.Name
                $tmpPOId = $keyTmp.Value


                #open DB connection and update based on AssetID
                try {
                    #For each Asset, get the AssetID from SQL and inject into new hash table
                    $conn = New-Object System.Data.SqlClient.SqlConnection $sqlConnStr
            
                    #Attempt to open the connection
                    $conn.Open()
                    if ($conn.State -eq "Open") {
                        # We have a successful connection here
                        # Notify of successful connection
                        $querySetPO = "Update dbo.tblAssetCustom SET dbo.tblAssetCustom.Custom1='$tmpPOId' where dbo.tblAssetCustom.AssetId=$tmpAssetID"
                        $returnValue = Invoke-Sqlcmd -ServerInstance LOCALHOST -Database lansweeperdb -Query $querySetPO -ErrorAction Stop
                        $conn.Close()
                    }
                    # We could not connect here
                    # Notify connection was not in the "open" state
            
                }
                catch {
                    # We could not connect here
                    # Notify there was an error connecting to the database
                    $currentTime = Get-LogTime
                    Write-Host "`n[$currentTime] Unable to connect to database." -ForegroundColor Red
                }
            }
        }
    }
}

Function Set-Log {

    $currentTime = Get-LogTime    
    $LogFileName = $currentTime+".log"
    Write-Host "`n[$currentTime] New log file $LogFileName created in $LogFileLocation." -ForegroundColor Green

}

Function Exit-App {
    Exit-PSSession
}


#Create new LogFile
Set-Log
#Open connection to the DB
Open-DatabaseConnection
#Read the static CSV file from static location
Get-CSVFile
#Builds a clean table for the db update function
Get-AssetNames
#Update database
Set-AssetPO
