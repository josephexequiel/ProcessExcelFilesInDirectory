$dateUnix = (Get-Date).ToString('MMddyyyy_hhmmss')
$sourceDirectory = "C:\Temp\DataSource\"
$outputSummaryFile = "C:\Temp\SummaryFile_$dateUnix.xlsx"

if (Test-Path $outputSummaryFile)
{
    try
    {
        Remove-Item $outputSummaryFile -ErrorAction Stop
        Write-Host "[INFO] Removed existing output file: $outputSummaryFile"
    }
    catch
    {
        Write-Host "[ERROR] Cannot remove existing output file: $outputSummaryFile"
        Read-Host -Prompt "[INFO] Press Enter to exit"
        exit -1
    }
}

class SummaryDataList {
    [datetime] $RecordDate
    [int] $TotalActive
    [int] $TotalInactive
    [int] $TotalUncategorized
    [int] $TotalOnboardingCustomers
    [int] $TotalSleepingCustomers
    [int] $TotalNonsleepingCustomers
    [int] $TotalUnfundedCustomers
    [int] $TotalOtherCustomers
}

class FinalRecordList {
    [string] $PartyID
    [string] $AccountType
    [string] $AccountClass
    [double] $BookedBalance
    [datetime] $DateOpened
    [datetime] $LastCreditTxnDate
    [datetime] $LastDebitTxnDate
    [datetime] $RecordDate
    [int] $InactiveDays
    [int] $SleepingDays
    [int] $AccountAge
}

try 
{
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $False
    $objExcel.DisplayAlerts = $False;

    # Initialize Data Array Instance from Class SummaryDataList
    $finalDataArr = [System.Collections.Generic.List[SummaryDataList]]::new()
    $finalDataArrCount = 0

    # Initialize Data Array Instance from Class FinalRecordList
    $finalRecordArr = [System.Collections.Generic.List[FinalRecordList]]::new()
    $finalRecordArrCount = 0

    Get-ChildItem $sourceDirectory\*.xlsx | ForEach-Object {
        $fileName = $_.Name
        $baseFileName = $_.BaseName
        $csvFileName = $sourceDirectory + $baseFileName + ".csv"
        Write-Output ("[INFO] File found: " + $fileName)
        $workBook = $objExcel.Workbooks.Open($_.FullName)
        $workSheet = $workBook.Sheets.Item("Sheet1")
        $workSheet.SaveAs($csvFileName, 6)
        $workBook.Close($False)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null

        # Import data to array
        $csvObject = Import-Csv $csvFileName
        $totalNoOfRecords = $csvObject.Count
        Remove-Item $csvFileName

        # Define Counter Containers
        $totalActive = 0
        $totalInactive = 0
        $totalUncategorized = 0
        $totalOthers = 0
        $totalOnboarding = 0
        $totalSleeping = 0
        $totalNonSleeping = 0
        $totalUnfunded = 0

        for ($num = 0; $num -lt $totalNoOfRecords; $num++) {
            $currentProgressPerct = [math]::Round((($num/$totalNoOfRecords)*100),2)
            Write-Progress -Activity "Processing file: $fileName" -Status "Progress: $currentProgressPerct% ($num/$totalNoOfRecords)" -PercentComplete $currentProgressPerct

            # Get Cell Value
            $getPartyId = $csvObject[$num]."Party ID"
            $getBookedBalance = $csvObject[$num]."Sum of OFBOOKEDBALANCE"
            $getDateOpened = [datetime]$csvObject[$num]."OFDATEOPENED"
            $getLastCreditTxnDate = [datetime]$csvObject[$num]."OFLASTCREDITTXNDATE"
            $getLastDebitTxnDate = [datetime]$csvObject[$num]."OFLASTDEBITTXNDATE"
            $getCurrentDate = [datetime]$csvObject[$num]."Currenct Date"
            $staticCurrentDate = $getCurrentDate.ToShortDateString()

            # Get Date Difference Between Transactions and Accounts
            $inactiveCountDebit = (New-TimeSpan -Start $getLastDebitTxnDate -End $getDateOpened).Days
            $sleepingCustomerDebit = (New-TimeSpan -Start $getLastDebitTxnDate -End $getCurrentDate).Days
            $inactiveCountCredit = (New-TimeSpan -Start $getLastCreditTxnDate -End $getDateOpened).Days
            $sleepingCustomerCredit = (New-TimeSpan -Start $getLastCreditTxnDate -End $getCurrentDate).Days
            $getAccountAge = (New-TimeSpan -Start $getDateOpened -End $getCurrentDate).Days

            # Check inactiveCount (Days of inactivity since account was opened) - Lower value means client is active
            # Check which one has the latest transaction between Debit and Credit
            # Select the oldest date (highest value) between transactions
            if ($inactiveCountDebit -gt $inactiveCountCredit)
            {
                $inactiveCount = $inactiveCountCredit
            }
            elseif ($inactiveCountDebit -lt $inactiveCountCredit)
            {
                $inactiveCount = $inactiveCountDebit
            }
            elseif ($inactiveCountDebit -eq $inactiveCountCredit)
            {
                $inactiveCount = $inactiveCountDebit
            }
    
            # Check sleepingCustomerCount - Higher value means more idle time
            # Check which one has the latest transaction between Debit and Credit
            # Select the newest date (lowest value) between transactions
            if ($sleepingCustomerDebit -lt $sleepingCustomerCredit)
            {
                $sleepingCustomerCount = $sleepingCustomerDebit
            }
            elseif ($sleepingCustomerDebit -gt $sleepingCustomerCredit)
            {
                $sleepingCustomerCount = $sleepingCustomerCredit
            }
            elseif ($sleepingCustomerDebit -eq $sleepingCustomerCredit)
            {
                $sleepingCustomerCount = $sleepingCustomerDebit
            }
            
            if (($getAccountAge -gt 1) -And ($getBookedBalance -ge 1) -And ($inactiveCount -lt 0) -And ($sleepingCustomerCount -le 90))
            {
                $showAccountType = "Active"
                $showAccountClass = "Non-Sleeping Customer"
                $totalActive++
                $totalNonSleeping++
            }
            elseif (($getAccountAge -gt 1) -And ($getBookedBalance -lt 1) -And ($inactiveCount -ge 0))
            {
                $showAccountType = "Inactive"
                $showAccountClass = "Unfunded Customer"
                $totalInactive++
                $totalUnfunded++
            }
            elseif (($getAccountAge -gt 1) -And ($getBookedBalance -lt 1) -And ($inactiveCount -lt 0) -And ($sleepingCustomerCount -gt 90))
            {
                $showAccountType = "Inactive"
                $showAccountClass = "Sleeping Customer"
                $totalInactive++
                $totalSleeping++
            }
            elseif (($getAccountAge -le 1) -And ($inactiveCount -ge 0))
            {
                $showAccountClass = "Onboarding Customer"
                $totalOnboarding++
                if ($getBookedBalance -ge 1)
                {
                    $showAccountType = "Active"
                    $totalActive++
                }
                elseif ($getBookedBalance -lt 1)
                {
                    $showAccountType = "Inactive"
                    $totalInactive++
                }
                else
                {
                    $totalUncategorized++
                }
            }
            else
            {
                $showAccountClass = "Other Customer"
                $totalOthers++
                if ($getBookedBalance -ge 1)
                {
                    $showAccountType = "Active"
                    $totalActive++
                }
                elseif ($getBookedBalance -lt 1)
                {
                    $showAccountType = "Inactive"
                    $totalInactive++
                }
                else
                {
                    $totalUncategorized++
                }
            }

            $finalRecordArr.Add([FinalRecordList]::new())
            $finalRecordArr[$finalRecordArrCount].PartyID = $getPartyId
            $finalRecordArr[$finalRecordArrCount].AccountType = $showAccountType
            $finalRecordArr[$finalRecordArrCount].AccountClass = $showAccountClass
            $finalRecordArr[$finalRecordArrCount].BookedBalance = $getBookedBalance
            $finalRecordArr[$finalRecordArrCount].DateOpened = $getDateOpened
            $finalRecordArr[$finalRecordArrCount].LastCreditTxnDate = $getLastCreditTxnDate
            $finalRecordArr[$finalRecordArrCount].LastDebitTxnDate = $getLastDebitTxnDate
            $finalRecordArr[$finalRecordArrCount].RecordDate = $getCurrentDate
            $finalRecordArr[$finalRecordArrCount].InactiveDays = $inactiveCount
            $finalRecordArr[$finalRecordArrCount].SleepingDays = $sleepingCustomerCount
            $finalRecordArr[$finalRecordArrCount].AccountAge = $getAccountAge
            $finalRecordArrCount++
        }

        Write-Progress -Activity "Processing file: $fileName" -Status "Progress: $currentProgressPerct% ($num/$totalNoOfRecords)" -Completed

        $finalDataArr.Add([SummaryDataList]::new())
        $finalDataArr[$finalDataArrCount].RecordDate = $staticCurrentDate
        $finalDataArr[$finalDataArrCount].TotalActive = $totalActive
        $finalDataArr[$finalDataArrCount].TotalInactive = $totalInactive
        $finalDataArr[$finalDataArrCount].TotalUncategorized = $totalUncategorized
        $finalDataArr[$finalDataArrCount].TotalOnboardingCustomers = $totalOnboarding
        $finalDataArr[$finalDataArrCount].TotalSleepingCustomers = $totalSleeping
        $finalDataArr[$finalDataArrCount].TotalNonsleepingCustomers = $totalNonSleeping
        $finalDataArr[$finalDataArrCount].TotalUnfundedCustomers = $totalUnfunded
        $finalDataArr[$finalDataArrCount].TotalOtherCustomers = $totalOthers
        $finalDataArrCount++
    }

    $finalDataArr | Export-Excel -WorkSheetName "Summary" -Path $outputSummaryFile
    $finalRecordArr | Export-Excel -WorkSheetName "RecordList" -Path $outputSummaryFile
    Write-Output ("[INFO] Records populated")

    $workBook = $objExcel.Workbooks.Open($outputSummaryFile)
    $workSheet = $workBook.Sheets.Item("Summary")
    $rowCount = $workSheet.UsedRange.rows.count
    $workSheet.Range("A2:A$rowCount").NumberFormat = 'MM/dd/yyyy'
    $workSheet.UsedRange.Columns.Autofit() | Out-Null

    $workSheet = $workBook.Sheets.Item("RecordList")
    $rowCount = $workSheet.UsedRange.rows.count
    $workSheet.Range("E2:E$rowCount").NumberFormat = 'MM/dd/yyyy'
    $workSheet.Range("F2:F$rowCount").NumberFormat = 'MM/dd/yyyy'
    $workSheet.Range("G2:G$rowCount").NumberFormat = 'MM/dd/yyyy'
    $workSheet.Range("H2:H$rowCount").NumberFormat = 'MM/dd/yyyy'
    $workSheet.UsedRange.Columns.Autofit() | Out-Null

    $workBook.SaveAs($outputSummaryFile)
    $workBook.Close($False)

    Write-Output ("[INFO] Workbook generated: $outputSummaryFile")
}
finally
{
    # Close excel object
    $objExcel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Read-Host -Prompt "[INFO] Processing Completed.`nPress Enter to continue"
