$testrun = $false

#### Script from here, don't change unless you know what you are doing ####

$defaultdownloads = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$workingdirectory = ".\Data\working directory"
$budgetlocation = ".\Data\Source\budget.csv"
$category_filter = ".\Data\Configuration\Categories.xlsx"

#List of bank accounts
$bankaccounts = import-csv ".\Data\Configuration\bank_statement_naming.csv" -Encoding unicode -Delimiter ';'
$exportoutgoing = @()

set-location $PSScriptRoot

## Import current budget
try {
    $budget = import-csv -path ".\Data\Source\budget.csv" -Delimiter ';' -ErrorAction SilentlyContinue
    
    ## Get newest date in document
    $budget = $budget | sort-object {[System.DateOnly]$_.Bogføringsdato} -Descending | Select-Object -first 1
}catch {}

if($null -eq $budget)
{
    $p = @{
        Bogføringsdato = get-date -Format "1990-01-01" 
        }
        $budget = New-Object -TypeName psobject -Property $p
}

#Cleanup working directory for old files
$cleanup = Get-ChildItem $workingdirectory -Recurse -ErrorAction SilentlyContinue
foreach($i in $cleanup)
{
    remove-item -Path $i -ErrorAction SilentlyContinue -Recurse -Confirm:$false -force
}

# Find all files in download folder that matches account statement
$folders = Get-ChildItem -Path $defaultdownloads
foreach($file in $folders)
{
    foreach($filename in $bankaccounts)
    {
        if($file.name -like $filename.name)
        {
            $copyfrom = $file.versioninfo.filename
            $copyto = $workingdirectory + "\" + $file.Name
            Copy-Item -Path $copyfrom -Destination $copyto
        }
    }
}

### Find new entries in the statement
$oldstatements = @()
$bankstatements = Get-ChildItem $workingdirectory 
$accountoverview = @()
$today = Get-Date
foreach($account in $bankstatements)
{
    $importedaccountstatement = import-csv $account.VersionInfo.filename -Delimiter ';'
    foreach($statement in $importedaccountstatement)
    {
        $lastupdate = ($budget.Bogføringsdato | get-date -f "yyyy-M-dd").ToString() 
        $bogføringsdato = ($statement.Bogføringsdato | get-date -f "yyyy-M-dd").ToString()
        $currentdate = (get-date).AddDays(-1).ToString("yyyy-M-dd")
        $differenceOlderthan2days = New-TimeSpan -start $bogføringsdato -End $currentdate
        $difference = New-TimeSpan -start $lastupdate -End $bogføringsdato
    
        if ($differenceOlderthan2days.Days -ge "0" -and $difference.Days -ge "1") {
            $statement | Add-Member -MemberType NoteProperty -Name "Account" -Value $account.name
            $accountoverview += $statement
        }
        else {
            $statement | Add-Member -MemberType NoteProperty -Name "Account" -Value $account.name
            $oldstatements += $statement
        }
    }
}

$categoriesdefinition = import-csv -path ".\Data\Configuration\category_filter.csv" -Encoding unicode -Delimiter ';'
foreach($entry in  $accountoverview)
{
    $månedbyname = get-date -Date $entry.Bogføringsdato -UFormat %B
    $yearbyname = get-date -Date $entry.Bogføringsdato -UFormat %Y

    if([int]$i.Beløb -le 0 )
    {$type = "udgående"}
    else 
    {$type = "Indgående"}
    $secondarycategory = "_unknown"
    $maincategory = "_unknown"
    $indexnumber = ""

    foreach($category in $categoriesdefinition)
    {
        if($entry.Beskrivelse -like $category.beskrivelse)
        {
            $secondarycategory = $category.kategori
            $maincategory = $category.maincategori
            $indexnumber = $category.indexnumber
        }
    }
    $p = [ordered]@{
        Maincategori            = $maincategory
        Kategori                = $secondarycategory
        Bogføringsdato          = $entry.Bogføringsdato
        Beløb                   = $entry.Beløb
        Beskrivelse             = $entry.Beskrivelse
        Måned                   = $månedbyname
        Year                    = $yearbyname
        Valuta                  = $entry.Valuta
        Type                    = $type
        Afsender                = $entry.Afsender
        Navn                    = $entry.Navn
        Modtager                = $entry.Modtager
        indexnumer              = $indexnumber
        }
        
        $objcmddata = New-Object -TypeName psobject -Property $p
        $exportoutgoing += $objcmddata
}

foreach($i in $exportoutgoing)
{
    $whatdayisit = $i.Bogføringsdato | get-date 

        $text = "dato: " + $i.Bogføringsdato + `
        "`nugeday: " + $whatdayisit.DayOfWeek + `
        "`nBeskrivelse: " + $i.Beskrivelse + `
        "`nBeløb: " + $i.Beløb + `
        "`nmaincategori: "+ $i.maincategori + `
        "`nkategori: "+ $i.kategori + `
        "`n________________________"

    if($whatdayisit.DayOfWeek -eq "saturday" -or $whatdayisit.DayOfWeek -eq "sunday")
    {
        Write-host $text -ForegroundColor red
     }else
    {
        Write-host $text -ForegroundColor Green
    }
}

Write-Output "Er resultatet ok? (ja/nej)"
$reply = Read-Host 
if($reply -eq "ja")
{
    #############################################
    #### Take backup of previous source file ####
    #############################################
    $runningdateandtimeforbackup = get-date -f "yyyy-M-dd - HH-mm-ss"
    $backuplocation = ".\Data\Backup\budget - " + $runningdateandtimeforbackup + ".csv"
    Copy-Item -Path $budgetlocation -Destination $backuplocation
    #start-sleep -Seconds 10
    $doesbackupexist = get-item $backuplocation -ErrorAction SilentlyContinue

    if($doesbackupexist)
    {
        ### Playing around with values, wan't the _unknowns ones to be at top.
        $allunknowns = $exportoutgoing | where {$_.Maincategori -eq "_unknown"}
        $everythingelse = $exportoutgoing | where {$_.Maincategori -ne "_unknown"}

        ### Merging current budget with new changes and sorting by date
        $allcurrent = import-csv $budgetlocation -Delimiter ';' -Encoding unicode
        $everythingelse += $allcurrent
        $everythingelse = $everythingelse | Sort-Object {[System.DateOnly]$_.Bogføringsdato} -Descending

        $allunknowns += $everythingelse
        $allunknowns | Select-Object Maincategori,Kategori,Bogføringsdato,Beløb,Beskrivelse,Måned,Year,Valuta,Type,Afsender,Navn,Modtager | Export-Csv $budgetlocation -Delimiter ';' -Encoding unicode -NoTypeInformation 
        Write-host "Tilføjet til budget" -ForegroundColor Green

        
        $opensourcefile = read-host "Ønsker du at åbne og rette i budgettet? (ja/nej) - default ja"
        if($opensourcefile -eq "ja" -or $opensourcefile -eq "")
        {
            Start-Process $budgetlocation
            start-process $category_filter
        }
    }
    else
    {
        Write-host "Afbrudt, data er ikke gemt i budgettet. Backup af data lykkedes ikke " -ForegroundColor red
    }
}
else
{
    Write-host "Afbrudt, data er ikke gemt i budgettet. " -ForegroundColor red
}