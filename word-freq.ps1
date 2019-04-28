# PowerShell: Counting words in a text file
# Returns a list of the most-used, longest words in a text file.
#
######
#
# Post from From: https://gist.github.com/arthurattwell/f6552158f17db18ad48d286146f533c7
# Modified by Lawrence Billinghurst
# Modification Data: April 2019
# and
# Adapted slightly from 
# https://blogs.technet.microsoft.com/josebda/2015/03/21/powershell-examples-counting-words-in-a-text-file/
# Modified for the http:\\2019.report
###### 
# From: https://gist.github.com/arthurattwell/f6552158f17db18ad48d286146f533c7
######
# Used color script from https://copdips.com/2018/05/grep-like-powershell-colorful-select-string.html
# and https://github.com/copdips/PSScripts/blob/master/Text/Select-ColorString.ps1
######
# Perfomance Bost
# https://www.mssqltips.com/sqlservertip/3359/powershell-and-text-mining-part-i-word-counts-positions-and-libraries/
######
## Uses the text from Alice in Wonderland 
# from http://www.gutenberg.org/ebooks/11.txt.utf-8
param 
(
    [Parameter(Mandatory=$False)] [string]$FileName= '.\Alice''sAdventureInWonderLand.txt'
    ,[Parameter(Mandatory=$False)] [string]$Search
    ,[Parameter(Mandatory=$False)] [string]$SortOrder
    ,[Parameter(Mandatory=$False)] [Alias("R")] [switch]$DisplayResults
    ,[Parameter(Mandatory=$False)] [Alias("I")] [switch]$Interactive
    ,[Parameter(Mandatory=$False)] [Alias("N")] [switch]$NoProgress
    ,[Parameter(Mandatory=$False)] [Alias("D")] [switch]$DontShow
    ,[Parameter(Mandatory=$False)] [Alias("E")] [switch]$ExportCSV
    ,[Parameter(Mandatory=$False)] [switch]$HashSwitch
)

#Stopwatch Start
$sw = [Diagnostics.Stopwatch]::StartNew()

#Define some Values
$NumberFound = 0
$WordCount = 0
$Longest = ""
$Dictionary = @{}
$LineCount = 0
$dashline="--------------------------------"


if ((Test-Path -path $FileName)-eq $false) { 
        $xMenuChoiceA  = "0"
        while (($xMenuChoiceA -ne "1") -and ((Test-Path -path $FileName)-eq $False)){
            Write-host $dashline
            [string]$xMenuChoiceA = read-host "Enter a file name or 1 to exit" 
            if ( $xMenuChoiceA -eq "1"){
                    exit
               } else {$FileName=$xMenuChoiceA}
         }
}

#Define Funciton print-search results
function print-search{

    if (($SearchWord -ne "") -and (-not $DontShow)) {
        $NumberFound = (get-content  $FileName| select-string -pattern $SearchWord).length
         Write-host $dashline
        if ($YesColors -eq $True) {
            Write-output "The word $SearchWord was found $NumberFound times." |Select-ColorString $SearchWord
            Write-host $dashline
            Select-String -Pattern $SearchWord $FileName |Select-ColorString $SearchWord} 
                else {
                    Write-Host "The word $SearchWord was found $NumberFound times."
                    Write-host $dashline
                    Select-String -Pattern $SearchWord $FileName
                    }
    } 
}

#Try to load select-colorstring function file
$FunctionFile=".\Select-ColorString.ps1"

if (Test-Path -path $FunctionFile) {. $FunctionFile
    $YesColors=$True}
$SearchWord = $Search.ToUpper()
$FileContents = Get-Content $FileName
$TotalLines = $FileContents.Count

if (-not $DontShow) {
    Write-Host "Reading file $FileName..." 
    Write-host $dashline
    Write-Host "$TotalLines lines read from the file."
    Write-host $dashline}
        

$FileContents | foreach {
    $Line = $_
    $LineCount++
    if (-not $NoProgress) {
        Write-Progress -Activity "Indexing Line ($LineCount of $TotalLines)..." -PercentComplete ($LineCount*100/$TotalLines) 
    }
    $Line.Split(" .,:;?!/()[]{}-```"") | foreach {
        $Word = $_.ToUpper()
        If ($Word[0] -ge 'A' -and $Word[0] -le "Z") {
            $WordCount++
            If ($Word.Contains($SearchWord)) { $Found++ }
            If ($Dictionary.ContainsKey($Word)) {
                if ($HashSwitch){
                    $Dictionary.$Word++   # Slow Method
                    } else {
                            $cnt = $Dictionary[$Word] + 1   
                            $Dictionary.Remove($Word)
                            $Dictionary.Add($Word, $cnt)
                            }
            } else {
                $Dictionary.Add($Word, 1)
            }
        }
    } 
}

if (-not $NoProgress) {
    Write-Progress -Activity "Indexing words..." -Completed
}

# Filter Word List to remove any single values and length greter than 2, then sort by Word name
$WordCountList=$($Dictionary.GetEnumerator()| ? {($_.Value -gt 1) -AND ($_.Name.Length -gt 2)} | Sort Name )
$DictWords = $WordCountList.Count
#$OutList = ($WordCountList | Select Name,Value -First 2)

if (-not $DontShow) {
    Write-Host "$WordCount total words in the text"
    Write-Host "$DictWords distinct words in the text"
}

#Call the print resule function
print-search

#Stopwatch Stop
$sw.Stop()
    switch ($SortOrder) {
       "1" {$WordCountList=$WordCountList|Sort Name ; break}
       "2" {$WordCountList=$WordCountList|Sort Name -Descending ; break}
       "3" {$WordCountList=$WordCountList|Sort Value ; break}
       "4" {$WordCountList=$WordCountList|Sort Value -Descending ; break}
       default {$WordCountList=$WordCountList|Sort Name ; break}
       }

if ($DisplayResults) { #Show Results when R Switch is activeated 
    $WordCountList|Select -First 5 |ft
}

if ($ExportCSV) { #Export file when E Switch is activeated 
    $WordCountFileName=(split-path -path $FileName) +"\"+ ((gci $FileName).BaseName) + "-wordcount.csv"
    $WordCountList|select name,value|Export-Csv $WordCountFileName -NoTypeInformation
}

if ($Interactive){
        $xMenuChoiceA  = "0"
        while ( $xMenuChoiceA -ne "1"){
            Write-host $dashline
      
            [string]$xMenuChoiceA = read-host "Enter word to search or 1 to exit" 

  
            if ( $xMenuChoiceA -ne "1"){
                      Write-host $dashline
                      $SearchWord = $xMenuChoiceA
                      #Call the print resule function 
                      print-search
                    }
               }
                
}else{

$LogFileName=(split-path -path $MyInvocation.MyCommand.Name) +".\"+ ((gci $MyInvocation.MyCommand.Name).BaseName) + ".log.txt"

        #Write timing to Log file, skip if Interactive
        $LogEntery= $(Get-Date -Format 'MM/dd/yyyy,hh:mm tt')+","+$HashSwitch+","+$sw.Elapsed.TotalSeconds+","+$TotalLines+","+$WordCount+","+$DictWords+","+$FileName 
        #Write-Host "Writing Log: "$LogEntery
        Add-Content $LogFileName $LogEntery
     }
