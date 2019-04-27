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
    [Parameter(Mandatory=$False)] [string]$FileName=".\Alice'sAdventureInWonderLand.txt"
    ,[Parameter(Mandatory=$False)] [string]$Search
    ,[Parameter(Mandatory=$False)] [string]$HashSwitch = 0
    ,[Parameter(Mandatory=$False)]  [Alias("I")] [switch]$Interactive
)

#Stopwatch Start
$sw = [Diagnostics.Stopwatch]::StartNew()
$NumberFound = 0
$WordCount = 0
$Longest = ""
$Dictionary = @{}
$LineCount = 0

#Define Funciton print-search results
function print-search{
    if ($SearchWord -ne "") {
        $NumberFound = (get-content  $FileName| select-string -pattern $SearchWord).length
        Write-host "--------------------------------"
        if ($YesColors -eq $True) {
            Write-output "The word $SearchWord was found $NumberFound times." |Select-ColorString $SearchWord
            Write-host "--------------------------------"
            Select-String -Pattern $SearchWord $FileName |Select-ColorString $SearchWord} 
                else {
                    Write-Host "The word $SearchWord was found $NumberFound times."
                    Write-host "--------------------------------" 
                    Select-String -Pattern $SearchWord $FileName
                    }
    } 
}
#Try to load select-colorstring function file
$FunctionFile=".\Select-ColorString.ps1"
if (Test-Path -path $FunctionFile) {. $FunctionFile
    $YesColors=$True}
$SearchWord = $Search.ToUpper()
#Clear-Host

Write-Host "Reading file $FileName..." 
Write-host "--------------------------------"
$FileContents = Get-Content $FileName
$TotalLines = $FileContents.Count
Write-Host "$TotalLines lines read from the file."
Write-host "--------------------------------"

$FileContents | foreach {
    $Line = $_
    $LineCount++
    Write-Progress -Activity "Indexing Line ($LineCount of $TotalLines)..." -PercentComplete ($LineCount*100/$TotalLines) 
    $Line.Split(" .,:;?!/()[]{}-```"") | foreach {
        $Word = $_.ToUpper()
        If ($Word[0] -ge 'A' -and $Word[0] -le "Z") {
            $WordCount++
            If ($Word.Contains($SearchWord)) { $Found++ }
            If ($Dictionary.ContainsKey($Word)) {
                if ($HashSwitch -eq 0){
                    $cnt = $Dictionary[$Word] + 1   
                    $Dictionary.Remove($Word)
                    $Dictionary.Add($Word, $cnt)
                    } else {
                            $Dictionary.$Word++   # Slow Method
                            }
            } else {
                $Dictionary.Add($Word, 1)
            }
        }
    } 
}

Write-Progress -Activity "Indexing words..." -Completed
# Filter Word List to remove any single values and length greter than 2, then sort by Word name
$WordCountList=$($Dictionary.GetEnumerator()| ? {($_.Value -gt 1) -AND ($_.Name.Length -gt 2)} | Sort Name )
$DictWords = $WordCountList.Count
$OutList = ($WordCountList | Select Name,Value -First 2)
Write-Host "$WordCount total words in the text"
Write-Host "$DictWords distinct words in the text"

#Call the print resule function
print-search

#Stopwatch Stop
$sw.Stop()

if ($Interactive){
        $xMenuChoiceA  = "0"
        while ( $xMenuChoiceA -ne "1"){
            Write-host "--------------------------------"
      
            [string]$xMenuChoiceA = read-host "Enter word to search or 1 to exit" 

  
            if ( $xMenuChoiceA -ne "1"){
                      Write-host "--------------------------------"
                      $SearchWord = $xMenuChoiceA
                      #Call the print resule function 
                      print-search
                    }
               }
                
}else{
        #Write timing to Log file, skip if Interactive
        $LogEntery= $(Get-Date -Format 'MM/dd/yyyy,hh:mm tt')+","+$HashSwitch+","+$sw.Elapsed.TotalSeconds+","+$TotalLines+","+$WordCount+","+$DictWords+","+$FileName 
        Write-Host "Writing Log: "$LogEntery
        Add-Content .\CountWords.log.txt $LogEntery
     }
