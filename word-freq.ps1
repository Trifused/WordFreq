# PowerShell: Counting words in a text file
# Returns a list of the most-used, longest words in a text file.
#
# Adapted slightly from 
# https://blogs.technet.microsoft.com/josebda/2015/03/21/powershell-examples-counting-words-in-a-text-file/
# 
# Run this with PowerShell in the folder containing your text, from PowerShell with:
# ./word-freq
# It looks for combined.txt. To get this from a set of markdown files,
# I use Cygwin (because native Windows CLIs have enciding issues) and run:
# cat *.md > combined.txt
#
# The very last number in the script sets the length of the word list you'll get.
#
Clear-Host
$FileName = "combined.txt"
Write-Host "Reading file $FileName..." 
$File = Get-Content $FileName
$TotalLines = $File.Count
Write-Host "$TotalLines lines read from the file."

$SearchWord = ""
$Found = 0
$WordCount = 0
$Longest = ""
$Dictionary = @{}
$LineCount = 0

$File | foreach {
    $Line = $_
    $LineCount++
    Write-Progress -Activity "Processing words..." -PercentComplete ($LineCount*100/$TotalLines) 
    $Line.Split(" .,:;?!/()[]{}-```"") | foreach {
        $Word = $_.ToUpper()
        If ($Word[0] -ge 'A' -and $Word[0] -le "Z") {
            $WordCount++
            If ($Word.Contains($SearchWord)) { $Found++ }
            If ($Word.Length -gt $Longest.Length) { $Longest = $Word }
            If ($Dictionary.ContainsKey($Word)) {
                $Dictionary.$Word++
            } else {
                $Dictionary.Add($Word, 1)
            }
        }
    } 
}

Write-Progress -Activity "Processing words..." -Completed
$DictWords = $Dictionary.Count
Write-Host "There were $WordCount total words in the text"
Write-Host "There were $DictWords distinct words in the text"
Write-Host "The word $SearchWord was found $Found times."
Write-Host "The longest word was $Longest" 
Write-Host
Write-Host "Most used words with more than 4 letters:"

$Dictionary.GetEnumerator() | ? { $_.Name.Length -gt 4 } | 
Sort Value -Descending | Select -First 200
