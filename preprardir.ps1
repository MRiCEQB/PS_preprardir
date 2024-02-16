if (Test-Path ".\x64") 
{
    Write-host "x64 folder found!" -f Cyan
}
else 
{
    Write-host "Creating x64 folder" -f Green
    New-Item -ItemType Directory -Path .\x64 | Out-Null
}

if (Test-Path ".\x86") 
{
    Write-host "x86 folder found!" -f Cyan
}
else 
{
    Write-host "Creating x86 folder" -f Green
    New-Item -ItemType Directory -Path .\x86 | Out-Null
}

foreach($setupFileName in Get-ChildItem -Filter *.exe)
{
	Write-host "`nProcessing $setupFileName" -f Cyan
	Write-host "Extracting Rar.exe from setup" -f Green
	& "C:\Program Files\WinRAR\Rar.exe" x $setupFileName "Rar.exe" "." -inul
	$exedate = (get-item .\Rar.exe).LastWriteTime.ToString("yyyy-MM-dd")

		If ($setupFileName.name -like "*-x64-*")
		{
			$rarVersion = ("$setupFileName".Substring(11).Trim(".exe"))
			$checkFileExists = ((".\x64\")+($exedate + "_rar" + $rarVersion + ".exe"))
			
			if (Test-Path -path $checkFileExists)
			{
				Write-Host ($exedate + "_rar" + $rarVersion + ".exe" + " already exists in x64 subdirectory - skipping!")  -f Red
				Remove-Item "Rar.exe"
			}
			else
			{
				Write-host "Moving Rar.exe to x64 subdirectory" -f Green
				Move-Item "Rar.exe" -Destination .\x64\
				Write-host (("Renaming Rar.exe to ") + ($exedate + "_rar" + $rarVersion + ".exe")) -f Green
				ren ".\x64\Rar.exe" ($exedate + "_rar" + $rarVersion + ".exe")
				Write-host "Done!" -f Cyan
			}
		}
		else
		{
			$rarVersion = ("$setupFileName".Substring(4).Trim(".exe"))
			$checkFileExists = ((".\x86\")+($exedate + "_rar" + $rarVersion + ".exe"))

			if (Test-Path -path $checkFileExists)
			{
				Write-Host ($exedate + "_rar" + $rarVersion + ".exe" + " already exists in x86 subdirectory - skipping!")  -f Red
				Remove-Item "Rar.exe"
			}
			else
			{
				Write-host "Moving Rar.exe to x86 subdirectory" -f Green
				Move-Item "Rar.exe" -Destination .\x86\			
				Write-host (("Renaming Rar.exe to ") + ($exedate + "_rar" + $rarVersion + ".exe")) -f Green
				ren ".\x86\Rar.exe" ($exedate + "_rar" + $rarVersion + ".exe")
				Write-host "Done!" -f Cyan
			}
		}
}

if ((Test-Path -path ".\x86\1996-09-02_rarar200.exe") -Or 
    (Test-Path -path ".\x86\1997-03-12_rarar201.exe") -Or
    (Test-Path -path ".\x86\1997-09-16_rarar202.exe") -Or 
    (Test-Path -path ".\x86\1998-03-19_rarar203.exe") -Or  
    (Test-Path -path ".\x86\1998-05-13_rarar204.exe") -Or  
    (Test-Path -path ".\x86\1998-09-21_rarar205.exe") -Or  
    (Test-Path -path ".\x86\1998-12-03_rarar206.exe") -Or  
    (Test-Path -path ".\x86\1999-02-02_rarB3A.EXE.exe"))
{
	Write-Host "`nOdd filenames detected!" -f Cyan
	Write-Host "Renaming..." -f Green
	Rename-Item -Path ".\x86\1996-09-02_rarar200.exe" -NewName "1996-09-02_rar200.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1997-03-12_rarar201.exe" -NewName "1997-03-12_rar201.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1997-09-16_rarar202.exe" -NewName "1997-09-16_rar202.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1998-03-19_rarar203.exe" -NewName "1998-03-19_rar203.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1998-05-13_rarar204.exe" -NewName "1998-05-13_rar204.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1998-09-21_rarar205.exe" -NewName "1998-09-21_rar205.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1998-12-03_rarar206.exe" -NewName "1998-12-03_rar206.exe" -erroraction 'silentlycontinue'
	Rename-Item -Path ".\x86\1999-02-02_rarB3A.EXE.exe" -NewName "1999-02-02_rar25b3.exe" -erroraction 'silentlycontinue'

	Write-Host "Deleting reappeared/orphaned duplicates..." -f Green
	Remove-Item -Path ".\x86\1996-09-02_rarar200.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1997-03-12_rarar201.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1997-09-16_rarar202.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1998-03-19_rarar203.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1998-05-13_rarar204.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1998-09-21_rarar205.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1998-12-03_rarar206.exe" -erroraction 'silentlycontinue'
	Remove-Item -Path ".\x86\1999-02-02_rarB3A.EXE.exe" -erroraction 'silentlycontinue'
	Write-Host "Process finished!" -f Cyan
}
else
{
	Write-Host "`nNo odd filenames detected!" -f Cyan
	Write-Host "Process finished!" -f Cyan
}