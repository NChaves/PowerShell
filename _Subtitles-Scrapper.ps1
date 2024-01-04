$ParentPath = "Q:\Completed Downloads\"
$Files = Get-ChildItem -Path $ParentPath -Recurse -Include *.nfo

############################# START 7zip #############################

$7zaBinary = "C:\Program Files\7-Zip\7z.exe"
Function Expand-7Zip {
    #[CmdletBinding(HelpUri='http://gavineke.com/PS7Zip/Expand-7Zip')]
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipelineByPropertyName=$True)]
        [ValidateScript({$_ | Test-Path -PathType Leaf})]
        [System.IO.FileInfo]$FullName,

        [Parameter()]
        [Alias('Destination')]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath,

        [Parameter()]
        [switch]$Remove
    )
    
    Begin {}
    
    Process {
        Write-Verbose -Message 'Extracting contents of compressed archive file'
        If ($DestinationPath) {
            & "$7zaBinary" x -o"$DestinationPath" "$FullName"
        } Else {
            & "$7zaBinary" x "$FullName"
        }

        If ($Remove) {
            Write-Verbose -Message 'Removing compressed archive file'
            Remove-Item -Path "$FullName" -Force
        }
    }
    
    End {}
    
}

############################# END 7zip #############################
	
foreach ($File in $Files) {
		$FullPath = $File | % { $_.FullName }
		$CreatedDate = (Get-ChildItem $File).LastWriteTime
		$Dir = $File.Directory.FullName
		$Directory = $Dir+'\'		
		$MediaFiles = Get-ChildItem -Path $Directory -Recurse -Include *.avi, *.mp4, *.mkv
		$MediaFileName = Get-ChildItem -Path $Directory -Recurse -Include *.avi, *.mp4, *.mkv | Select BaseName
		$SubsPT = Get-ChildItem -Path $Directory -Recurse -Include *.pt.srt, *.pt.sub, *.pt.ass, *.pt.smi, *.pt.ssa
		$SubsEN = Get-ChildItem -Path $Directory -Recurse -Include *.en.srt, *.en.sub, *.en.ass, *.en.smi, *.en.ssa
		$SubsES = Get-ChildItem -Path $Directory -Recurse -Include *.es.srt, *.es.sub, *.es.ass, *.es.smi, *.es.ssa
		$SubsBR = Get-ChildItem -Path $Directory -Recurse -Include *.pt-br.srt, *.pt-br.sub, *.pt-br.ass, *.pt-br.smi, *.pt-br.ssa
		#=======================================================
		$TitleStart = '<title>'
		$TitleEnd = '</title>'
		$YearStart = '<year>'
		$YearEnd = '</year>'
		$DateStart = '<premiered>'
		$DateEnd = '</premiered>'
		$IMDBIDStart = '<uniqueid default="true" type="imdb">'
		$IMDBIDEnd = '</uniqueid>'
		#Get content from file
		$nfo = Get-Content $File
		#===========================Get Title============================= 
		#Regex pattern to compare two strings
		$titlematch = "$TitleStart(.*?)$TitleEnd"
		#Perform the opperation
		$title = [regex]::Match($nfo,$titlematch).Groups[1].Value
		#===========================Get Year==============================
		#Regex pattern to compare two strings
		$yearmatch = "$YearStart(.*?)$YearEnd"
		#Perform the opperation
		$year = [regex]::Match($nfo,$yearmatch).Groups[1].Value
		#===========================Get Date==============================
		#Regex pattern to compare two strings
		$datematch = "$DateStart(.*?)$DateEnd"
		#Perform the opperation
		$date = [regex]::Match($nfo,$datematch).Groups[1].Value
		#===========================Get IMDB ID==============================
		#Regex pattern to compare two strings
		$IMDBmatch = "$IMDBIDStart(.*?)$IMDBIDEnd"
		#Perform the opperation
		$IMDB = [regex]::Match($nfo,$IMDBMatch).Groups[1].Value
		#===========================Return Results========================
		
		Invoke-RestMethod https://yts-subs.com/movie-imdb/$IMDB | Out-File -File "$Directory\$IMDB.txt"
		$subLink = get-content "$Directory\$IMDB.txt" -ReadCount 1000 | foreach { $_ -match "$year-portuguese-yify-"} | Out-File -File "$Directory\$IMDB-links.txt"
		$subLink = get-content "$Directory\$IMDB.txt" -ReadCount 1000 | foreach { $_ -match "$year-english-yify-"} | Out-File -File "$Directory\$IMDB-links.txt" -Append
		$subLink = get-content "$Directory\$IMDB.txt" -ReadCount 1000 | foreach { $_ -match "$year-spanish-yify-"} | Out-File -File "$Directory\$IMDB-links.txt" -Append
		$subLink = get-content "$Directory\$IMDB.txt" -ReadCount 1000 | foreach { $_ -match "$year-brazilian-portuguese-yify-"} | Out-File -File "$Directory\$IMDB-links.txt" -Append
		$subLink = get-content "$Directory\$IMDB-links.txt"
		$subLink.Trim() | Out-File -File "$Directory\$IMDB-links.txt"
		$subLink = get-content "$Directory\$IMDB-links.txt"
		$subLink2 = $subLink | Where-Object {$_ -notmatch "subtitle-download"} | Out-File -File "$Directory\$IMDB-links.txt"
		$test = get-content "$Directory\$IMDB-links.txt" 
		$test.Replace('<a href="/subtitles','') | Out-File -File "$Directory\$IMDB-links.txt"
		$test = get-content "$Directory\$IMDB-links.txt" 
		$test.Replace('">','') | Out-File -File "$Directory\$IMDB-links.txt"

        #### https://yts-subs.com/movie-imdb/tt2249221
		#### https://yifysubtitles.ch/subtitle/top-gun-1986-brazilian-portuguese-yify-133353.zip
		#### https://yifysubtitles.ch/subtitle/peter-rabbit-2-the-runaway-2021-english-yify-340413.zip

### brazilian-portuguese-


		#Return result
		#Write-Host $title
		#Write-Host $year
		#Write-Host $date
		#Write-Host $Directory
		#Write-Host $IMDB
		if ($SubsPT -like "*.srt" -or $SubsPT -like "*.sub" -or $SubsPT -like "*.ass" -or $SubsPT -like "*.smi") {
		Write-Host "has PT sub: $File"
		} else {
            $subsPT = get-content "$Directory\$IMDB-links.txt" -ReadCount 1000 | foreach { $_ -match "$year-portuguese-yify"} | Select-Object -First 1 -ErrorAction SilentlyContinue
            $subsPTLenght = $subsPT.Length
            if ($subsPTLenght -gt 10) {
            foreach ($subPT in $SubsPT) {
            Write-Host $subPT
            $subPTzipname = $IMDB +"pt.zip"
            $subPTdir = $IMDB +"pt"
            $newFileNamePT = $MediaFileName.basename
            $DownLink = "https://yifysubtitles.ch/subtitle$subPT.zip"
            $statusCode = wget $DownLink | % {$_.StatusCode} -ErrorAction SilentlyContinue
            if ($statuscode -eq 200) {
            Invoke-WebRequest -Uri $DownLink -OutFile "$Directory\$subPTzipname"
            #Start-Process $DownLink
            New-Item -Path "$Directory" -Name "$subPTdir" -ItemType "directory"
            Start-Sleep -Milliseconds 1000
            $subArchivePT = "$Directory" + "$subPTzipname"
            $subDir = "$Directory" + "$subPTdir"
            Expand-7Zip $subArchivePT -DestinationPath $subDir
            #Expand-Archive -LiteralPath $subArchivePT -DestinationPath $subDir
            Start-Sleep -Milliseconds 2000
            $GetExpandedSubName = Get-ChildItem -Path $subDir -Recurse -Include *.srt, *.sub, *.ass, *.smi, *.ssa
            $ExpandedSubName = $GetExpandedSubName.Name
            $ExpandedSubExtension = $GetExpandedSubName.Extension
            $ExpandedSubDir = "$subDir\$ExpandedSubName"
            $subDestination = "$Directory$newFileNamePT.pt" + $ExpandedSubExtension
            #Move-Item "$ExpandedSubDir" -Destination "$subDestination"
            Get-ChildItem $subDir | Move-Item  -Destination $subDestination
            #$SubContent = Get-Content "$ExpandedSubDir" | Out-File $subDestination
            Start-Sleep -Milliseconds 1000
            $subDestination | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            Get-ChildItem -Path $subDir | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$subDir" -Force -ErrorAction SilentlyContinue
            }
            $statusCode = "Error 404: $DownLink" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            } else {
            $statusCode = "no PT hyperlink: $File" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            }

		if ($SubsEN -like "*.srt" -or $SubsEN -like "*.sub" -or $SubsEN -like "*.ass" -or $SubsEN -like "*.smi") {
		Write-Host "has EN sub: $File"
		} else {
            $subsEN = get-content "$Directory\$IMDB-links.txt" -ReadCount 1000 | foreach { $_ -match "$year-english-yify"} | Select-Object -First 1 -ErrorAction SilentlyContinue
            $subsENLenght = $subsEN.Length
            if ($subsENLenght -gt 10) {
            foreach ($subEN in $SubsEN) {
            Write-Host $subEN
            $subENzipname = $IMDB +"en.zip"
            $subENdir = $IMDB +"en"
            $newFileNameEN = $MediaFileName.basename
            $DownLink = "https://yifysubtitles.ch/subtitle$subEN.zip"
            $statusCode = wget $DownLink | % {$_.StatusCode} -ErrorAction SilentlyContinue
            if ($statuscode -eq 200) {
            Invoke-WebRequest -Uri $DownLink -OutFile "$Directory\$subENzipname"
            #Start-Process $DownLink
            New-Item -Path "$Directory" -Name "$subENdir" -ItemType "directory"
            Start-Sleep -Milliseconds 1000
            $subArchiveEN = "$Directory" + "$subENzipname"
            $subDir = "$Directory" + "$subENdir"
            Expand-7Zip $subArchiveEN -DestinationPath $subDir
            #Expand-Archive -LiteralPath $subArchivePT -DestinationPath $subDir
            Start-Sleep -Milliseconds 2000
            $GetExpandedSubName = Get-ChildItem -Path $subDir -Recurse -Include *.srt, *.sub, *.ass, *.smi, *.ssa
            $ExpandedSubName = $GetExpandedSubName.Name
            $ExpandedSubExtension = $GetExpandedSubName.Extension
            $ExpandedSubDir = "$subDir\$ExpandedSubName"
            $subDestination = "$Directory$newFileNameEN.en" + $ExpandedSubExtension
            #Move-Item "$ExpandedSubDir" -Destination "$subDestination"
            Get-ChildItem $subDir | Move-Item  -Destination $subDestination
            #$SubContent = Get-Content "$ExpandedSubDir" | Out-File $subDestination
            Start-Sleep -Milliseconds 1000
            $subDestination | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            Get-ChildItem -Path $subDir | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$subDir" -Force -ErrorAction SilentlyContinue
            }
            $statusCode = "Error 404: $DownLink" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            } else {
            $statusCode = "no EN hyperlink: $File" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            }

		if ($SubsES -like "*.srt" -or $SubsES -like "*.sub" -or $SubsES -like "*.ass" -or $SubsES -like "*.smi") {
		Write-Host "has ES sub: $File"
		} else {
            $subsES = get-content "$Directory\$IMDB-links.txt" -ReadCount 1000 | foreach { $_ -match "$year-spanish-yify"} | Select-Object -First 1 -ErrorAction SilentlyContinue
            $subsESLenght = $subsES.Length
            if ($subsESLenght -gt 10) {
            foreach ($subES in $SubsES) {
            Write-Host $subES
            $subESzipname = $IMDB +"es.zip"
            $subESdir = $IMDB +"es"
            $newFileNameES = $MediaFileName.basename
            $DownLink = "https://yifysubtitles.ch/subtitle$subES.zip"
            $statusCode = wget $DownLink | % {$_.StatusCode} -ErrorAction SilentlyContinue
            if ($statuscode -eq 200) {
            Invoke-WebRequest -Uri $DownLink -OutFile "$Directory\$subESzipname"
            #Start-Process $DownLink
            New-Item -Path "$Directory" -Name "$subESdir" -ItemType "directory"
            Start-Sleep -Milliseconds 1000
            $subArchiveES = "$Directory" + "$subESzipname"
            $subDir = "$Directory" + "$subESdir"
            Expand-7Zip $subArchiveES -DestinationPath $subDir
            #Expand-Archive -LiteralPath $subArchiveES -DestinationPath $subDir
            Start-Sleep -Milliseconds 2000
            $GetExpandedSubName = Get-ChildItem -Path $subDir -Recurse -Include *.srt, *.sub, *.ass, *.smi, *.ssa
            $ExpandedSubName = $GetExpandedSubName.Name
            $ExpandedSubExtension = $GetExpandedSubName.Extension
            $ExpandedSubDir = "$subDir\$ExpandedSubName"
            $subDestination = "$Directory$newFileNameES.es" + $ExpandedSubExtension
            #Move-Item "$ExpandedSubDir" -Destination "$subDestination"
            Get-ChildItem $subDir | Move-Item  -Destination $subDestination
            #$SubContent = Get-Content "$ExpandedSubDir" | Out-File $subDestination
            Start-Sleep -Milliseconds 1000
            $subDestination | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            Get-ChildItem -Path $subDir | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$subDir" -Force -ErrorAction SilentlyContinue
            }
            $statusCode = "Error 404: $DownLink" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            } else {
            $statusCode = "no ES hyperlink: $File" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            }

            if ($SubsBR -like "*.srt" -or $SubsBR -like "*.sub" -or $SubsBR -like "*.ass" -or $SubsBR -like "*.smi" -or $SubsPT -like "*.srt" -or $SubsPT -like "*.sub" -or $SubsPT -like "*.ass" -or $SubsPT -like "*.smi") {
            Write-Host "has BR or PT sub: $File"
            } else {
            $subsBR = get-content "$Directory\$IMDB-links.txt" -ReadCount 1000 | foreach { $_ -match "$year-brazilian-portuguese-yify"} | Select-Object -First 1 -ErrorAction SilentlyContinue
            $subsBRLenght = $subsBR.Length
            if ($subsBRLenght -gt 10) {
            foreach ($subBR in $SubsBR) {
            Write-Host $subBR
            $subBRzipname = $IMDB +"pt-br.zip"
            $subBRdir = $IMDB +"pt-br"
            $newFileNameBR = $MediaFileName.basename
            $DownLink = "https://yifysubtitles.ch/subtitle$subBR.zip"
            $statusCode = wget $DownLink | % {$_.StatusCode} -ErrorAction SilentlyContinue
            if ($statuscode -eq 200) {
            Invoke-WebRequest -Uri $DownLink -OutFile "$Directory\$subBRzipname"
            #Start-Process $DownLink
            New-Item -Path "$Directory" -Name "$subBRdir" -ItemType "directory"
            Start-Sleep -Milliseconds 1000
            $subArchiveBR = "$Directory" + "$subBRzipname"
            $subDir = "$Directory" + "$subBRdir"
            Expand-7Zip $subArchiveBR -DestinationPath $subDir
            #Expand-Archive -LiteralPath $subArchiveBR -DestinationPath $subDir
            Start-Sleep -Milliseconds 2000
            $GetExpandedSubName = Get-ChildItem -Path $subDir -Recurse -Include *.srt, *.sub, *.ass, *.smi, *.ssa
            $ExpandedSubName = $GetExpandedSubName.Name
            $ExpandedSubExtension = $GetExpandedSubName.Extension
            $ExpandedSubDir = "$subDir\$ExpandedSubName"
            $subDestination = "$Directory$newFileNameBR.pt-br" + $ExpandedSubExtension
            #Move-Item "$ExpandedSubDir" -Destination "$subDestination"
            Get-ChildItem $subDir | Move-Item  -Destination $subDestination
            #$SubContent = Get-Content "$ExpandedSubDir" | Out-File $subDestination
            Start-Sleep -Milliseconds 1000
            $subDestination | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            Get-ChildItem -Path $subDir | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$subDir" -Force -ErrorAction SilentlyContinue
            }
            $statusCode = "Error 404: $DownLink" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            } else {
            $statusCode = "no BR hyperlink: $File" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
            }
            }

            Remove-Item "$Directory\$IMDB.txt"
            Remove-Item "$Directory\$IMDB-links.txt"
            $got_ptSubs = Get-ChildItem -Path $Directory -Recurse -Include *pt.srt, *pt.sub, *pt.ass, *pt.smi, *pt.ssa
            if ($got_ptSubs.Length -gt 1) {
            $pr_BRSub = Get-ChildItem -Path $Directory -Recurse -Include *pt-br.srt, *pt-br.sub, *pt-br.ass, *pt-br.smi, *pt-br.ssa | Remove-Item
            }
}

    $zips = Get-ChildItem -Path $ParentPath -Recurse -Include *.zip
    #$Files 

    foreach ($file in $zips) {
        if ($file -like "*.zip") {
        Write-Host $file
        Remove-Item $file
        }
    }

    $subs = Get-ChildItem -Path $ParentPath -Recurse -Include *.srt, *.sub, *.ass, *.smi, *.ssa

    foreach ($file in $subs) {
    $size = $file.Length
                #Write-Host $size
                if ($size -lt 2000) {
                #Write-Host "rubish: $file $size" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
                $result = "rubish: $file $size" | Out-File -File "P:\Completed Downloads\EmptySubsList.txt" -Append
                Remove-Item $file
                    }
    }

    # $statusCode = wget "https://yifysubtitles.ch/subtitle/anomalous-2016-english-yify-348573.zip" | % {$_.StatusCode}
    # $statusCode = wget "https://yifysubtitles.ch/subtitle/peter-rabbit-2-the-runaway-2021-english-yify-340413.zip" | % {$_.StatusCode}
    # $statusCode = wget $DownLink | % {$_.StatusCode}