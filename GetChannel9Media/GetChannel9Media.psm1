

function Get-Channel9Media {
<#  
.SYNOPSIS  
   	Downloads  MP4's and PPTX from channel 9 site, Based on input CSV

.DESCRIPTION  
    Downloads Teched MP4's and PPTX from channel 9 site, Based on input CSV.
    CSV format Code,Name, Where Code is session code and Name is the name you want to call the download file.

.NOTES  
    Modified for Channel9 Downloads NA by Tom Arbuthnot lyncdup.com
    Original provided by blog.SCOMfaq.ch / Stefan Roth
    credit http://blog.scomfaq.ch/2012/06/13/teched-2012-orlando-download-sessions-offline-viewing/
    Credit: Pat Richard for New-Download Function http://www.ehloworld.com

    Microsoft handily use a simple format of URL/SessionID.MP4/PPTX, so in future techeds this method will also likely
    work
    
    Use completely at your own risk

.LINK  
	http://tomtalks.uk

.EXAMPLE
    Get-Channel9Media -SessionCSV .\InputCSV.csv -DownloadTargetDirectory C:\Downloads
    Get-Channel9Media -SessionCSV .\InputCSV.csv -DownloadTargetDirectory C:\Downloads 

.INPUTS
    None. You cannot pipe objects to this script.
#>

	[CmdletBinding(SupportsShouldProcess = $True)]
Param(
        [parameter(mandatory=$true)]
        [String]$SessionCSV,

        [parameter(mandatory=$true)]
        [String]$DownloadTargetDirectory,

        [Switch]$DownloadMP3Only
        
        )

Begin {
Write-host ' '
Write-host ' '
Write-host ' '
Write-host ' '
Write-host ' '
Write-host ' '
Write-host ' '
Write-host ' '
Write-Host '###############################################################################################'  -ForegroundColor Yellow
Write-Host '                                                                                               '  -ForegroundColor Yellow
Write-Host '                    Get-Channel9Media by Tom Arbuthnot http://tomtalks.uk                      '  -ForegroundColor Green
Write-Host '                                                                                               '  -ForegroundColor Yellow
Write-Host '            Credit: Original 2012 downloader by http://blog.scomfaq.ch (c) 2012 Stefan Roth    '  -ForegroundColor Green
Write-Host '                                                                                               '  -ForegroundColor Yellow
Write-Host '                 Credit: Pat Richard for New-Download Function http://www.ehloworld.com/       '  -ForegroundColor Green
Write-Host '                                                                                               '  -ForegroundColor Yellow
Write-Host '###############################################################################################'  -ForegroundColor Yellow


$sessions = Import-csv $SessionCSV

$TotalDownloads =  $($sessions.count)

           }


Process {
    Foreach($session in $sessions){ #ForEach Loop 1

            $i = $i+1


        try { # Try 1


                $name = ($session.Name)
                

                # Clean Up names for characters that can't be in filenames
                $name = $name.Replace('\',' ')
                $name = $name.Replace('/',' ')
                $name = $name.Replace(':',' ')
                $name = $name.Replace("`*",' ')
                $name = $name.Replace("`?",' ')
                $name = $name.Replace('|',' ')

           

                    # Download MP4 or MP3

                    $ToDownload = $($session.URL)

                    $DestinationFile = "$($name)" + '.mp4'
                    
                    If ($DownloadMP3Only) # change the file name for the mp3s
                        {

               
                        $MP4URL = $($session.URL)

                        $ToDownload = $MP4URL.Replace('_high.mp4','.mp3')


                        $DestinationFile = "$($name)" + '.mp3'

                        Write-Host "File URL: $ToDownload "
                        }
       

                    write-host -ForegroundColor Yellow ('Queueing Session ...' + $($name) + ' Number: ' + $i + ' out of ' + $TotalDownloads)
                    New-FileDownload -SourceFile $ToDownload -DestFolder $DownloadTargetDirectory -DestFile $DestinationFile


    } # close Try1
     
    catch {
            write-host "`n"
	        write-host 'Download hit an error, Caught in Catch1 in Get-Channel9Media'
            Write-Host "Result from function is $Global:NewFileDownloadResult"
            write-host "`n"

           

        } # Close Catch


            } # Close Foreach Loop1

write-host ' '
write-host -ForegroundColor Green 'Finished!'

} # Close Process Block

} # Close Function

# Credit Pat Richard

# 0.4 added rety loop for 400 fail, seem to get this often on some machines

function New-FileDownload {
	[CmdletBinding(SupportsShouldProcess = $True)]
	param(
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[ValidateNotNullOrEmpty()]
		[string]$SourceFile,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFolder,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFile
	)
	
    [bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

	if (!($DestFolder)){$DestFolder = $TargetFolder}
	Write-Host 'Checking for BitsModule'
    Set-ModuleStatus -name BitsTransfer
	
    if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf('/') + 1)}
	if (Test-Path $DestFolder){
		Write-Host "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..."
		New-Item $DestFolder -type Directory | Out-Null
		Write-Host 'Done! ' -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host -ForegroundColor Yellow "File: $DestFile already exists."
        #write finish result to global
        $Global:NewFileDownloadResult = $?
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." 
			
            Try {
                # Forcing the error output to  a custom variable, as it was the only way to catch the non-terminating error

                # clear down error
                $bitserror = $null
                Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction SilentlyContinue
                # Write-Host "Done! " -ForegroundColor Green

                # This sends the result variable of the last command run to the global scope
                # $? is true if command ran successfully and false if it didn't
                
                # Show-ErrorDetails $bitserror
                
                # Write if this was successful or not to a global variable
                $Global:NewFileDownloadResult = $?
                $Global:NewFileDownloadError = $bitserror
                 
                If ($BitsError -like '*HTTP status 400*' )
                            {
                                # loop three times
                                    While ($loop -lt '3' -and $Global:NewFileDownloadResult -eq $false) 
                                    {

                                    $bitserror = $null
                                    Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction SilentlyContinue 
                                    $Global:NewFileDownloadResult = $?
                                    Start-Sleep -Seconds 5
                                    $loop = $loop + 1
                                    Write-Host "Download Rety attempt $loop of 3"
                                    
                                    }
                             

                            } # Close If error 400


                }
	       catch
                {
                Write-Host 'Hit Generic Catch on New-FileDownload'
                Write-Host $bitserror
                
                }
		


			
		}else{
			Write-Host 'Internet access not detected. Please resolve and try again.' -ForegroundColor red
		}
	}
} # end function New-FileDownload

# Credit Pat Richard
function Set-ModuleStatus { 
	[CmdletBinding(SupportsShouldProcess = $True)]
	param	(
		[parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available
			return $false
		}
	}else {
		# module was already imported
		# Write-Host "$name module already imported"
		return $true
	}
} # end function Set-ModuleStatus

