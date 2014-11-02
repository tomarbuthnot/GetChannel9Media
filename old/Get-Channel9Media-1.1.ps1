
# Test for how script was loaded to allow warning if loaded directly rather than
# via import-module
if ($MyInvocation.InvocationName -eq ‘&‘) {
    #write-host “Called using operator“
} 
    elseif ($MyInvocation.InvocationName -eq ‘.‘) {
    #write-host “Dot sourced“
    } 
    elseif ((Resolve-Path -Path $MyInvocation.InvocationName).ProviderPath -eq $MyInvocation.MyCommand.Path) {
    write-host “ Script Called using path $($MyInvocation.InvocationName)“
    $calledusingpath = $true
}

IF ($calledusingpath)
{
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host " Hi"
Write-Host " "
Write-Host " It looks like you've tried to run this as a script but it's actually a module (a package of scripts)"
write-Host " "
write-Host " Don't panic, just Run Import-Module followed by the folder path to the module.psm1 file"
write-Host " "
write-host " e.g. Import-Module c:\mydownloads\GetChannel9MediaPSModule\GetChannel9MediaPSModule.psm1"
write-host " "
write-Host " One the module is loaded you can use the cmdlet in the normal way"
write-host " "
write-host " For Example, Get-Channel9Media -SessionCSV .\InputCSV.csv -DownloadTargetDirectory C:\Downloads"
Write-Host " "
}


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
	www.lyncdup.com

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
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-Host "###############################################################################################"  -ForegroundColor Yellow
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                    Get-Channel9Media by Tom Arbuthnot http://lyncdup.com                      "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "            Credit: Original 2012 downloader by http://blog.scomfaq.ch (c) 2012 Stefan Roth    "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                 Credit: Pat Richard for New-Download Function http://www.ehloworld.com/       "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "###############################################################################################"  -ForegroundColor Yellow


$sessions = Import-csv $SessionCSV

$TotalDownloads =  $($sessions.count)

           }


Process {
    Foreach($session in $sessions){ #ForEach Loop 1

            $i = $i+1


        try { # Try 1


                $name = ($session.Name)
                

                # Clean Up names for characters that can't be in filenames
                $name = $name.Replace("\"," ")
                $name = $name.Replace("/"," ")
                $name = $name.Replace(":"," ")
                $name = $name.Replace("`*"," ")
                $name = $name.Replace("`?"," ")
                $name = $name.Replace("|"," ")

           

                    # Download MP4 or MP3

                    $ToDownload = $($session.URL)

                    $DestinationFile = "$($name)" + ".mp4"
                    
                    If ($DownloadMP3Only) # change the file name for the mp3s
                        {

               
                        $MP4URL = $($session.URL)

                        $ToDownload = $MP4URL.Replace("_high.mp4",".mp3")


                        $DestinationFile = "$($name)" + ".mp3"

                        Write-Host "File URL: $ToDownload "
                        }
       

                    write-host -ForegroundColor Yellow ("Queueing Session ..." + $($name) + " Number: " + $i + " out of " + $TotalDownloads)
                    New-FileDownload -SourceFile $ToDownload -DestFolder $DownloadTargetDirectory -DestFile $DestinationFile


    } # close Try1
     
    catch {
            write-host "`n"
	        write-host "Download hit an error, Caught in Catch1 in Get-Channel9Media"
            Write-Host "Result from function is $Global:NewFileDownloadResult"
            write-host "`n"

           

        } # Close Catch


            } # Close Foreach Loop1

write-host " "
write-host -ForegroundColor Green "Finished!"

} # Close Process Block

} # Close Function

# SIG # Begin signature block
# MIIQGQYJKoZIhvcNAQcCoIIQCjCCEAYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUDhiuDxLXMe8q8ao4rsR5VX+h
# yESggg1eMIIGozCCBYugAwIBAgIQD6hJBhXXAKC+IXb9xextvTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTEwMjExMTIwMDAwWhcNMjYwMjEwMTIwMDAwWjBvMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnHz5
# oI8KyolLU5o87BkifwzL90hE0D8ibppP+s7fxtMkkf+oUpPncvjxRoaUxasX9Hh/
# y3q+kCYcfFMv5YPnu2oFKMygFxFLGCDzt73y3Mu4hkBFH0/5OZjTO+tvaaRcAS6x
# ZummuNwG3q6NYv5EJ4KpA8P+5iYLk0lx5ThtTv6AXGd3tdVvZmSUa7uISWjY0fR+
# IcHmxR7J4Ja4CZX5S56uzDG9alpCp8QFR31gK9mhXb37VpPvG/xy+d8+Mv3dKiwy
# RtpeY7zQuMtMEDX8UF+sQ0R8/oREULSMKj10DPR6i3JL4Fa1E7Zj6T9OSSPnBhbw
# JasB+ChB5sfUZDtdqwIDAQABo4IDQzCCAz8wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMIIBwwYDVR0gBIIBujCCAbYwggGyBghghkgBhv1sAzCC
# AaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1jcHMt
# cmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBz
# AGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBv
# AG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAg
# AHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAg
# AHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBt
# AGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0
# AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABo
# AGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wEgYDVR0TAQH/
# BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUe2jOKarAF75JeuHl
# P9an90WPNTIwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZI
# hvcNAQEFBQADggEBAHtyHWT/iMg6wbfp56nEh7vblJLXkFkz+iuH3qhbgCU/E4+b
# gxt8Q8TmjN85PsMV7LDaOyEleyTBcl24R5GBE0b6nD9qUTjetCXL8KvfxSgBVHkQ
# RiTROA8moWGQTbq9KOY/8cSqm/baNVNPyfI902zcI+2qoE1nCfM6gD08+zZMkOd2
# pN3yOr9WNS+iTGXo4NTa0cfIkWotI083OxmUGNTVnBA81bEcGf+PyGubnviunJmW
# eNHNnFEVW0ImclqNCkojkkDoht4iwpM61Jtopt8pfwa5PA69n8SGnIJHQnEyhgmZ
# cgl5S51xafVB/385d2TxhI2+ix6yfWijpZCxDP8wggazMIIFm6ADAgECAhAHg+Of
# aoJDPCCKgFvnHm/HMA0GCSqGSIb3DQEBBQUAMG8xCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xLjAs
# BgNVBAMTJURpZ2lDZXJ0IEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBLTEwHhcN
# MTMwNzAxMDAwMDAwWhcNMTQwNzA5MTIwMDAwWjB/MQswCQYDVQQGEwJHQjEWMBQG
# A1UECBMNSGVydGZvcmRzaGlyZTESMBAGA1UEBxMJU3RldmVuYWdlMSEwHwYDVQQK
# ExhUaG9tYXMgQ2hhcmxlcyBBcmJ1dGhub3QxITAfBgNVBAMTGFRob21hcyBDaGFy
# bGVzIEFyYnV0aG5vdDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALHE
# 1KWj6eAh2E54UiDcHmcmw817ohO4sZ5zMirY5CFJx4G/IIfEg6JHneIXtNrY9QbH
# 2gBvoCJ/j+rMLUiG0G8jw2n0mOAyWEcBDga57SDzI6OHyKM3n+OkC5D6wQSS0lH5
# e90Suegs5bxLfZFTSFWVRKsHhoCtKFVevaEKIbt2S8wE5Fdss2BCsmgf7RcIrj4r
# Zcxg3OZ1UDtDwCPIncryM0j/BC+81j/QPTJ4fu2rfSVEKELHR89JN+MAdrcJbWLH
# Zl9SgsVGDWG15wQUiVYB+A1Mz6ZwT+3St7/iJgWGFvZdcI+A7sEWZZSIyJMre8/s
# CYEeO1bRImqVbbS1RX8CAwEAAaOCAzkwggM1MB8GA1UdIwQYMBaAFHtozimqwBe+
# SXrh5T/Wp/dFjzUyMB0GA1UdDgQWBBQ4SFcQ1DpIYCaa1vJbAgMEUO7/UTAOBgNV
# HQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwcwYDVR0fBGwwajAzoDGg
# L4YtaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL2Fzc3VyZWQtY3MtMjAxMWEuY3Js
# MDOgMaAvhi1odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDEx
# YS5jcmwwggHEBgNVHSAEggG7MIIBtzCCAbMGCWCGSAGG/WwDATCCAaQwOgYIKwYB
# BQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1jcHMtcmVwb3NpdG9y
# eS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYA
# IAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkA
# dAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAA
# RABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAA
# UgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAA
# dwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4A
# ZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkA
# bgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wgYIGCCsGAQUFBwEBBHYwdDAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEwGCCsGAQUFBzAC
# hkBodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURD
# b2RlU2lnbmluZ0NBLTEuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQEFBQAD
# ggEBAHgP9yvgRLYzST2TX1EyULaVCbCHskIGU492MMofP18wz0V6+k1xJ0oql+Oy
# Ph5gJBOnAeKOio1dyzxy6UHYTCEmFHgvI58KpJzy930szpCWCoaOIUBegy2zoYd+
# EKB0H1pA4FD93bkt3T48HlP/54FBkSeiDL/Q8Hw1ar7acZx0GOAfHOLa2QjUhzJK
# W1Zp9S2nWX2FSvM5HotQeQDp0UVqIgPd7d7FD16GiRZkPdSWoQ/bQcS+kpQzG9n6
# ePMe1HpHx0FFB78MBYd3LDpPs4XnZlw9pQGuAoL7T4lsoUNMH+SA0io+jRgtLUzB
# XUEZC8y0ESYBXxTteMmbzmUkxT8xggIlMIICIQIBATCBgzBvMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QS0xAhAHg+OfaoJDPCCKgFvnHm/HMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEM
# MQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQB
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRpUgOZCee2Ylgh
# y9png6NriaXdnTANBgkqhkiG9w0BAQEFAASCAQBKP7RXX0J5A5DEzJvdjI5KYyhx
# M19auU1Zl5BlHI5dn2+/YEFj/UPw/6OJ/uh2a0U58Fc29dpDS182O1J/5MRMyZE+
# vDNyblhO1jPWVwhMDVSEa9eKJDa4cc73u8JghK+6NKoT+4Y73AL8oWXLHPvJQ5sg
# UhEGCyKzaCRXKspOmpOafvI+PCIs8M1LT6c4yXnLT1X0RbTRi7WFe8aU6GrRwRDp
# uoktJsTICur+kuxXEMj8VOe6qrjm5B8rBCs9ObFUokUjRmXpgl2Y9b2Jno2Qk5hN
# E91UqCWL8oGZjx44nmn+WZUh0gKDFW5daeSHllOpuMAB27XBO5Qi6UTWiFwc
# SIG # End signature block
