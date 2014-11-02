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
	Write-Host "Checking for BitsModule"
    Set-ModuleStatus -name BitsTransfer
	
    if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Host "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..."
		New-Item $DestFolder -type Directory | Out-Null
		Write-Host "Done! " -ForegroundColor Green
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
                 
                If ($BitsError -like "*HTTP status 400*" )
                            {
                                # loop three times
                                    While ($loop -lt "3" -and $Global:NewFileDownloadResult -eq $false) 
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
                Write-Host "Hit Generic Catch on New-FileDownload"
                Write-Host $bitserror
                
                }
		


			
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # end function New-FileDownload
# SIG # Begin signature block
# MIIQGQYJKoZIhvcNAQcCoIIQCjCCEAYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUo1uyAKaLdycdbRDal1Y8jfBv
# EqSggg1eMIIGozCCBYugAwIBAgIQD6hJBhXXAKC+IXb9xextvTANBgkqhkiG9w0B
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
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSsIZxptftLzrSj
# c+/o81Ue8Gg7UDANBgkqhkiG9w0BAQEFAASCAQAKyBAOY9xSTkESJ4sKuWUkqhnZ
# q7hnhNy/+9Jx1U5aI9FXJb86gruXnW662ZzeSRZtxjk0LfozmqT/m0RRhzO+uaO0
# hPDUJLLuTmLUY7ecQvokTVYeewqJPpZxt2qh1sR88xeZg6Y4i6UcaMJe/zI6oMDj
# 44SRGgx9x2mSyw4bcBM2lfA2kJrvEtoP6NNgaDIh3Lc0Dp59Vau6rwEhSOsO8WLB
# fswIRcG7NLWhMczPAyTmmz4nlhpASY+c43drcVIeAsoMJKwPghCZabEqcM0PBsKf
# iyWLbaNyE7VAEUwJAGa1o/pZZDnaW/fEi1TbwMrO6k3U7w23IcJta3Qdclda
# SIG # End signature block
