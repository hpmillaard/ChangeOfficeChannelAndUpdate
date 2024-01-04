If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {Start powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit}
$Channels = @("Beta Channel","Current Channel (Preview)","Current Channel","Monthly Enterprise Channel","Semi-Annual Enterprise Channel (Preview)","Semi-Annual Enterprise Channel")
$Channel = $Channels | ogv -Title "Select the Channel you want to use" -OutputMode Single
If ($Channel){
	switch ($Channel) {
		"Beta Channel" {$CDN = "5440fd1f-7ecb-4221-8110-145efaa6372f"}
		"Current Channel (Preview)" {$CDN = "64256afe-f5d9-4f86-8936-8840a6a4f5be"}
		"Current Channel" {$CDN = "492350f6-3a01-4f97-b9c0-c7c6ddf67d60"}
		"Monthly Enterprise Channel" {$CDN = "55336b82-a18d-4dd6-b5f6-9e5095c314a6"}
		"Semi-Annual Enterprise Channel (Preview)" {$CDN = "b8f9b850-328d-4355-9145-c59439a0c4cf"}
		"Semi-Annual Enterprise Channel" {$CDN = "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"}
	}
	$Key = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$URL = "http://officecdn.microsoft.com/pr/$CDN"
	sp $Key -Name AudienceId -Value $CDN
	sp $Key -Name CDNBaseUrl -Value $URL
	sp $Key -Name UpdateChannel -Value $URL
	start "$ENV:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" "/update user"
}