<# 

.DESCRIPTION 
 Lets you know you Windows version information using Toast.
 
 Creates an image with a Base64 conversion.
 
 Currently, running this script will run the below example.
 
 Function ``Export-ToastNotification`` is in-progress.
 
.EXAMPLE
 Invoke-ToastNotification -Arguments $(Get-UpdateInfo) -Image $(New-ToastImage -ImageBase64 $ImageBase64)
#> 
Param()

$ImageBase64 = "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"

#region Functions
function New-ToastImage{
    <# 
    .DESCRIPTION
    https://eddiejackson.net/wp/?p=23393
    Uses the above $ImageBase64 string, and re-constitutes it into a temporary file.
    Link above describes that process.
    So far PNG and BMP files work.
    #>
    param(
        [string]$ImageBase64
    )
    #$ImageFile = Join-Path $env:TEMP -ChildPath 'usps.png'
    $ImageFile = New-TemporaryFile | Select-Object -ExpandProperty FullName
    [byte[]]$Bytes = [convert]::FromBase64String($ImageBase64)
    [System.IO.File]::WriteAllBytes($ImageFile,$Bytes)
    $ImageFile
}

function Get-UpdateInfo{
    $WindowsQuery = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property Caption,InstallDate,Organization,Version,BuildNumber,OSArchitecture
    $WindowsReleaseId = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' | Select-Object -ExpandProperty ReleaseId
    $OfficeQuery = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail* | Select-Object DisplayName, DisplayVersion

    $Report = [ordered]@{
        Windows = [ordered]@{
            Name = $WindowsQuery.Caption
            Installed = $($WindowsQuery.InstallDate | Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Organization = $WindowsQuery.Organization
            Version = $WindowsQuery.Version
            BuildNumber = $WindowsQuery.BuildNumber
            OSArchitecture = $WindowsQuery.OSArchitecture
            ReleaseId = $WindowsReleaseId
        }
        Office = [ordered]@{
            Name = $OfficeQuery.DisplayName
            Version = $OfficeQuery.DisplayVersion
        }
    }
    $Report
}

function Invoke-ToastNotification{
    <#
    Invoke-ToastNotification -Arguments $(Get-UpdateInfo) -Image $(New-ToastImage -ImageBase64 $ImageBase64)
    #>
    param(
        [psobject]$Arguments,
        [string]$Image
    )
    $Text = @(
        "Welcome to Windows $($Arguments.Windows.ReleaseId)!",
        "$($Arguments.Windows.Name) was installed on $($Arguments.Windows.Installed) $(if($Arguments.Office.Name){ ' with '+$($Arguments.Office.Name)+' v'+$($Arguments.Office.Version) })",
        "Build $($Arguments.Windows.BuildNumber) ($($Arguments.Windows.OSArchitecture)) v$($Arguments.Windows.Version)"
    )
    New-BurntToastNotification -AppLogo $Image -Text $Text
}

function Export-ToastNotification{
    <#
    Export-ToastNotification -Arguments $(Get-UpdateInfo) -Image $(New-ToastImage -ImageBase64 $ImageBase64)
    #>
    param(
        [psobject]$Arguments,
        [string]$Image
    )
    $Text1 = New-BTText -Content "Welcome to Windows $($Arguments.Windows.ReleaseId)!"
    $Text2 = New-BTText -Content "$($Arguments.Windows.Name) was installed on $($Arguments.Windows.Installed) with $($Arguments.Office.Name) v$($Arguments.Office.Version)"
    $Audio1 = New-BTAudio -Source 'ms-winsoundevent:Notification.Looping.Alarm'
    $Binding1 = New-BTBinding -Children $Text1, $Text2
    $Visual1 = New-BTVisual -BindingGeneric $Binding1
    $Content1 = New-BTContent -Audio $Audio1 -Visual $Visual1
    $CleanContent = $Content1.GetContent().Replace('<text>{', '<text>')
    $CleanContent = $CleanContent.Replace('}</text>', '</text>')
    $CleanContent = $CleanContent.Replace('="{', '="')
    $CleanContent = $CleanContent.Replace('}" ', '" ')
    $CleanContent
}
#endregion

Invoke-ToastNotification -Arguments $(Get-UpdateInfo) -Image $(New-ToastImage -ImageBase64 $ImageBase64)
