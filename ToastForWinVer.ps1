<# 

.DESCRIPTION 
 Creates a success notification for 1809/OP+ upgrade using Toast. 

#> 
Param()

$ImageBase64 = "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"

#region Functions
function New-ToastImage{
    <# https://eddiejackson.net/wp/?p=23393 #>
    param(
        [string]$ImageBase64
    )
    $ImageFile = Join-Path $env:TEMP -ChildPath 'usps.png'
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
        "$($Arguments.Windows.Name) was installed on $($Arguments.Windows.Installed) with $($Arguments.Office.Name) v$($Arguments.Office.Version)",
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
