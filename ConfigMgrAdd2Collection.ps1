

# Set the source directory
$OS = (Get-CimInstance -ClassName Win32_OperatingSystem -Property OSArchitecture).OSArchitecture
If ($OS -eq "32-bit")
{
    $ProgramFiles = $env:ProgramFiles
}
If ($OS -eq "64-bit")
{
    $ProgramFiles = ${env:ProgramFiles(x86)}
}

$Source = "$ProgramFiles\SMSAgent\ConfigMgr Add2Collection"
#$Source = "C:\Users\tjones\Desktop\POSH Projects\ConfigMgr Add2Collection"

# Load the required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -Path "$Source\bin\MaterialDesignColors.dll"
Add-Type -Path "$Source\bin\MaterialDesignThemes.Wpf.dll"

# Load in the function library
. "$Source\bin\FunctionLibrary.ps1"

# Do PS version check
If ($PSVersionTable.PSVersion.Major -lt 5)
{
  $Content = "ConfigMgr Add2Collection cannot start because it requires PowerShell 5 or greater. Please upgrade your PowerShell version."
  New-WPFMessageBox -Content $Content -Title "Oops!" -TitleBackground Orange -TitleTextForeground Yellow -TitleFontSize 20 -TitleFontWeight Bold -BorderThickness 1 -BorderBrush Orange -Sound 'Windows Exclamation'
  Break
}

# File hash checks
$XAMLFiles = @(
    "About.xaml"
    "App.xaml"
    "Help.xaml"
    "HelpFlowDocument.xaml"
    "Settings.xaml"
)

$PSFiles = @(
    "ClassLibrary.ps1"
    "EventLibrary.ps1"
    "FunctionLibrary.ps1"
    "Show-CollectionsTreeView.ps1"
)

$Hashes = @{
    "ClassLibrary.ps1" = '24D335AF2F3C0C318AC03D5D1CFBF1BF8056A66381A9F507F14D51FCB26CBF1D'
    "EventLibrary.ps1" = '122C74D611236FDF3C3F146C2518E94705DF7B4CAE593A0088A6535427D203F7'
    "FunctionLibrary.ps1" = '8DC3F7A514312CEB76052FB349528B78B161DA8C2BA4273B4298A1EC4127801B'
    "About.xaml" = 'A9994056D48CE205A14987372304E6FDB460AE5BBE362C235E39D4A9C73DA038'
    "App.xaml" = 'B4BB3F9005EA28E60543B00C0A41BA6CF2C16873983FA0A0F3D2710E379692BE'
    "Help.xaml" = 'E52DC0F561D41A74522EBDC7660A77A89606E324BCE74769F27492EAD7797812'
    "HelpFlowDocument.xaml" = 'C11CD14C0B554C986E310AA0B2E2DB2571D6FFD59AB79D334B6E01C03EB219B1'
    "Settings.xaml" = '4E9071DAC7371F235E23FA15908D12B985938952EEE49CABA26EBEF0B33CA08B'
    "Show-CollectionsTreeView.ps1" = 'A0214C055FFE658C9BF62950BC8B5A2A64E890343F09B959D8C1499354F53267'
}

$XAMLFiles | foreach {

    If ((Get-FileHash -Path "$Source\XAML Files\$_").Hash -ne $Hashes.$_)
    {
        $Content = "One or more installation files failed a hash check. As a security measure, the installation files cannot be altered to prevent running unauthorized code. Please revert the changes or reinstall the application."
        New-WPFMessageBox -Content $Content -Title "Oops!" -TitleBackground Orange -TitleTextForeground Yellow -TitleFontSize 20 -TitleFontWeight Bold -BorderThickness 1 -BorderBrush Orange -Sound 'Windows Exclamation'
        Break
    }
}

$PSFiles | foreach {

    If ((Get-FileHash -Path "$Source\bin\$_").Hash -ne $Hashes.$_)
    {
        $Content = "One or more installation files failed a hash check. As a security measure, the installation files cannot be altered to prevent running unauthorized code. Please revert the changes or reinstall the application."
        New-WPFMessageBox -Content $Content -Title "Oops!" -TitleBackground Orange -TitleTextForeground Yellow -TitleFontSize 20 -TitleFontWeight Bold -BorderThickness 1 -BorderBrush Orange -Sound 'Windows Exclamation'
        Break
    }
}



# Define the XAML code for the main window
[XML]$Xaml = [System.IO.File]::ReadAllLines("$Source\XAML files\App.xaml") 

# Create a synchronized hash table and add the WPF window and its named elements to it
$Global:UI = [System.Collections.Hashtable]::Synchronized(@{})
$UI.Host = $Host
$UI.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $UI.$($_.Name) = $UI.Window.FindName($_.Name)
    }

# Set the window icon
$UI.Window.Icon = "$source\bin\col.png"

# Load in the code libraries
. "$Source\bin\ClassLibrary.ps1"
. "$Source\bin\EventLibrary.ps1"

# OC for common session data
$UI.SessionData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.SessionData.Add($null) # SQL Server
$UI.SessionData.Add($null) # Database
$UI.SessionData.Add($null) # Site server
$UI.SessionData.Add("False")
$UI.SessionData.Add("HKCU:\SOFTWARE\SMSAgent\ConfigMgr Add2Collection")  # Reg branch
$UI.SessionData.Add($null)
$UI.SessionData.Add([double]1.0) # current version
$UI.SessionData.Add($Source)
$UI.SessionData.Add($null) # SQLServer
$UI.SessionData.Add($null) # Database
$UI.SessionData.Add($null) # AdminUIServer
$UI.SessionData.Add($null) # Changes xml
$UI.SessionData.Add($null) # Change table
$UI.Window.DataContext = $UI.SessionData

# OC for collection info
$UI.CollectionInfo = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.WrapPanel.DataContext = $UI.CollectionInfo

# OC for Add Resource results
$UI.Results = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.Results.Add($null)
$UI.ResultsGrid.DataContext = $UI.Results

# OC for status bar
$UI.StatusBarData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.StatusBarData.Add("Idle")
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBar.DataContext = $UI.StatusBarData


# Register an event that will be called by another thread to open the TechNet page when an update is available
Register-EngineEvent -SourceIdentifier "InvokeUpdate" -Action {Start-Process "https://gallery.technet.microsoft.com/ConfigMgr-Add2Collection-fe63fe15"} | Out-Null

# If code is running in ISE, use ShowDialog() to display...
if ($psISE)
{
    $null = $UI.window.Dispatcher.InvokeAsync{$UI.window.ShowDialog()}.Wait()
}
# ...otherwise run as an application
Else
{
    # Make PowerShell Disappear
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -Name Win32ShowWindowAsync -Namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

    $app = New-Object -TypeName Windows.Application
    $app.Run($UI.Window)

}