param(
    [Parameter(Mandatory=$false)]
    [string]$Server = $env:COMPUTERNAME,

    [Parameter(Mandatory=$false)]
    [string]$Username,
    
    [Parameter(Mandatory=$false)]
    [SecureString]$Password,

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)
# Function to check and fix the status of the WinRM service
function Ensure-WinRMConfiguration {
    Write-Host "Checking WinRM service..." -ForegroundColor Yellow

    # 1. Check the status of the WinRM service
    $winrmStatus = Get-Service -Name WinRM
    if ($winrmStatus.Status -ne "Running") {
        Write-Host "WinRM service is not active, activating..." -ForegroundColor Yellow
        Start-Service -Name WinRM
        Set-Service -Name WinRM -StartupType Automatic
    } else {
        Write-Host "WinRM service is already active." -ForegroundColor Green
    }

    # 2. Check and apply WinRM configuration
    Write-Host "Checking and applying WinRM configuration..." -ForegroundColor Yellow
    winrm quickconfig -quiet

    # 3. Activate PowerShell Remoting
    Write-Host "Activating PowerShell Remoting..." -ForegroundColor Yellow
    Enable-PSRemoting -Force

    # 4. Check and open firewall rule
    $firewallRule = Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)" -ErrorAction SilentlyContinue
    if (-not $firewallRule) {
        Write-Host "Adding firewall rule for WinRM..." -ForegroundColor Yellow
        New-NetFirewallRule -Name "WinRM-HTTP" -DisplayName "Windows Remote Management (HTTP-In)" -Enabled True -Direction Inbound -Protocol TCP -LocalPort 5985 -Action Allow
    } else {
        Write-Host "WinRM firewall rule already exists." -ForegroundColor Green
    }

    # 5. Configure TrustedHosts
    $trustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
    if ($trustedHosts -ne "*") {
        Write-Host "Configuring TrustedHosts..." -ForegroundColor Yellow
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
    } else {
        Write-Host "TrustedHosts is already set to '*'." -ForegroundColor Green
    }

    # 6. Restart the WinRM service
    Write-Host "Restarting WinRM service..." -ForegroundColor Yellow
    Restart-Service WinRM

    Write-Host "`n✅ All operations completed successfully! WinRM service is active and ready for remote connections." -ForegroundColor Green
}




# Function to check and install the Active Directory Module if needed
function Ensure-ActiveDirectoryModule {
    Write-Host "Checking system readiness..." -ForegroundColor Yellow

    # Check if the Active Directory module is installed
    $moduleInstalled = Get-Module -ListAvailable -Name ActiveDirectory
    if (-not $moduleInstalled) {
        Write-Host "Active Directory module not found. Installing module..." -ForegroundColor Yellow

        # Install RSAT Active Directory module
        try {
            Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
            Write-Host "Active Directory module installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error occurred while installing Active Directory module: $_" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Active Directory module is already installed." -ForegroundColor Green
    }

    # Import the module
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "Active Directory module loaded." -ForegroundColor Green
    } catch {
        Write-Host "Error occurred while loading Active Directory module: $_" -ForegroundColor Red
        exit 1
    }
}

# Function to check the PowerShell version
function Ensure-PowerShellVersion {
    $requiredVersion = [Version]"5.1"
    if ($PSVersionTable.PSVersion -lt $requiredVersion) {
        Write-Host "PowerShell version $($PSVersionTable.PSVersion) is below the required $requiredVersion." -ForegroundColor Red
        Write-Host "You need to download [Windows Management Framework] to update PowerShell." -ForegroundColor Yellow
        exit 1
    } else {
        Write-Host "PowerShell version is at the required level ($($PSVersionTable.PSVersion))." -ForegroundColor Green
    }
}

# Function to add a server to TrustedHosts
function Ensure-TrustedHosts {
    param([string]$Server)

    $trustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
    if ($trustedHosts -notlike "*$Server*") {
        Write-Host "Adding server ($Server) to TrustedHosts..." -ForegroundColor Yellow
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$trustedHosts,$Server" -Force
        Write-Host "Server added to TrustedHosts." -ForegroundColor Green
    } else {
        Write-Host "Server ($Server) is already in TrustedHosts." -ForegroundColor Green
    }
}

# Call functions
Ensure-PowerShellVersion
Ensure-ActiveDirectoryModule


Write-Host "System successfully prepared. Now running the main script..." -ForegroundColor Cyan

# Main script starts here

function Test-IsDomainController {
    try {
        $computerSystem = Get-WmiObject Win32_ComputerSystem
        return $computerSystem.DomainRole -ge 4
    } catch {
        return $false
    }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

if (Test-IsDomainController) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "Active Directory module loaded" -ForegroundColor Green
    } catch {
        Write-Host "Active Directory module could not be loaded: $_" -ForegroundColor Red
        exit 1
    }
} else {

$loginXaml = @"


<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:av="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="av"
    Title="Domain Controller Connection" 
    Height="400" Width="440"
    WindowStartupLocation="CenterScreen"
    WindowStyle="ToolWindow"
    Background="#FF424242">
    <Window.Resources>
        <Style x:Key="ModernTextBox" TargetType="{x:Type TextBox}">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#E5E7EB"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#4F46E5"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernPasswordBox" TargetType="{x:Type PasswordBox}">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#E5E7EB"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>

        <Style x:Key="ModernButton" TargetType="{x:Type Button}">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="20,0"/>
            <Setter Property="Background" Value="#4F46E5"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="6" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#4338CA"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="30,30,30,-102">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="191*"/>
            <ColumnDefinition Width="159*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="1" Margin="0,-21,0,20" Grid.ColumnSpan="2">
            <TextBlock Text="Server (IP or DNS):" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <TextBox x:Name="ServerTextBox" 
                     Style="{StaticResource ModernTextBox}" Text=""/>
        </StackPanel>

        <StackPanel Grid.Row="2" Margin="0,0,0,20" Grid.ColumnSpan="2">
            <TextBlock Text="Username:" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <TextBox x:Name="UsernameTextBox" 
                     Style="{StaticResource ModernTextBox}"/>
        </StackPanel>

        <StackPanel Grid.Row="3" Margin="0,0,0,30" Grid.ColumnSpan="2">
            <TextBlock Text="Password:" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <PasswordBox x:Name="PasswordBox" 
                        Style="{StaticResource ModernPasswordBox}"/>
        </StackPanel>

        <Button Grid.Row="4" 
                x:Name="ConnectButton"
                Content="Connect" 
                Style="{StaticResource ModernButton}"
                Width="200" Grid.ColumnSpan="2" Height="40" Margin="75,-9,75,130" Foreground="White" Background="#FF29298E"/>
    </Grid>
</Window>



"@

    $reader = [System.IO.StringReader]::new($loginXaml)
    $xmlReader = [System.Xml.XmlReader]::Create($reader)
    $loginWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

    $serverTextBox = $loginWindow.FindName("ServerTextBox")
    $usernameTextBox = $loginWindow.FindName("UsernameTextBox")
    $passwordBox = $loginWindow.FindName("PasswordBox")
    $connectButton = $loginWindow.FindName("ConnectButton")

    $connectButton.Add_Click({
        $server = $serverTextBox.Text
        $username = $usernameTextBox.Text
        $password = $passwordBox.SecurePassword
    
        if ([string]::IsNullOrWhiteSpace($server) -or 
            [string]::IsNullOrWhiteSpace($username)) {
            [System.Windows.MessageBox]::Show(
                "Please fill in all fields",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
    
        try {
            $credential = New-Object System.Management.Automation.PSCredential($username, $password)
            
            # First attempt to connect normally
            try {
                $session = New-PSSession -ComputerName $server -Credential $credential -ErrorAction Stop
            } catch {
                # If it's an IP address, add to TrustedHosts
                if ($server -match "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$") {
                    Write-Host "Adding $server to TrustedHosts..." -ForegroundColor Yellow
                    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server -Force
                    
                    # Retry after adding to TrustedHosts
                    $session = New-PSSession -ComputerName $server -Credential $credential
                }
                else {
                    throw $_
                }
            }
            
            Invoke-Command -Session $session -ScriptBlock {
                Import-Module ActiveDirectory
            }
            Import-PSSession -Session $session -Module ActiveDirectory -AllowClobber
    
            $loginWindow.DialogResult = $true
            $loginWindow.Close()
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Connection error: $_`n`nPlease check the server name and login credentials.",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    })

    $result = $loginWindow.ShowDialog()
    if (-not $result) {
        exit
    }
}

$script:allUsers = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]

$script:requiredProperties = @(
    'DisplayName', 'EmailAddress', 'Department', 'LastLogonDate',
    'Enabled', 'PasswordLastSet', 'PasswordNeverExpires', 'Title',
    'telephoneNumber', 'mobile', 'MemberOf', 'LockedOut',
    'UserAccountControl', 'WhenCreated', 'SamAccountName'
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Active Directory module loaded" -ForegroundColor Green

    $totalTimer = [System.Diagnostics.Stopwatch]::StartNew()

    $script:computerAccessCache = @{}
    $script:groupMembershipCache = @{}

    $adUsers = Get-ADUser -Filter * -Properties $script:requiredProperties | Sort-Object DisplayName
    Write-Host "Data successfully retrieved: $($adUsers.Count) users" -ForegroundColor Green

    $script:stats = @{
        TotalUsers = $adUsers.Count
        ActiveUsers = 0
        DeactivatedUsers = 0
        ExpiredPasswords = 0
        NoPasswords = 0
        ValidPasswords = 0
    }
    $script:rowNumber = 1

    foreach ($user in $adUsers) {
        $passwordStatus = if (-not $user.PasswordLastSet) {
            $script:stats.NoPasswords++
            "No password set"
        } elseif ($user.PasswordNeverExpires) {
            $script:stats.ValidPasswords++
            "Never expires"
        } else {
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
            $passwordAge = ((Get-Date) - $user.PasswordLastSet).Days

            if ($passwordAge -gt $maxPasswordAge) {
                $script:stats.ExpiredPasswords++
                "Expired"
            } else {
                $script:stats.ValidPasswords++
                "$($maxPasswordAge - $passwordAge) days remaining"
            }
        }

        if ($user.Enabled) {
            $script:stats.ActiveUsers++
        } else {
            $script:stats.DeactivatedUsers++
        }

           # Extract OU from DistinguishedName
           $ou = ($user.DistinguishedName -split ",") | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
           $ou = $ou -join ", "  # Join multiple OUs
       
           # Retrieve groups
           $groups = $user.MemberOf | ForEach-Object {
               try {
                   (Get-ADGroup $_).Name
               } catch {
                   $null
               }
           } | Where-Object { $_ -ne $null }
       
                # Add user to allUsers collection
                $script:allUsers.Add([PSCustomObject]@{
                   RowNumber = $script:rowNumber
                   DisplayName = $user.DisplayName
                   SamAccountName = $user.SamAccountName
                   EmailAddress = $user.EmailAddress
                   Department = $user.Department
                   LastLogonDate = if ($user.LastLogonDate) {
                       $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm")
                   } else { "Never" }
                   Status = if ($user.Enabled) { "Active" } else { "Inactive" }
                   PasswordStatus = $passwordStatus
                   Groups = ($groups -join ", ")
                   Title = $user.Title
                   Phone = $user.telephoneNumber
                   Mobile = $user.mobile
                   Created = $user.WhenCreated.ToString("dd.MM.yyyy")
                   AccountStatus = if ($user.Enabled) { "Active" } else { "Inactive" }
                   OU = $ou  # Add OU
               })
       
               $script:rowNumber++
           }

    Write-Host "Data processed" -ForegroundColor Green
    Write-Host "Total users: $($script:stats.TotalUsers)" -ForegroundColor Cyan
    Write-Host "Active users: $($script:stats.ActiveUsers)" -ForegroundColor Green
    Write-Host "Deactivated accounts: $($script:stats.DeactivatedUsers)" -ForegroundColor Yellow
    Write-Host "Password expired: $($script:stats.ExpiredPasswords)" -ForegroundColor Red
    Write-Host "No password set: $($script:stats.NoPasswords)" -ForegroundColor Magenta  
    Write-Host "Valid passwords: $($script:stats.ValidPasswords)" -ForegroundColor Blue

    $totalTimer.Stop()
    Write-Host "Total execution time: $($totalTimer.Elapsed.TotalSeconds) seconds" -ForegroundColor Cyan

} catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
    exit
}

$xaml = @"



<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Active Directory User Report" 
    Height="700" Width="1282"
    WindowStartupLocation="CenterScreen"
    WindowState="Normal"
    Background="#1E1E1E" Foreground="#FF1E1E1E">

    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#BB86FC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#7B1FA2"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#4A148C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#2c2c2c"/>
            <Setter Property="Foreground" Value="#f2f2f2"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="CanUserSortColumns" Value="True"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Background" Value="#333333"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#424242"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#4A148C"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#f2f2f2"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#f2f2f2"/>
        </Style>

    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="1" Margin="0,0,0,20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="TotalUsersPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Total Users" FontWeight="SemiBold" Foreground="#4CAF50"/>
                    <TextBlock Name="TotalUsersText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#A5D6A7"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="1" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ActiveUsersPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Active Users" FontWeight="SemiBold" Foreground="#2196F3"/>
                    <TextBlock Name="ActiveUsersText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#90CAF9"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="2" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="DeactivatedPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Deactivated Accounts" FontWeight="SemiBold" Foreground="#FF9800"/>
                    <TextBlock Name="DeactivatedText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#FFCC80"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="3" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ExpiredPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Password Expired" FontWeight="SemiBold" Foreground="#F44336"/>
                    <TextBlock Name="ExpiredPasswordText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#EF9A9A"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="4" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="NoPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="No Password Set" FontWeight="SemiBold" Foreground="#E91E63"/>
                    <TextBlock Name="NoPasswordText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#F48FB1"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="5" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ValidPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Valid Passwords" FontWeight="SemiBold" Foreground="#9C27B0"/>
                    <TextBlock Name="ValidPasswordText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#CE93D8"/>
                </StackPanel>
            </Border>

        </Grid>

        <Grid Grid.Row="2" Margin="0,0,0,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#424242" 
                        CornerRadius="25" 
                        Margin="0,0,0,10"
                        Padding="15,0">
                <TextBox Name="SearchBox" 
                             Height="30"
                             Background="Transparent"
                             Foreground="#f2f2f2"  
                             VerticalContentAlignment="Center"
                             FontSize="14">
                    <TextBox.Style>
                        <Style TargetType="TextBox">
                            <Setter Property="Text" Value="Search..."/>
                            <Setter Property="Foreground" Value="#9e9e9e"/>
                            <Setter Property="FontStyle" Value="Italic"/>
                            <Style.Triggers>
                                <Trigger Property="IsFocused" Value="True">
                                    <Setter Property="Text" Value=""/>
                                    <Setter Property="FontStyle" Value="Normal"/>
                                    <Setter Property="Foreground" Value="#f2f2f2"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </TextBox.Style>
                </TextBox>
            </Border>

            <Expander Grid.Row="1" 
                          Name="FilterExpander" 
                          Header="Filters" 
                          Margin="0,0,0,10"
                          Background="#424242"
                          Foreground="#f2f2f2"  
                          BorderBrush="#616161">
                <Grid Margin="10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Column="0" Margin="5">
                        <TextBlock Text="Status" FontWeight="Bold" Margin="0,0,0,5"/>
                        <CheckBox Name="ActiveFilter" Content="Active" Margin="0,2"/>
                        <CheckBox Name="InactiveFilter" Content="Inactive" Margin="0,2"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Margin="5">
                        <TextBlock Text="Password Status" FontWeight="Bold" Margin="0,0,0,5"/>
                        <CheckBox Name="PasswordExpiredFilter" Content="Expired" Margin="0,2"/>
                        <CheckBox Name="PasswordNeverExpiresFilter" Content="Never expires" Margin="0,2"/>
                        <CheckBox Name="NoPasswordFilter" Content="No password set" Margin="0,2"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Margin="5">
                        <TextBlock Text="Last Login" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ComboBox Name="LastLoginFilter" Margin="0,2">
                            <ComboBoxItem Content="All"/>
                            <ComboBoxItem Content="Today"/>
                            <ComboBoxItem Content="Last 7 days"/>
                            <ComboBoxItem Content="Last 30 days"/>
                            <ComboBoxItem Content="Never"/>
                        </ComboBox>
                    </StackPanel>

                    <StackPanel Grid.Column="4" Margin="5" VerticalAlignment="Bottom">
                        <Button Name="ApplyFilterButton" 
                                    Content="Apply Filter" 
                                    Style="{StaticResource ModernButton}"
                                    Margin="0,0,0,5" Foreground="#FFF6F7F9" Background="#FF2909E4"/>
                        <Button Name="ClearFilterButton" 
                                    Content="Clear Filter" 
                                    Style="{StaticResource ModernButton}" Foreground="White" Background="#FF2909E4"/>
                    </StackPanel>
                </Grid>
            </Expander>

            <DataGrid Grid.Row="2" Name="UsersGrid" 
                          IsReadOnly="True"
                          SelectionMode="Extended" Foreground="White" BorderBrush="#FF010306" Margin="0,0,0,-29">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Row Number" Binding="{Binding RowNumber}" Width="100" SortDirection="Ascending">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="HorizontalContentAlignment" Value="Center"/>
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                <Setter Property="VerticalAlignment" Value="Center"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="130">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Email" Binding="{Binding EmailAddress}" Width="120">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="130">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Last Login" Binding="{Binding LastLogonDate}" Width="140">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                        <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                <Setter Property="VerticalAlignment" Value="Center"/>
                            </Style>
                        </DataGridTextColumn.ElementStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Password Status" Binding="{Binding PasswordStatus}" Width="140">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="OU" Binding="{Binding OU}" Width="120">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Groups" Binding="{Binding Groups}" Width="SizeToCells">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>

        <Grid Grid.Row="3" Margin="0,20,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="ExportButton"
                        Content="Export"
                        Style="{StaticResource ModernButton}"
                        Width="100"
                        Margin="0,0,10,0" Background="#FF2909E4"/>
                <Button Name="RefreshButton" 
                        Content="Refresh" 
                        Style="{StaticResource ModernButton}"
                        Width="100" Background="#FF2909E4"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>





"@

try {
    $reader = [System.IO.StringReader]::new($xaml)
    $xmlReader = [System.Xml.XmlReader]::Create($reader)
    $window = [Windows.Markup.XamlReader]::Load($xmlReader)
} catch {
    Write-Host "Error loading XAML: $_" -ForegroundColor Red
    exit
}

$usersGrid = $window.FindName("UsersGrid")
$searchBox = $window.FindName("SearchBox")
$totalUsersText = $window.FindName("TotalUsersText")
$activeUsersText = $window.FindName("ActiveUsersText")
$deactivatedText = $window.FindName("DeactivatedText")
$expiredPasswordText = $window.FindName("ExpiredPasswordText")
$noPasswordText = $window.FindName("NoPasswordText")
$validPasswordText = $window.FindName("ValidPasswordText")
$exportButton = $window.FindName("ExportButton")
$refreshButton = $window.FindName("RefreshButton")
$totalUsersPanel = $window.FindName("TotalUsersPanel")
$activeUsersPanel = $window.FindName("ActiveUsersPanel")
$deactivatedPanel = $window.FindName("DeactivatedPanel")
$expiredPasswordPanel = $window.FindName("ExpiredPasswordPanel")
$noPasswordPanel = $window.FindName("NoPasswordPanel")  
$validPasswordPanel = $window.FindName("ValidPasswordPanel")
$filterExpander = $window.FindName("FilterExpander")
$applyFilterButton = $window.FindName("ApplyFilterButton")
$clearFilterButton = $window.FindName("ClearFilterButton")
$lastLoginFilter = $window.FindName("LastLoginFilter")
$activeFilter = $window.FindName("ActiveFilter")
$inactiveFilter = $window.FindName("InactiveFilter")
$passwordExpiredFilter = $window.FindName("PasswordExpiredFilter")
$passwordNeverExpiresFilter = $window.FindName("PasswordNeverExpiresFilter")
$noPasswordFilter = $window.FindName("NoPasswordFilter")

function Update-UIStatistics {
    $window.Dispatcher.Invoke({
        $totalUsersText.Text = $script:allUsers.Count
        $activeUsersText.Text = ($script:allUsers | Where-Object {$_.Status -eq "Active"}).Count
        $deactivatedText.Text = ($script:allUsers | Where-Object {$_.Status -eq "Inactive"}).Count
        $expiredPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -match "Expired"}).Count
        $noPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -eq "No password set"}).Count
        $validPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -notmatch "Expired|No password set"}).Count
    })
}

$usersGrid.ItemsSource = $script:allUsers
Update-UIStatistics

$contextMenuXaml = @"
<ContextMenu xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <MenuItem Header="User Details"/>
    <MenuItem Header="Reset Password"/>
    <MenuItem Header="Lock/Unlock Account"/>
</ContextMenu>
"@

try {
    $contextMenuReader = [System.IO.StringReader]::new($contextMenuXaml)
    $xmlReader = [System.Xml.XmlReader]::Create($contextMenuReader)
    $contextMenu = [Windows.Markup.XamlReader]::Load($xmlReader)
} catch {
    Write-Host "Error creating Context Menu: $_" -ForegroundColor Red
}

function Show-UserDetails {
    param([PSCustomObject]$User)

    if ($User -eq $null) {
        [System.Windows.MessageBox]::Show(
            "No user selected.", 
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    $detailsXaml = @"
    
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Details"
    Height="400"
    Width="494"
    Background="#FF1E1E1E"
    WindowStartupLocation="CenterOwner">

    <Grid Margin="10,0,10,4">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->

        <!-- Main Content -->
        <Border Grid.Row="1" CornerRadius="8" Background="#2C2C2C" Padding="15" Margin="0,0,0,-5">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Username -->
                    <StackPanel Grid.Row="0" Margin="0,5">
                        <TextBlock Text="Username" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding SamAccountName}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Full Name -->
                    <StackPanel Grid.Row="1" Margin="0,5">
                        <TextBlock Text="Name" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding DisplayName}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Email -->
                    <StackPanel Grid.Row="2" Margin="0,5">
                        <TextBlock Text="Email" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding EmailAddress}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Department -->
                    <StackPanel Grid.Row="3" Margin="0,5">
                        <TextBlock Text="Department" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Department}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Title -->
                    <StackPanel Grid.Row="4" Margin="0,5">
                        <TextBlock Text="Title" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Title}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Phone -->
                    <StackPanel Grid.Row="5" Margin="0,5">
                        <TextBlock Text="Phone" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Phone}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Mobile -->
                    <StackPanel Grid.Row="6" Margin="0,5">
                        <TextBlock Text="Mobile" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Mobile}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Status -->
                    <StackPanel Grid.Row="7" Margin="0,5">
                        <TextBlock Text="Status" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Status}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Last Login -->
                    <StackPanel Grid.Row="8" Margin="0,5">
                        <TextBlock Text="Last Login Date" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding LastLogonDate}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- Password Status -->
                    <StackPanel Grid.Row="9" Margin="0,5">
                        <TextBlock Text="Password Status" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding PasswordStatus}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12"/>
                    </StackPanel>

                    <!-- OU -->
                    <StackPanel Grid.Row="10" Margin="0,5">
                        <TextBlock Text="Organizational Unit (OU)" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding OU}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12" TextWrapping="Wrap"/>
                    </StackPanel>

                    <!-- Groups -->
                    <StackPanel Grid.Row="11" Margin="0,5">
                        <TextBlock Text="Groups" FontSize="12" FontWeight="SemiBold" Foreground="#A5D6A7" Margin="0,0,0,3"/>
                        <TextBox Text="{Binding Groups}" IsReadOnly="True" Background="#333333" Foreground="White" Padding="3" FontSize="12" TextWrapping="Wrap"/>
                    </StackPanel>
                </Grid>
            </ScrollViewer>
        </Border>

        <!-- Footer Buttons -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="CloseDetailsButton" Content="Close" Width="90" Background="#F44336" Foreground="White" FontSize="12"/>
        </StackPanel>
    </Grid>
</Window>

"@ 
    $detailsReader = [System.IO.StringReader]::new($detailsXaml)
    $detailsXmlReader = [System.Xml.XmlReader]::Create($detailsReader)
    $detailsWindow = [Windows.Markup.XamlReader]::Load($detailsXmlReader)
    $detailsWindow.DataContext = $User
    
    $closeButton = $detailsWindow.FindName("CloseDetailsButton")
    $closeButton.Add_Click({
        $detailsWindow.Close()
    })

    $detailsWindow.ShowDialog()
}

function Reset-UserPassword {
    param([PSCustomObject]$User)

    if ($User -eq $null) {
        [System.Windows.MessageBox]::Show(
            "No user selected.",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    $resetPasswordXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Reset Password"
 Height="220" Width="300"
 WindowStartupLocation="CenterOwner"
 ResizeMode="NoResize"
 WindowStyle="ToolWindow"
 Background="#FF1E1E1E">

    <Grid Margin="10,10,10,0" VerticalAlignment="Top">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- New Password -->
        <StackPanel Grid.Row="1" Margin="0,0,0,10">
            <TextBlock Text="New Password:" Margin="0,0,0,5" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
            <PasswordBox Name="NewPasswordBox" Height="30" Padding="5" FontSize="14" BorderBrush="#1a73e8" Background="White"/>
        </StackPanel>

        <!-- Confirm Password -->
        <StackPanel Grid.Row="2" Margin="0,0,0,10">
            <TextBlock Text="Confirm Password:" Margin="0,0,0,5" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
            <PasswordBox Name="ConfirmPasswordBox" Height="30" Padding="5" FontSize="14" BorderBrush="#1a73e8" Background="White"/>
        </StackPanel>

        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Content="Save" 
                Width="90" 
                Height="30" 
                x:Name="SaveResetPasswordButton"
                Background="#FF087929" 
                Foreground="White" 
                BorderBrush="#FF087929" 
                FontWeight="Bold"
                Margin="0,0,10,0"
                Cursor="Hand"/>
            <Button Content="Close" 
                Width="90" 
                Height="30" 
                x:Name="CloseResetPasswordButton"
                Background="#FF7B0000" 
                Foreground="White" 
                BorderBrush="#FF7B0000" 
                FontWeight="Bold"
                Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $resetPasswordReader = [System.IO.StringReader]::new($resetPasswordXaml)
    $resetPasswordXmlReader = [System.Xml.XmlReader]::Create($resetPasswordReader)
    $resetPasswordWindow = [Windows.Markup.XamlReader]::Load($resetPasswordXmlReader)

    $newPasswordBox = $resetPasswordWindow.FindName("NewPasswordBox")
    $confirmPasswordBox = $resetPasswordWindow.FindName("ConfirmPasswordBox")
    $saveButton = $resetPasswordWindow.FindName("SaveResetPasswordButton")
    $closeButton = $resetPasswordWindow.FindName("CloseResetPasswordButton")

    $saveButton.Add_Click({
        $newPassword = $newPasswordBox.SecurePassword
        $confirmPassword = $confirmPasswordBox.SecurePassword
    
        if ($newPassword.Length -eq 0 -or $confirmPassword.Length -eq 0) {
            [System.Windows.MessageBox]::Show(
                "Password cannot be empty.",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
    
        $newPassPlain = ([System.Net.NetworkCredential]::new("",$newPassword)).Password
        $confirmPassPlain = ([System.Net.NetworkCredential]::new("",$confirmPassword)).Password
    
        if ($newPassPlain -ne $confirmPassPlain) {
            [System.Windows.MessageBox]::Show(
                "Passwords don't match.",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
    
        $adminUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    
        try {
            # Check if the password meets the requirements
            $domainPolicy = Get-ADDefaultDomainPasswordPolicy
            $passwordLengthValid = ($newPassPlain.Length -ge $domainPolicy.MinPasswordLength)
            $passwordComplexityValid = $newPassPlain -match "^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^\w\d]).+$"
    
            if (-not $passwordLengthValid) {
                [System.Windows.MessageBox]::Show(
                    "Password length does not meet the minimum requirements. Please choose a longer password.",
                    "Password Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                return
            }
    
            if (-not $passwordComplexityValid) {
                [System.Windows.MessageBox]::Show(
                    "Password does not meet complexity requirements. Please use uppercase and lowercase letters, numbers, and special characters.",
                    "Password Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                return
            }
    
            # Reset the password
            Set-ADAccountPassword -Identity $User.SamAccountName -NewPassword $newPassword -Reset
    
            # Write data for the log file
            $logBasePath = "C:\AD-Reports-logs"
            if (-not (Test-Path -Path $logBasePath)) {
                New-Item -ItemType Directory -Path $logBasePath -Force
            }
    
            $dateString = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $fileName = "$($User.SamAccountName)_PasswordChanged_$dateString.log"
            $logFilePath = Join-Path -Path $logBasePath -ChildPath $fileName
    
            $logContent = "@
    Action: PasswordChanged
    Target User: $($User.SamAccountName)
    Performed By: $adminUser
    Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    @"
    
            Set-Content -Path $logFilePath -Value $logContent -Force
    
            [System.Windows.MessageBox]::Show(
                "Password successfully reset and operation logged.",
                "Information",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            $resetPasswordWindow.Close()
        } catch {
            # Analyze the error and display the appropriate message
            $errorMessage = $_.Exception.Message
            [System.Windows.MessageBox]::Show(
                "Error occurred while resetting password: $errorMessage",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
    
            # Write the error to the log file
            $logBasePath = "C:\AD-Reports-logs"
            $errorLogFileName = "$($User.SamAccountName)_PasswordChangeError_$dateString.log"
            $errorLogFilePath = Join-Path -Path $logBasePath -ChildPath $errorLogFileName
    
            $errorLogContent = "@
    Action: PasswordChangeError
    Target User: $($User.SamAccountName)
    Performed By: $adminUser
    Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Error: $errorMessage
    @"
    
            Set-Content -Path $errorLogFilePath -Value $errorLogContent -Force
        }
    })
    
    $closeButton.Add_Click({
        $resetPasswordWindow.Close()
    })

    $resetPasswordWindow.ShowDialog()
}


function Toggle-AccountLock {
    param([PSCustomObject]$User)

    if ($User -eq $null) {
        [System.Windows.MessageBox]::Show(
            "No user selected.",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    # Get the admin username
    $adminUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    try {
        # Base path for the log folder
        $logBasePath = "C:\AD-Reports-logs"
        if (-not (Test-Path -Path $logBasePath)) {
            New-Item -ItemType Directory -Path $logBasePath -Force
        }

        # Format the date for use in the log file name
        $dateString = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $fileName = "$($User.SamAccountName)_AccountStatusChange_$dateString.log"
        $logFilePath = Join-Path -Path $logBasePath -ChildPath $fileName

        if ($User.Status -eq "Active") {
            # Deactivate the account
            Disable-ADAccount -Identity $User.SamAccountName
            [System.Windows.MessageBox]::Show(
                "Account successfully deactivated.",
                "Information",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Write log data
            $logContent = "@
Action: AccountDeactivated
Target User: $($User.SamAccountName)
Performed By: $adminUser
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
@"
            Set-Content -Path $logFilePath -Value $logContent -Force
        } else {
            # Activate the account
            Enable-ADAccount -Identity $User.SamAccountName
            [System.Windows.MessageBox]::Show(
                "Account successfully activated.",
                "Information",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Write log data
            $logContent = "@
Action: AccountActivated
Target User: $($User.SamAccountName)
Performed By: $adminUser
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
@"
            Set-Content -Path $logFilePath -Value $logContent -Force
        }

        Refresh-ADUserData
    } catch {
        # Write the error to the log file
        $logBasePath = "C:\AD-Reports-logs"
        $errorFileName = "$($User.SamAccountName)_AccountStatusChangeError_$dateString.log"
        $errorLogFilePath = Join-Path -Path $logBasePath -ChildPath $errorFileName

        $errorLogContent = "@
Action: AccountStatusChangeError
Target User: $($User.SamAccountName)
Performed By: $adminUser
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Error: $_
@"
        Set-Content -Path $errorLogFilePath -Value $errorLogContent -Force

        # Display the error message
        [System.Windows.MessageBox]::Show(
            "Error occurred while changing account status: $_",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Refresh-ADUserData {
    try {
        $window.Dispatcher.Invoke({
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
            $view.Filter = $null
        })

        $script:allUsers.Clear()
        $adUsers = Get-ADUser -Filter * -Properties $script:requiredProperties | Sort-Object DisplayName

        $script:stats = @{
            TotalUsers = $adUsers.Count
            ActiveUsers = 0
            DeactivatedUsers = 0
            ExpiredPasswords = 0
            NoPasswords = 0
            ValidPasswords = 0
        }
        $script:rowNumber = 1

        foreach ($user in $adUsers) {
            $passwordStatus = if (-not $user.PasswordLastSet) {
                $script:stats.NoPasswords++
                "No password set"
            } elseif ($user.PasswordNeverExpires) {
                $script:stats.ValidPasswords++
                "Never expires"
            } else {
                $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
                $passwordAge = ((Get-Date) - $user.PasswordLastSet).Days

                if ($passwordAge -gt $maxPasswordAge) {
                    $script:stats.ExpiredPasswords++
                    "Expired"
                } else {
                    $script:stats.ValidPasswords++
                    "$($maxPasswordAge - $passwordAge) days remaining"
                }
            }

            if ($user.Enabled) {
                $script:stats.ActiveUsers++
            } else {
                $script:stats.DeactivatedUsers++
            }

            # Extract OU from DistinguishedName
            $ou = ($user.DistinguishedName -split ",") | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
            $ou = $ou -join ", "  # Join multiple OUs

            # Retrieve groups
            $groups = $user.MemberOf | ForEach-Object {
                try {
                    (Get-ADGroup $_).Name
                } catch {
                    $null
                }
            } | Where-Object { $_ -ne $null }

            # Add user to allUsers collection
            $script:allUsers.Add([PSCustomObject]@{
                RowNumber = $script:rowNumber
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                EmailAddress = $user.EmailAddress
                Department = $user.Department
                LastLogonDate = if ($user.LastLogonDate) {
                    $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm")
                } else { "Never" }
                Status = if ($user.Enabled) { "Active" } else { "Inactive" }
                PasswordStatus = $passwordStatus
                Groups = ($groups -join ", ")
                Title = $user.Title
                Phone = $user.telephoneNumber
                Mobile = $user.mobile
                Created = $user.WhenCreated.ToString("dd.MM.yyyy")
                OU = $ou  # Add OU
            })

            $script:rowNumber++
        }

        $window.Dispatcher.Invoke({
            $usersGrid.ItemsSource = $script:allUsers
            $usersGrid.Items.Refresh()
            Update-UIStatistics
        })
    } catch {
        [System.Windows.MessageBox]::Show(
            "Error occurred while refreshing data: $_",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Apply-Filter {
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
    $view.Filter = {
        param($item)

        $statusMatch = if ($activeFilter.IsChecked -or $inactiveFilter.IsChecked) {
            ($activeFilter.IsChecked -and $item.Status -eq "Active") -or
            ($inactiveFilter.IsChecked -and $item.Status -eq "Inactive") 
        } else { $true }

        $passwordMatch = if ($passwordExpiredFilter.IsChecked -or 
                            $passwordNeverExpiresFilter.IsChecked -or 
                            $noPasswordFilter.IsChecked) {
            ($passwordExpiredFilter.IsChecked -and $item.PasswordStatus -match "Expired") -or
            ($passwordNeverExpiresFilter.IsChecked -and $item.PasswordStatus -eq "Never expires") -or
            ($noPasswordFilter.IsChecked -and $item.PasswordStatus -eq "No password set")
        } else { $true }

        $lastLoginMatch = switch ($lastLoginFilter.SelectedItem.Content) {
            "Today" { 
                if ($item.LastLogonDate -eq "Never") { $false }
                else {
                    try {
                        $today = (Get-Date).ToString("dd.MM.yyyy")
                        $item.LastLogonDate.StartsWith($today)
                    } catch {
                        $false
                    }
                }
            }
            "Last 7 days" { 
                if ($item.LastLogonDate -eq "Never") { $false }
                else {
                    try {
                        $date = [DateTime]::ParseExact($item.LastLogonDate.Split()[0], "dd.MM.yyyy", $null)
                        $date -ge (Get-Date).AddDays(-7)
                    } catch {
                        $false
                    }
                }
            }
            "Last 30 days" { 
                if ($item.LastLogonDate -eq "Never") { $false }
                else {
                    try {
                        $date = [DateTime]::ParseExact($item.LastLogonDate.Split()[0], "dd.MM.yyyy", $null)
                        $date -ge (Get-Date).AddDays(-30)
                    } catch {
                        $false
                    }
                }
            }
            "Never" { $item.LastLogonDate -eq "Never" }
            default { $true }
        }

        return $statusMatch -and $passwordMatch -and $lastLoginMatch
    }

    $rowNumber = 1
    $filteredItems = $view.GetEnumerator() | ForEach-Object { $_ }
    foreach ($item in $filteredItems) {
        $item.RowNumber = $rowNumber++
    }

    Update-UIStatistics
}

function Clear-Filter {
    $activeFilter.IsChecked = $false
    $inactiveFilter.IsChecked = $false
    $passwordExpiredFilter.IsChecked = $false
    $passwordNeverExpiresFilter.IsChecked = $false
    $noPasswordFilter.IsChecked = $false
    $lastLoginFilter.SelectedIndex = 0

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
    $view.Filter = $null

    $rowNumber = 1
    foreach ($item in $script:allUsers) {
        $item.RowNumber = $rowNumber++
    }

    Update-UIStatistics
}

$applyFilterButton.Add_Click({
    Apply-Filter
})

$clearFilterButton.Add_Click({
    Clear-Filter
})

$searchBox.Add_TextChanged({
    $searchText = $searchBox.Text.ToLower()
    if ($searchText -eq "search...") { return }

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
    $view.Filter = {
        param($item)
        ($item.DisplayName -like "*$searchText*") -or
        ($item.SamAccountName -like "*$searchText*") -or
        ($item.EmailAddress -like "*$searchText*") -or
        ($item.Department -like "*$searchText*")
    }

    $rowNumber = 1
    $filteredItems = $view.GetEnumerator() | ForEach-Object { $_ }
    foreach ($item in $filteredItems) {
        $item.RowNumber = $rowNumber++
    }

    Update-UIStatistics
})

$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq "Search...") {
        $searchBox.Clear()
        $searchBox.FontStyle = "Normal"
        $searchBox.Foreground = "Black"
    }
})

$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Text = "Search..."
        $searchBox.FontStyle = "Italic"
        $searchBox.Foreground = "#666"
    }
})

$exportButton.Add_Click({
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV file (*.csv)|*.csv"
    $saveDialog.DefaultExt = ".csv"

    if ($saveDialog.ShowDialog()) {
        try {
            # Retrieve data from DataGrid
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
            $filteredData = @()
            foreach ($item in $view) {
                $filteredData += $item
            }

            # Select columns in the correct order and name them in English
            $orderedData = $filteredData | Select-Object @{
                    Name='Row Number'; Expression={$_.RowNumber}
                }, @{
                    Name='Display Name'; Expression={$_.DisplayName}
                }, @{
                    Name='Username'; Expression={$_.SamAccountName}
                }, @{
                    Name='Email'; Expression={$_.EmailAddress}
                }, @{
                    Name='Department'; Expression={$_.Department}
                }, @{
                    Name='Title'; Expression={$_.Title}
                }, @{
                    Name='Phone'; Expression={$_.Phone}
                }, @{
                    Name='Mobile'; Expression={$_.Mobile}
                }, @{
                    Name='Last Login'; Expression={$_.LastLogonDate}
                }, @{
                    Name='Status'; Expression={$_.Status}
                }, @{
                    Name='Password Status'; Expression={$_.PasswordStatus}
                }, @{
                    Name='OU'; Expression={$_.OU}
                }, @{
                    Name='Groups'; Expression={$_.Groups}
                }

            # Export to CSV file
            $orderedData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

            [System.Windows.MessageBox]::Show(
                "Successfully exported to CSV file: $($saveDialog.FileName)", 
                "Export Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Export Error: $_", 
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
})

$refreshButton.Add_Click({
    try {
        $filterExpander.IsExpanded = $false

        $searchBox.Text = "Search..."
        $searchBox.FontStyle = "Italic"
        $searchBox.Foreground = "#666"

        Clear-Filter
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
        $view.Filter = $null

        $usersGrid.Items.SortDescriptions.Clear()
        $sortDescription = New-Object System.ComponentModel.SortDescription("RowNumber", "Ascending")
        $usersGrid.Items.SortDescriptions.Add($sortDescription)

        $rowNumber = 1
        foreach ($item in $script:allUsers) {
            $item.RowNumber = $rowNumber++
        }

        $usersGrid.Items.Refresh()
        Update-UIStatistics
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Error occurred: $_", 
            "Error", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

function Set-Filter {
    param([string]$FilterType)
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
    switch ($FilterType) {
        "Total" {
            $view.Filter = $null
        }
        "Active" {
            $view.Filter = { param($item) $item.Status -eq "Active" }
        }
        "Deactivated" {
            $view.Filter = { param($item) $item.Status -eq "Inactive" }
        }
        "PasswordExpired" {
            $view.Filter = { param($item) $item.PasswordStatus -match "Expired" }
        }
        "NoPassword" {
            $view.Filter = { param($item) $item.PasswordStatus -eq "No password set" }
        }
        "ValidPassword" {
            $view.Filter = { param($item) $item.PasswordStatus -notmatch "Expired|No password set" }
        }
    }

    $rowNumber = 1
    $filteredItems = $view.GetEnumerator() | ForEach-Object { $_ }
    foreach ($item in $filteredItems) {
        $item.RowNumber = $rowNumber++
    }

    Update-UIStatistics
}

$totalUsersPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "Total"
})

$activeUsersPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "Active"
})

$deactivatedPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "Deactivated"
})

$expiredPasswordPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "PasswordExpired"
})

$noPasswordPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "NoPassword"
})

$validPasswordPanel.Add_MouseLeftButtonDown({
    Set-Filter -FilterType "ValidPassword"
})

$usersGrid.Add_ContextMenuOpening({
    if ($usersGrid.SelectedItem -eq $null) {
        $usersGrid.ContextMenu.IsEnabled = $false
    } else {
        $usersGrid.ContextMenu.IsEnabled = $true
    }
})

foreach ($menuItem in $contextMenu.Items) {
    switch ($menuItem.Header) {
        "User Details" {
            $menuItem.Add_Click({ Show-UserDetails -User $usersGrid.SelectedItem })
        }
        "Reset Password" {
            $menuItem.Add_Click({ Reset-UserPassword -User $usersGrid.SelectedItem })
        }
        "Lock/Unlock Account" {
            $menuItem.Add_Click({ Toggle-AccountLock -User $usersGrid.SelectedItem })
        }
    }
}

$usersGrid.ContextMenu = $contextMenu

$window.ShowDialog()
