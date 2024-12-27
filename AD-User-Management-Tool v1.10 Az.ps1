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
        Write-Host "Active Directory modulu yükləndi" -ForegroundColor Green
    } catch {
        Write-Host "Active Directory modulu yüklənə bilmədi: $_" -ForegroundColor Red
        exit 1
    }
} else {

$loginXaml = @"


<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:av="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="av"
    Title="Domain Controller Qoşulması" 
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
            <TextBlock Text="Server (IP və ya DNS):" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <TextBox x:Name="ServerTextBox" 
                     Style="{StaticResource ModernTextBox}" Text=""/>
        </StackPanel>

        <StackPanel Grid.Row="2" Margin="0,0,0,20" Grid.ColumnSpan="2">
            <TextBlock Text="İstifadəçi adı:" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <TextBox x:Name="UsernameTextBox" 
                     Style="{StaticResource ModernTextBox}"/>
        </StackPanel>

        <StackPanel Grid.Row="3" Margin="0,0,0,30" Grid.ColumnSpan="2">
            <TextBlock Text="Şifrə:" 
                      FontSize="14" 
                      FontWeight="Medium" 
                      Foreground="White" 
                      Margin="0,0,0,8"/>
            <PasswordBox x:Name="PasswordBox" 
                        Style="{StaticResource ModernPasswordBox}"/>
        </StackPanel>

        <Button Grid.Row="4" 
                x:Name="ConnectButton"
                Content="Qoşul" 
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
            
            # Əvvəlcə normal qoşulma cəhdi
            try {
                $session = New-PSSession -ComputerName $server -Credential $credential -ErrorAction Stop
            } catch {
                # IP ünvanıdırsa, TrustedHosts-a əlavə et
                if ($server -match "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$") {
                    Write-Host "Adding $server to TrustedHosts..." -ForegroundColor Yellow
                    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server -Force
                    
                    # TrustedHosts-a əlavədən sonra yenidən cəhd et
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
    Write-Host "Active Directory modulu yükləndi" -ForegroundColor Green

    $totalTimer = [System.Diagnostics.Stopwatch]::StartNew()

    $script:computerAccessCache = @{}
    $script:groupMembershipCache = @{}

    $adUsers = Get-ADUser -Filter * -Properties $script:requiredProperties | Sort-Object DisplayName
    Write-Host "Məlumatlar uğurla alındı: $($adUsers.Count) istifadəçi" -ForegroundColor Green

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
            "Parol təyin edilməyib"
        } elseif ($user.PasswordNeverExpires) {
            $script:stats.ValidPasswords++
            "Vaxtı bitmir"
        } else {
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
            $passwordAge = ((Get-Date) - $user.PasswordLastSet).Days

            if ($passwordAge -gt $maxPasswordAge) {
                $script:stats.ExpiredPasswords++
                "Vaxtı bitib"
            } else {
                $script:stats.ValidPasswords++
                "$($maxPasswordAge - $passwordAge) gün qalıb"
            }
        }

        if ($user.Enabled) {
            $script:stats.ActiveUsers++
        } else {
            $script:stats.DeactivatedUsers++
        }

           # OU-nu DistinguishedName-dən çıxarırıq
           $ou = ($user.DistinguishedName -split ",") | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
           $ou = $ou -join ", "  # Bir neçə OU-ni birləşdiririk
       
           # Qrupları alırıq
           $groups = $user.MemberOf | ForEach-Object {
               try {
                   (Get-ADGroup $_).Name
               } catch {
                   $null
               }
           } | Where-Object { $_ -ne $null }
       
                # İstifadəçini allUsers kolleksiyasına əlavə edirik
                $script:allUsers.Add([PSCustomObject]@{
                   RowNumber = $script:rowNumber
                   DisplayName = $user.DisplayName
                   SamAccountName = $user.SamAccountName
                   EmailAddress = $user.EmailAddress
                   Department = $user.Department
                   LastLogonDate = if ($user.LastLogonDate) {
                       $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm")
                   } else { "Heç vaxt" }
                   Status = if ($user.Enabled) { "Aktiv" } else { "Deaktiv" }
                   PasswordStatus = $passwordStatus
                   Groups = ($groups -join ", ")
                   Title = $user.Title
                   Phone = $user.telephoneNumber
                   Mobile = $user.mobile
                   Created = $user.WhenCreated.ToString("dd.MM.yyyy")
                   AccountStatus = if ($user.Enabled) { "Aktiv" } else { "Deaktiv" }
                   OU = $ou  # OU-ni əlavə edirik
               })
       
               $script:rowNumber++
           }

    Write-Host "Məlumatlar emal edildi" -ForegroundColor Green
    Write-Host "Ümumi istifadəçilər: $($script:stats.TotalUsers)" -ForegroundColor Cyan
    Write-Host "Aktiv istifadəçilər: $($script:stats.ActiveUsers)" -ForegroundColor Green
    Write-Host "Deaktiv hesablar: $($script:stats.DeactivatedUsers)" -ForegroundColor Yellow
    Write-Host "Parol vaxtı bitmişlər: $($script:stats.ExpiredPasswords)" -ForegroundColor Red
    Write-Host "Parol təyin edilməyənlər: $($script:stats.NoPasswords)" -ForegroundColor Magenta  
    Write-Host "Etibarlı parollar: $($script:stats.ValidPasswords)" -ForegroundColor Blue

    $totalTimer.Stop()
    Write-Host "Ümumi icra müddəti: $($totalTimer.Elapsed.TotalSeconds) saniyə" -ForegroundColor Cyan

} catch {
    Write-Host "Xəta baş verdi: $_" -ForegroundColor Red
    exit
}

$xaml = @"



<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Active Directory İstifadəçi Hesabatı" 
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
                    <TextBlock Text="Ümumi İstifadəçilər" FontWeight="SemiBold" Foreground="#4CAF50"/>
                    <TextBlock Name="TotalUsersText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#A5D6A7"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="1" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ActiveUsersPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Aktiv İstifadəçilər" FontWeight="SemiBold" Foreground="#2196F3"/>
                    <TextBlock Name="ActiveUsersText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#90CAF9"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="2" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="DeactivatedPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Deaktiv Hesablar" FontWeight="SemiBold" Foreground="#FF9800"/>
                    <TextBlock Name="DeactivatedText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#FFCC80"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="3" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ExpiredPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Parol Vaxtı Bitənlər" FontWeight="SemiBold" Foreground="#F44336"/>
                    <TextBlock Name="ExpiredPasswordText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#EF9A9A"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="4" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="NoPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Parol Təyin Edilməyənlər" FontWeight="SemiBold" Foreground="#E91E63"/>
                    <TextBlock Name="NoPasswordText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#F48FB1"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="5" Background="#212121" Margin="5" Padding="15" CornerRadius="5"
                        Name="ValidPasswordPanel" Cursor="Hand">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.3"/>
                </Border.Effect>
                <StackPanel>
                    <TextBlock Text="Parol Vaxtı Bitməyənlər" FontWeight="SemiBold" Foreground="#9C27B0"/>
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
                            <Setter Property="Text" Value="Axtarış..."/>
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
                          Header="Filtrlər" 
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
                        <CheckBox Name="ActiveFilter" Content="Aktiv" Margin="0,2"/>
                        <CheckBox Name="InactiveFilter" Content="Deaktiv" Margin="0,2"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Margin="5">
                        <TextBlock Text="Parol Status" FontWeight="Bold" Margin="0,0,0,5"/>
                        <CheckBox Name="PasswordExpiredFilter" Content="Vaxtı bitib" Margin="0,2"/>
                        <CheckBox Name="PasswordNeverExpiresFilter" Content="Vaxtı bitmir" Margin="0,2"/>
                        <CheckBox Name="NoPasswordFilter" Content="Parol təyin edilməyib" Margin="0,2"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Margin="5">
                        <TextBlock Text="Son Giriş" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ComboBox Name="LastLoginFilter" Margin="0,2">
                            <ComboBoxItem Content="Bütün"/>
                            <ComboBoxItem Content="Bu gün"/>
                            <ComboBoxItem Content="Son 7 gün"/>
                            <ComboBoxItem Content="Son 30 gün"/>
                            <ComboBoxItem Content="Heç vaxt"/>
                        </ComboBox>
                    </StackPanel>

                    <StackPanel Grid.Column="4" Margin="5" VerticalAlignment="Bottom">
                        <Button Name="ApplyFilterButton" 
                                    Content="Filtri Tətbiq Et" 
                                    Style="{StaticResource ModernButton}"
                                    Margin="0,0,0,5" Foreground="#FFF6F7F9" Background="#FF2909E4"/>
                        <Button Name="ClearFilterButton" 
                                    Content="Filtri Təmizlə" 
                                    Style="{StaticResource ModernButton}" Foreground="White" Background="#FF2909E4"/>
                    </StackPanel>
                </Grid>
            </Expander>

            <DataGrid Grid.Row="2" Name="UsersGrid" 
                          IsReadOnly="True"
                          SelectionMode="Extended" Foreground="White" BorderBrush="#FF010306" Margin="0,0,0,-29">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="№" Binding="{Binding RowNumber}" Width="100" SortDirection="Ascending">
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

                    <DataGridTextColumn Header="İstifadəçi Adı" Binding="{Binding SamAccountName}" Width="130">
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
                    <DataGridTextColumn Header="Şöbə" Binding="{Binding Department}" Width="130">
                        <DataGridTextColumn.HeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#424242"/>
                                <Setter Property="Foreground" Value="#f2f2f2"/>
                                <Setter Property="Padding" Value="10,5"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Style>
                        </DataGridTextColumn.HeaderStyle>
                    </DataGridTextColumn>

                    <DataGridTextColumn Header="Son Giriş" Binding="{Binding LastLogonDate}" Width="140">
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

                    <DataGridTextColumn Header="Parol Statusu" Binding="{Binding PasswordStatus}" Width="140">
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

                    <DataGridTextColumn Header="Qruplar" Binding="{Binding Groups}" Width="SizeToCells">
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
                        Content="Yenilə" 
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
    Write-Host "XAML yüklənməsində xəta: $_" -ForegroundColor Red
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
        $activeUsersText.Text = ($script:allUsers | Where-Object {$_.Status -eq "Aktiv"}).Count
        $deactivatedText.Text = ($script:allUsers | Where-Object {$_.Status -eq "Deaktiv"}).Count
        $expiredPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -match "bitib"}).Count
        $noPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -eq "Parol təyin edilməyib"}).Count
        $validPasswordText.Text = ($script:allUsers | Where-Object {$_.PasswordStatus -notmatch "bitib|təyin edilməyib"}).Count
    })
}

$usersGrid.ItemsSource = $script:allUsers
Update-UIStatistics

$contextMenuXaml = @"
<ContextMenu xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <MenuItem Header="İstifadəçi Detalları"/>
    <MenuItem Header="Parolu Sıfırla"/>
    <MenuItem Header="Hesabı Kilidlə/Kiliddən Çıxar"/>
</ContextMenu>
"@

try {
    $contextMenuReader = [System.IO.StringReader]::new($contextMenuXaml)
    $xmlReader = [System.Xml.XmlReader]::Create($contextMenuReader)
    $contextMenu = [Windows.Markup.XamlReader]::Load($xmlReader)
} catch {
    Write-Host "Context Menu-nun yaradılmasında xəta: $_" -ForegroundColor Red
}

function Show-UserDetails {
    param([PSCustomObject]$User)

    if ($User -eq $null) {
        [System.Windows.MessageBox]::Show(
            "Seçilmiş istifadəçi yoxdur.", 
            "Xəta",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    $detailsXaml = @"
    
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="İstifadəçi Məlumatları" 
    Height="435" Width="518"
    WindowStartupLocation="CenterOwner"
    WindowStyle="ToolWindow"
    Background="#FF1E1E1E">

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Main Content -->
        <Border Grid.Row="1" 
                BorderBrush="#e0e0e0" 
                BorderThickness="1" 
                CornerRadius="4"
                Background="#FFFFFF">
            <ScrollViewer VerticalScrollBarVisibility="Auto" Foreground="#FF333333" Background="#FF424242">
                <Grid Margin="15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Sol Panel -->
                    <StackPanel Grid.Column="0" Margin="0,0,10,0">
                        <TextBlock Text="Əsas Məlumatlar" 
                                 FontWeight="SemiBold" 
                                 Margin="0,0,0,10" Foreground="White"/>

                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Logini -->
                            <StackPanel Grid.Row="0" Margin="0,0,0,8">
                                <TextBlock Text="User Name" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding SamAccountName}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Email -->
                            <StackPanel Grid.Row="1" Margin="0,0,0,8">
                                <TextBlock Text="Email" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding EmailAddress}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Şöbə -->
                            <StackPanel Grid.Row="2" Margin="0,0,0,8">
                                <TextBlock Text="Şöbə" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding Department}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Vəzifə -->
                            <StackPanel Grid.Row="3" Margin="0,0,0,8">
                                <TextBlock Text="Vəzifə" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding Title}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Telefon -->
                            <StackPanel Grid.Row="4" Margin="0,0,0,8">
                                <TextBlock Text="Telefon" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding Phone}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Mobil -->
                            <StackPanel Grid.Row="5" Margin="0,0,0,8">
                                <TextBlock Text="Mobil" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding Mobile}" Margin="0,2" Foreground="White"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>

                    <!-- Sağ Panel -->
                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                        <TextBlock Text="Sistem Məlumatları" 
                                 FontWeight="SemiBold" 
                                 Margin="0,0,0,10" Foreground="White"/>
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Status -->
                            <StackPanel Grid.Row="0" Margin="0,0,0,8">
                                <TextBlock Text="Status" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding Status}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Son Giriş -->
                            <StackPanel Grid.Row="1" Margin="0,0,0,8">
                                <TextBlock Text="Son Giriş" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding LastLogonDate}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Parol Statusu -->
                            <StackPanel Grid.Row="2" Margin="0,0,0,8">
                                <TextBlock Text="Parol Statusu" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding PasswordStatus}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- OU -->
                            <StackPanel Grid.Row="3" Margin="0,0,0,8">
                                <TextBlock Text="OU" Foreground="White" FontSize="12"/>
                                <TextBlock Text="{Binding OU}" Margin="0,2" Foreground="White"/>
                            </StackPanel>

                            <!-- Qruplar -->
                            <StackPanel Grid.Row="4" Margin="0,0,0,8">
                                <TextBlock Text="Qruplar" Foreground="#FFFFFDFD" FontSize="12"/>
                                <TextBlock Text="{Binding Groups}" Margin="0,2" Foreground="White" TextWrapping="Wrap"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </Grid>
            </ScrollViewer>
        </Border>

        <!-- Footer Button -->
        <Button Grid.Row="2" 
                x:Name="CloseDetailsButton"
                Content="Bağla"
                Width="100"
                Height="30"
                Margin="0,15,0,0"
                Background="#FF7B0000"
                BorderBrush="#FFB3363C"
                HorizontalAlignment="Right"
                Foreground="White"/>
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
            "Seçilmiş istifadəçi yoxdur.",
            "Xəta",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    $resetPasswordXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Parolu Sıfırla"
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

        <!-- Yeni Parol -->
        <StackPanel Grid.Row="1" Margin="0,0,0,10">
            <TextBlock Text="Yeni Parol:" Margin="0,0,0,5" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
            <PasswordBox Name="NewPasswordBox" Height="30" Padding="5" FontSize="14" BorderBrush="#1a73e8" Background="White"/>
        </StackPanel>

        <!-- Parolu Təsdiqlə -->
        <StackPanel Grid.Row="2" Margin="0,0,0,10">
            <TextBlock Text="Parolu Təsdiqlə:" Margin="0,0,0,5" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
            <PasswordBox Name="ConfirmPasswordBox" Height="30" Padding="5" FontSize="14" BorderBrush="#1a73e8" Background="White"/>
        </StackPanel>

        <!-- Buttonlar -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Content="Yadda Saxla" 
                Width="90" 
                Height="30" 
                x:Name="SaveResetPasswordButton"
                Background="#FF087929" 
                Foreground="White" 
                BorderBrush="#FF087929" 
                FontWeight="Bold"
                Margin="0,0,10,0"
                Cursor="Hand"/>
            <Button Content="Bağla" 
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
                "Parol boş ola bilməz.",
                "Xəta",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
    
        $newPassPlain = ([System.Net.NetworkCredential]::new("",$newPassword)).Password
        $confirmPassPlain = ([System.Net.NetworkCredential]::new("",$confirmPassword)).Password
    
        if ($newPassPlain -ne $confirmPassPlain) {
            [System.Windows.MessageBox]::Show(
                "Parollar uyğun gəlmir.",
                "Xəta",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
    
        $adminUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    
        try {
            # Parolun tələblərə uyğun olub olmadığını yoxlayaq
            $domainPolicy = Get-ADDefaultDomainPasswordPolicy
            $passwordLengthValid = ($newPassPlain.Length -ge $domainPolicy.MinPasswordLength)
            $passwordComplexityValid = $newPassPlain -match "^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^\w\d]).+$"
    
            if (-not $passwordLengthValid) {
                [System.Windows.MessageBox]::Show(
                    "Parolun uzunluğu minimum tələblərə uyğun deyil. Zəhmət olmasa, daha uzun bir parol seçin.",
                    "Parol Xətası",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                return
            }
    
            if (-not $passwordComplexityValid) {
                [System.Windows.MessageBox]::Show(
                    "Parol komplekslik tələblərinə cavab vermir. Zəhmət olmasa, böyük və kiçik hərflər, rəqəmlər və xüsusi simvollar istifadə edin.",
                    "Parol Xətası",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                return
            }
    
            # Parolun sıfırlanması
            Set-ADAccountPassword -Identity $User.SamAccountName -NewPassword $newPassword -Reset
    
            # Log faylı üçün məlumatların yazılması
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
                "Parol uğurla sıfırlandı və əməliyyat qeyd edildi.",
                "Məlumat",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            $resetPasswordWindow.Close()
        } catch {
            # Xətanı təhlil et və müvafiq mesajı göstər
            $errorMessage = $_.Exception.Message
            [System.Windows.MessageBox]::Show(
                "Parol sıfırlanarkən xəta baş verdi: $errorMessage",
                "Xəta",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
    
            # Xətanı log faylına yazırıq
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
            "Seçilmiş istifadəçi yoxdur.",
            "Xəta",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    # Admin istifadəçi adını əldə edirik
    $adminUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    try {
        # Log qovluğu üçün əsas yol
        $logBasePath = "C:\AD-Reports-logs"
        if (-not (Test-Path -Path $logBasePath)) {
            New-Item -ItemType Directory -Path $logBasePath -Force
        }

        # Tarixi log faylı adında istifadə etmək üçün formatlayırıq
        $dateString = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $fileName = "$($User.SamAccountName)_AccountStatusChange_$dateString.log"
        $logFilePath = Join-Path -Path $logBasePath -ChildPath $fileName

        if ($User.Status -eq "Aktiv") {
            # Hesabı deaktiv edirik
            Disable-ADAccount -Identity $User.SamAccountName
            [System.Windows.MessageBox]::Show(
                "Hesab uğurla deaktiv edildi.",
                "Məlumat",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Log məlumatını yazırıq
            $logContent = "@
Action: AccountDeactivated
Target User: $($User.SamAccountName)
Performed By: $adminUser
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
@"
            Set-Content -Path $logFilePath -Value $logContent -Force
        } else {
            # Hesabı aktiv edirik
            Enable-ADAccount -Identity $User.SamAccountName
            [System.Windows.MessageBox]::Show(
                "Hesab uğurla aktiv edildi.",
                "Məlumat",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Log məlumatını yazırıq
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
        # Xətanı log faylına yazırıq
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

        # Xəta mesajını göstəririk
        [System.Windows.MessageBox]::Show(
            "Hesabın statusunu dəyişərkən xəta baş verdi: $_",
            "Xəta",
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
                "Parol təyin edilməyib"
            } elseif ($user.PasswordNeverExpires) {
                $script:stats.ValidPasswords++
                "Vaxtı bitmir"
            } else {
                $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
                $passwordAge = ((Get-Date) - $user.PasswordLastSet).Days

                if ($passwordAge -gt $maxPasswordAge) {
                    $script:stats.ExpiredPasswords++
                    "Vaxtı bitib"
                } else {
                    $script:stats.ValidPasswords++
                    "$($maxPasswordAge - $passwordAge) gün qalıb"
                }
            }

            if ($user.Enabled) {
                $script:stats.ActiveUsers++
            } else {
                $script:stats.DeactivatedUsers++
            }

            # OU-nu DistinguishedName-dən çıxarırıq
            $ou = ($user.DistinguishedName -split ",") | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
            $ou = $ou -join ", "  # Bir neçə OU-ni birləşdiririk

            # Qrupları alırıq
            $groups = $user.MemberOf | ForEach-Object {
                try {
                    (Get-ADGroup $_).Name
                } catch {
                    $null
                }
            } | Where-Object { $_ -ne $null }

            # İstifadəçini allUsers kolleksiyasına əlavə edirik
            $script:allUsers.Add([PSCustomObject]@{
                RowNumber = $script:rowNumber
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                EmailAddress = $user.EmailAddress
                Department = $user.Department
                LastLogonDate = if ($user.LastLogonDate) {
                    $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm")
                } else { "Heç vaxt" }
                Status = if ($user.Enabled) { "Aktiv" } else { "Deaktiv" }
                PasswordStatus = $passwordStatus
                Groups = ($groups -join ", ")
                Title = $user.Title
                Phone = $user.telephoneNumber
                Mobile = $user.mobile
                Created = $user.WhenCreated.ToString("dd.MM.yyyy")
                OU = $ou  # OU-ni əlavə edirik
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
            "Məlumatları yeniləyərkən xəta baş verdi: $_",
            "Xəta",
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
            ($activeFilter.IsChecked -and $item.Status -eq "Aktiv") -or
            ($inactiveFilter.IsChecked -and $item.Status -eq "Deaktiv") 
        } else { $true }

        $passwordMatch = if ($passwordExpiredFilter.IsChecked -or 
                            $passwordNeverExpiresFilter.IsChecked -or 
                            $noPasswordFilter.IsChecked) {
            ($passwordExpiredFilter.IsChecked -and $item.PasswordStatus -match "bitib") -or
            ($passwordNeverExpiresFilter.IsChecked -and $item.PasswordStatus -eq "Vaxtı bitmir") -or
            ($noPasswordFilter.IsChecked -and $item.PasswordStatus -eq "Parol təyin edilməyib")
        } else { $true }

        $lastLoginMatch = switch ($lastLoginFilter.SelectedItem.Content) {
            "Bu gün" { 
                if ($item.LastLogonDate -eq "Heç vaxt") { $false }
                else {
                    try {
                        $today = (Get-Date).ToString("dd.MM.yyyy")
                        $item.LastLogonDate.StartsWith($today)
                    } catch {
                        $false
                    }
                }
            }
            "Son 7 gün" { 
                if ($item.LastLogonDate -eq "Heç vaxt") { $false }
                else {
                    try {
                        $date = [DateTime]::ParseExact($item.LastLogonDate.Split()[0], "dd.MM.yyyy", $null)
                        $date -ge (Get-Date).AddDays(-7)
                    } catch {
                        $false
                    }
                }
            }
            "Son 30 gün" { 
                if ($item.LastLogonDate -eq "Heç vaxt") { $false }
                else {
                    try {
                        $date = [DateTime]::ParseExact($item.LastLogonDate.Split()[0], "dd.MM.yyyy", $null)
                        $date -ge (Get-Date).AddDays(-30)
                    } catch {
                        $false
                    }
                }
            }
            "Heç vaxt" { $item.LastLogonDate -eq "Heç vaxt" }
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
    if ($searchText -eq "axtarış...") { return }

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
    if ($searchBox.Text -eq "Axtarış...") {
        $searchBox.Clear()
        $searchBox.FontStyle = "Normal"
        $searchBox.Foreground = "Black"
    }
})

$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Text = "Axtarış..."
        $searchBox.FontStyle = "Italic"
        $searchBox.Foreground = "#666"
    }
})

$exportButton.Add_Click({
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV faylı (*.csv)|*.csv"
    $saveDialog.DefaultExt = ".csv"

    if ($saveDialog.ShowDialog()) {
        try {
            # DataGrid-dən məlumatları əldə edirik
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($usersGrid.ItemsSource)
            $filteredData = @()
            foreach ($item in $view) {
                $filteredData += $item
            }

            # Sütunları düzgün sırada seçirik və Azərbaycan dilində adlandırırıq
            $orderedData = $filteredData | Select-Object @{
                    Name='№'; Expression={$_.RowNumber}
                }, @{
                    Name='Ad Soyad'; Expression={$_.DisplayName}
                }, @{
                    Name='İstifadəçi Adı'; Expression={$_.SamAccountName}
                }, @{
                    Name='Email'; Expression={$_.EmailAddress}
                }, @{
                    Name='Şöbə'; Expression={$_.Department}
                }, @{
                    Name='Vəzifə'; Expression={$_.Title}
                }, @{
                    Name='Telefon'; Expression={$_.Phone}
                }, @{
                    Name='Mobil'; Expression={$_.Mobile}
                }, @{
                    Name='Son Giriş'; Expression={$_.LastLogonDate}
                }, @{
                    Name='Status'; Expression={$_.Status}
                }, @{
                    Name='Parol Statusu'; Expression={$_.PasswordStatus}
                }, @{
                    Name='OU'; Expression={$_.OU}
                }, @{
                    Name='Qruplar'; Expression={$_.Groups}
                }

            # CSV faylına export edirik
            $orderedData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

            [System.Windows.MessageBox]::Show(
                "Məlumatlar uğurla CSV faylına export edildi: $($saveDialog.FileName)", 
                "Uğurlu İxrac",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Export zamanı xəta baş verdi: $_", 
                "Xəta",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
})

$refreshButton.Add_Click({
    try {
        $filterExpander.IsExpanded = $false

        $searchBox.Text = "Axtarış..."
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
            "Xəta baş verdi: $_", 
            "Xəta", 
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
            $view.Filter = { param($item) $item.Status -eq "Aktiv" }
        }
        "Deactivated" {
            $view.Filter = { param($item) $item.Status -eq "Deaktiv" }
        }
        "PasswordExpired" {
            $view.Filter = { param($item) $item.PasswordStatus -match "bitib" }
        }
        "NoPassword" {
            $view.Filter = { param($item) $item.PasswordStatus -eq "Parol təyin edilməyib" }
        }
        "ValidPassword" {
            $view.Filter = { param($item) $item.PasswordStatus -notmatch "bitib|təyin edilməyib" }
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
        "İstifadəçi Detalları" {
            $menuItem.Add_Click({ Show-UserDetails -User $usersGrid.SelectedItem })
        }
        "Parolu Sıfırla" {
            $menuItem.Add_Click({ Reset-UserPassword -User $usersGrid.SelectedItem })
        }
        "Hesabı Kilidlə/Kiliddən Çıxar" {
            $menuItem.Add_Click({ Toggle-AccountLock -User $usersGrid.SelectedItem })
        }
    }
}

$usersGrid.ContextMenu = $contextMenu

$window.ShowDialog() 