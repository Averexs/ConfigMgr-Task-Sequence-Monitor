#==========================================================================
# GLOBAL CONFIGURATION VARIABLES
#==========================================================================
$SiteCode = "XXX" # Site code 
$ProviderMachineName = "PRIMARY_SITE_SERVER_FQDN" # SMS Provider machine name
$DBName = (Get-CimInstance -ComputerName $ProviderMachineName -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_SCI_SiteDefinition).SQLDatabaseName
$SQLServerName =  (Get-CimInstance -ComputerName $ProviderMachineName -Namespace "ROOT\SMS\site_$SiteCode" -ClassName SMS_SCI_SiteDefinition).SQLServerName
$GlobalSQLServer = $SQLServerName
$GlobalDatabase  = $DBName

# Reports output directory (defaults to <ScriptRoot>\reports, or set your own custom path)
$ScriptRootPath = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
$ReportsPath = Join-Path $ScriptRootPath "reports" 
$GlobalReportsFolder = $ReportsPath
#==========================================================================

#region Add Assemblies
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

# Win32 Icon extraction (guarded to prevent duplicate compilation in ISE/VS Code)
if (-not ("System.IconExtractor" -as [type]))
{
    $code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
namespace System
{
    public class IconExtractor
    {
        public static Icon Extract(string file, int number, bool largeIcon)
        {
            IntPtr large;
            IntPtr small;
            ExtractIconEx(file, number, out large, out small, 1);
            try
            {
                return Icon.FromHandle(largeIcon ? large : small);
            }
            catch
            {
                return null;
            }
        }
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
    }
}
"@
    Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing
}
#endregion

#region GUI and Variables
### Main Window (Native WPF) ###
[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ConfigMgr Task Sequence Monitor" Height="685" Width="1347" WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip">
    <Grid VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="0,0,2,0" Width="auto">
        <Grid.RowDefinitions>
            <RowDefinition Height="503*" />
            <RowDefinition Height="10" />
            <RowDefinition Height="141*" />
        </Grid.RowDefinitions>
        <GroupBox Header="ConfigMgr" HorizontalAlignment="Stretch" Margin="10,10,0,0" VerticalAlignment="Top" Height="115" Width="1317">
            <Grid HorizontalAlignment="Left" Height="100" Margin="0,0,-2,-1" VerticalAlignment="Top" Width="1299">
                <Label Content="Task Sequence:" HorizontalAlignment="Left" Margin="0,7,0,0" VerticalAlignment="Top" Width="96"/>
                <ComboBox x:Name="TaskSequence" HorizontalAlignment="Left" Margin="96,7,0,0" VerticalAlignment="Top" Width="356" Height="26"/>
                <Label Content="Time Period &#xD;&#xA;(Hours):" HorizontalAlignment="Left" Margin="457,0,0,0" VerticalAlignment="Top" Width="73" Height="46"/>
                <TextBox x:Name="TimePeriod" HorizontalAlignment="Left" Height="26" Margin="535,7,0,0" TextWrapping="Wrap" Text="24" VerticalAlignment="Top" Width="48" TextAlignment="Center" VerticalContentAlignment="Center"/>
                <Label Content="Errors &#xD;&#xA;Only:" HorizontalAlignment="Left" Margin="855,0,0,0" VerticalAlignment="Top" Width="41" Height="46"/>
                <CheckBox x:Name="ErrorsOnly" Content="" HorizontalAlignment="Left" Margin="901,18,0,0" VerticalAlignment="Top"/>
                <Label Content="ComputerName:" HorizontalAlignment="Left" Margin="588,8,0,0" VerticalAlignment="Top"/>
                <ComboBox x:Name="ComputerName" HorizontalAlignment="Left" Margin="690,7,0,0" VerticalAlignment="Top" Width="160" Height="26"/>
                <Label Content="Refresh Period &#xD;&#xA;(Minutes):" HorizontalAlignment="Left" Margin="1073,0,0,0" VerticalAlignment="Top" Height="46"/>
                <TextBox x:Name="RefreshPeriod" HorizontalAlignment="Left" Height="26" Margin="1165,8,0,0" TextWrapping="Wrap" Text="1" VerticalAlignment="Top" Width="33" TextAlignment="Center" VerticalContentAlignment="Center"/>
                <Button x:Name="RefreshNow" Content="Refresh Now!" HorizontalAlignment="Left" Margin="1204,7,0,0" VerticalAlignment="Top" Width="95" Height="29"/>
                <Button x:Name="SettingsButton" Content="Settings" HorizontalAlignment="Left" Margin="1100,45,0,0" VerticalAlignment="Top" Width="95" Height="29"/>
                <Button x:Name="ReportButton" Content="Generate &#xA;  Report" HorizontalAlignment="Left" Margin="1204,45,0,0" VerticalAlignment="Top" Width="95" Height="38"/>
                <Label Content="Error Count:" HorizontalAlignment="Left" Margin="921,10,0,0" VerticalAlignment="Top"/>
                <TextBox x:Name="ErrorCount" HorizontalAlignment="Left" Height="26" Margin="1000,8,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="37" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" IsReadOnly="True"/>
            </Grid>
        </GroupBox>

        <DataGrid x:Name="DataGrid" AutoGenerateColumns="False" HorizontalAlignment="Stretch" Margin="10,135,10,0" VerticalAlignment="Stretch" Height="Auto" IsReadOnly="True" HorizontalGridLinesBrush="#FF297566" VerticalGridLinesBrush="#FF489183">
            <DataGrid.Columns>
                <DataGridTemplateColumn Width="SizeToCells">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Image Source="{Binding Path=Icon}" Width="15" Height="15" />
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="ComputerName" Binding="{Binding Path=ComputerName}" />
                <DataGridTextColumn Header="GUID" Binding="{Binding Path=GUID}" Visibility="Hidden"/>
                <DataGridTextColumn Header="ExecutionTime" Binding="{Binding Path=ExecutionTime}" />
                <DataGridTextColumn Header="Step" Binding="{Binding Path=Step}" />
                <DataGridTextColumn Header="ActionName" Binding="{Binding Path=ActionName}" />
                <DataGridTextColumn Header="GroupName" Binding="{Binding Path=GroupName}" />
                <DataGridTextColumn Header="LastStatusMsgName" Binding="{Binding Path=LastStatusMsgName}" />
                <DataGridTextColumn Header="ExitCode" Binding="{Binding Path=ExitCode}"/>
                <DataGridTextColumn Header="Record" Binding="{Binding Path=Record}" Visibility="Hidden"/>
            </DataGrid.Columns>
        </DataGrid>
        <GridSplitter Grid.Row="1" HorizontalAlignment="Stretch" Width="auto" Height="8" Background="White" ToolTip="Resize" />
        <TextBox x:Name="ActionOutput" Grid.Row="2" HorizontalAlignment="Stretch" Margin="10,26,10,10" TextWrapping="Wrap" VerticalAlignment="Stretch" ScrollViewer.VerticalScrollBarVisibility="Auto" IsReadOnly="True"/>
        <Label Content="Action Output:" Grid.Row="2" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="26" Width="88"/>
    </Grid>
</Window>
"@

$hash = [hashtable]::Synchronized(@{})
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml)
$hash.Window = [Windows.Markup.XamlReader]::Load( $reader )
$global:PSInstances = @()
$Global:Timezones = @()
$hash.TaskSequence = $hash.Window.FindName('TaskSequence')
$hash.TimePeriod = $hash.Window.FindName('TimePeriod')
$hash.ErrorsOnly = $hash.Window.FindName('ErrorsOnly')
$hash.ComputerName = $hash.Window.FindName('ComputerName')
$hash.RefreshPeriod = $hash.Window.FindName('RefreshPeriod')
$hash.RefreshNow = $hash.Window.FindName('RefreshNow')
$hash.DataGrid = $hash.Window.FindName('DataGrid')
$hash.ActionOutput = $hash.Window.FindName('ActionOutput')
$hash.DeploymentStatus = $hash.Window.FindName('DeploymentStatus')
$hash.CurrentStep = $hash.Window.FindName('CurrentStep')
$hash.StepName = $hash.Window.FindName('StepName')
$hash.PercentComplete = $hash.Window.FindName('PercentComplete')
$hash.SettingsButton = $hash.Window.FindName('SettingsButton')
$hash.ReportButton = $hash.Window.FindName('ReportButton')
$hash.ErrorCount = $hash.Window.FindName('ErrorCount')
$hash.DeploymentStatusLabel = $hash.Window.FindName('DeploymentStatusLabel')
$hash.CurrentStepLabel = $hash.Window.FindName('CurrentStepLabel')
$hash.StepNameLabel = $hash.Window.FindName('StepNameLabel')
$hash.PercentCompleteLabel = $hash.Window.FindName('PercentCompleteLabel')
$hash.StartLabel = $hash.Window.FindName('StartLabel')
$hash.EndLabel = $hash.Window.FindName('EndLabel')
$hash.ElapsedLabel = $hash.Window.FindName('ElapsedLabel')
$hash.ProgressLabel = $hash.Window.FindName('ProgressLabel')
$hash.ProgressBar = $hash.Window.FindName('ProgressBar')

### Settings Window (Native WPF) ###
[xml]$xaml2 = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Settings" Height="290" Width="520.986" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid>
        <TabControl HorizontalAlignment="Left" Height="245" Margin="10,10,0,0" VerticalAlignment="Top" Width="497">
            <TabItem x:Name="SettingsTab" Header="Settings">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-2">
                    <Label Content="SQL Server:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
                    <Label Content="Database:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="SQLServer" HorizontalAlignment="Left" Height="23" Margin="140,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="206" VerticalContentAlignment="Center"/>
                    <TextBox x:Name="Database" HorizontalAlignment="Left" Height="23" Margin="140,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="87" VerticalContentAlignment="Center"/>
                    <Button x:Name="ConnectSQL" Content="Connect SQL" HorizontalAlignment="Left" Margin="381,10,0,0" VerticalAlignment="Top" Width="100" Height="29"/>
                    <Label x:Name="Runasadmin" Content="Note:  Please run the application as administrator to save these&#xD;&#xA;settings to the registry!" HorizontalAlignment="Left" Margin="140,125,0,0" VerticalAlignment="Top" Height="43" Visibility="Hidden"/>
                    <Label Content="Display Date/Time in:" HorizontalAlignment="Left" Margin="10,103,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="DTFormat" HorizontalAlignment="Left" Margin="140,103,0,0" VerticalAlignment="Top" Width="206"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="ReportTab" Header="Summary Report">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Task Sequence:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
                    <Label Content="Start Date:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
                    <Label Content="End Date:" HorizontalAlignment="Left" Margin="10,72,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="TSList" HorizontalAlignment="Left" Margin="106,14,0,0" VerticalAlignment="Top" Width="375" IsReadOnly="True"/>
                    <DatePicker x:Name="StartDate" HorizontalAlignment="Left" Margin="106,43,0,0" VerticalAlignment="Top" Width="114"/>
                    <DatePicker x:Name="EndDate" HorizontalAlignment="Left" Margin="106,74,0,0" VerticalAlignment="Top" Width="114"/>

                    <Label Content="Report Format:" HorizontalAlignment="Left" Margin="10,105,0,0" VerticalAlignment="Top"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="106,108,0,0" VerticalAlignment="Top">
                        <RadioButton x:Name="ReportFormatHTML" Content="HTML Report" IsChecked="True" GroupName="ReportFormatGroup" Margin="0,0,15,0"/>
                        <RadioButton x:Name="ReportFormatCSV" Content="CSV Report" GroupName="ReportFormatGroup"/>
                    </StackPanel>

                    <Button x:Name="GenerateReport" Content="Generate Report" HorizontalAlignment="Left" Margin="10,145,0,0" VerticalAlignment="Top" Width="118" Height="30"/>
                    <Label x:Name="Working" Content="" HorizontalAlignment="Left" Margin="146,145,0,0" VerticalAlignment="Top" Height="30" VerticalContentAlignment="Center" FontStyle="Italic"/>
                    <Label Content="Create a report summarizing all task sequence&#xD;&#xA;steps executed in a given time period.&#xD;&#xA;Files export to configured reports folder." HorizontalAlignment="Left" Margin="243,43,0,0" VerticalAlignment="Top" Width="240" Height="65"/>
                    <ProgressBar x:Name="ReportProgress" HorizontalAlignment="Left" Height="20" Margin="224,150,0,0" VerticalAlignment="Top" Width="238" Minimum="0" Maximum="100" Visibility="Hidden"/>
                </Grid>
            </TabItem>
            <TabItem Header="About" HorizontalAlignment="Left" VerticalAlignment="Top">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-4">
                    <RichTextBox HorizontalAlignment="Left" Height="171" VerticalAlignment="Top" Width="491" IsDocumentEnabled="True" IsReadOnly="True" IsReadOnlyCaretVisible="True">
                        <FlowDocument>
                            <Paragraph>
                                <Run Text="ConfigMgr Task Sequence Monitor" FontFamily="Calibri" FontSize="18"/>
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="This application enables you to monitor or review task sequence executions in Configuration Manager."/>
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="Original tool created by Trevor Jones." />
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="Rewritten as a standalone script by Kip Bartholomew." />
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="Version 2.0 (Standalone Script Edition)" />
                            </Paragraph>
                        </FlowDocument>
                    </RichTextBox>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
$hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
$hash.SQLServer = $hash.Window2.FindName('SQLServer')
$hash.Database = $hash.Window2.FindName('Database')
$hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
$hash.TSList = $hash.Window2.FindName('TSList')
$hash.StartDate = $hash.Window2.FindName('StartDate')
$hash.EndDate = $hash.Window2.FindName('EndDate')
$hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
$hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
$hash.ReportTab = $hash.Window2.FindName('ReportTab')
$hash.Tabs = $hash.Window2.FindName('Tabs')
$hash.Working = $hash.Window2.FindName('Working')
$hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
$hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
$hash.DTFormat = $hash.Window2.FindName('DTFormat')
$hash.ReportFormatHTML = $hash.Window2.FindName('ReportFormatHTML')
$hash.ReportFormatCSV  = $hash.Window2.FindName('ReportFormatCSV')

# Pre-populate defaults from global configuration block
$hash.SQLServer.Text = $GlobalSQLServer
$hash.Database.Text  = $GlobalDatabase
$script:SQLServer = $hash.SQLServer.Text
$Script:Database  = $hash.Database.Text
#endregion

#region Icons and Runspacepool
$icons = @()
$global:greentickiconpath = "$env:temp\GreenTick.bmp"
$icons += $greentickiconpath 
$global:redcrossiconpath = "$env:temp\RedCross.bmp"
$icons += $redcrossiconpath

if (!(Test-Path $greentickiconpath))
{
    $global:greentickicon = [System.IconExtractor]::Extract('comres.dll',8,$true).ToBitmap()
    $greentickicon.save("$greentickiconpath")
}
if (!(Test-Path $redcrossiconpath))
{
    $global:redcrossicon = [System.IconExtractor]::Extract('comres.dll',10,$true).ToBitmap()
    $redcrossicon.save("$redcrossiconpath")
}

$script:RunspacePool = [runspacefactory]::CreateRunspacePool()
$RunspacePool.ApartmentState = 'STA'
$RunspacePool.ThreadOptions = 'ReUseThread'
$RunspacePool.Open()
#endregion

#region Functions
Function Get-DateTimeFormat 
{
    if ([System.TimeZone]::CurrentTimeZone.IsDaylightSavingTime($(Get-Date)))
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.DaylightName
    }
    Else 
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    }
    $obj = New-Object -TypeName psobject -Property @{ TimeZone = 'UTC' }
    $Global:Timezones = [Array]$Timezones + $obj
    $obj = New-Object -TypeName psobject -Property @{ TimeZone = $TimeZone }
    $Global:Timezones = [Array]$Timezones + $obj
}

Function Get-TaskSequenceList 
{
    $script:SQLServer = $hash.SQLServer.Text
    $Script:Database = $hash.Database.Text
    if ([string]::IsNullOrWhiteSpace($SQLServer) -or $SQLServer -eq '<SQLServer\Instance>')
    {
        $hash.ActionOutput.Text = 'No SQL Server defined. Click Settings, and set your SQL Server parameters.'
        return
    }
    try
    {
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        $hash.ActionOutput.Text = "Connected to SQL Server ($SQLServer). Select a Task Sequence to begin."
    }
    catch 
    {
        $hash.ActionOutput.Text = "[ERROR] Could not connect to SQL Server database ($SQLServer - $Database): $($_.Exception.Message)"
        return
    }
    $Query = "
        SELECT DISTINCT summ.SoftwareName AS 'Task Sequence'
        FROM vDeploymentSummary summ
        WHERE (summ.FeatureType=7)
        ORDER BY summ.SoftwareName
    "
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $result = $command.ExecuteReader()
    $table = New-Object -TypeName 'System.Data.DataTable'
    $table.Load($result)
    $connection.Close()
    $global:Views = @()
    Foreach ($Row in $table.Rows)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'TS' -Value $Row.'Task Sequence'
        $global:Views = [Array]$Views + $obj
    }
    $hash.Window.Dispatcher.Invoke([action]{
        $hash.TaskSequence.ItemsSource = [Array]$Views.TS
    })
}

Function Get-TaskSequenceData 
{
    param ($hash,$RunspacePool)
    $code = 
    {
        param($hash,$SQLServer,$Database,$TimePeriod,$ErrorsOnly,$ComputerName,$TS,$MDTIntegrated,$URL,$DTFormat)
        $hash.Window.Dispatcher.Invoke([action]{
            $hash.ActionOutput.Text = 'Retrieving data...'
            $hash.DataGrid.ItemsSource = ''
        })
        if ($MyGUID) { Remove-Variable -Name MyGuid }
        if ($Unknowns) { Remove-Variable -Name Unknowns }
        
        $ExitFilter = if ($ErrorsOnly -eq 'True' -or $ErrorsOnly -eq $true) { "and ExitCode <> 0" } else { "" }
        $SQLComputerName = if ($ComputerName -eq '-All-' -or [string]::IsNullOrWhiteSpace($ComputerName)) { '%' } else { $ComputerName }
        $greentickiconpath = "$env:temp\GreenTick.bmp"
        $redcrossiconpath = "$env:temp\RedCross.bmp"
        try
        {
            $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
            $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
            $connection.ConnectionString = $connectionString
            $connection.Open()
        }
        catch 
        {
            $MyError = $_.Exception.Message
            $hash.Window.Dispatcher.Invoke([action]{
                $hash.ActionOutput.Text = "[ERROR] Could not connect to SQL Server database! $MyError"
            })
            return
        }

        if ($MDTIntegrated -eq 'True' -or $MDTIntegrated -eq $true)
        {
            $Query = "
                Select Distinct Name0,
                SMBIOS_GUID0 as 'GUID'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                $ExitFilter
                ORDER BY Name0 Desc
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $result = $command.ExecuteReader()
            $table = New-Object -TypeName 'System.Data.DataTable'
            $table.Load($result)
            $UnknownComputers = @()
            Foreach ($Row in $table.Rows | Where-Object { $_.Name0 -eq 'Unknown' })
            {
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
                $UnknownComputers += $obj
            }
            if ($UnknownComputers.Count -ge 1 -and !([string]::IsNullOrWhiteSpace($URL)))
            {
                $URL1 = $URL.Replace('Computers','ComputerIdentities')
                function GetMDTIDs 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            ID         = $($property.ID.'#text')
                            Identifier = $($property.Identifier)
                        }
                    } 
                }
                $MDTIDs = GetMDTIDs -URL $URL1 | Where-Object { $_.Identifier -like '*-*' } | Sort-Object -Property ID
                $MDTComputerIDs = @()
                Foreach ($Computer in $UnknownComputers)
                {
                    $MDTComputerID = $MDTIDs | Where-Object { $_.Identifier -eq $Computer.GUID } | Select-Object -Property ID, Identifier 
                    $MDTComputerIDs += $MDTComputerID
                }
                function GetMDTComputerNames 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            Name = $($property.Name)
                            ID   = $($property.ID.'#text')
                        } 
                    } 
                } 
                $MDTComputers = GetMDTComputerNames -URL $URL | Sort-Object -Property ID
                $ResolvedComputerNames = @()
                Foreach ($MDTComputerID in $MDTComputerIDs)
                {
                    $MDTComputerName = $MDTComputers | Where-Object { $_.ID -eq $MDTComputerID.ID } | Select-Object -ExpandProperty Name
                    $GUID = $MDTIDs | Where-Object { $_.ID -eq $MDTComputerID.ID } | Select-Object -ExpandProperty Identifier
                    $obj = New-Object -TypeName PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ComputerName -Value $MDTComputerName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name GUID -Value $GUID
                    $ResolvedComputerNames += $obj
                }
            }
            foreach ($Computer in $ResolvedComputerNames)
            {
                if ($ComputerName -eq $Computer.ComputerName) { $MyGUID = $Computer.GUID }
            }
        }

        if ($MyGUID)
        {
            $Query = "
                Select Distinct Name0 as 'Computer Name',
                sys.SMBIOS_GUID0 as 'GUID',
                Name as 'Task Sequence',
                ExecutionTime,
                Step,
                ActionName,
                GroupName,
                tes.LastStatusMsgName,
                ExitCode,
                ActionOutput
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and sys.SMBIOS_GUID0 = '$MyGUID'
                $ExitFilter
                ORDER BY ExecutionTime Desc
            "
            $ErrQuery = "
                Select Count(Name0) as 'Count'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and sys.SMBIOS_GUID0 = '$MyGUID'
                and ExitCode <> 0
            "
        }
        else
        {
            $Query = "
                Select Distinct Name0 as 'Computer Name',
                sys.SMBIOS_GUID0 as 'GUID',
                Name as 'Task Sequence',
                ExecutionTime,
                Step,
                ActionName,
                GroupName,
                tes.LastStatusMsgName,
                ExitCode,
                ActionOutput
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and Name0 like '$SQLComputerName'
                $ExitFilter
                ORDER BY ExecutionTime Desc
            "
            $ErrQuery = "
                Select Count(Name0) as 'Count'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and Name0 like '$SQLComputerName'
                and ExitCode <> 0
            "
        }

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)

        $command = $connection.CreateCommand()
        $command.CommandText = $ErrQuery
        $erresult = $command.ExecuteReader()
        $errtable = New-Object -TypeName 'System.Data.DataTable'
        $errtable.Load($erresult)
        $connection.Close()

        if ($table.rows.Count -lt 1)
        {
            $hash.Window.Dispatcher.Invoke([action]{
                $hash.ActionOutput.Text = 'No results found.'
            })
            return
        }
        $global:Results = @()
        $i = 0
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            $i++
            $iconVal = if ($Row.ExitCode -eq 0) { $greentickiconpath } else { $redcrossiconpath }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $iconVal
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.'Computer Name'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.'GUID'
            if ($DTFormat -eq 'UTC')
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $Row.'ExecutionTime'
            }
            Else 
            {
                $extime = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($Row.'ExecutionTime' | Get-Date))
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $extime
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Step' -Value $Row.'Step'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value $Row.'ActionName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GroupName' -Value $Row.'GroupName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'LastStatusMsgName' -Value $Row.'LastStatusMsgName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExitCode' -Value $Row.'ExitCode'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionOutput' -Value $Row.'ActionOutput'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Record' -Value $i
            $Results += $obj
        }
        $FilteredResults = $Results | Select-Object -Property Icon, ComputerName, GUID, ExecutionTime, Step, ActionName, GroupName, LastStatusMsgName, ExitCode, Record
        $hash.Window.Dispatcher.Invoke([action]{
            $hash.DataGrid.ItemsSource = $FilteredResults
            $hash.ErrorCount.Text = $errtable.Count
            $hash.ActionOutput.Text = 'Click any step to view its Action Output.'
        })
    }
    $MDTInt = if ($hash.MDTIntegrated) { $hash.MDTIntegrated.IsChecked } else { $false }
    $MDTU   = if ($hash.MDTURL) { $hash.MDTURL.Text } else { '' }
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($hash.SQLServer.Text).AddArgument($hash.Database.Text).AddArgument($hash.TimePeriod.Text).AddArgument($hash.ErrorsOnly.IsChecked).AddArgument($hash.ComputerName.SelectedItem).AddArgument($hash.TaskSequence.SelectedItem).AddArgument($MDTInt).AddArgument($MDTU).AddArgument($hash.DTFormat.SelectedItem)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ActionOutput 
{
    param ($hash,$RunspacePool)
    $code = 
    {
        param($hash,$Record)
        $msg = $Results | Where-Object { $_.Record -eq $Record }
        $hash.Window.Dispatcher.Invoke([action]{
            $hash.ActionOutput.Text = $msg.ActionOutput
        })
    }
    $Record = $hash.DataGrid.SelectedItem.Record
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($Record)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ComputerNames 
{
    param ($hash,$RunspacePool)
    $code = 
    {
        param($hash,$SQLServer,$Database,$TimePeriod,$ErrorsOnly,$TS,$MDTIntegrated,$URL)
        $ExitFilter = if ($ErrorsOnly -eq 'True' -or $ErrorsOnly -eq $true) { "and ExitCode <> 0" } else { "" }
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        $Query = "
            Select Distinct Name0,
            SMBIOS_GUID0 as 'GUID'
            from vSMS_TaskSequenceExecutionStatus tes
            inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
            $ExitFilter
            ORDER BY Name0 Desc
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $connection.Close()

        $PCResults = @()
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
            $PCResults += $obj
        }

        if ($MDTIntegrated -eq 'true' -or $MDTIntegrated -eq $true)
        {
            $URL1 = $URL.Replace('Computers','ComputerIdentities')
            function GetMDTIDs 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        ID         = $($property.ID.'#text')
                        Identifier = $($property.Identifier)
                    }
                } 
            }
            $MDTIDs = GetMDTIDs -URL $URL1 | Where-Object { $_.Identifier -like '*-*' } | Sort-Object -Property ID
            $UnknownComputers = $PCResults | Where-Object { $_.ComputerName -eq 'Unknown' }
            $MDTComputerIDs = @()
            Foreach ($Computer in $UnknownComputers)
            {
                $MDTComputerID = $MDTIDs | Where-Object { $_.Identifier -eq $Computer.GUID } | Select-Object -Property ID 
                $MDTComputerIDs += $MDTComputerID
            }
            function GetMDTComputerNames 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name = $($property.Name)
                        ID   = $($property.ID.'#text')
                    } 
                } 
            } 
            $MDTComputers = GetMDTComputerNames -URL $URL | Sort-Object -Property ID
            $AdditionalComputerNames = @()
            Foreach ($MDTComputerID in $MDTComputerIDs)
            {
                $MDTComputerName = $MDTComputers | Where-Object { $_.ID -eq $MDTComputerID.ID } | Select-Object -ExpandProperty Name
                $AdditionalComputerNames += $MDTComputerName.ToUpper()
            }
            $ConfigMgrList = $PCResults | Select-Object -Property ComputerName | Where-Object { $_.ComputerName -ne 'Unknown' }
            $FinalComputerNameList = @()
            $FinalComputerNameList += $ConfigMgrList.ComputerName
            $FinalComputerNameList += $AdditionalComputerNames
        }  
        else
        {
            $FinalComputerNameList = @()
            if ($PCResults)
            {
                $validNames = $PCResults | Select-Object -ExpandProperty ComputerName | Where-Object { $_ -and $_ -ne 'Unknown' } | Select-Object -Unique | Sort-Object
                if ($validNames) { $FinalComputerNameList += $validNames }
            }
        }

        $FinalComputerNameList += '-All-'
        $hash.Window.Dispatcher.Invoke([action]{
            $hash.ComputerName.ItemsSource = [Array]$FinalComputerNameList
            if ($hash.ComputerName.SelectedItem -notin $FinalComputerNameList)
            {
                $hash.ComputerName.SelectedItem = '-All-'
            }
        })
    }
    $MDTInt = if ($hash.MDTIntegrated) { $hash.MDTIntegrated.IsChecked } else { $false }
    $MDTU   = if ($hash.MDTURL) { $hash.MDTURL.Text } else { '' }
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($hash.SQLServer.Text).AddArgument($hash.Database.Text).AddArgument($hash.TimePeriod.Text).AddArgument($hash.ErrorsOnly.IsChecked).AddArgument($hash.TaskSequence.SelectedItem).AddArgument($MDTInt).AddArgument($MDTU)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Enable-MDT 
{
    $hash.DeploymentStatus.IsEnabled = 'True'
    $hash.CurrentStep.IsEnabled = 'True'
    $hash.StepName.IsEnabled = 'True'
    $hash.PercentComplete.IsEnabled = 'True'
    $hash.MDTStartTime.IsEnabled = 'True'
    $hash.MDTEndTime.IsEnabled = 'True'
    $hash.MDTElapsedTime.IsEnabled = 'True'
    $hash.DeploymentStatusLabel.IsEnabled = 'True'
    $hash.CurrentStepLabel.IsEnabled = 'True'
    $hash.StepNameLabel.IsEnabled = 'True'
    $hash.PercentCompleteLabel.IsEnabled = 'True'
    $hash.StartLabel.IsEnabled = 'True'
    $hash.EndLabel.IsEnabled = 'True'
    $hash.ElapsedLabel.IsEnabled = 'True'
    $hash.ProgressLabel.IsEnabled = 'True'
}

Function Disable-MDT 
{
    $hash.DeploymentStatus.IsEnabled = $false
    $hash.CurrentStep.IsEnabled = $false
    $hash.StepName.IsEnabled = $false
    $hash.PercentComplete.IsEnabled = $false
    $hash.MDTStartTime.IsEnabled = $false
    $hash.MDTEndTime.IsEnabled = $false
    $hash.MDTElapsedTime.IsEnabled = $false
    $hash.DeploymentStatusLabel.IsEnabled = $false
    $hash.CurrentStepLabel.IsEnabled = $false
    $hash.StepNameLabel.IsEnabled = $false
    $hash.PercentCompleteLabel.IsEnabled = $false
    $hash.StartLabel.IsEnabled = $false
    $hash.EndLabel.IsEnabled = $false
    $hash.ElapsedLabel.IsEnabled = $false
    $hash.ProgressLabel.IsEnabled = $false
}

Function Get-MDTData 
{
    param ($hash,$RunspacePool)
    $code = 
    {
        param($hash,$URL,$ComputerName,$IsMDTIntegrated,$DTFormat)
        $hash.Window.Dispatcher.Invoke([action]{
                $hash.DeploymentStatus.Text = ''
                $hash.CurrentStep.Text = ''
                $hash.StepName.Text = ''
                $hash.PercentComplete.Text = ''
                $hash.MDTStartTime.Text = ''
                $hash.MDTEndTime.Text = ''
                $hash.MDTElapsedTime.Text = ''
                $hash.ProgressBar.Value = 0
        })
        if (($IsMDTIntegrated -eq $true -or $IsMDTIntegrated -eq 'True') -and $ComputerName -ne '-All-' -and !([string]::IsNullOrWhiteSpace($ComputerName)) -and !([string]::IsNullOrWhiteSpace($URL)))
        {
            function GetMDTData2 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name             = $($property.Name)
                        PercentComplete  = $($property.PercentComplete.'#text')
                        CurrentStep      = $($property.CurrentStep.'#text')
                        StepName         = $($property.StepName)
                        Warnings         = $($property.Warnings.'#text')
                        Errors           = $($property.Errors.'#text')
                        DeploymentStatus = $( 
                            Switch ($property.DeploymentStatus.'#text') { 
                                1 { 'Active/Running' } 
                                2 { 'Failed' } 
                                3 { 'Successfully completed' } 
                                Default { 'Unknown' } 
                            } 
                        )
                        StartTime        = $($property.StartTime.'#text') -replace 'T', ' '
                        EndTime          = $($property.EndTime.'#text') -replace 'T', ' '
                    } 
                } 
            } 
            try 
            {
                $MDT = GetMDTData2 -URL $URL | Sort-Object -Property Name | Where-Object { $_.Name -eq $ComputerName }
                if ($MDT)
                {
                    $MDTServer = $URL.Split('//')[2].Split(':')[0]
                    $Start = $MDT.StartTime | Get-Date
                    if ($DTFormat -ne 'UTC')
                    {
                        $Start = [System.TimeZone]::CurrentTimeZone.ToLocalTime($Start)
                    }
                    if (!$MDT.EndTime)
                    {
                        $MDTDate = Invoke-Command -ComputerName $MDTServer -ScriptBlock { (Get-Date).ToUniversalTime() }
                        if ($DTFormat -ne 'UTC')
                        {
                            $MDTDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($MDTDate)
                        }
                        $Elapsed = $MDTDate - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                    }
                    if ($MDT.EndTime)
                    {
                        $End = $MDT.EndTime | Get-Date
                        if ($DTFormat -ne 'UTC')
                        {
                            $End = [System.TimeZone]::CurrentTimeZone.ToLocalTime($End)
                        }
                        $Elapsed = $End - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                    }
                    $hash.Window.Dispatcher.Invoke([action]{
                            $hash.DeploymentStatus.Text = $MDT.DeploymentStatus
                            $hash.CurrentStep.Text = $MDT.CurrentStep
                            $hash.StepName.Text = $MDT.StepName
                            $hash.PercentComplete.Text = $MDT.PercentComplete
                            $hash.ProgressBar.Value = $MDT.PercentComplete
                            $hash.MDTStartTime.Text = $Start
                            $hash.MDTEndTime.Text = $End
                            $hash.MDTElapsedTime.Text = $Elapsed
                    })
                }
                Else 
                {
                    $hash.Window.Dispatcher.Invoke([action]{
                        $hash.DeploymentStatus.Text = 'No data found'
                    })
                }
            }
            catch
            {
                $hash.Window.Dispatcher.Invoke([action]{
                    $hash.ActionOutput.Text = '[ERROR] Could not connect to MDT Web Service'
                })
            }
        }
    }
    $MDTInt = if ($hash.MDTIntegrated) { $hash.MDTIntegrated.IsChecked } else { $false }
    $MDTU   = if ($hash.MDTURL) { $hash.MDTURL.Text } else { '' }
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($MDTU).AddArgument($hash.ComputerName.SelectedItem).AddArgument($MDTInt).AddArgument($hash.DTFormat.SelectedItem)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

function Dispose-PSInstances 
{
    foreach ($PSinstance in $PSInstances)
    {
        if ($PSinstance.InvocationStateInfo.State -eq 'Completed')
        {
            $PSinstance.Dispose()
        }
    }
}

Function Create-Timer 
{
    $global:Timer = New-Object -TypeName System.Windows.Forms.Timer
    $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
}

Function Start-Timer { if ($timer) { $timer.Start() } }
Function Stop-Timer  { if ($timer) { $timer.Stop()  } }

Function Update-Registry 
{
    param($hash)
    If (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
    {
        if (!(Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor'))
        {
            New-Item -Path 'HKLM:\SOFTWARE\SMSAgent' -Name 'ConfigMgr Task Sequence Monitor' -Force | Out-Null
        }
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -Value $hash.SQLServer.Text
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -Value $hash.Database.Text
        if ($hash.MDTURL) {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -Value $hash.MDTURL.Text
        }
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -Value $hash.DTFormat.SelectedItem
    }
}

Function Read-Registry 
{
    if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SQLServer -ErrorAction SilentlyContinue
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Database -ErrorAction SilentlyContinue
        $regmdt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty MDTURL -ErrorAction SilentlyContinue
        $regdtformat = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DTFormat -ErrorAction SilentlyContinue
        if ($regsql) { $hash.SQLServer.Text = $regsql }
        if ($regdb)  { $hash.Database.Text  = $regdb  }
        if ($regmdt -and $hash.MDTURL) { $hash.MDTURL.Text = $regmdt }
        if ($regdtformat -eq 'UTC') { $Global:CurrentDateTimeF = 'UTC' }
    }
}

Function Generate-Report 
{
    param ($hash,$RunspacePool)
    $code = 
    {
        param($hash,$SQLServer,$Database,$StartDate,$EndDate,$TS,$DTFormat,$ReportFormat,$ReportsFolder)
        $Results = @()
        if ($DTFormat -ne 'UTC')
        {
            [datetime]$StartDate = $StartDate.ToUniversalTime()
            [datetime]$EndDate = $EndDate.ToUniversalTime()
        }
        $SQLStart = $StartDate | Get-Date -Format s
        $SQLEnd = $EndDate | Get-Date -Format s
        $hash.Window.Dispatcher.Invoke([action]{
                $hash.Working.Content = 'Working...'
                $hash.ReportProgress.Visibility = 'Visible'
                $hash.ReportProgress.Value = 10
        })
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # 1. Query distinct ResourceIDs for the selected Task Sequence and Time Period
        $Query = "
            select distinct tes.ResourceID
            from vSMS_TaskSequenceExecutionStatus tes
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)
        $hash.Window.Dispatcher.Invoke([action]{
                $hash.ReportProgress.Value = 20
        })

        # 2. Build Execution Summary dataset per machine
        foreach ($ResourceID in $table.Rows.ResourceID)
        {
            $Query = "
                Select (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart' 
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine started execution of a task sequence'
                and Step = 0
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Start',
                (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart' 
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine successfully completed a task sequence'
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Finish',
                (Select name0 from v_R_System sys where sys.ResourceID = $ResourceID) as 'ComputerName',
                (select Model0 from v_GS_Computer_System comp where comp.ResourceID = $ResourceID) as 'Model'
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $reader = $command.ExecuteReader()
            $table2 = New-Object -TypeName 'System.Data.DataTable'
            $table2.Load($reader)
            if ($table2.rows[0].Start.GetType().Name -eq 'DBNull') { $Start = '' }
            else 
            {
                $Start = if ($DTFormat -eq 'UTC') { $table2.rows[0].Start } else { [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table2.rows[0].Start | Get-Date)) }
            }
            if ($table2.rows[0].Finish.GetType().Name -eq 'DBNull') { $Finish = '' }
            else 
            {
                $Finish = if ($DTFormat -eq 'UTC') { $table2.rows[0].Finish } else { [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table2.rows[0].Finish | Get-Date)) }
            }
            $diff = if ($Start -eq '' -or $Finish -eq '') { $Null } else { $Finish - $Start }
            $PC = New-Object -TypeName psobject
            Add-Member -InputObject $PC -MemberType NoteProperty -Name ComputerName -Value $table2.rows[0].ComputerName
            Add-Member -InputObject $PC -MemberType NoteProperty -Name StartTime -Value $Start
            Add-Member -InputObject $PC -MemberType NoteProperty -Name FinishTime -Value $Finish
            $depTime = if ($Start -eq '' -or $Finish -eq '') { '' } else { "$($diff.hours) hours $($diff.minutes) minutes" }
            Add-Member -InputObject $PC -MemberType NoteProperty -Name DeploymentTime -Value $depTime
            Add-Member -InputObject $PC -MemberType NoteProperty -Name Model -Value $table2.rows[0].Model
            $Results += $PC
        }
        $Results = $Results | Sort-Object -Property ComputerName
        $hash.Window.Dispatcher.Invoke([action]{
                $hash.ReportProgress.Value = 50
        })

        # 3. Query ALL steps executed across machines (without exit code filtering)
        $Query = "
            select sys.Name0 as 'ComputerName',
            tsp.Name 'Task Sequence',
            comp.Model0 as Model,
            tes.ExecutionTime,
            tes.Step,
            tes.GroupName,
            tes.ActionName,
            tes.LastStatusMsgName,
            tes.ExitCode,
            tes.ActionOutput
            from vSMS_TaskSequenceExecutionStatus tes
            left join v_R_System sys on tes.ResourceID = sys.ResourceID
            left join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            left join v_GS_COMPUTER_SYSTEM comp on tes.ResourceID = comp.ResourceID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
            Order by tes.ExecutionTime desc
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table3 = New-Object -TypeName 'System.Data.DataTable'
        $table3.Load($reader)
        $connection.Close()

        # 4. Convert step rows into custom PowerShell objects to ensure full row-by-row export
        $AllSteps = @()
        foreach ($row in $table3.Rows)
        {
            $execTime = $row.ExecutionTime
            if ($DTFormat -ne 'UTC' -and $execTime -ne $null -and $execTime -ne [DBNull]::Value)
            {
                $execTime = [System.TimeZone]::CurrentTimeZone.ToLocalTime($execTime)
            }

            $stepObj = New-Object -TypeName PSObject -Property ([ordered]@{
                'ComputerName'      = $row['ComputerName']
                'Task Sequence'     = $row['Task Sequence']
                'Model'             = $row['Model']
                'ExecutionTime'     = $execTime
                'Step'              = $row['Step']
                'GroupName'         = $row['GroupName']
                'ActionName'        = $row['ActionName']
                'LastStatusMsgName' = $row['LastStatusMsgName']
                'ExitCode'          = $row['ExitCode']
                'ActionOutput'      = $row['ActionOutput']
            })
            $AllSteps += $stepObj
        }

        $hash.Window.Dispatcher.Invoke([action]{
                $hash.ReportProgress.Value = 80
        })

        # Ensure target reports folder exists
        if (!(Test-Path -Path $ReportsFolder))
        {
            New-Item -ItemType Directory -Path $ReportsFolder -Force | Out-Null
        }

        # 5. Output based on format selection
        if ($ReportFormat -eq 'CSV')
        {
            # Export CSV Summary Report (per machine)
            $CsvSummaryFile = Join-Path $ReportsFolder "TSReport_Summary.csv"
            $Results | Select-Object -Property ComputerName, StartTime, FinishTime, DeploymentTime, Model | Export-Csv -Path $CsvSummaryFile -NoTypeInformation -Force

            # Export CSV All Steps Report (detailed steps)
            $CsvStepsFile = Join-Path $ReportsFolder "TSReport_AllSteps.csv"
            $AllSteps | Select-Object -Property ComputerName, 'Task Sequence', Model, ExecutionTime, Step, GroupName, ActionName, LastStatusMsgName, ExitCode, ActionOutput | Export-Csv -Path $CsvStepsFile -NoTypeInformation -Force

            $hash.Window.Dispatcher.Invoke([action]{
                $hash.Working.Content = ''
                $hash.ReportProgress.Value = 100
            })

            # Open both the summary and detailed steps reports
            Invoke-Item -Path $CsvSummaryFile
            Invoke-Item -Path $CsvStepsFile
        }
        else
        {
            # HTML Report (includes Summary and All Steps)
            $style = @"
<style>
body { color:#012E34; font-family:Calibri,Tahoma; font-size: 10pt; }
h1 { text-align:center; }
h2 { border-top:1px solid #666666; }
th { font-weight:bold; color:#012E34; background-color:#69969C; }
</style>
"@
            $Headers = "<H1>Task Sequence Execution Summary Report</H1><H3>Starting Date: $StartDate</H3><H3>End Date: $EndDate</H3><H3>Task Sequence: $TS</H3><H3>TimeZone for Date/Time: $DTFormat</H3>"
            $body1 = $Results | Select-Object -Property ComputerName, StartTime, FinishTime, DeploymentTime, Model | ConvertTo-Html -Head $style -Body "<H2>Task Sequence Executions Summary ($($Results.Count))</H2>" | Out-String
            $body2 = $AllSteps | Select-Object -Property ComputerName, 'Task Sequence', Model, ExecutionTime, Step, GroupName, ActionName, LastStatusMsgName, ExitCode | ConvertTo-Html -Head $style -Body "<H2>Task Sequence Execution Steps ($($AllSteps.Count))</H2>" | Out-String
            $Body = $Headers + $body1 + $body2

            $hash.Window.Dispatcher.Invoke([action]{
                $hash.Working.Content = ''
                $hash.ReportProgress.Value = 100
            })

            $HtmlFile = Join-Path $ReportsFolder "TSReport.htm"
            $Body | Out-File -FilePath $HtmlFile -Force
            Invoke-Item -Path $HtmlFile
        }
    }

    [datetime]$StartDate = $hash.StartDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    [datetime]$EndDate = $hash.EndDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    $EndDate = $EndDate.AddDays(1).AddSeconds(-1)
    $TS = $hash.TSList.SelectedItem
    $ReportFormat = if ($hash.ReportFormatCSV.IsChecked -eq $true) { 'CSV' } else { 'HTML' }

    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($hash.SQLServer.Text).AddArgument($hash.Database.Text).AddArgument($StartDate).AddArgument($EndDate).AddArgument($TS).AddArgument($hash.DTFormat.SelectedItem).AddArgument($ReportFormat).AddArgument($GlobalReportsFolder)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

#region Event Handlers
$hash.Window.Add_ContentRendered({
        Read-Registry
        Get-DateTimeFormat
        Get-TaskSequenceList
})
$hash.TaskSequence.Add_SelectionChanged({
        Dispose-PSInstances
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Stop-Timer
        Create-Timer
        $timer.add_Tick({
                Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
                Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        })
        Start-Timer
        $Global:CurrentTS = $hash.TaskSequence.SelectedItem
})
$hash.ErrorsOnly.Add_Checked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})
$hash.ErrorsOnly.Add_Unchecked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})
$hash.DataGrid.Add_SelectionChanged({
        Dispose-PSInstances
        Populate-ActionOutput -hash $hash -RunspacePool $RunspacePool
})
$hash.RefreshNow.Add_Click({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})
$hash.ComputerName.Add_SelectionChanged({
        if ($hash.TaskSequence.SelectedItem -eq $CurrentTS)
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Start-Timer
        }
})
$hash.TimePeriod.Add_KeyDown({
        if ($_.Key -eq 'Return')
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
            $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
            Start-Timer
        }
})
$hash.RefreshPeriod.Add_TextChanged({
        Stop-Timer
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})

$ConfigureSettingsWindow = {
        $reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
        $hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
        $hash.SQLServer = $hash.Window2.FindName('SQLServer')
        $hash.Database = $hash.Window2.FindName('Database')
        $hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
        $hash.TSList = $hash.Window2.FindName('TSList')
        $hash.StartDate = $hash.Window2.FindName('StartDate')
        $hash.EndDate = $hash.Window2.FindName('EndDate')
        $hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
        $hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
        $hash.ReportTab = $hash.Window2.FindName('ReportTab')
        $hash.Tabs = $hash.Window2.FindName('Tabs')
        $hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
        $hash.Working = $hash.Window2.FindName('Working')
        $hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
        $hash.DTFormat = $hash.Window2.FindName('DTFormat')
        $hash.ReportFormatHTML = $hash.Window2.FindName('ReportFormatHTML')
        $hash.ReportFormatCSV  = $hash.Window2.FindName('ReportFormatCSV')
        $hash.SQLServer.Text = $GlobalSQLServer
        $hash.Database.Text  = $GlobalDatabase
        Read-Registry
        If (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
        {
            $hash.Runasadmin.Visibility = 'Visible'
        }
        $hash.TSList.ItemsSource = $Views.TS
        $hash.DTFormat.ItemsSource = $Timezones.TimeZone
        $hash.DTFormat.SelectedIndex = if ($CurrentDateTimeF -eq 'UTC') { 0 } else { 1 }

        $hash.ConnectSQL.Add_Click({
                Update-Registry -hash $hash
                $global:GlobalSQLServer = $hash.SQLServer.Text
                $global:GlobalDatabase  = $hash.Database.Text
                Get-TaskSequenceList
        })
        $hash.GenerateReport.Add_Click({ Generate-Report -hash $hash -RunspacePool $RunspacePool })
        $hash.DTFormat.Add_SelectionChanged({ $Global:CurrentDateTimeF = $hash.DTFormat.SelectedItem })
        $hash.Window2.Add_Closed({ Update-Registry -hash $hash })
}

$hash.SettingsButton.Add_Click({
        & $ConfigureSettingsWindow
        $hash.SettingsTab.Focus()
        $Null = $hash.Window2.ShowDialog()
})

$hash.ReportButton.Add_Click({
        & $ConfigureSettingsWindow
        $hash.ReportTab.Focus()
        $Null = $hash.Window2.ShowDialog()
})

$hash.Window.Add_Closed({
        Stop-Timer
        Dispose-PSInstances
        $RunspacePool.close()
        $RunspacePool.Dispose()
})
#endregion

#region Show UI
# Ensure an Application instance exists in this AppDomain without re-instantiating it
if ([System.Windows.Application]::Current -eq $null) {
    [void](New-Object Windows.Application)
}
# Display the window modally on the current PowerShell STA thread
[void]$hash.Window.ShowDialog()
#endregion
