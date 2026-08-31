Add-Type -AssemblyName PresentationFramework -ErrorAction Stop;Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaxRowsPerFile=200000;$MaxFileBytes=100MB;$MaxIssues=1000

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Timeline Creator" Width="1040" Height="820" MinWidth="900" MinHeight="760"
        WindowStartupLocation="CenterScreen" Background="#F4F6F8" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <Style x:Key="Card" TargetType="Border"><Setter Property="Background" Value="White"/><Setter Property="BorderBrush" Value="#D7DEE8"/><Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="8"/><Setter Property="Padding" Value="22"/></Style>
        <Style x:Key="Input" TargetType="TextBox"><Setter Property="Height" Value="38"/><Setter Property="Padding" Value="10,7"/><Setter Property="BorderBrush" Value="#CBD5E1"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
        <Style x:Key="Primary" TargetType="Button"><Setter Property="Height" Value="38"/><Setter Property="Background" Value="#2563EB"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderThickness" Value="0"/><Setter Property="FontWeight" Value="SemiBold"/></Style>
        <Style x:Key="Secondary" TargetType="Button"><Setter Property="Height" Value="36"/><Setter Property="Padding" Value="15,0"/><Setter Property="Background" Value="White"/><Setter Property="BorderBrush" Value="#CBD5E1"/><Setter Property="Foreground" Value="#334155"/></Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions><RowDefinition Height="88"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <Border Grid.Row="0" Background="White" BorderBrush="#D7DEE8" BorderThickness="0,0,0,1"><TextBlock Text="Timeline Creator" FontSize="23" FontWeight="SemiBold" Foreground="#0F172A" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
        <Grid Grid.Row="1" Margin="28,22,28,28">
            <Grid.ColumnDefinitions><ColumnDefinition Width="370"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Grid.Column="0" Style="{StaticResource Card}">
                <StackPanel>
                    <TextBlock Text="INPUT FILES" FontSize="11" FontWeight="Bold" Foreground="#64748B" Margin="0,0,0,12"/>
                    <StackPanel Orientation="Horizontal"><Button x:Name="AddButton" Content="Add TXT / CSV / XLSX" Width="190" Style="{StaticResource Primary}"/><Button x:Name="RemoveButton" Content="Remove" Margin="8,0,0,0" Style="{StaticResource Secondary}"/></StackPanel>
                    <ListBox x:Name="FileList" Height="95" Margin="0,10,0,0" BorderBrush="#CBD5E1"/>
                    <TextBlock Text="Source timezone for selected file" FontSize="11" Foreground="#64748B" Margin="0,12,0,6"/>
                    <ComboBox x:Name="ZoneBox" Height="38" Padding="8,5" BorderBrush="#CBD5E1"/>
                    <Border Height="1" Background="#E2E8F0" Margin="0,16"/>
                    <TextBlock Text="INVESTIGATED USER" FontSize="11" FontWeight="Bold" Foreground="#64748B"/>
                    <TextBlock Text="Primary user ID or email address" FontSize="11" Foreground="#64748B" Margin="0,6,0,6"/>
                    <TextBox x:Name="UserInput" Style="{StaticResource Input}" MaxLength="256"/>
                    <TextBlock Text="Exact aliases, one per line" FontSize="11" Foreground="#64748B" Margin="0,10,0,6"/>
                    <TextBox x:Name="AliasInput" Height="52" Padding="10,7" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" BorderBrush="#CBD5E1" MaxLength="2048"/>
                    <TextBlock Text="Optional Central time range" FontSize="11" Foreground="#64748B" Margin="0,10,0,6"/>
                    <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="8"/><ColumnDefinition/></Grid.ColumnDefinitions><StackPanel Grid.Column="0"><TextBlock Text="Start" FontSize="10" Foreground="#64748B"/><TextBox x:Name="StartInput" Style="{StaticResource Input}" ToolTip="yyyy-MM-dd HH:mm:ss"/></StackPanel><StackPanel Grid.Column="2"><TextBlock Text="End" FontSize="10" Foreground="#64748B"/><TextBox x:Name="EndInput" Style="{StaticResource Input}" ToolTip="yyyy-MM-dd HH:mm:ss"/></StackPanel></Grid>
                    <Button x:Name="BuildButton" Content="Build Timeline" Margin="0,14,0,0" Style="{StaticResource Primary}"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="1" Margin="20,0,0,0" Style="{StaticResource Card}">
                <Grid>
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                    <Grid><TextBlock Text="TIMELINE OUTPUT" FontSize="11" FontWeight="Bold" Foreground="#64748B"/><Border x:Name="StatusBadge" HorizontalAlignment="Right" Background="#F1F5F9" CornerRadius="12" Padding="11,4"><TextBlock x:Name="StatusBadgeText" Text="READY" FontSize="10" FontWeight="Bold" Foreground="#475569"/></Border></Grid>
                    <TextBlock x:Name="ResultTitle" Grid.Row="1" Text="How to build a timeline" Margin="0,22,0,10" FontSize="18" FontWeight="SemiBold"/>
                    <TextBox x:Name="ResultText" Grid.Row="2" IsReadOnly="True" Text="1. Add TXT, CSV, or XLSX log files.&#x0A;&#x0A;2. Source timezone tells the tool how to interpret timestamps without a UTC marker or offset.&#x0A;&#x0A;3. Investigated User is the exact primary user ID or email address to correlate across every file.&#x0A;&#x0A;4. Exact aliases are other user IDs or email addresses that belong to the same person. Enter one per line.&#x0A;&#x0A;5. Start and End are optional inclusive filters in Central time (CST/CDT). Use yyyy-MM-dd HH:mm:ss, or leave both blank to include all matching events.&#x0A;&#x0A;6. Build Timeline creates the Excel workbook and asks where to save it." TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Padding="14" Background="#F8FAFC" BorderBrush="#D7DEE8" FontFamily="Consolas" FontSize="12.5"/>
                    <ProgressBar x:Name="Progress" Grid.Row="3" Height="8" Minimum="0" Maximum="100" Value="0" Margin="0,16,0,0"/>
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0"><Button x:Name="OpenFileButton" Content="Open File" IsEnabled="False" Style="{StaticResource Secondary}"/><Button x:Name="OpenFolderButton" Content="Open Folder" IsEnabled="False" Margin="8,0,0,0" Style="{StaticResource Secondary}"/></StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Parse($xaml)
$AddButton=$window.FindName('AddButton'); $RemoveButton=$window.FindName('RemoveButton'); $FileList=$window.FindName('FileList'); $ZoneBox=$window.FindName('ZoneBox')
$UserInput=$window.FindName('UserInput'); $AliasInput=$window.FindName('AliasInput'); $StartInput=$window.FindName('StartInput'); $EndInput=$window.FindName('EndInput')
$BuildButton=$window.FindName('BuildButton'); $StatusBadge=$window.FindName('StatusBadge'); $StatusBadgeText=$window.FindName('StatusBadgeText'); $ResultTitle=$window.FindName('ResultTitle')
$ResultText=$window.FindName('ResultText'); $Progress=$window.FindName('Progress'); $OpenFileButton=$window.FindName('OpenFileButton'); $OpenFolderButton=$window.FindName('OpenFolderButton')
$script:Files=New-Object 'System.Collections.Generic.List[object]';$script:Issues=New-Object 'System.Collections.Generic.List[object]';$script:OutputPath='';$script:ChangingZone=$false;$script:SuppressedIssues=0;$script:KeyCache=@{};$script:NameCache=@{}
$zones=@([pscustomobject]@{Name='UTC';Id='UTC'},[pscustomobject]@{Name='Central (CST/CDT)';Id='Central Standard Time'},[pscustomobject]@{Name='Eastern (EST/EDT)';Id='Eastern Standard Time'},[pscustomobject]@{Name='Mountain (MST/MDT)';Id='Mountain Standard Time'},[pscustomobject]@{Name='Pacific (PST/PDT)';Id='Pacific Standard Time'})
$profileFields=@{
    'Microsoft Entra'=@{Timestamp=@('Date','CreatedDateTime','CreationTime','TimeGenerated','ActivityDateTime');User=@('Username','UserPrincipalName','User Principal Name','Sign-in identifier','Identity','User');Host=@('DeviceName','Device ID','DeviceId');Target=@('Application','AppDisplayName','Resource','ResourceDisplayName');Ip=@('IP address','IPAddress','ClientIP');Code=@('Sign-in error code','ResultType','ErrorCode');Action=@('Sign-in event types','Client app','ClientAppUsed','Authentication requirement');Outcome=@('Status','Result','Conditional Access status','ConditionalAccessStatus');Message=@('Failure reason','ResultDescription','Status');Process=@('');Parent=@('')}
    'Windows Security'=@{Timestamp=@('Date and Time','TimeCreated','DateTime','EventTime','TimeGenerated');User=@('TargetUserName','SubjectUserName','AccountName','UserName','User');Host=@('Computer','MachineName','ComputerName','Server','WorkstationName');Target=@('TargetServerName','ResourceName','ObjectName');Ip=@('IpAddress','IP Address','Source Network Address','ClientAddress');Code=@('Event ID','EventID','Id');Action=@('Task Category','TaskCategory','Opcode','Action');Outcome=@('LevelDisplayName','Keywords','Status','Result');Message=@('Message','Description','EventDescription','FailureReason');Process=@('NewProcessName','ProcessName','Image','Executable');Parent=@('ParentProcessName','CreatorProcessName','ParentImage')}
    'Generic'=@{Timestamp=@('Timestamp','TimeCreated','DateTime','CreatedDateTime','CreationTime','EventTime','TimeGenerated','Logged','Date','DateUTC','Date (UTC)','StartTime','EventDateTime');User=@('User','UserName','Account','AccountName','UPN','UserPrincipalName','Identity','Actor');Host=@('Device','DeviceName','Computer','ComputerName','Host','HostName','Server','Workstation');Target=@('Resource','ResourceName','Application','AppDisplayName','Target');Ip=@('IPAddress','IP Address','SourceIP','ClientIP');Code=@('EventID','EntryID','Id','RecordId','ActivityId');Action=@('Operation','OperationName','Action','EventType','ActivityDisplayName');Outcome=@('Outcome','Result','Status');Message=@('Description','Message','Details','EventDescription','ResultDescription','Activity','FailureReason');Process=@('ProcessName','Image','Executable');Parent=@('ParentProcessName','ParentImage')}
}
$ZoneBox.DisplayMemberPath='Name'; $ZoneBox.ItemsSource=$zones
$brush=New-Object Windows.Media.BrushConverter
$styles=@{Ready=@('#F1F5F9','#475569','READY');Processing=@('#DBEAFE','#1D4ED8','PROCESSING');Success=@('#DCFCE7','#166534','COMPLETE');Warning=@('#FEF3C7','#92400E','REVIEW');Error=@('#FEE2E2','#991B1B','ERROR')}
function Set-State($State,$Title,$Text){$style=$styles[$State];$StatusBadge.Background=$brush.ConvertFromString($style[0]);$StatusBadgeText.Foreground=$brush.ConvertFromString($style[1]);$StatusBadgeText.Text=$style[2];$ResultTitle.Text=$Title;$ResultText.Text=$Text}
function Update-Files([int]$Selected=-1){$FileList.Items.Clear();foreach($file in $script:Files){[void]$FileList.Items.Add("[$($file.Status)] [$($file.ZoneName)] $($file.Name)")};if($Selected-ge 0-and$Selected-lt$script:Files.Count){$FileList.SelectedIndex=$Selected}}
function Invoke-WindowPump{$window.Dispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background)}
function Release-ComObject($Value){if($null-ne$Value-and[Runtime.InteropServices.Marshal]::IsComObject($Value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)}}
function Add-Issue($File,$Row,$Issue,$Columns=''){if($script:Issues.Count-lt$MaxIssues){[void]$script:Issues.Add([pscustomobject]@{File=$File;Row=$Row;Issue=$Issue;'Detected Columns'=$Columns})}else{$script:SuppressedIssues++}}
function ConvertTo-NormalizedKey([string]$Value){if(-not$script:KeyCache.ContainsKey($Value)){$script:KeyCache[$Value]=($Value-replace'[^a-zA-Z0-9]','').ToLowerInvariant()};return $script:KeyCache[$Value]}
function Get-Delimiter([string]$Line){$best='';$count=0;foreach($candidate in @("`t",',',';','|')){$hits=([regex]::Matches($Line,[regex]::Escape($candidate))).Count;if($hits-gt$count){$best=$candidate;$count=$hits}};if(-not$best){throw 'No supported delimiter was found. TXT and CSV files must contain a delimited table.'};return $best}
function Get-ProfileName([string[]]$Headers){$keys=@($Headers|ForEach-Object{ConvertTo-NormalizedKey $_});if(($keys-contains'eventid'-or$keys-contains'taskcategory')-and($keys-contains'message'-or$keys-contains'source'-or$keys-contains'providername'-or$keys-contains'computer')){return 'Windows Security'};if(($keys-contains'userprincipalname'-or$keys-contains'username'-or$keys-contains'user')-and($keys-contains'application'-or$keys-contains'appdisplayname'-or$keys-contains'resource')-and($keys-contains'status'-or$keys-contains'resulttype'-or$keys-contains'signinerrorcode')){return 'Microsoft Entra'};return 'Generic'}
function Add-RowMetadata($Record,$Row,$Sheet,$ProfileName){$Record['__TimelineRow']=$Row;$Record['__TimelineSheet']=$Sheet;$Record['__TimelineProfile']=$ProfileName;return [pscustomobject]$Record}
function Read-DelimitedFile($Path){
    $reader=[IO.StreamReader]::new($Path,[Text.Encoding]::UTF8,$true)
    try{do{$first=$reader.ReadLine()}while($null-ne$first-and[string]::IsNullOrWhiteSpace($first))}finally{$reader.Dispose()}
    if($null-eq$first){throw 'The file is empty.'}
    $delimiter=Get-Delimiter $first
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.Encoding]::UTF8,$true)
    $result=New-Object 'System.Collections.Generic.List[object]'
    try{
        $parser.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string[]]@($delimiter));$parser.HasFieldsEnclosedInQuotes=$true;$parser.TrimWhiteSpace=$false
        $headers=@($parser.ReadFields());if($headers.Count-lt 2){throw 'The file does not contain at least two columns.'}
        $seen=@{}
        for($i=0;$i-lt$headers.Count;$i++){
            $base=if([string]::IsNullOrWhiteSpace($headers[$i])){"Column$($i+1)"}else{$headers[$i].Trim()}
            $name=$base;$suffix=2;while($seen.ContainsKey($name)){$name="$base$suffix";$suffix++}
            $seen[$name]=$true;$headers[$i]=$name
        }
        $ProfileName=Get-ProfileName $headers;$row=1
        while(-not$parser.EndOfData){
            if($result.Count-ge$MaxRowsPerFile){throw "The file exceeds the $MaxRowsPerFile row limit."}
            $values=@($parser.ReadFields());$row++;$record=[ordered]@{}
            for($column=0;$column-lt$headers.Count;$column++){$record[$headers[$column]]=if($column-lt$values.Count){$values[$column]}else{$null}}
            if($values.Count-ne$headers.Count){Add-Issue ([IO.Path]::GetFileName($Path)) $row "Expected $($headers.Count) column(s), found $($values.Count)." ($headers-join', ')}
            [void]$result.Add((Add-RowMetadata $record $row '' $ProfileName))
        }
        return $result.ToArray()
    }catch [Microsoft.VisualBasic.FileIO.MalformedLineException]{throw "Malformed delimited data near line $($parser.ErrorLineNumber)."}finally{$parser.Close()}
}
function Read-ExcelFile($Path,$Excel){$book=$null;$sheets=$null;$result=New-Object 'System.Collections.Generic.List[object]';try{$book=$Excel.Workbooks.Open($Path,0,$true);$sheets=$book.Worksheets;for($sheetIndex=1;$sheetIndex-le$sheets.Count;$sheetIndex++){$sheet=$null;$used=$null;try{$sheet=$sheets.Item($sheetIndex);$used=$sheet.UsedRange;$rows=[int]$used.Rows.Count;$columns=[int]$used.Columns.Count;if($rows-lt 2-or$columns-lt 2){continue};if($result.Count+$rows-1-gt$MaxRowsPerFile){throw "The workbook exceeds the $MaxRowsPerFile row limit."};$values=$used.Value2;$headers=@();$seen=@{};for($column=1;$column-le$columns;$column++){$base=[string]$values[1,$column];if([string]::IsNullOrWhiteSpace($base)){$base="Column$column"};$name=$base;$suffix=2;while($seen.ContainsKey($name)){$name="$base$suffix";$suffix++};$seen[$name]=$true;$headers+=$name};$ProfileName=Get-ProfileName $headers
                for($row=2;$row-le$rows;$row++){$record=[ordered]@{};$hasValue=$false;for($column=1;$column-le$columns;$column++){$value=$values[$row,$column];$record[$headers[$column-1]]=$value;if($null-ne$value-and-not[string]::IsNullOrWhiteSpace([string]$value)){$hasValue=$true}};if($hasValue){[void]$result.Add((Add-RowMetadata $record ([int]$used.Row+$row-1) ([string]$sheet.Name) $ProfileName))}}
            }finally{Release-ComObject $used;Release-ComObject $sheet}};return $result.ToArray()}finally{Release-ComObject $sheets;if($book){$book.Close($false);Release-ComObject $book}}}
function Get-Names($ProfileName){if(-not$script:NameCache.ContainsKey($ProfileName)){$fields=@{};foreach($field in $profileFields['Generic'].Keys){$names=@($profileFields[$ProfileName][$field]);if($ProfileName-ne'Generic'){$names+=@($profileFields['Generic'][$field])};$fields[$field]=@($names|Where-Object{$_}|ForEach-Object{ConvertTo-NormalizedKey $_}|Select-Object -Unique)};$script:NameCache[$ProfileName]=$fields};return $script:NameCache[$ProfileName]}
function Get-RowMap($Row){$map=@{};foreach($property in $Row.PSObject.Properties){if($property.Name-notlike'__Timeline*'){$key=ConvertTo-NormalizedKey $property.Name;if(-not$map.ContainsKey($key)){$map[$key]=$property.Value}}};return $map}
function Get-MapValue($Map,[string[]]$Names){foreach($key in $Names){if($Map.ContainsKey($key)){$value=$Map[$key];if($null-ne$value-and-not[string]::IsNullOrWhiteSpace([string]$value)){return $value}}};return $null}
function Get-MapValues($Map,[string[]]$Names){$values=@();foreach($key in $Names){if($Map.ContainsKey($key)){$value=[string]$Map[$key];if(-not[string]::IsNullOrWhiteSpace($value)){$values+=$value.Trim()}}};return @($values|Select-Object -Unique)}
function Get-RawLog($Row){$parts=New-Object 'System.Collections.Generic.List[string]';foreach($property in $Row.PSObject.Properties){if($property.Name-notlike'__Timeline*'){$value=([string]$property.Value)-replace"`r?`n",' \n ';[void]$parts.Add("$($property.Name)=$value")}};return ($parts-join' | ')}
function Limit-Text($Value,[int]$Maximum=30000){$text=[string]$Value;if($text.Length-le$Maximum){return $text};return $text.Substring(0,$Maximum)+' [TRUNCATED - review source file]'}
function Convert-Stamp($Value,$ZoneId){if($Value-is[double]-or$Value-is[decimal]-or$Value-is[int]-or$Value-is[long]){$local=[datetime]::FromOADate([double]$Value)}else{$text=[string]$Value;if([string]::IsNullOrWhiteSpace($text)){throw 'Timestamp is empty.'};$text=$text.Trim();if($text-match'(?i)(Z|UTC|GMT|[+-]\d{2}:?\d{2})$'){$offset=[datetimeoffset]::MinValue;if(-not[datetimeoffset]::TryParse($text,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$offset)){throw "Timestamp '$text' is invalid."};return $offset.UtcDateTime};$local=[datetime]::MinValue;if(-not[datetime]::TryParse($text,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$local)-and-not[datetime]::TryParse($text,[Globalization.CultureInfo]::CurrentCulture,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$local)){throw "Timestamp '$text' is invalid."}};$local=[datetime]::SpecifyKind($local,[DateTimeKind]::Unspecified);$zone=[TimeZoneInfo]::FindSystemTimeZoneById($ZoneId);if($zone.IsInvalidTime($local)){throw "Timestamp '$local' does not exist in the selected timezone."};if($zone.IsAmbiguousTime($local)){throw "Timestamp '$local' is ambiguous in the selected timezone."};return [TimeZoneInfo]::ConvertTimeToUtc($local,$zone)}
function Format-Central([datetime]$Utc){$zone=[TimeZoneInfo]::FindSystemTimeZoneById('Central Standard Time');$value=[TimeZoneInfo]::ConvertTimeFromUtc([datetime]::SpecifyKind($Utc,[DateTimeKind]::Utc),$zone);$label=if($zone.IsDaylightSavingTime($value)){'CDT'}else{'CST'};return "$($value.ToString('yyyy-MM-dd HH:mm:ss')) $label"}
function Format-Interval([double]$Seconds){$seconds=[Math]::Max(0,[Math]::Round($Seconds));if($seconds-lt 60){return "$seconds seconds"};if($seconds-lt 3600){$minutes=[Math]::Floor($seconds/60);$rest=$seconds%60;return $(if($rest){"$minutes min $rest sec"}else{"$minutes min"})};$hours=[Math]::Round($seconds/3600,1);return $(if($hours-eq 1){'1 hour'}else{"$hours hours"})}
function Get-Explanation($TimelineEvent){
    $user=$TimelineEvent.User;$hostText=if($TimelineEvent.Host){" on '$($TimelineEvent.Host)'"}else{''};$ip=if($TimelineEvent.Ip){" from '$($TimelineEvent.Ip)'"}else{''}
    if($TimelineEvent.Profile-eq'Microsoft Entra'){
        $state=if($TimelineEvent.Outcome-match'(?i)success|succeeded|^0$'){'successful'}elseif($TimelineEvent.Outcome-match'(?i)fail|denied|interrupt'-or($TimelineEvent.Code-match'^\d+$'-and$TimelineEvent.Code-ne'0')){'failed'}else{'recorded'}
        $target=if($TimelineEvent.Target){" to '$($TimelineEvent.Target)'"}else{''};$text="Microsoft Entra recorded a $state sign-in for '$user'$target$ip."
        if($state-eq'failed'-and$TimelineEvent.Message-and$TimelineEvent.Message-ne$TimelineEvent.Outcome){$text+=" The source reported: $($TimelineEvent.Message)."}
        return [pscustomobject]@{Text=(Limit-Text $text 2000);Rule='ENTRA-SIGNIN-v1'}
    }
    if($TimelineEvent.Profile-eq'Windows Security'){
        $code=([string]$TimelineEvent.Code-replace'[^0-9]','')
        switch($code){
            '4624'{$text="Windows recorded a successful logon for '$user'$hostText$ip."}
            '4625'{$text="Windows recorded a failed logon for '$user'$hostText$ip."}
            '4634'{$text="Windows recorded a logoff for '$user'$hostText."}
            '4648'{$text="Windows recorded an explicit-credentials logon attempt by '$user'$hostText."}
            '4672'{$text="Windows assigned special privileges to '$user'$hostText."}
            '4688'{$process=if($TimelineEvent.Process){"'$($TimelineEvent.Process)'"}else{'an unspecified process'};$text="Windows created $process for '$user'$hostText.";if($TimelineEvent.Parent){$text+=" Parent process: '$($TimelineEvent.Parent)'."}}
            default{$label=if($code){" event $code"}else{' an event'};$text="Windows Security recorded$label for '$user'$hostText. Review the log for details."}
        }
        $rule=if($code-in@('4624','4625','4634','4648','4672','4688')){"WIN-$code-v1"}else{'WIN-GENERIC-v1'}
        return [pscustomobject]@{Text=(Limit-Text $text 2000);Rule=$rule}
    }
    $action=if($TimelineEvent.Action){" '$($TimelineEvent.Action)'"}else{' an event'};$outcome=if($TimelineEvent.Outcome){" Outcome: '$($TimelineEvent.Outcome)'."}else{''}
    return [pscustomobject]@{Text=(Limit-Text "The source recorded$action for '$user'$hostText.$outcome Review the log for details." 2000);Rule='GENERIC-v1'}
}
function Get-TimelineRows($Events){$groups=@{};foreach($TimelineEvent in $Events){$key="$($TimelineEvent.Utc.Ticks)|$($TimelineEvent.Profile)|$($TimelineEvent.User)|$($TimelineEvent.RawLog)";if(-not$groups.ContainsKey($key)){$groups[$key]=[pscustomobject]@{Utc=$TimelineEvent.Utc;First=$TimelineEvent;Items=(New-Object 'System.Collections.Generic.List[object]')}};[void]$groups[$key].Items.Add($TimelineEvent)};$output=New-Object 'System.Collections.Generic.List[object]';$previous=$null;foreach($group in @($groups.Values|Sort-Object Utc,@{Expression={$_.First.Source}})){$items=$group.Items.ToArray();$TimelineEvent=$group.First;$elapsed=if($previous){Format-Interval (($TimelineEvent.Utc-$previous).TotalSeconds)}else{'Start'};$explanation=$TimelineEvent.WhatHappened;if($items.Count-gt 1){$explanation+=" Repeated $($items.Count) times in the imported evidence."};[void]$output.Add([pscustomobject]@{Time=Format-Central $TimelineEvent.Utc;Source=(@($items.Source|Select-Object -Unique)-join'; ');Log=Limit-Text $TimelineEvent.Log;'Time Elapsed Since Last Log'=$elapsed;'What Happened'=Limit-Text $explanation});$previous=$TimelineEvent.Utc};return $output.ToArray()}
function Protect-ExcelValue($Value){if($Value-is[string]-and$Value-match'^[=+\-@]'){return "'$Value"};return $Value}
function Write-Worksheet($Sheet,[string[]]$Headers,$Rows,$TableName){$items=@($Rows);$rowCount=$items.Count+1;$columnCount=$Headers.Count;$data=[Array]::CreateInstance([object],[int[]]@($rowCount,$columnCount));$range=$null;$table=$null;try{for($column=0;$column-lt$columnCount;$column++){$data.SetValue($Headers[$column],0,$column)};for($row=0;$row-lt$items.Count;$row++){for($column=0;$column-lt$columnCount;$column++){$data.SetValue((Protect-ExcelValue $items[$row].PSObject.Properties[$Headers[$column]].Value),$row+1,$column)}};$range=$Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item($rowCount,$columnCount));$range.NumberFormat='@';$range.Value2=$data;$Sheet.Rows.Item(1).Font.Bold=$true;$Sheet.Rows.Item(1).Interior.Color=0xE8D7C1;$range.WrapText=$true;[void]$range.Columns.AutoFit();for($column=1;$column-le$columnCount;$column++){if($Sheet.Columns.Item($column).ColumnWidth-gt 60){$Sheet.Columns.Item($column).ColumnWidth=60}};if($items.Count){$table=$Sheet.ListObjects.Add(1,$range,$null,1);$table.Name=$TableName;$table.TableStyle='TableStyleMedium2'};$Sheet.Activate();$Sheet.Application.ActiveWindow.SplitRow=1;$Sheet.Application.ActiveWindow.FreezePanes=$true}finally{Release-ComObject $table;Release-ComObject $range}}

$FileList.Add_SelectionChanged({if($FileList.SelectedIndex-ge 0){$script:ChangingZone=$true;$ZoneBox.SelectedItem=@($zones|Where-Object Id-eq$script:Files[$FileList.SelectedIndex].ZoneId)[0];$script:ChangingZone=$false}})
$ZoneBox.Add_SelectionChanged({if(-not$script:ChangingZone-and$FileList.SelectedIndex-ge 0-and$ZoneBox.SelectedItem){$index=$FileList.SelectedIndex;$script:Files[$index].ZoneId=$ZoneBox.SelectedItem.Id;$script:Files[$index].ZoneName=$ZoneBox.SelectedItem.Name;$script:Files[$index].Status='PENDING';Update-Files $index}})
$AddButton.Add_Click({$dialog=New-Object Microsoft.Win32.OpenFileDialog;$dialog.Multiselect=$true;$dialog.Filter='Supported files (*.txt;*.csv;*.xlsx)|*.txt;*.csv;*.xlsx|TXT files (*.txt)|*.txt|CSV files (*.csv)|*.csv|Excel workbooks (*.xlsx)|*.xlsx';if($dialog.ShowDialog()){foreach($path in $dialog.FileNames){if(-not@($script:Files|Where-Object Path-eq$path).Count){[void]$script:Files.Add([pscustomobject]@{Path=$path;Name=[IO.Path]::GetFileName($path);ZoneId='UTC';ZoneName='UTC';Status='PENDING'})}};Update-Files ($script:Files.Count-1);Set-State Ready 'Files added' "$($script:Files.Count) file(s) added. Select each source timezone before building."}})
$RemoveButton.Add_Click({$index=$FileList.SelectedIndex;if($index-ge 0){$script:Files.RemoveAt($index);Update-Files ([Math]::Min($index,$script:Files.Count-1))}})
$BuildButton.Add_Click({
    if(-not$script:Files.Count){Set-State Warning 'Files required' 'Add at least one TXT, CSV, or XLSX file.';return};$identity=$UserInput.Text.Trim();if(-not$identity){Set-State Warning 'User required' 'Enter the investigated user.';return}
    try{$startUtc=if($StartInput.Text.Trim()){Convert-Stamp $StartInput.Text.Trim() 'Central Standard Time'}else{$null};$endUtc=if($EndInput.Text.Trim()){Convert-Stamp $EndInput.Text.Trim() 'Central Standard Time'}else{$null};if($startUtc-and$endUtc-and$startUtc-gt$endUtc){throw 'The Central start time must be before the end time.'}}catch{Set-State Warning 'Time range needs review' $_.Exception.Message;return}
    $save=New-Object Microsoft.Win32.SaveFileDialog;$save.Filter='Excel workbook (*.xlsx)|*.xlsx';$save.FileName="meridiantimeline_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx";$downloads=Join-Path([Environment]::GetFolderPath('UserProfile'))'Downloads';if(Test-Path $downloads){$save.InitialDirectory=$downloads};if(-not$save.ShowDialog()){return}
    $BuildButton.IsEnabled=$false;$OpenFileButton.IsEnabled=$false;$OpenFolderButton.IsEnabled=$false;$Progress.Value=2;Set-State Processing 'Building timeline' 'Preparing secure Excel processing.';Invoke-WindowPump
    $excel=$null;$book=$null;$timelineSheet=$null;$evidenceSheet=$null;$sourceSheet=$null;$issueSheet=$null
    try{if(-not[type]::GetTypeFromProgID('Excel.Application')){throw 'Microsoft Excel is required to read and create workbooks.'};$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3;$script:Issues.Clear();$script:SuppressedIssues=0
        $aliases=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase);[void]$aliases.Add($identity);foreach($alias in($AliasInput.Text-split'[\r\n;,]+'|Where-Object{$_.Trim()})){[void]$aliases.Add($alias.Trim())}
        $events=New-Object 'System.Collections.Generic.List[object]';$sourceFiles=New-Object 'System.Collections.Generic.List[object]'
        for($fileIndex=0;$fileIndex-lt$script:Files.Count;$fileIndex++){
            $file=$script:Files[$fileIndex];$Progress.Value=5+[int](55*($fileIndex/$script:Files.Count));Set-State Processing 'Reading and normalizing' "$($file.Name)`r`nSource timezone: $($file.ZoneName)";Invoke-WindowPump
            $hash='';$sizeMb=0;$rowsRead=0;$matchedBefore=$events.Count;$profileNames='';$status='REVIEW';$columns=@()
            try{
                $item=Get-Item -LiteralPath $file.Path;$sizeMb=[Math]::Round($item.Length/1MB,2)
                if($item.Length-gt$MaxFileBytes){throw "The file exceeds the $([int]($MaxFileBytes/1MB)) MB size limit."}
                $hash=(Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash
                $rows=@(switch([IO.Path]::GetExtension($file.Path).ToLowerInvariant()){'.txt'{Read-DelimitedFile $file.Path};'.csv'{Read-DelimitedFile $file.Path};'.xlsx'{Read-ExcelFile $file.Path $excel};default{throw 'Unsupported file type.'}})
                if(-not$rows.Count){throw 'The file does not contain data rows.'}
                $rowsRead=$rows.Count;$columns=@($rows[0].PSObject.Properties.Name|Where-Object{$_-notlike'__Timeline*'});$profiles=New-Object 'System.Collections.Generic.HashSet[string]';$recognized=$false
                foreach($row in $rows){
                    $ProfileName=[string]$row.__TimelineProfile;[void]$profiles.Add($ProfileName);$map=Get-RowMap $row;$names=Get-Names $ProfileName;$stamp=Get-MapValue $map $names.Timestamp;$users=@(Get-MapValues $map $names.User)
                    if($null-ne$stamp-and$users.Count){$recognized=$true};$user=@($users|Where-Object{$aliases.Contains($_)}|Select-Object -First 1);if(-not$user.Count){continue}
                    try{$utc=Convert-Stamp $stamp $file.ZoneId}catch{Add-Issue $file.Name $row.__TimelineRow $_.Exception.Message;continue}
                    if(($startUtc-and$utc-lt$startUtc)-or($endUtc-and$utc-gt$endUtc)){continue};$raw=Get-RawLog $row
                    if($raw.Length-gt 30000){Add-Issue $file.Name $row.__TimelineRow 'The raw event was truncated in Excel; review the hashed source file for the complete row.'}
                    $source=$file.Name;if($row.__TimelineSheet){$source+=" / $($row.__TimelineSheet)"}
                    $TimelineEvent=[pscustomobject]@{Utc=$utc;OriginalTimestamp=[string]$stamp;User=[string]$user[0];Host=[string](Get-MapValue $map $names.Host);Target=[string](Get-MapValue $map $names.Target);Ip=[string](Get-MapValue $map $names.Ip);Code=[string](Get-MapValue $map $names.Code);Action=[string](Get-MapValue $map $names.Action);Outcome=[string](Get-MapValue $map $names.Outcome);Message=[string](Get-MapValue $map $names.Message);Process=[string](Get-MapValue $map $names.Process);Parent=[string](Get-MapValue $map $names.Parent);Source=$source;Profile=$ProfileName;Zone=$file.ZoneName;Row=$row.__TimelineRow;RawLog=(Limit-Text $raw);WhatHappened='';ExplanationRule='';Log=''}
                    $explanation=Get-Explanation $TimelineEvent;$TimelineEvent.WhatHappened=$explanation.Text;$TimelineEvent.ExplanationRule=$explanation.Rule
                    $log=if($TimelineEvent.Message){$TimelineEvent.Message}elseif($TimelineEvent.Action){$TimelineEvent.Action}else{$TimelineEvent.RawLog};$TimelineEvent.Log=Limit-Text $log;[void]$events.Add($TimelineEvent)
                }
                if(-not$recognized){throw 'Required timestamp and user columns were not recognized.'};$profileNames=@($profiles)-join', ';$status='READY';$file.Status='READY'
            }catch{$file.Status='REVIEW';Add-Issue $file.Name '' $_.Exception.Message ($columns-join', ')}finally{[void]$sourceFiles.Add([pscustomobject]@{File=$file.Name;'SHA-256'=$hash;'Size MB'=$sizeMb;'Source Timezone'=$file.ZoneName;'Detected Profile'=$profileNames;'Rows Read'=$rowsRead;'Matching Events'=($events.Count-$matchedBefore);Status=$status});Update-Files $fileIndex;Invoke-WindowPump}
        }
        $ordered=@($events|Sort-Object Utc,Source,Row);$timeline=if($ordered.Count){@(Get-TimelineRows $ordered)}else{@()};if(-not$ordered.Count){Add-Issue 'Investigation' '' 'No events matched the investigated user with a valid timestamp in the selected time range.'};if($timeline.Count-gt 60){Add-Issue 'Investigation' '' "The timeline contains $($timeline.Count) distinct moments. No evidence was discarded; narrow the time range to produce 60 or fewer moments."};if($script:SuppressedIssues){[void]$script:Issues.Add([pscustomobject]@{File='Investigation';Row='';Issue="$($script:SuppressedIssues) additional import issue(s) were suppressed after the $MaxIssues issue limit.";'Detected Columns'=''})}
        $evidence=New-Object 'System.Collections.Generic.List[object]';$previous=$null;foreach($TimelineEvent in $ordered){$delta=if($previous){Format-Interval(($TimelineEvent.Utc-$previous).TotalSeconds)}else{'Start'};[void]$evidence.Add([pscustomobject]@{'Original Timestamp'=$TimelineEvent.OriginalTimestamp;'Central Time'=Format-Central $TimelineEvent.Utc;'UTC Time'=$TimelineEvent.Utc.ToString('yyyy-MM-dd HH:mm:ss')+' UTC';'Time Since Previous'=$delta;'Observed User'=$TimelineEvent.User;'Source Profile'=$TimelineEvent.Profile;'Source Host'=$TimelineEvent.Host;Target=$TimelineEvent.Target;'Source IP'=$TimelineEvent.Ip;'Event Code'=$TimelineEvent.Code;Action=$TimelineEvent.Action;Outcome=$TimelineEvent.Outcome;Log=$TimelineEvent.Log;'What Happened'=$TimelineEvent.WhatHappened;'Explanation Rule'=$TimelineEvent.ExplanationRule;Source=$TimelineEvent.Source;'Source Timezone'=$TimelineEvent.Zone;'Source Row'=$TimelineEvent.Row;'Raw Event'=$TimelineEvent.RawLog});$previous=$TimelineEvent.Utc}
        $Progress.Value=75;Set-State Processing 'Creating workbook' 'Writing Timeline, Evidence, Source Files, and Import Issues.';Invoke-WindowPump;$book=$excel.Workbooks.Add();while($book.Worksheets.Count-gt 1){$book.Worksheets.Item($book.Worksheets.Count).Delete()};$timelineSheet=$book.Worksheets.Item(1);$timelineSheet.Name='Timeline';Write-Worksheet $timelineSheet @('Time','Source','Log','Time Elapsed Since Last Log','What Happened') $timeline 'InvestigationTimeline';$evidenceSheet=$book.Worksheets.Add();$evidenceSheet.Name='Evidence';Write-Worksheet $evidenceSheet @('Original Timestamp','Central Time','UTC Time','Time Since Previous','Observed User','Source Profile','Source Host','Target','Source IP','Event Code','Action','Outcome','Log','What Happened','Explanation Rule','Source','Source Timezone','Source Row','Raw Event') $evidence.ToArray() 'InvestigationEvidence';$sourceSheet=$book.Worksheets.Add();$sourceSheet.Name='Source Files';Write-Worksheet $sourceSheet @('File','SHA-256','Size MB','Source Timezone','Detected Profile','Rows Read','Matching Events','Status') $sourceFiles.ToArray() 'InvestigationSources';$issueSheet=$book.Worksheets.Add();$issueSheet.Name='Import Issues';Write-Worksheet $issueSheet @('File','Row','Issue','Detected Columns') $script:Issues.ToArray() 'TimelineImportIssues';$timelineSheet.Move($book.Worksheets.Item(1));$timelineSheet.Activate();$book.SaveAs($save.FileName,51);$script:OutputPath=$save.FileName;$Progress.Value=100;$OpenFileButton.IsEnabled=$true;$OpenFolderButton.IsEnabled=$true
        if($script:Issues.Count){Set-State Warning 'Timeline ready for review' "$($timeline.Count) timeline moment(s) created from $($ordered.Count) matching event(s). Review Import Issues.`r`n`r`n$($script:OutputPath)"}else{Set-State Success 'Timeline ready' "$($timeline.Count) timeline moment(s) created from $($ordered.Count) matching event(s).`r`n`r`n$($script:OutputPath)"}
    }catch{$Progress.Value=0;Set-State Error 'Timeline could not be created' $_.Exception.Message}finally{if($book){$book.Close($false)};foreach($object in @($issueSheet,$sourceSheet,$evidenceSheet,$timelineSheet,$book)){Release-ComObject $object};if($excel){$excel.Quit();Release-ComObject $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();$BuildButton.IsEnabled=$true}
})
$OpenFileButton.Add_Click({if(Test-Path -LiteralPath $script:OutputPath){Start-Process -FilePath $script:OutputPath}})
$OpenFolderButton.Add_Click({if(Test-Path -LiteralPath $script:OutputPath){Start-Process -FilePath (Split-Path -Parent $script:OutputPath)}})
$window.Add_ContentRendered({$UserInput.Focus()|Out-Null})
[void]$window.ShowDialog()
