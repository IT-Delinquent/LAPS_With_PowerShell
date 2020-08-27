#Enter your domain and domain controller below :)
$script:domainController = "DC.DOMAIN.LAN"
$script:domainRoot = "DOMAIN.LAN"

#LOADING ASSEMBLIES
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

#ICON FOR FORM
[string]$base64=@'
BASE64 DATA HERE
'@

#CREATING THE IMAGE FROM BASE64 DATA
$bitmap = New-Object System.Windows.Media.Imaging.BitMapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()
$bitmap.Freeze()

#LAPS WINDOW XML
[xml]$LAPSXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="LAPS UI" Height="400" Width="400" MinHeight="400" MinWidth="400" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2"/>
            <ColumnDefinition/>
            <ColumnDefinition Width="Auto" MinWidth="75"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto" MinHeight="7"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Label Content="ComputerName:" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="1" FontSize="14"/>
        <TextBox Name="Computer_Textbox" VerticalContentAlignment="Center" HorizontalAlignment="Stretch" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Stretch" Margin="3" Grid.Column="1" FontSize="14"/>
        <Button Name="Search_Button" Content="Search" Grid.Column="2" HorizontalAlignment="Stretch" Grid.Row="1" VerticalAlignment="Stretch" Margin="0,3,5,3"/>
        <Label Content="Password" Grid.Column="1" HorizontalAlignment="Stretch" Grid.Row="2" VerticalAlignment="Stretch" FontSize="14"/>
        <TextBox Name="Password_Textbox" Grid.Column="1" HorizontalAlignment="Stretch" Grid.Row="3" TextWrapping="Wrap" Margin="3" VerticalAlignment="Stretch" IsReadOnly="True" FontSize="14"/>
        <Button Name="Copy_Button" Content="Copy" Grid.Column="2" HorizontalAlignment="Stretch" Grid.Row="3" Margin="0,3,5,3" VerticalAlignment="Stretch"/>
        <Label Content="Password Expires" Grid.Column="1" HorizontalAlignment="Stretch" Grid.Row="4" VerticalAlignment="Stretch" FontSize="14"/>
        <TextBox Name="Password_Ex_Textbox" Grid.Column="1" IsReadOnly="True" HorizontalAlignment="Stretch" Grid.Row="5" TextWrapping="Wrap" VerticalAlignment="Stretch" Margin="3" FontSize="14"/>
        <Label Content="New Expiration" Grid.Column="1" HorizontalAlignment="Stretch" Grid.Row="6" VerticalAlignment="Stretch" FontSize="14"/>
        <DatePicker Name="Date_Picker" Grid.Column="1" HorizontalAlignment="Stretch" Grid.Row="7" VerticalAlignment="Stretch" Margin="3" FontSize="14"/>
        <Button Name="Set_Button" Content="Set" Grid.Column="2" HorizontalAlignment="Stretch" Grid.Row="7" VerticalAlignment="Stretch" Margin="0,5,5,5"/>
        <GridSplitter IsEnabled="False" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Row="8" Grid.Column="1" Margin="5,2,5,2" Grid.ColumnSpan="2"/>
        <TextBox Name="Output_Textbox" VerticalScrollBarVisibility="Auto" IsReadOnly="True" HorizontalAlignment="Stretch" Grid.Row="9" TextWrapping="Wrap" Margin="1,5,1,1" VerticalAlignment="Stretch" Grid.ColumnSpan="3" FontSize="12"/>
    </Grid>
</Window>
"@

#LOADING XAML
$LAPSReader=(New-Object System.Xml.XmlNodeReader $LAPSXaml)
$LAPSWindow=[Windows.Markup.XamlReader]::Load($LAPSReader)
$LAPSWindow.Icon = $bitmap

#ASSIGNING CONTROLS
$Computer_Textbox = $LAPSWindow.FindName("Computer_Textbox")
$Search_Button = $LAPSWindow.FindName("Search_Button")
$Password_Textbox = $LAPSWindow.FindName("Password_Textbox")
$Copy_Button = $LAPSWindow.FindName("Copy_Button")
$Password_Ex_Textbox = $LAPSWindow.FindName("Password_Ex_Textbox")
$Date_Picker = $LAPSWindow.FindName("Date_Picker")
$Set_Button = $LAPSWindow.FindName("Set_Button")
$Output_Textbox = $LAPSWindow.FindName("Output_Textbox")

#FUNCTION TO SET OUTPUT TEXTBOX
function set-output-textbox{
    param(
        [string]$value,
        [bool]$date
    )
    if ($date){
        $Output_Textbox.Text = ("[$(Get-Date)] - $value `r`n")
    }else{
        $Output_Textbox.Text = $value
    }
}

#FUNCTION TO UPDATE OUTPUT TEXTBOX
function update-output-textbox{
    param(
        [string]$value,
        [bool]$date
    )
    if ($date){
        $Output_Textbox.AppendText("[$(Get-Date)] - $value `r`n")
    }else{
        $Output_Textbox.AppendText("     $value `r`n")
    }
    $Output_Textbox.ScrollToEnd()
}

#FUNCTION TO UPDATE FORM
function update-form{
    [System.Windows.Forms.Application]::DoEvents()
}

#FUNCTION TO UPDATE PASSWORD TEXTBOX
function update-password-textbox($value){
    $Password_Textbox.Text = $value
}

#FUNCTION TO UPDATE PASSWORD EX TEXTBOX
function update-passwordex-texbox($value){
    $Password_Ex_Textbox.Text = $value
}

#FUNCTION TO SET CONTROLS
function set-controls{
    param(
        [bool]$switcher,
        [bool]$setswitcher
    )
    $Search_Button.IsEnabled = $switcher
    $Set_Button.IsEnabled = $setswitcher
    $Date_Picker.IsEnabled = $setswitcher
}

#DECIDE IF COPY BUTTON SHOULD BE ENABLED
$Copy_Button.IsEnabled = $false
$Password_Textbox.Add_TextChanged({
    if ($Password_Textbox.Text.Length -gt 0){
        $Copy_Button.IsEnabled = $true
    }else{
        $Copy_Button.IsEnabled = $false
    }
})

#MAKING COMPUTER NAME UPPERCASE ON FOCUS LOST
$Computer_Textbox.Add_LostFocus({
    $Computer_Textbox.Text = $Computer_Textbox.Text.ToUpper()
})

#COPY BUTTON LOGIC
$Copy_Button.Add_Click({
    Set-Clipboard -Value $Password_Textbox.Text
})

#COMPUTER TEXTBOX KEYDOWN LOGIC
$Computer_Textbox.Add_KeyDown({
    if ($args.Key -eq 'Enter'){
        $Search_Button.RaiseEvent((New-Object -TypeName System.Windows.RoutedEventArgs $([System.Windows.Controls.Button]::ClickEvent)))
    }
})

#DISABLING CONTROLS ON FORM LOAD
set-controls -switcher $true -setswitcher $false

#WELCOME MESSAGE ON FORM LOAD
$Output_Textbox.HorizontalContentAlignment="Center"
$Output_Textbox.VerticalContentAlignment="Center"
set-output-textbox -date $false -value "Welcome to version 3 of this form! It is now responsive and a lot cleaner in the background. Nothing you ever had to worry about though :)"

#SEARCH BUTTON LOGIC
$Search_Button.Add_Click({

    #DISABLING CONTROLS ON BUTTON PRESS
    $Output_Textbox.HorizontalContentAlignment="Left"
    $Output_Textbox.VerticalContentAlignment="Top"
    set-controls -switcher $false -setswitcher $false
    update-password-textbox -value $null
    update-passwordex-texbox -value $null
    $Date_Picker.Text = $null

    if ($Computer_Textbox.Text.Length -le 0){
        #OUTPUT IF EMPTY SEARCH AND ENABLING CONTROLS
        set-output-textbox -date $true -value "Input cannot be empty"
        set-controls -switcher $true -setswitcher $false    
    }else{
        set-output-textbox -date $true -value "Please Wait"
        
        #PUTTING INPUT INTO VARIABLE
        $script:computerName = $Computer_Textbox.Text

        #CREATING A SYNCHRONISED HASHTABLE
        $script:syncHash = [hashtable]::Synchronized(@{})

        #CREATING SEARCH RUNSPACE
        $searchRunspace = [runspacefactory]::CreateRunspace()
        $searchRunspace.ApartmentState = "STA"
        $searchRunspace.ThreadOptions = "ReuseThread"
        $searchRunspace.Open()
        $searchRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
        $searchRunspace.SessionStateProxy.SetVariable("computerName",$computerName)
        $searchRunspace.SessionStateProxy.SetVariable("domainController",$domainController)

        #POWERSHELL TO BE RAN IN RUNSPACE
        $searchPowerShell = [powershell]::Create().AddScript({
            $syncHash.searchADComputer = Get-ADComputer -Identity $computerName
            $syncHash.searchInvoke = Invoke-Command -ComputerName $domainController -ScriptBlock { Get-AdmPwdPassword -ComputerName $args[0] } -ArgumentList $computerName | Select-Object Password, ExpirationTimeStamp
        })

        #ASSIGNING RUNSPACE TO POWERSHELL
        $searchPowerShell.Runspace = $searchRunspace
        #STARTING THE RUNSPACE AND POWERSHELL
        $searchObject = $searchPowerShell.BeginInvoke()

        #REFRESHING UNTIL POWERSHELL IS COMPLETE
        do{
            Start-Sleep -Milliseconds 100
            update-form
        }while (!$searchObject.IsCompleted)

        #ENDING POWERSHELL INVOKE AND DISPOSING OF RUNSPACE
        $searchPowerShell.EndInvoke($searchObject)
        $searchPowerShell.Dispose()
    
        if ($syncHash.searchADComputer){
            #COMPUTER IS FOUND ON DOMAIN
            if ($syncHash.searchInvoke){
                #INVOKE SUCCESSFUL
                $admpwdPassword = $syncHash.searchInvoke.password
                $admpwdPasswordExpiration = $syncHash.searchInvoke.ExpirationTimeStamp
                $admpwdPasswordExpirationFormatted = $admpwdPasswordExpiration.ToString("dd/MM/yyyy hh:mm:ss")

                #UPDATING FIELDS
                update-output-textbox -date $true -value "Information retrieved"
                update-password-textbox -value $admpwdPassword
                update-passwordex-texbox -value $admpwdPasswordExpirationFormatted
                set-controls -switcher $true -setswitcher $true
            }else{
                #INVOKE FAILED
                update-output-textbox -date $true -value "Failded to retrieve password information"
                update-password-textbox -value $null
                update-passwordex-texbox -value $null
                set-controls -switcher $true -setswitcher $false
            }
        }else{
            #COMPUTER NOT FOUND ON DOMAIN
            update-output-textbox -date $true -value "Host not found on domain"
            update-password-textbox -value $null
            update-passwordex-texbox -value $null
            set-controls -switcher $true -setswitcher $false
        }
    }
})

#SET EXPIRATION BUTTON LOGIC
$Set_Button.Add_Click({
    
    #DISABLING CONTROLS ON BUTTON PRESS
    set-controls -switcher $false -setswitcher $false

    if ($Date_Picker.Text.Length -le 0){
        #OUTPUT IF EMPTY DATE AND ENABLING CONTROLS
        update-output-textbox -date $true -value "No date selected"
        set-controls -switcher $true -setswitcher $true
    }else{
        #GETTING NEW DATES FOR EXPIRATION
        $newExpirationString = $Date_Picker.SelectedDate.ToString("MM/dd/yyyy")
        $script:newExpirationDate = [datetime]::ParseExact($newExpirationString, 'MM/dd/yyyy', $null)
        
        #OUTPUTTING FRIENDLY EXPIRATION TO OUTPUT TEXTBOX
        update-output-textbox -date $true -value "Setting expiration to $newExpirationString..."

        #CREATING SEARCH RUNSPACE
        $setRunspace = [runspacefactory]::CreateRunspace()
        $setRunspace.ApartmentState = "STA"
        $setRunspace.ThreadOptions = "ReuseThread"
        $setRunspace.Open()
        $setRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
        $setRunspace.SessionStateProxy.SetVariable("computerName",$computerName)
        $setRunspace.SessionStateProxy.SetVariable("domainController",$domainController)
        $setRunspace.SessionStateProxy.SetVariable("newExpirationDate",$newExpirationDate)

        #POWERSHELL TO BE RAN IN RUNSPACE
        $setPowerShell = [powershell]::Create().AddScript({
            try{
                $syncHash.setInvoke = Invoke-Command -ComputerName $domainController -ScriptBlock {Reset-AdmPwdPassword -ComputerName $args[0] -WhenEffective $args[1] } -ArgumentList $computerName, $newExpirationDate -ErrorAction Stop
                try{
                    Invoke-GPUpdate -Computer $computerName -ErrorAction Stop
                    $syncHash.setGPUpdate = $true
                }catch{
                    #GP UPDATE FAILED
                    $syncHash.setGPUpdate = $null
                }
            }catch{
                #CHANGING EXPIRATION FAILED
                $syncHash.setInvoke = $null
            }
        })

        #ASSIGNING RUNSPACE TO POWERSHELL
        $setPowerShell.Runspace = $setRunspace
        #STARTING THE RUNSPACE AND POWERSHELL
        $setObject = $setPowerShell.BeginInvoke()

        #REFRESHING UNTIL POWERSHELL IS COMPLETE
        do{
            Start-Sleep -Milliseconds 100
            update-form
        }while (!$setObject.IsCompleted)

        #ENDING POWERSHELL INVOKE AND DISPOSING OF RUNSPACE
        $setPowerShell.EndInvoke($setObject)
        $setPowerShell.Dispose()

        #CHECKING PASSWORD EXPIRATION SUCCESS
        if ($syncHash.setInvoke){
            update-output-textbox -date $true -value "Successfully reset password expiration date"
            #CHECKING GP UPDATE SUCCESS
            if ($syncHash.setGPUpdate){
                update-output-textbox -date $true -value "Succesfully ran GP update"
            }else{
                update-output-textbox -date $true -value "Failed to run GP update, this is probably due to permissions"
            }
        }else{
            update-output-textbox -date $true -value "Failed to reset password expiration date"
        }

        #RESETTING CONTROLS
        set-controls -switcher $true -setswitcher $true
    }
})

#CHECK FOR AD MODULE AND TEST IF ON LOCAL DOMAIN/NETWORK
if ( Test-Connection $domainRoot -Count 1 -Quiet){
    #DOMAIN IS ACCESSIBLE
    if (Get-Module -List ActiveDirectory ){
        #AD MODULE INSTALLED
        #FORM WILL BE DISPLAYED WITHOUT ANY MODIFICATIONS
    }else{
        #AD MODULE NOT INSTALLED
        set-output-textbox -date $false -value "Install the AD module and restart"
        set-controls -switcher $false -setswitcher $false
        $Computer_Textbox.IsEnabled = $false
    }
}else{
    #DOMAIN ISN'T ACCESSIBLE
    set-output-textbox -date $false -value "Domain is not accessible"
    set-controls -switcher $false -setswitcher $false
    $Computer_Textbox.IsEnabled = $false
}   

#REMOVING PROCESS ON FORM CLOSE
$LAPSWindow.Add_Closing({
    try{
        $syncHash.Clear() | Out-Null
    }catch{}
    
    Stop-Process -Name "LAPS" -ErrorAction SilentlyContinue
})

#DISPLAY FORM WHILST TESTING
$app = [Windows.Application]::new()
$app.run($LAPSWindow)
