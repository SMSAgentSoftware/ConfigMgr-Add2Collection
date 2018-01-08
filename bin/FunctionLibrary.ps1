# Function to get collection details
Function Populate-CollectionInfo {
    [CmdletBinding()]
    
    $Query = "
    Select * from v_Collections
    where CollectionName = '$($UI.SelectedCollection)'
    "
    $Result = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()
    
    $UI.CollectionInfo[0] = "Name: " + $Result.Rows[0].CollectionName + "   | "
    If ($Result.Rows[0].CollectionType -eq 1)
    {
        $UI.CollectionInfo[1] = "Type: User" + "   | "
    }
    Else
    {
        $UI.CollectionInfo[1] = "Type: Device" + "   | "
    }
    $UI.CollectionInfo[2] = "Collection ID: " + $Result.Rows[0].SiteID + "   | "
    $UI.CollectionInfo[3] = "Member Count: " + $Result.Rows[0].MemberCount + "   | "
    $UI.CollectionInfo[4] = "Limiting Collection: " + $Result.Rows[0].LimitToCollectionName + "   | "
    $UI.CollectionInfo[5] = "Referenced Collections: " + $Result.Rows[0].IncludeExcludeCollectionsCount + "   | "
    $UI.CollectionInfo[6] = "Membership Change Time: " + $Result.Rows[0].LastMemberChangeTime + "   | "
    $UI.CollectionInfo[7] = "Last Refresh Time: " + $Result.Rows[0].LastRefreshTime

}

# Function to get the current collection membership
Function Populate-Members {
    [CmdletBinding()]

    $Query = "
    Select Name,IsDirect from vCollectionMembers
    Where SiteID = '$($UI.CollectionInfo[2].Split()[2])'
    Order By Name
    "
    $Result = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

    $UI.MembersGrid.ItemsSource = $Result.DefaultView
    $UI.CurrentMembers = $UI.MembersGrid.ItemsSource.Name
}

# Function to add resources to collection
Function Add-ResourcesToCollection {
    [CmdletBinding()]

    # Convert the resource list text into an array
    [array]$ResourceArray = ($ui.Resources.Text -split "\n").Trim() | where  {-not [string]::IsNullOrEmpty($_)}

    $UI.StatusBarData[1] = $ResourceArray.Count

    # If it's a device collection
    If ($UI.CollectionInfo[1] -match "Device")
    {
        # Get a list of resources that exist in ConfigMgr out of the device list provided
        $Query = "
        Select Name0,ResourceID from v_R_System
        where Name0 in ($("'" + ($ResourceArray -join "','") + "'"))
        "
        $Resources = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

        # Create a datatable to hold this info
        $DataTable = New-Object System.Data.DataTable
        [void]$DataTable.Columns.AddRange(@('Resource','Status'))

        # Set the initial status of each device
        Foreach ($Resource in $Resources)
        {
            If ($UI.CurrentMembers -match $Resource.Name0)
            {               
                [void]$DataTable.Rows.Add($Resource.Name0,"Already in collection")
            }
            Else
            {
                [void]$DataTable.Rows.Add($Resource.Name0,"Pending")
            }
        }
        
        # If the resource is not found in ConfigMgr
        Foreach ($Resource in $ResourceArray)
        {
            If ($Resources.Rows.Name0 -notcontains $Resource)
            {
                [void]$DataTable.Rows.Add($Resource,"Resource not found!")
                $UI.StatusBarData[3] ++
            }
        }

        # Update UI
        $UI.Results[0] = $DataTable

        # Remove any devices that are already present in the collection
        $ResourcesTemp = New-Object System.Data.DataTable
        $ResourcesTemp = $Resources.Copy()
        Foreach ($Row in $Resources)
        {
            If ($UI.CurrentMembers -match $Row.Name0)
            {
                $RowToRemove = $ResourcesTemp.Select("Name0 = '$($Row.Name0)'")
                $ResourcesTemp.Rows.Remove($RowToRemove[0])
            }
        }
        $Resources = $ResourcesTemp.Copy()

    }

    # If it's a user collection
    If ($UI.CollectionInfo[1] -match "User")
    {
        # Get a list of resources that exist in ConfigMgr out of the users provided
        $Query = "
        Select User_Name0,ResourceID from v_R_User
        where User_Name0 in ($("'" + ($ResourceArray -join "','") + "'"))
        "
        $UserResources = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

        # Get a list of resources that exist in ConfigMgr out of the user groups list provided
        $Query = "
        Select Usergroup_Name0,ResourceID from v_R_UserGroup
        where Usergroup_Name0 in ($("'" + ($ResourceArray -join "','") + "'"))
        "
        $GroupResources = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

        # Create a datatable to hold this info
        $DataTable = New-Object System.Data.DataTable
        [void]$DataTable.Columns.AddRange(@('Resource','Type','Status'))

        # Set the initial status of each user
        If ($UserResources.Rows.Count -ge 1)
        {
            Foreach ($Resource in $UserResources)
            {
                If ($UI.CurrentMembers -match $Resource.User_Name0)
                {                    
                    [void]$DataTable.Rows.Add($Resource.User_Name0,"User","Already in collection")
                }
                Else
                {
                    [void]$DataTable.Rows.Add($Resource.User_Name0,"User","Pending")
                }
            }
        }

        # Set the initial status of each user group
        If ($GroupResources.Rows.Count -ge 1)
        {
            Foreach ($Resource in $GroupResources)
            {
                If ($UI.CurrentMembers -match $Resource.Usergroup_Name0)
                {
                    [void]$DataTable.Rows.Add($Resource.Usergroup_Name0,"User Group","Already in collection")                    
                }
                Else
                {
                    [void]$DataTable.Rows.Add($Resource.Usergroup_Name0,"User Group","Pending")
                }
            }
        }

        # If the resource is not found in ConfigMgr
        Foreach ($Resource in $ResourceArray)
        {
            If ($GroupResources.Rows.Usergroup_Name0 -notcontains $Resource -and $UserResources.Rows.User_Name0 -notcontains $Resource)
            {
                [void]$DataTable.Rows.Add($Resource,"N/A","Resource not found!")
                $UI.StatusBarData[3] ++
            }
        }

        # Update UI
        $UI.Results[0] = $DataTable

        # Remove any users that are already present in the collection
        $UserResourcesTemp = New-Object System.Data.DataTable
        $UserResourcesTemp = $UserResources.Copy()
        Foreach ($Row in $UserResources)
        {
            If ($UI.CurrentMembers -match $Row.User_Name0)
            {
                $RowToRemove = $UserResourcesTemp.Select("User_Name0 = '$($Row.User_Name0)'")
                $UserResourcesTemp.Rows.Remove($RowToRemove[0])
            }
        }
        $UserResources = $UserResourcesTemp.Copy()

        # Remove any user groups that are already present in the collection
        $GroupResourcesTemp = New-Object System.Data.DataTable
        $GroupResourcesTemp = $GroupResources.Copy()
        Foreach ($Row in $GroupResources)
        {
            If ($UI.CurrentMembers -match $Row.Usergroup_Name0)
            {
                $RowToRemove = $GroupResourcesTemp.Select("Usergroup_Name0 = '$($Row.Usergroup_Name0)'")
                $GroupResourcesTemp.Rows.Remove($RowToRemove[0])
            }
        }
        $GroupResources = $GroupResourcesTemp.Copy()

    }

      
    # Code to run in background job for device resources
    $DeviceCode = {
        Param($Resources,$CollectionID,$UI)

        # Create remote session to system with ConfigMgr Console
        If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
        {
            Try
            {
                $PSSession = New-PSSession -ComputerName $UI.SessionData[2] -ErrorAction Stop
            }
            Catch
            {
                $DataTable = New-Object System.Data.DataTable
                $DataTable.Columns.Add('Error')
                $DataTable.Rows.Add("Error connecting to $($UI.SessionData[2]): $_")
                $UI.Results[0] = $DataTable
                $UI.SessionData[3] ="False"
                $UI.StatusBarData[0] = "Idle"
                Return

            }
        }

        # Process each resource
        Foreach ($Resource in $Resources)
        {
            $UI.StatusBarData[0] = "Processing $($Resource.Name0)"
            Try
            {
                $Code = {
                    Param($Resource,$CollectionID)
                    
                    # Add a new collection rule using the ConfigMgr dll classes
                    Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\adminui.wqlqueryengine.dll"
                    Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ManagementProvider.dll"

                    $SmsNamedValuesDictionary = New-Object Microsoft.ConfigurationManagement.ManagementProvider.SmsNamedValuesDictionary
                    $WQLConnectionManager = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager($SmsNamedValuesDictionary)
                    $WQLConnectionManager.Connect($env:COMPUTERNAME)
                    $Collection = $WQLConnectionManager.GetInstance("SMS_Collection.CollectionID='" + $CollectionID + "'");

                    $collectionRule = $WQLConnectionManager.CreateEmbeddedObjectInstance("SMS_CollectionRuleDirect");
                    $collectionRule["ResourceClassName"].StringValue = "SMS_R_System";
                    $collectionRule["ResourceID"].IntegerValue = $Resource.ResourceID

                    $Dictionary = New-Object 'System.Collections.Generic.Dictionary[[string],[object]]'
                    $Dictionary.Add("collectionRule", $collectionRule);
                    $Result = $collection.ExecuteMethod("AddMembershipRule", $Dictionary)
                    Return $Result.PropertyList.ReturnValue
                
                }

                If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                {
                    $Command = Invoke-Command -Session $PSSession -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                }
                Else
                {
                    $Command = Invoke-Command -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                }

            }
            Catch
            {
                # Error
                $Row = $UI.Results[0].Select("Resource = '$($Resource.Name0)'")
                $Row[0].Status = "Error: $_"
                $UI.StatusBarData[4] ++
                $UI.SessionData[3] ="False"
                $UI.StatusBarData[0] = "Idle"
                If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                {
                    Remove-PSSession $PSSession
                }
                Return
            }
            If ($Command -eq 0)
            {
                # Success
                $Row = $UI.Results[0].Select("Resource = '$($Resource.Name0)'")
                $Row[0].Status = "Success"
                $UI.StatusBarData[2] ++
            }
            Else
            {
                # Non-zero return code
                $Row = $UI.Results[0].Select("Resource = '$($Resource.Name0)'")
                $Row[0].Status = "Error: ReturnCode $Command"
                $UI.StatusBarData[4] ++
            }
        }

        # Cleanup
        $UI.SessionData[3] ="False"
        $UI.StatusBarData[0] = "Idle"
        If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
        {
            Remove-PSSession $PSSession
        }
    }

    # Code to run in background job for user resources
    $UserCode = {
        Param($UserResources,$GroupResources,$CollectionID,$UI)

        # Create remote session to system with ConfigMgr Console
        If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
        {
            Try
            {
                $PSSession = New-PSSession -ComputerName $UI.SessionData[2] -ErrorAction Stop
            }
            Catch
            {
                $DataTable = New-Object System.Data.DataTable
                $DataTable.Columns.Add('Error')
                $DataTable.Rows.Add("Error connecting to $($UI.SessionData[2]): $_")
                $UI.Results[0] = $DataTable
                $UI.SessionData[3] ="False"
                $UI.StatusBarData[0] = "Idle"
                Return
            }
        }

        # Add user-based collection rule
        If ($UserResources.Rows.Count -ge 1 -or $GroupResources.ItemArray.Count -ge 2)
        {
            # Process each resource
            Foreach ($Resource in $UserResources)
            {
                $UI.StatusBarData[0] = "Processing $($Resource.User_Name0)"
                Try
                {
                    $Code = {
                        Param($Resource,$CollectionID)

                        # Add a new collection rule using the ConfigMgr dll classes
                        Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\adminui.wqlqueryengine.dll"
                        Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ManagementProvider.dll"

                        $SmsNamedValuesDictionary = New-Object Microsoft.ConfigurationManagement.ManagementProvider.SmsNamedValuesDictionary
                        $WQLConnectionManager = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager($SmsNamedValuesDictionary)
                        $WQLConnectionManager.Connect($env:COMPUTERNAME)
                        $Collection = $WQLConnectionManager.GetInstance("SMS_Collection.CollectionID='" + $CollectionID + "'");

                        $collectionRule = $WQLConnectionManager.CreateEmbeddedObjectInstance("SMS_CollectionRuleDirect");
                        $collectionRule["ResourceClassName"].StringValue = "SMS_R_User";
                        $collectionRule["ResourceID"].IntegerValue = $Resource.ResourceID

                        $Dictionary = New-Object 'System.Collections.Generic.Dictionary[[string],[object]]'
                        $Dictionary.Add("collectionRule", $collectionRule);
                        $Result = $collection.ExecuteMethod("AddMembershipRule", $Dictionary)
                        Return $Result.PropertyList.ReturnValue             
                    }

                    If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                    {
                        $Command = Invoke-Command -Session $PSSession -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                    }
                    Else
                    {
                        $Command = Invoke-Command -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                    }
                }
                Catch
                {
                    # Error
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.User_Name0)'")
                    $Row[0].Status = "Error: $_"
                    $UI.StatusBarData[4] ++
                    $UI.SessionData[3] ="False"
                    $UI.StatusBarData[0] = "Idle"
                    If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                    {
                        Remove-PSSession $PSSession
                    }
                    Return
                }
                If ($Command -eq 0)
                {
                    # Success
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.User_Name0)'")
                    $Row[0].Status = "Success"
                    $UI.StatusBarData[2] ++
                }
                Else
                {
                    # Non-zero return code
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.User_Name0)'")
                    $Row[0].Status = "Error: ReturnCode $Command"
                    $UI.StatusBarData[4] ++
                }
            }
        }

        # Add user group-based collection rule
        If ($GroupResources.Rows.Count -ge 1 -or $GroupResources.ItemArray.Count -ge 2)
        {
            # Process each resource
            Foreach ($Resource in $GroupResources)
            {
                $UI.StatusBarData[0] = "Processing $($Resource.Usergroup_Name0)"
                Try
                {
                    $Code = {
                        Param($Resource,$CollectionID)

                        # Add a new collection rule using the ConfigMgr dll classes
                        Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\adminui.wqlqueryengine.dll"
                        Add-Type -Path "$(Split-Path $env:SMS_ADMIN_UI_PATH)\Microsoft.ConfigurationManagement.ManagementProvider.dll"

                        $SmsNamedValuesDictionary = New-Object Microsoft.ConfigurationManagement.ManagementProvider.SmsNamedValuesDictionary
                        $WQLConnectionManager = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager($SmsNamedValuesDictionary)
                        $WQLConnectionManager.Connect($env:COMPUTERNAME)
                        $Collection = $WQLConnectionManager.GetInstance("SMS_Collection.CollectionID='" + $CollectionID + "'");

                        $collectionRule = $WQLConnectionManager.CreateEmbeddedObjectInstance("SMS_CollectionRuleDirect");
                        $collectionRule["ResourceClassName"].StringValue = "SMS_R_UserGroup";
                        $collectionRule["ResourceID"].IntegerValue = $Resource.ResourceID

                        $Dictionary = New-Object 'System.Collections.Generic.Dictionary[[string],[object]]'
                        $Dictionary.Add("collectionRule", $collectionRule);
                        $Result = $collection.ExecuteMethod("AddMembershipRule", $Dictionary)
                        Return $Result.PropertyList.ReturnValue               
                    }

                    If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                    {
                        $Command = Invoke-Command -Session $PSSession -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                    }
                    Else
                    {
                        $Command = Invoke-Command -ScriptBlock $Code -ArgumentList $Resource,$CollectionID -ErrorAction Stop
                    }

                }
                Catch
                {
                    # Error
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.Usergroup_Name0)'")
                    $Row[0].Status = "Error: $_"
                    $UI.StatusBarData[4] ++
                    $UI.SessionData[3] ="False"
                    $UI.StatusBarData[0] = "Idle"
                    If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
                    {
                        Remove-PSSession $PSSession
                    }
                    Return
                }
                If ($Command -eq 0)
                {
                    # Success
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.Usergroup_Name0)'")
                    $Row[0].Status = "Success"
                    $UI.StatusBarData[2] ++
                }
                Else
                {
                    # Non-zero return code
                    $Row = $UI.Results[0].Select("Resource = '$($Resource.Usergroup_Name0)'")
                    $Row[0].Status = "Error: ReturnCode $Command"
                    $UI.StatusBarData[4] ++
                }
            }
        }

        # Cleanup
        $UI.SessionData[3] ="False"
        $UI.StatusBarData[0] = "Idle"
        If ($UI.SessionData[2] -ne $env:COMPUTERNAME)
        {
            Remove-PSSession $PSSession
        }
    }
    
    # Create a background thread (PS instance)
    $PowerShell = [PowerShell]::Create()
    
    If ($DataTable.Columns.Caption -contains "Type")
    {
        # Add code and parameters for user collection
        $PowerShell.AddScript($UserCode)
        $PowerShell.AddParameter('UserResources',$UserResources).AddParameter('GroupResources',$GroupResources).AddParameter('CollectionID',$UI.CollectionInfo[2].Split()[2]).AddParameter('UI',$UI)
    }
    Else
    {
        # Add code and parameters for device collection
        $PowerShell.AddScript($DeviceCode)
        $PowerShell.AddArgument($Resources).AddArgument($UI.CollectionInfo[2].Split()[2]).AddArgument($UI)
    }

    # Thunderbirds are go!
    $PowerShell.BeginInvoke()

}

# Function to display a custom messagebox
Function New-WPFMessageBox {

    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore

        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segui"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
            New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}


# Function to create the required registry keys
Function Create-RegistryKeys
{
    If (!(Test-Path -Path $UI.SessionData[4]))
    {
        New-Item -Path $UI.SessionData[4] -Force | out-null
        New-ItemProperty -Path $UI.SessionData[4] -Name SQLServer -Value "" | out-null
        New-ItemProperty -Path $UI.SessionData[4] -Name Database -Value "" | out-null
        New-ItemProperty -Path $UI.SessionData[4] -Name AdminUIServer -Value "" | out-null
    }
}


# Function to update registry keys
Function Update-Registry 
{
    param($SQLServer, $Database, $AdminUIServer)

    Set-ItemProperty -Path $UI.SessionData[4] -Name SQLServer -Value $SQLServer | out-null
    Set-ItemProperty -Path $UI.SessionData[4] -Name Database -Value $Database | Out-Null
    Set-ItemProperty -Path $UI.SessionData[4] -Name AdminUIServer -Value $AdminUIServer | Out-Null
}


# Function to read the registry keys
Function Read-Registry 
{
    $UI.SessionData[0] = Get-ItemProperty -Path $UI.SessionData[4] -Name SQLServer |  Select-Object -ExpandProperty SQLServer
    $UI.SessionData[1] = Get-ItemProperty -Path $UI.SessionData[4] -Name Database | Select-Object -ExpandProperty Database
    $UI.SessionData[2] = Get-ItemProperty -Path $UI.SessionData[4] -Name AdminUIServer | Select-Object -ExpandProperty AdminUIServer

    # Prompt user to enter the SQL Server and Database info if not yet populated
    If (!($ui.SessionData[0]) -or !($UI.SessionData[1]) -or !($UI.SessionData[2]))
    {
        Get-Settings
    }
}

# Function to display settings window
Function Get-Settings {

    # Create the Settings window
    [XML]$Xaml3 = [System.IO.File]::ReadAllLines("$Source\XAML files\Settings.xaml") 
    $UI.SettingsWindow = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml3))
    $xaml3.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $UI.$($_.Name) = $UI.SettingsWindow.FindName($_.Name)
    }
    $UI.SettingsWindow.Icon = "$source\bin\Settings.png"
    $UI.SettingsWindow.DataContext = $UI.SessionData
    $UI.SettingsWindow.Owner = $UI.Window

    # Event: Save button clicked
    $UI.Btn_SettingsOK.Add_Click({

        # Update the registry with the [new] values
        Update-Registry -SQLServer $UI.SQLServer.text -Database $UI.Database.text -AdminUIServer $UI.AdminUIServer.Text

        # Read the registry again to make sure valid values are set
        Read-Registry

        # Close the Settings window
        $UI.SettingsWindow.Close()

    })

    # Show the Settings window
    $null = $UI.SettingsWindow.ShowDialog()
}


# Function to display "About" window
Function Display-About {

    # Create the Display window
    [XML]$Xaml4 = [System.IO.File]::ReadAllLines("$Source\XAML files\About.xaml") 
    $UI.AboutWindow = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml4))
    $xaml4.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $UI.$($_.Name) = $UI.AboutWindow.FindName($_.Name)
    }
    $UI.AboutWindow.Icon = "$Source\bin\information-outline.png"
    $UI.AboutWindow.DataContext = $UI.SessionData
    $UI.AboutWindow.Owner = $UI.Window

    # Set events to open the hyperlinks
    $UI.BlogLink, $UI.MDLink, $UI.GitLink, $UI.PayPalLink | Foreach {
        $_.Add_Click({
            Start-Process $This.NavigateURI
        })
    }

    # Show the About window
    $null = $UI.AboutWindow.ShowDialog()
}


# Function to display "Help" window
Function Display-Help {

    # Create the Help window
    [XML]$Xaml5 = [System.IO.File]::ReadAllLines("$Source\XAML files\Help.xaml") 
    $UI.HelpWindow = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml5))
    $xaml5.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $UI.$($_.Name) = $UI.HelpWindow.FindName($_.Name)
    }
    $UI.HelpWindow.Icon = "$Source\bin\help-circle.png"
    $UI.HelpWindow.DataContext = $UI.SessionData
    $UI.HelpWindow.Owner = $UI.Window

    # Read the FlowDocument content
    [XML]$HelpFlow = [System.IO.File]::ReadAllLines("$Source\XAML Files\HelpFlowDocument.xaml")
    $Reader = New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $HelpFlow
    $XamlDoc = [System.Windows.Markup.XamlReader]::Load($Reader)

    # Add the FlowDocument to the Window
    $UI.HelpWindow.AddChild($XamlDoc)

    # Handle the ContextMenuOpening event and essentially cancel it to prevent the context menu displaying.
    # This is due to a bug in the materialdesigninxaml code that causes the app to crash when right-clicking and loading a context menu (doesn't happen in ISE though?)
    $UI.HelpWindow.Add_ContextMenuOpening({
        [System.Object]$sender = $this
		[System.Windows.RoutedEventArgs]$e = $_
        # or
        #[System.Object]$sender = $args[0]
        #[System.Windows.RoutedEventArgs]$e = $args[1]
        $e.Handled = "true"
    })

    # Show thw Help window
    $UI.HelpWindow.ShowDialog()
}


# Function to check if a new version has been released
Function Check-CurrentVersion {
    Param($UI)

    # Download XML from internet
    Try
    {
        # Use the raw.gihubusercontent.com/... URL
        $URL = "https://raw.githubusercontent.com/SMSAgentSoftware/ConfigMgr-Add2Collection/master/Versions/Add2Collection_Current.xml"
        $WebClient = New-Object System.Net.WebClient
        $webClient.UseDefaultCredentials = $True
        $ByteArray = $WebClient.DownloadData($Url)
        $WebClient.DownloadFile($url, "$env:USERPROFILE\AppData\Local\Temp\Add2Collection.xml")
        $Stream = New-Object System.IO.MemoryStream($ByteArray, 0, $ByteArray.Length)
        $XMLReader = New-Object System.Xml.XmlTextReader -ArgumentList $Stream
        $XMLDocument = New-Object System.Xml.XmlDocument
        [void]$XMLDocument.Load($XMLReader)
        $Stream.Dispose()
    }
    Catch
    {
        Return
    }

    # Add version history to OC
    $UI.SessionData[11] = $XMLDocument

    # Create a datatable for the version history
    $Table = New-Object -TypeName 'System.Data.DataTable'
    [void]$Table.Columns.Add('Version')
    [void]$Table.Columns.Add('Release Date')
    [void]$Table.Columns.Add('Changes')

    # Add a row for each version
    $XMLDocument.Add2Collection.Versions.Version | sort Value -Descending | foreach {
    
        # The changes are put into an array, then converted to a string with each change on a new line for correct display
        [array]$Changes = $_.Changes.Change
        $ofs = "`r`n"
        $Table.Rows.Add($_.Value, $_.ReleaseDate, [string]$Changes)
    
    }

    # Set the source of the datagrid
    $UI.SessionData[12] = $Table

    # Get Current version number
    [double]$CurrentVersion = $XMLDocument.Add2Collection.Versions.Version.Value | Sort -Descending | Select -First 1

    # Enable the "Update" menu item to notify user
    If ($CurrentVersion -gt $UI.SessionData[6])
    {
        Show-BalloonTip -Text "A new version is available. Click to update!" -Title "ConfigMgr Add2Collection" -Icon Info -UI $UI
    }

    # Cleanup temp file
    If (Test-Path -Path "$env:USERPROFILE\AppData\Local\Temp\Add2Collection.xml")
    {
        Remove-Item -Path "$env:USERPROFILE\AppData\Local\Temp\Add2Collection.xml" -Force -Confirm:$false
    }

}


# Function to display a notification tip
function Show-BalloonTip  
{
 
  [CmdletBinding(SupportsShouldProcess = $true)]
  param
  (
    [Parameter(Mandatory=$true)]
    $Text,
   
    [Parameter(Mandatory=$true)]
    $Title,
   
    [ValidateSet('None', 'Info', 'Warning', 'Error')]
    $Icon = 'Info',
    $Timeout = 30000,
    $UI
  )
 
  Add-Type -AssemblyName System.Windows.Forms

  $Form = New-Object System.Windows.Forms.Form
  $Form.ShowInTaskbar = $false
  $Form.WindowState = "Minimized"

  $balloon = New-Object System.Windows.Forms.NotifyIcon

  $path                    = Get-Process -id $pid | Select-Object -ExpandProperty Path
  $balloon.Icon            = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
  $balloon.BalloonTipIcon  = $Icon
  $balloon.BalloonTipText  = $Text
  $balloon.BalloonTipTitle = $Title
  $balloon.Visible         = $true

  $Balloon.Add_BalloonTipClicked({
    $UI.Host.Runspace.Events.GenerateEvent("InvokeUpdate",$null,$null, "InvokeUpdate")
    $This.Dispose()
    $Form.Dispose()
  })

  $Balloon.Add_BalloonTipClosed({
    $This.Dispose()
    $Form.Dispose()
  })

  $balloon.ShowBalloonTip($Timeout)

  $Form.ShowDialog()

  # Can run as app but generate event won't work (different context)
  #$App = [System.Windows.Application]::new()
  #$app.Run($Form)

} 