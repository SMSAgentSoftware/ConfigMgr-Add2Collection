## Script to create a WPF TreeView containing ConfigMgr Collections

# Function to query SQL server for any nested containers
Function Query-SubContainers {
    [CmdletBinding()]
    Param($ParentContainerNodeID)

    $Query = "
    Select * from dbo.Folders
    Where ParentContainerNodeID = '$ParentContainerNodeID'
    and IsDeleted != 'true'
    Order By Name
    "

    $Result = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()
    Return $Result

}

# Function to query collections for container
Function Script:Query-Collections {
    [CmdletBinding()]
    Param($ObjectPath,$CollectionType)

    $Query = "
    Select CollectionName 
    from v_Collections
    where ObjectPath = '$ObjectPath'
    and CollectionType = '$CollectionType'
    and CollectionID in ($RBAC_CollectionIDs)
    Order By CollectionName
    "

    $Result = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()
    Return $Result
}

# Function to filter collections using RBAC
Function Query-RBACPermissions {
    [CmdletBinding()]

    # Get the list of current RBAC admins
    $Query = "
    select AdminID,AdminSID from dbo.RBAC_Admins
    "

    $RBAC_Admins = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

    # Get the user/group claims of the current user context and convert the SIDs to binary form
    $Claims = @()
    [Security.Principal.WindowsIdentity]::GetCurrent().Claims | foreach {
        $Claim = New-Object psobject
        try
        {
            $SSID = New-Object Security.Principal.SecurityIdentifier -ArgumentList $_.Value
            $Binary = New-Object 'byte[]' $SSID.BinaryLength
            $SSID.GetBinaryForm($Binary,0)
            Add-Member -InputObject $Claim -MemberType NoteProperty -Name Binary -Value $Binary
            $Claims += $Claim
        }
        Catch {}
        }
    
    # Match any SIDs in RBAC to SIDs in the users claim
    $AdminIDs = @()
    Foreach ($Claim in $Claims)
    {
        Foreach ($Row in $RBAC_Admins)
        {
            If (($Row.AdminSID -join " ") -eq ($Claim.Binary -join " "))
            {
                $AdminIDs += $Row.AdminID
            }
        }

    }

    $AdminIDs = "'" + ($AdminIDs -join "','") + "'"

    # Find collections where the user has modify resource permissions as a minimum using RBAC
    # Granted operations are calculated on the collection instance from values in the vRBAC_AvailableOperations view, ie:
    # 1 (Read) + 2 (Modify) + 128 (Modify Resource) + 4096 (Read Resource) = 4227, the minimum permissions required to add resources to collections.
    $Query = "
    select col.CollectionID from dbo.RBAC_InstancePermissions ip
    left join dbo.v_Collections col on ip.ObjectKey = col.SiteID
    where ip.AdminID in ($AdminIDs)
    and ip.GrantedOperations >= 4227
    and ip.ObjectTypeID = 1
    "

    $Result = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()

    Return $Result
}

# Function to add an item to the treeView
Function Add-TreeViewItem {

    Param($Text,$ImageSource,$Tag,$Parent,$CollectionType)

    $TreeViewItem = New-Object System.Windows.Controls.TreeViewItem

    # Create a stackpanel
    $StackPanel = New-Object System.Windows.Controls.StackPanel
    $StackPanel.Orientation = "Horizontal"

    $Image = New-Object System.Windows.Controls.Image
    $image.Height = 16
    $Image.Width = 16
    $Image.Margin = "5,0,0,0"
    $Image.Source = $ImageSource

    # Create a textblock
    $TextBlock = New-Object System.Windows.Controls.TextBlock
    $TextBlock.Text = $Text
    $TextBlock.VerticalAlignment = "Center"
    $TextBlock.Padding = 5
    $TextBlock.Margin = 5

    # Add these to treeviewItem
    $StackPanel.AddChild($Image)
    $StackPanel.AddChild($TextBlock)
    $TreeViewItem.Header = $StackPanel
    $TreeViewItem.Tag = $Tag
    If ($CollectionType -eq 1)
    {
        $TreeViewItem.Name = "User"
    }
    Else
    {
        $TreeViewItem.Name = "Device"
    }

    # Add to treeView
   
    [void]$Parent.AddChild($TreeViewItem)

    Return $TreeViewItem
}

Try
{
    $Results = Query-RBACPermissions -ErrorAction Stop
}
Catch
{
    $Customerror = $_.Exception.Message.Replace('"',"'")
    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
    Return
}
$script:RBAC_CollectionIDs = "'" + ($Results.CollectionID -join "','") + "'"


#region WPFWindow
# Create the WPF window 
Add-Type -AssemblyName PresentationFramework

$Window = New-Object System.Windows.Window
$Window.Width = 650
$Window.Height = 620
$Window.WindowStartupLocation = "CenterOwner"
$Window.Title = "ConfigMgr Collections"
$Window.Icon = "$source\bin\col.png"
$Window.Owner = $UI.Window

# Create a Grid container
$Grid = New-Object System.Windows.Controls.Grid
$ColumnDefinition = New-Object System.Windows.Controls.ColumnDefinition
$ColumnDefinition.Width = "1*"#"275*"
$Grid.ColumnDefinitions.Add($ColumnDefinition)
$ColumnDefinition = New-Object System.Windows.Controls.ColumnDefinition
$ColumnDefinition.Width = "4"
$Grid.ColumnDefinitions.Add($ColumnDefinition)
$ColumnDefinition = New-Object System.Windows.Controls.ColumnDefinition
$ColumnDefinition.Width = "2*"#"375*"
$Grid.ColumnDefinitions.Add($ColumnDefinition)
$RowDefinition = New-Object System.Windows.Controls.RowDefinition
$RowDefinition.Height = "30"
$Grid.RowDefinitions.Add($RowDefinition)
$RowDefinition = New-Object System.Windows.Controls.RowDefinition
$RowDefinition.Height = "30*"
$Grid.RowDefinitions.Add($RowDefinition)

# Add a gridsplitter
$GridSplitter = New-Object System.Windows.Controls.GridSplitter
$GridSplitter.HorizontalAlignment = "Stretch"
$GridSplitter.VerticalAlignment = "Stretch"
$GridSplitter.Width = "NaN"
$GridSplitter.Height = "NaN"
$GridSplitter.Background = "White"
[System.Windows.Controls.Grid]::SetColumn($GridSplitter, 1)
[System.Windows.Controls.Grid]::SetRowSpan($GridSplitter,2)
$Grid.AddChild($GridSplitter)

# Add a TreeView control
$TreeView = New-Object System.Windows.Controls.TreeView
$TreeView.Margin = 2
$TreeView.Background = "#E8EAF6"#"#C5CAE9"#"#e5ebf9"
$TreeView.MinWidth = 200
$TreeView.Width = "NaN"
$TreeView.HorizontalAlignment = "Stretch"
$TreeView.BorderThickness = 0
[System.Windows.Controls.Grid]::SetColumn($TreeView, 0)
[System.Windows.Controls.Grid]::SetRowSpan($TreeView,2)
$Grid.AddChild($TreeView)

# Create a stackpanel
$DockPanel = New-Object System.Windows.Controls.DockPanel
$DockPanel.LastChildFill = "False"
$DockPanel.Width ="NaN"
$DockPanel.HorizontalAlignment = "Stretch"

$ClearFilterButton = New-Object System.Windows.Controls.Button
$Image = New-Object System.Windows.Controls.Image
$Image.Source = "$Source\bin\Icon21.bmp"
$Image.Height = 20
$Image.Width = 20
$ClearFilterButton.Content = $Image#"x"
$ClearFilterButton.Width = "26"
$ClearFilterButton.Height = "26"
$ClearFilterButton.Background = "White"
$ClearFilterButton.HorizontalContentAlignment = "Center"
$ClearFilterButton.VerticalContentAlignment = "Center"
$ClearFilterButton.HorizontalAlignment = "Right"
$ClearFilterButton.BorderThickness = 0
$ClearFilterButton.Add_Click({
    $FilterBox.Text = "Filter..."
    $FilterBox.Foreground = [System.Windows.Media.Brushes]::LightGray
})
$ClearFilterButton.Visibility = "Hidden"

$FilterBox = New-Object System.Windows.Controls.TextBox
$FilterBox.Height = 26
$FilterBox.MinWidth = "200"
$FilterBox.HorizontalAlignment = "Stretch"
$FilterBox.VerticalAlignment = "Center"
$FilterBox.VerticalContentAlignment = "Center"
$FilterBox.Text = "Filter..."
$FilterBox.Foreground = [System.Windows.Media.Brushes]::LightGray
$FilterBox.Margin = 2
$FilterBox.Padding = "5,0,0,0"
$FilterBox.BorderThickness = 0
$FilterBox.Add_GotFocus({
    If ($This.Text -eq "Filter...")
    {
        $This.Clear()
        $This.Foreground = [System.Windows.Media.Brushes]::Black
        $ClearFilterButton.Visibility = "Visible"
    }
})
$FilterBox.Add_LostFocus({
    If ($This.Text -eq "")
    {
        $This.Text = "Filter..."
        $This.Foreground = [System.Windows.Media.Brushes]::LightGray
        $ClearFilterButton.Visibility = "Hidden"
    }
})
$FilterBox.Add_TextChanged({
    If ($This.Text -ne "Filter...")
    {
        [System.Windows.Data.CollectionViewSource]::GetDefaultView($List.Items).Filter = [Predicate[Object]]{             
            Try {
                $args[0] -match [regex]::Escape($This.Text)
            } Catch {
                $True
            }
        } 
    }
})
$DockPanel.AddChild($FilterBox)
$DockPanel.AddChild($ClearFilterButton)

$Border = New-Object System.Windows.Controls.Border
$Border.Height = 26
$Border.Width = "NaN"
$Border.HorizontalAlignment = "Stretch"
$Border.Margin = 2
$Border.BorderThickness = 1
$Border.BorderBrush = [System.Windows.Media.Brushes]::Gray
$Border.IsHitTestVisible = "False"
[System.Windows.Controls.Grid]::SetColumn($Border, 2)
[System.Windows.Controls.Grid]::SetRow($Border,0)
$Border.AddChild($DockPanel)
$Grid.AddChild($Border)

# Add a ListView Control
$List = New-Object System.Windows.Controls.ListView
$List.Margin = 2
$List.MinWidth = 300
$List.Width = "NaN"
$List.Height = "NaN"
$List.HorizontalAlignment = "Stretch"
$List.SelectionMode = [System.Windows.Controls.SelectionMode]::Single
[System.Windows.Controls.Grid]::SetColumn($List, 2)
[System.Windows.Controls.Grid]::SetRow($List,1)
$Grid.AddChild($List)

# Add a listitem to advise that collection info is being loaded
$ListItem = New-Object System.Windows.Controls.ListViewItem
$ListItem.Content = "Loading collection info..."
$ListItem.FontSize = 16
$ListItem.FontStyle = [System.Windows.FontStyles]::Italic
$List.AddChild($ListItem)

# Assemble the window
$Window.AddChild($Grid)

# Event to populate the listview with the collections for the selected container
$TreeView.Add_SelectedItemChanged({
    If ($this.SelectedItem.Name -eq "User")
    {
        $List.Items.Clear()
        $Items = @(
            Try
            {
                Query-Collections -ObjectPath $($TreeView.SelectedItem.Tag) -CollectionType 1 -ErrorAction Stop | Select -ExpandProperty CollectionName
            }
            Catch
            {
                $Customerror = $_.Exception.Message.Replace('"',"'")
                New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                Return
            }
        )
        Foreach ($Item in $Items)
        {
            $ListItem = New-Object System.Windows.Controls.ListViewItem
            $ListItem.Padding = 5
            $ListItem.Content = $Item
            $List.Items.Add($ListItem)
        }
    }
    Else
    {
        $List.Items.Clear()
        $Items = @(
            Try
            {
                Query-Collections -ObjectPath $($TreeView.SelectedItem.Tag) -CollectionType 2 -ErrorAction Stop | Select -ExpandProperty CollectionName
            }
            Catch
            {
                $Customerror = $_.Exception.Message.Replace('"',"'")
                New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                Return
            }
        )
        Foreach ($Item in $Items)
        {
            $ListItem = New-Object System.Windows.Controls.ListViewItem
            $ListItem.Padding = 5
            $ListItem.Content = $Item
            $List.Items.Add($ListItem)
        }
    }
})

# Event to close the window and return the selected collection on double-click
$List.Add_MouseDoubleClick({
    $Window.Close()
    $UI.Results[0] = $Null
    $UI.Resources.Text = $Text
    $UI.StatusBarData[1] = 0
    $UI.StatusBarData[2] = 0
    $UI.StatusBarData[3] = 0
    $UI.StatusBarData[4] = 0
    $UI.SelectedCollection = $This.SelectedItem.Content
    Try
    {
        Populate-CollectionInfo -ErrorAction Stop
    }
    Catch
    {
        $Customerror = $_.Exception.Message.Replace('"',"'")
        New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
        Return
    }
    
})

# Add the user and device collection nodes at top level
"User Collections","Device Collections" | foreach {
    $TreeViewItem = New-Object System.Windows.Controls.TreeViewItem

    # Create a stackpanel
    $StackPanel = New-Object System.Windows.Controls.StackPanel
    $StackPanel.Orientation = "Horizontal"

    $Image = New-Object System.Windows.Controls.Image
    $image.Height = 16
    $Image.Width = 16
    $Image.Margin = "5,0,0,0"

    If ($_ -eq "Device Collections")
    {
        $Image.Source = "$source\bin\Icon96.bmp"
        $CollectionType = "Device"
    }
    Else
    {
        $Image.Source = "$source\bin\Icon194.bmp"
        $CollectionType = "User"
    }

    # Create a textblock
    $TextBlock = New-Object System.Windows.Controls.TextBlock
    $TextBlock.Text = $_
    $TextBlock.VerticalAlignment = "Center"
    $TextBlock.Padding = 5
    $TextBlock.Margin = 5

    # Add these to treeviewItem
    $StackPanel.AddChild($Image)
    $StackPanel.AddChild($TextBlock)
    $TreeViewItem.Header = $StackPanel
    $TreeViewItem.Name = $CollectionType
    $TreeViewItem.Tag = "/"
    $TreeViewItem.IsExpanded = "True"

    # Add to treeView
    [void]$treeView.Items.Add($TreeViewItem)
}
#endregion

# Query for Device collection folders
$Query = "
Select * from dbo.Folders
where ObjectType = 5000
and ParentContainerNodeID = 0
and IsDeleted != 'true'
Order By Name
"

Try
{
    $DeviceContainers = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()
}
Catch
{
    $Customerror = $_.Exception.Message.Replace('"',"'")
    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
    Return
}

# Query for User collection folders
$Query = "
Select * from dbo.Folders
where ObjectType = 5001
and ParentContainerNodeID = 0
and IsDeleted != 'true'
Order By Name
"

Try
{
    $UserContainers = ([SQLQuery]::new($UI.SessionData[0],$UI.SessionData[1],$Query)).Execute()
}
Catch
{
    $Customerror = $_.Exception.Message.Replace('"',"'")
    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
    Return
}

$Window.Add_ContentRendered({

# Populate user containers nesting up to 4 levels
If ($UserContainers)
{
    Foreach ($Row in $UserContainers.Rows)
    {
        $UserTreeViewItem1 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $treeView.Items[0] -CollectionType 1

        Try
        {
            $UserSubContainers = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
        }
        Catch
        {
            $Customerror = $_.Exception.Message.Replace('"',"'")
            New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
            Return
        }

        If ($UserSubContainers.Rows.Count -ge 1 -or $UserSubContainers.Table.Rows.Count -ge 1)
        {
            Foreach ($Row in $UserSubContainers)
            {
                $UserTreeViewItem2 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $UserTreeViewItem1 -CollectionType 1
                
                Try
                {
                    $UserSubContainers1 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                }
                Catch
                {
                    $Customerror = $_.Exception.Message.Replace('"',"'")
                    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                    Return
                }

                If ($UserSubContainers1.Rows.Count -ge 1 -or $UserSubContainers1.Table.Rows.Count -ge 1)
                {
                    Foreach ($Row in $UserSubContainers1)
                    {
                        $UserTreeViewItem3 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $UserTreeViewItem2 -CollectionType 1
                        
                        Try
                        {
                            $UserSubContainers2 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                        }
                        Catch
                        {
                            $Customerror = $_.Exception.Message.Replace('"',"'")
                            New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                            Return
                        }

                        If ($UserSubContainers2.Rows.Count -ge 1 -or $UserSubContainers2.Table.Rows.Count -ge 1)
                        {
                            Foreach ($Row in $UserSubContainers2)
                            {
                                $UserTreeViewItem4 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $UserTreeViewItem3 -CollectionType 1
                                
                                Try
                                {
                                    $UserSubContainers3 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                                }
                                Catch
                                {
                                    $Customerror = $_.Exception.Message.Replace('"',"'")
                                    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                                    Return
                                }

                                If ($UserSubContainers3.Rows.Count -ge 1 -or $UserSubContainers3.Table.Rows.Count -ge 1)
                                {
                                    Foreach ($Row in $UserSubContainers3)
                                    {
                                        $UserTreeViewItem5 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $UserTreeViewItem4 -CollectionType 1             
                                    }
                                }               
                            }
                        }               
                    }
                }              
            }
        }
    }
}

# Populate device containers nesting up to 4 levels
If ($DeviceContainers)
{
    Foreach ($Row in $DeviceContainers.Rows)
    {
        $DevicesTreeViewItem1 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $treeView.Items[1] -CollectionType 2

        Try
        {
            $DeviceSubContainers1 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
        }
        Catch
        {
            $Customerror = $_.Exception.Message.Replace('"',"'")
            New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
            Return
        }

        If ($DeviceSubContainers1.Rows.Count -ge 1 -or $DeviceSubContainers1.Table.Rows.Count -ge 1)
        {
            Foreach ($Row in $DeviceSubContainers1)
            {
                $DevicesTreeViewItem2 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $DevicesTreeViewItem1 -CollectionType 2
                
                Try
                {
                    $DeviceSubContainers2 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                }
                Catch
                {
                    $Customerror = $_.Exception.Message.Replace('"',"'")
                    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                    Return
                }

                If ($DeviceSubContainers2.Rows.Count -ge 1 -or $DeviceSubContainers2.Table.Rows.Count -ge 1)
                {
                    Foreach ($Row in $DeviceSubContainers2)
                    {
                        $DevicesTreeViewItem3 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $DevicesTreeViewItem2 -CollectionType 2
                        
                        Try
                        {
                            $DeviceSubContainers3 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                        }
                        Catch
                        {
                            $Customerror = $_.Exception.Message.Replace('"',"'")
                            New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                            Return
                        }

                        If ($DeviceSubContainers3.Rows.Count -ge 1 -or $DeviceSubContainers3.Table.Rows.Count -ge 1)
                        {
                            Foreach ($Row in $DeviceSubContainers3)
                            {
                                $DevicesTreeViewItem4 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $DevicesTreeViewItem3 -CollectionType 2
                                
                                Try
                                {
                                    $DeviceSubContainers4 = Query-SubContainers -ParentContainerNodeID $Row.ContainerNodeID -ErrorAction Stop
                                }
                                Catch
                                {
                                    $Customerror = $_.Exception.Message.Replace('"',"'")
                                    New-WPFMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
                                    Return
                                }

                                If ($DeviceSubContainers4.Rows.Count -ge 1 -or $DeviceSubContainers4.Table.Rows.Count -ge 1)
                                {
                                    Foreach ($Row in $DeviceSubContainers4)
                                    {
                                        $DevicesTreeViewItem5 = Add-TreeViewItem -Text $Row.Name -ImageSource "$source\bin\Icon0.bmp" -Tag $Row.FolderPath -Parent $DevicesTreeViewItem4 -CollectionType 2              
                                    }
                                }               
                            }
                        }               
                    }
                }              
            }
        }
    }
}

# Clear the current list
$List.Items[0].Content = "Double-click a collection to select"

})

# Display the window
#$UI.CollectionWindow = $Window
$null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()