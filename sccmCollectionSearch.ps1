<#
sccmCollectionSearch.ps1
Quickly access information on computers, collections and collection rules
Search by computer name or collection name
SDK based. Uses WMI queries. Does not require SCCM cmdlets

#>


function SCCMCollectionSearch ($ConnectionDetails) {

    # Import the Assemblies
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    $Username = $ConnectionDetails.Username
    $Password = $ConnectionDetails.Password
    $Server = $ConnectionDetails.Server
    $Site = $ConnectionDetails.Site
    $Collections = @()
    $Rules = @()
    $SecPasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential ($Username, $SecPasswd)

    # Functions
    function ClearComputers{
        $TextBox1.Text = ""
        $ListBox1.items.Clear()
        $ListBox2.items.Clear()
        $ListBox3.items.Clear()

        $Label2.Text = ""
        $Label3.Text = ""
        $Label4.Text = ""
        $Label5.Text = ""
        $Label6.Text = ""
        $Label7.Text = ""
        $Button4.Text = "Search"
    }

    function SearchComputers{
        $SearchTerm = $TextBox1.Text

        $ListBox1.items.Clear()
        $ListBox2.items.Clear()
        $ListBox3.items.Clear()

        $Label2.Text = ""
        $Label3.Text = ""
        $Label4.Text = ""
        $Label5.Text = ""
        $Label6.Text = ""
        $Label7.Text = ""

        if ($SearchTerm)
        {
            $Query = "SELECT Name FROM SMS_R_System WHERE Name like `'%$SearchTerm%`'"
            $Result = Get-WmiObject -Query $Query -ComputerName $Server -credential $Cred -namespace "root\sms\site_$Site" | Sort-Object -Property Name
            # populate listBox1 with the names of the items
            foreach ($Item in $Result.Name)
            {   
                $ListBox1.items.Add($item)
            }
            $ListBox1.Show()
            $Button4.Text = "Filter"        
        }

    }

    function ShowComputerDetails {
        $Label2.Text = ""
        $Label3.Text = ""
        
        $ComputerName = $ListBox1.SelectedItem
        if ($computername)
        {
            $Label2.Text = "Computer: $ComputerName"
            $query = 
        @"
            select SMS_R_System.ADSiteName
                , SMS_R_System.LastLogonUserName
                , SMS_R_System.OperatingSystemNameandVersion 
                , SMS_G_System_COMPUTER_SYSTEM.Manufacturer
                , SMS_G_System_COMPUTER_SYSTEM.Model
            from  SMS_R_System 
            inner join SMS_G_System_COMPUTER_SYSTEM 
            on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId 
            where SMS_R_System.Name = `'$computerName`'
"@
            $result = Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site"
            $ADSite = $result.SMS_R_System.ADSiteName
            $LastUser = $result.SMS_R_System.LastLogonUserName
            $OS = $result.SMS_R_System.OperatingSystemNameandVersion
            $Manufacturer = $result.SMS_G_System_COMPUTER_SYSTEM.Manufacturer
            $Model = $result.SMS_G_System_COMPUTER_SYSTEM.Model
            
            $ComputerInfo = 
        @"
            Manufacturer: $Manufacturer 
            Model: $Model 
            Operating System: $OS 
            Location: $ADSite 
            Last User: $LastUser 
"@
            $label3.Text = $ComputerInfo       
        }
    }

    function ShowCollections {
        $Button4Text = $button4.Text
        if ($Button4Text -eq "Search")
        {
            SearchCollections
        }
        elseif ($Button4Text -eq "Filter") 
        {
            ShowCollectionMembership
        }
    }

    function ShowCollectionMembership {
        $ComputerName = $ListBox1.SelectedItem
        $ListBox2.Items.Clear()
        $ListBox3.Items.Clear()
        $Label4.Text = ""
        $Label5.Text = ""
        $Label6.Text = ""
        $Label7.Text = ""

        if ($ComputerName)
        {
            $Query = "SELECT ResourceID FROM SMS_R_System WHERE Name = `'$ComputerName`'"
            $ResourceID = (Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site").ResourceID
            $CheckBox1Checked = $CheckBox1.Checked
            if ($CheckBox1Checked)
            {
                $query = @"
        SELECT DISTINCT SMS_Collection.Name 
        FROM SMS_Collection
        RIGHT JOIN SMS_FullCollectionMembership 
            ON SMS_Collection.CollectionID = SMS_FullCollectionMembership.CollectionID  
        LEFT JOIN SMS_DeploymentSummary 
            ON SMS_Collection.Name = SMS_DeploymentSummary.CollectionName
        WHERE SMS_FullCollectionMembership.ResourceID = $ResourceID
        AND SMS_DeploymentSummary.FeatureType = 1
"@
            }
            else
            {
            $Query = @"
        SELECT SMS_Collection.Name 
        FROM SMS_Collection 
        JOIN SMS_FullCollectionMembership ON SMS_Collection.CollectionID = SMS_FullCollectionMembership.CollectionID
        WHERE SMS_FullCollectionMembership.ResourceID = $ResourceID
"@ 
            }

            $result = Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site"

            $collections = $result.Name | Sort-Object

            $FilterMemberOf = $TextBox2.Text

            if ($FilterMemberOf) 
            {
                $filtered = @($collections | Where-Object {$_ -like "*$Filtermemberof*"})
                foreach ($item in $filtered)
                {
                    $ListBox2.items.Add($item)
                }
                $ListBox2.Show()
            }

            else {
                foreach ($Collection in $collections)
                {
                    $ListBox2.items.Add($collection)
                }
            }        
        }
    }

    function ShowCollectionDetails {
        $CollectionName =  $ListBox2.SelectedItem
        $Label4.Text = ""
        $Label5.Text = ""
        
        if ($CollectionName)
        {
            $query = 
    @"
            SELECT SMS_Collection.IncludeExcludeCollectionsCount
                , SMS_Collection.LimitToCollectionName
                , SMS_Collection.MemberCount
                , SMS_Collection.Comment
                , SMS_DeploymentInfo.TargetName
            FROM SMS_Collection
            FULL JOIN  SMS_DeploymentInfo
            ON  SMS_DeploymentInfo.CollectionID = SMS_Collection.CollectionID
            WHERE SMS_Collection.Name = `'$CollectionName`'
"@
            $result = @(Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site")
            $LimitingColl = $result.SMS_Collection.LimitToCollectionName | Select-Object -First 1
            $MemberCount = $result.SMS_Collection.MemberCount | Select-Object -First 1
            $Comment = $result.SMS_Collection.Comment | Select-Object -First 1
            $TargetName = $result.SMS_DeploymentInfo.TargetName

            $Label4.Text = "Collection: $CollectionName"
            $CollectionInfo = 
    @"
            Member Count: $MemberCount
            Comment: $Comment
            Limiting Collection: $LimitingColl
            Deployments:
"@
            if ($TargetName)
            {
                foreach ($Deployment in $TargetName)
                {
                    $CollectionInfo = $CollectionInfo + "`n           $Deployment"
                }           
            }

            $Label5.Text = $CollectionInfo
        }
    }

    function FilterMemberOf {

        $FilterMemberOf = $TextBox2.Text

        if ($FilterMemberOf) {
            # clear listBox2
            $ListBox2.Items.Clear()
            # get list of collections from collection array 
            # filter collections by name against $FilterMemberOf
            $filtered = @($collections | Where-Object {$_ -like "*$Filtermemberof*"})
            # populate listBox2 with list of collection names
            foreach ($item in $filtered)
            {
                $ListBox2.items.Add($item)
            }
            $ListBox2.Show()
        }
    }

    function SearchCollections {
        # search all collections for string
        $SearchCollection = $TextBox2.Text
        if ($SearchCollection) {
            # clear listBox2
            $ListBox2.Items.Clear()
            # get list of collections from query

            if ($CheckBox1.Checked)
            {
                $Query = 
    @"
                SELECT DISTINCT SMS_Collection.Name 
                FROM SMS_Collection
                LEFT JOIN SMS_DeploymentSummary 
                    ON SMS_Collection.Name = SMS_DeploymentSummary.CollectionName
                WHERE SMS_Collection.Name like `'%$SearchCollection%`'
                AND SMS_DeploymentSummary.FeatureType = 1
"@            
            }

            else {
                $Query = 
    @"
                SELECT Name FROM SMS_Collection
                WHERE Name like `'%$SearchCollection%`'
"@            
            }



            $result = @(Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site")
            $result
            foreach ($Item in $Result)
            {
                $Name = $Item.Name
                $ListBox2.items.Add($Name)
            }
            $ListBox2.Show()
        }
    }


    Function FilterSearchCollections {   
        $Button4Text = $Button4.Text
        if ($Button4Text -eq "Filter")
        {
            ShowCollectionMembership
        }
        elseif ($Button4Text -eq "Search")
        {
            SearchCollections
        }
        
    }

    function GetCollectionRules {
        $CollectionName = $ListBox2.SelectedItem
        if ($CollectionName)
        {
            $query = "SELECT * FROM SMS_Collection WHERE Name = `'$CollectionName`'"
            $result = Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site"
            $result.get()
            $CollectionRules = @($result.CollectionRules) | Sort-Object -Property $RuleName
            
            $RuleType = $ComboBox1.SelectedItem
            $Rules = @()

            switch ($RuleType) {
                "All" {$Rules =  $CollectionRules }
                "Direct" {$Rules = $CollectionRules | Where-Object {$Null -ne $_.ResourceID}} 
                "Include Collection" {$Rules = $CollectionRules | Where-Object {$Null -ne $_.IncludeCollectionID}}  
                "Exclude Collection" {$Rules = $CollectionRules | Where-Object {$Null -ne $_.ExcludeCollectionID}}  
                "Query" {$Rules = $CollectionRules | Where-Object {$Null -ne $_.QueryID}}  
                Default {$Rules = $CollectionRules}
            }
            # All -> * 
            # Direct -> ResourceID
            # Include -> IncludeCollectionID
            # Exclude -> ExcludeCollectionID
            # Query -> QueryID
            return $Rules
        }
    }


    function ShowCollectionRules {
        $Rules = GetCollectionRules
        $ListBox3.Items.Clear()    
        $Label6.Text = ""
        $Label7.Text = ""
        $ListBox3.Items.Clear()

        foreach ($Rule in $Rules) {
            $ListBox3.Items.Add(($Rule.RuleName))
        }
        $ListBox3.Show()
    }
    function FilterRule {
        $Rules = GetCollectionRules
        $FilterRule = $TextBox3.Text
        $ListBox3.Items.Clear()
        $Label6.Text = ""
        $Label7.Text = ""

        if ($FilterRule)
        {
            $filtered = $Rules | Where-Object {$_.RuleName -like "*$FilterRule*"}
            foreach ($Rule in $filtered)
            {
                $ListBox3.items.Add(($Rule.RuleName))
            }
        }
    }

    function ShowRuleDetails {
        $CollectionName = $ListBox2.SelectedItem
        $RuleName = $ListBox3.SelectedItem
        $Label6.Text = ""
        $Label7.Text = ""
        $query = "SELECT * FROM SMS_Collection WHERE Name = `'$CollectionName`'"
        $result = Get-WmiObject -Query $query -ComputerName $server -credential $cred -namespace "root\sms\site_$site"
        $result.get()
        $CollectionRule = $result.CollectionRules | Where-Object {$_.RuleName -eq "$RuleName"}
        $Label6.Text = "Membership Rule: $RuleName"
        if ($CollectionRule.ResourceID)
        {
            $Direct = $CollectionRule.ResourceID
            $Label7.Text = "        Direct membership: $Direct"
        }

        elseif ($CollectionRule.IncludeCollectionID)
        {
            $Include = $CollectionRule.IncludeCollectionID
            $Label7.Text = "        Include Collection: $Include"
        }

        elseif ($CollectionRule.ExcludeCollectionID)
        {
            $Exclude = $CollectionRule.ExcludeCollectionID
            $Label7.Text = "        Exclude Collection: $Exclude"
        }

        elseif ($CollectionRule.QueryExpression)
        {
            $QueryExpression = $CollectionRule.QueryExpression
            $Label7.Text = "        Query:`n        $QueryExpression"
        }
    }

    function CopyLabelText {
        $Label = $ContextMenuStrip1.SourceControl
        $LabelText = $Label.Text
        [System.Windows.Forms.Clipboard]::SetText($LabelText)
    }

    # -- Main Form --
    # Top
    $Form1 = New-Object System.Windows.Forms.Form
    $Label1 = New-Object System.Windows.Forms.Label

    # Left
    $GroupBox1 = New-Object System.Windows.Forms.GroupBox
    $TextBox1 = New-Object System.Windows.Forms.TextBox
    $Button2 = New-Object System.Windows.Forms.Button
    $button3 = New-Object System.Windows.Forms.Button
    $ListBox1 = New-Object System.Windows.Forms.ListBox

    # Middle
    $GroupBox2 = New-Object System.Windows.Forms.GroupBox
    $CheckBox1 = New-Object System.Windows.Forms.CheckBox
    $TextBox2 = New-Object System.Windows.Forms.TextBox
    $button4 = New-Object System.Windows.Forms.Button
    $ListBox2 = New-Object System.Windows.Forms.ListBox

    # Right
    $GroupBox3 = New-Object System.Windows.Forms.GroupBox
    $ComboBox1 = New-Object System.Windows.Forms.ComboBox
    $TextBox3 = New-Object System.Windows.Forms.TextBox
    $button5 = New-Object System.Windows.Forms.Button
    $ListBox3 = New-Object System.Windows.Forms.ListBox

    # Bottom
    $FlowLayoutPanel1 = New-Object System.Windows.Forms.FlowLayoutPanel
    $Label2 = New-Object System.Windows.Forms.Label
    $Label3 = New-Object System.Windows.Forms.Label
    $Label4 = New-Object System.Windows.Forms.Label
    $Label5 = New-Object System.Windows.Forms.Label
    $Label6 = New-Object System.Windows.Forms.Label
    $Label7 = New-Object System.Windows.Forms.Label


    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    #Events
    $SearchButtonClicked = 
    {
        SearchComputers
    }

    $ClearButtonClicked =
    {
        ClearComputers
    }

    $ListBox1_selected =
    {
        ShowComputerDetails
        ShowCollectionMembership
    }

    $handler_checkBox1_CheckedChanged = 
    {
        ShowCollections
    }

    $FilterSearchCollectionButtonClick = 
    {   
        FilterSearchCollections
    }

    $ListBox2_selected =
    {
        ShowCollectionDetails
        ShowCollectionRules
    }

    $ComboBoxSelectedIndexChanged =
    {
        ShowCollectionRules
    }

    $FilterRuleButtonClick = 
    {
        FilterRule
    }

    $ListBox3_selected =
    {
        ShowRuleDetails
    }

    $ClickCopyMenu = 
    {
        #[System.Windows.Forms.Clipboard]::SetText($label.Text)
        CopyLabelText
    }

    $OnLoadForm_StateCorrection =
    {
        $Form1.WindowState = $InitialFormWindowState
    }

    #----------------------------------------------
    # TOP
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 640
    $System_Drawing_Size.Width = 760
    $Form1.ClientSize = $System_Drawing_Size

    $Form1.FormBorderStyle = 1
    $Form1.Name = "Form1"
    $Form1.Text = "SCCM Collection Search"

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 12
    $System_Drawing_Point.Y = 12
    $Label1.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 30
    $System_Drawing_Size.Width = 670
    $ConnectionMessage = "Connected to Server $server with SCCM Site $Site as user account $username "
    $Label1.Text = $ConnectionMessage
    $Label1.Size = $System_Drawing_Size
    $Form1.Controls.Add($Label1)


    # LEFT
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 12
    $System_Drawing_Point.Y = 42
    $GroupBox1.Location = $System_Drawing_Point
    $GroupBox1.Name = "groupBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 300
    $System_Drawing_Size.Width = 240
    $GroupBox1.Size = $System_Drawing_Size
    $GroupBox1.TabStop = $False
    $GroupBox1.Text = "Computers"
    $Form1.Controls.Add($GroupBox1)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 49
    $TextBox1.Location = $System_Drawing_Point
    $TextBox1.Name = "textBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 160
    $TextBox1.Size = $System_Drawing_Size
    $TextBox1.TabIndex = 2
    $TextBox1.add_KeyUp({
            if($_.KeyCode -eq 'Enter')
            {
                &$SearchButtonClicked
            }
        })

    $GroupBox1.Controls.Add($TextBox1)


    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 175
    $System_Drawing_Point.Y = 21
    $button2.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 54
    $button2.Size = $System_Drawing_Size
    $button2.TabIndex = 3
    $button2.Text = "Clear"
    $button2.UseVisualStyleBackColor = $True
    $button2.add_Click($ClearButtonClicked)
    $GroupBox1.Controls.Add($Button2)


    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 175
    $System_Drawing_Point.Y = 49
    $button3.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 54
    $button3.Size = $System_Drawing_Size
    $button3.TabIndex = 4
    $button3.Text = "Search"
    $button3.UseVisualStyleBackColor = $True
    $button3.add_Click($SearchButtonClicked)
    $GroupBox1.Controls.Add($button3)

    $ListBox1.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 79
    $ListBox1.Location = $System_Drawing_Point
    $ListBox1.Name = "listBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 212
    $System_Drawing_Size.Width = 220
    $ListBox1.Size = $System_Drawing_Size
    $ListBox1.TabIndex = 5
    $ListBox1.HorizontalScrollbar = $True
    $ListBox1.add_SelectedIndexChanged($ListBox1_selected)
    $GroupBox1.Controls.Add($ListBox1)

    # MIDDLE
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 260
    $System_Drawing_Point.Y = 42
    $GroupBox2.Location = $System_Drawing_Point
    $GroupBox2.Name = "groupBox2"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 300
    $System_Drawing_Size.Width = 240
    $GroupBox2.Size = $System_Drawing_Size
    $GroupBox2.TabStop = $False
    $GroupBox2.Text = "Collections"
    $Form1.Controls.Add($GroupBox2)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 6
    $System_Drawing_Point.Y = 20
    $CheckBox1.Location = $System_Drawing_Point
    $CheckBox1.Name = "checkBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 180
    $CheckBox1.Size = $System_Drawing_Size
    $CheckBox1.TabIndex = 6
    $CheckBox1.Text = "Deploys Application"
    $CheckBox1.UseVisualStyleBackColor = $True
    $CheckBox1.add_CheckedChanged($handler_checkBox1_CheckedChanged)
    $CheckBox1.Checked = $True
    $GroupBox2.Controls.Add($CheckBox1)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 6
    $System_Drawing_Point.Y = 48
    $TextBox2.Location = $System_Drawing_Point
    $TextBox2.Name = "textBox2"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 160
    $TextBox2.Size = $System_Drawing_Size
    $TextBox2.TabIndex = 7
    $TextBox2.add_KeyUp({
            if($_.KeyCode -eq 'Enter')
            {
                &$FilterSearchCollectionButtonClick
            }
        })
    $GroupBox2.Controls.Add($TextBox2)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 48
    $button4.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 54
    $button4.Size = $System_Drawing_Size
    $button4.TabIndex = 8
    $button4.Text = "Search"
    $button4.UseVisualStyleBackColor = $True
    $button4.add_Click($FilterSearchCollectionButtonClick)
    $GroupBox2.Controls.Add($button4)

    $ListBox2.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 7
    $System_Drawing_Point.Y = 79
    $ListBox2.Location = $System_Drawing_Point
    $ListBox2.Name = "listBox2"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 212
    $System_Drawing_Size.Width = 220
    $ListBox2.Size = $System_Drawing_Size
    $ListBox2.TabIndex = 9
    $ListBox2.HorizontalScrollbar = $True
    $ListBox2.add_SelectedIndexChanged($ListBox2_selected)
    $GroupBox2.Controls.Add($ListBox2)

    # RIGHT
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 510
    $System_Drawing_Point.Y = 42
    $GroupBox3.Location = $System_Drawing_Point
    $GroupBox3.Name = "groupBox3"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 300
    $System_Drawing_Size.Width = 240
    $GroupBox3.Size = $System_Drawing_Size
    $GroupBox3.TabStop = $False
    $GroupBox3.Text = "Membership Rules"
    $Form1.Controls.Add($GroupBox3)

    $ComboBox1.DropDownStyle = 2
    $ComboBox1.FormattingEnabled = $True
    $ComboBox1.Items.Add("All")|Out-Null
    $ComboBox1.Items.Add("Direct")|Out-Null
    $ComboBox1.Items.Add("Include Collection")|Out-Null
    $ComboBox1.Items.Add("Exclude Collection")|Out-Null
    $ComboBox1.Items.Add("Query")|Out-Null
    $ComboBox1.SelectedIndex = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 7
    $System_Drawing_Point.Y = 21
    $ComboBox1.Location = $System_Drawing_Point
    $ComboBox1.Name = "comboBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 21
    $System_Drawing_Size.Width = 160
    $ComboBox1.Size = $System_Drawing_Size
    $ComboBox1.TabIndex = 10
    $ComboBox1.add_SelectedIndexChanged($ComboBoxSelectedIndexChanged)
    $GroupBox3.Controls.Add($ComboBox1)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 6
    $System_Drawing_Point.Y = 48
    $TextBox3.Location = $System_Drawing_Point
    $TextBox3.Name = "textBox3"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 160
    $TextBox3.Size = $System_Drawing_Size
    $TextBox3.TabIndex = 11
    $TextBox3.add_KeyUp({
            if($_.KeyCode -eq 'Enter')
            {
                &$FilterRuleButtonClick
            }
        })
    $GroupBox3.Controls.Add($TextBox3)

    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 49
    $button5.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 54
    $button5.Size = $System_Drawing_Size
    $button5.TabIndex = 12
    $button5.Text = "Filter"
    $button5.UseVisualStyleBackColor = $True
    $button5.add_Click($FilterRuleButtonClick)
    $GroupBox3.Controls.Add($button5)

    $ListBox3.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 7
    $System_Drawing_Point.Y = 79
    $ListBox3.Location = $System_Drawing_Point
    $ListBox3.Name = "listBox3"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 212
    $System_Drawing_Size.Width = 220
    $ListBox3.Size = $System_Drawing_Size
    $ListBox3.TabIndex = 13
    $ListBox3.HorizontalScrollbar = $True
    $ListBox3.add_SelectedIndexChanged($ListBox3_selected)
    $GroupBox3.Controls.Add($ListBox3)

    # BOTTOM
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 12
    $System_Drawing_Point.Y = 356
    $FlowLayoutPanel1.Location = $System_Drawing_Point
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 258
    $System_Drawing_Size.Width = 736
    $FlowLayoutPanel1.Size = $System_Drawing_Size
    $FlowLayoutPanel1.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $FlowLayoutPanel1.Margin = 6
    $FlowLayoutPanel1.AutoSize = $False
    $FlowLayoutPanel1.AutoScroll = $True
    $Form1.Controls.Add($FlowLayoutPanel1)

    $ContextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
    [System.Windows.Forms.ToolStripItem]$toolStripItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripItem1.Text = "Copy"
    $toolStripItem1.add_Click($ClickCopyMenu)
    $contextMenuStrip1.Items.Add($toolStripItem1)


    $Label2.Text = ""
    $Label2.AutoSize = $True
    $Label2.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label2)

    $Label3.Text = ""
    $Label3.AutoSize = $True
    $Label3.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label3)

    $Label4.Text = ""
    $Label4.AutoSize = $True
    $Label4.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label4)

    $Label5.Text = ""
    $Label5.AutoSize = $True
    $Label5.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label5)

    $Label6.Text = ""
    $Label6.AutoSize = $True
    $Label6.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label6)

    $Label7.Text = ""
    $Label7.AutoSize = $True
    $Label7.ContextMenuStrip = $ContextMenuStrip1
    $FlowLayoutPanel1.Controls.Add($Label7)

    #Save the initial state of the form
    $InitialFormWindowState = $Form1.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $Form1.add_Load($OnLoadForm_StateCorrection)

    #Show the Form
    $Form1.ShowDialog() | Out-Null
}
function ConnectionDetails 
{
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

	# Variables
	$ConnectionDetails = New-Object -TypeName psobject 
	$username = $env:USERNAME
	$Domain = $env:USERDOMAIN
	$UserDomainName = "$Domain\$UserName"
	$RegKey = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\SMS\DP
	[array]$ManagementPoints =  $regkey.ManagementPoints

	if ($ManagementPoints)
	{
		$Server = $ManagementPoints[0]
	}

	else {
		$Server = ""
	}

	if (-not ($Site = $RegKey.SiteCode))
	{
		$Site = ""
	}

	# Objects
	$ConForm1 = New-Object System.Windows.Forms.Form
	$ConLabel6 = New-Object System.Windows.Forms.Label
	$ConButton3 = New-Object System.Windows.Forms.Button
	$ConTextBox4 = New-Object System.Windows.Forms.TextBox
	$ConTextBox3 = New-Object System.Windows.Forms.TextBox
	$ConTextBox2 = New-Object System.Windows.Forms.TextBox
	$ConTextBox1 = New-Object System.Windows.Forms.TextBox
	$ConLabel4 = New-Object System.Windows.Forms.Label
	$ConLabel3 = New-Object System.Windows.Forms.Label
	$ConLabel2 = New-Object System.Windows.Forms.Label
	$ConLabel1 = New-Object System.Windows.Forms.Label
	$ConButton2 = New-Object System.Windows.Forms.Button
	$ConButton1 = New-Object System.Windows.Forms.Button
	$ConInitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

	$handler_button3_Click= 
	{
		# Confirm
		$Username = $ConTextBox1.Text
		$Password = $ConTextBox2.Text
		$Server = $ConTextBox3.Text
		$Site = $ConTextBox4.Text

		$Properties = @{
			Username = $Username
			Password = $Password
			Server = $Server
			Site = $Site
		}
		$ConnectionDetails | add-member $Properties
		$ConForm1.Close()
		#Return $ConnectionDetails
	}

	$handler_button2_Click = 
	{
		# Test
		$Message = ""
		$Username = (($ConTextBox1.Text).Split("\"))[-1]
		$Password = $ConTextBox2.Text
		$Server = $ConTextBox3.Text
		$Site = $ConTextBox4.Text

		$CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
		$domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)
		if ($null -eq $domain.name)
	{
		$Message = "Incorrect User or Password"
		$ConLabel6.Text = $Message
		$ConButton3.Enabled = $false
	}

		else {
			$namespace = "root\sms\site_$Site"
			$WbemLocator = New-Object -ComObject "WbemScripting.SWbemLocator"
			try {
				$WbemLocator.ConnectServer($Server, $Namespace, $username, $Password)
				$Message = "Success!"
				$ConLabel6.Text = $Message
				$ConButton3.Enabled = $True			
			}
			catch {
                $ErrorCode = $_.Exception.ErrorCode
                Write-Host ($ErrorCode.Tostring())
                if ($ErrorCode -eq -2147023174){
                    $Message = "Unable to connect to server"
                }

                elseif ($ErrorCode -eq -2147217394){
                    $Message = "Incorrect site"
                }
                else {
                    $Message = $_.Exception.Message

                }

                $ConLabel6.Text = $Message
                $ConButton3.Enabled = $false

                
			}
	}
	
	}

	$handler_button1_Click = 
	{
		# Cancel
        $ConForm1.Close()
	}

	$OnLoadForm_StateCorrection =
	{
		$ConForm1.WindowState = $ConInitialFormWindowState
	}


    $Handler_Textbox_Enter =
    {   
        $Password = ""
        $Password = $ConTextBox2.Text
        if ($Password){
            if ($ConButton3.Enabled){
                &$handler_button3_Click
            }
            else {
                &$handler_button2_Click
            }            
        }

        
    }

	# Widget Layout

	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 200
	$System_Drawing_Size.Width = 300
	$ConForm1.ClientSize = $System_Drawing_Size
	$ConForm1.DataBindings.DefaultDataSourceUpdateMode = 0
	$ConForm1.Text = "SCCM Connection"

	$ConLabel6.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 114
	$System_Drawing_Point.Y = 125
	$ConLabel6.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 36
	$System_Drawing_Size.Width = 106
	$ConLabel6.Size = $System_Drawing_Size
	$ConLabel6.Text = "Not tested"
	$ConForm1.Controls.Add($ConLabel6)

	$ConButton3.DataBindings.DefaultDataSourceUpdateMode = 0
	$ConButton3.Enabled = $false
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 164
	$ConButton3.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 75
	$ConButton3.Size = $System_Drawing_Size
	$ConButton3.TabIndex = 5
	$ConButton3.Text = "OK"
	$ConButton3.UseVisualStyleBackColor = $True
	$ConButton3.add_Click($handler_button3_Click)

	$ConForm1.Controls.Add($ConButton3)

    # Site
    $ConTextBox4.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 114
	$System_Drawing_Point.Y = 91
	$ConTextBox4.Location = $System_Drawing_Point
	$ConTextBox4.Text = $Site
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 174
	$ConTextBox4.Size = $System_Drawing_Size
    $ConTextBox4.TabIndex = 3
    $ConTextBox4.add_KeyUp(
        {
            if($_.KeyCode -eq 'Enter')
            {
                &$Handler_Textbox_Enter
            }
        }
    )



	$ConForm1.Controls.Add($ConTextBox4)
    
    # Server
	$ConTextBox3.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 114
	$System_Drawing_Point.Y = 68
	$ConTextBox3.Location = $System_Drawing_Point
	$ConTextBox3.Text = $Server
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 174
	$ConTextBox3.Size = $System_Drawing_Size
    $ConTextBox3.TabIndex = 2
    $ConTextBox3.add_KeyUp(
        {
            if($_.KeyCode -eq 'Enter')
            {
                &$Handler_Textbox_Enter
            }
        }
    )

	$ConForm1.Controls.Add($ConTextBox3)

    # Password
    $ConTextBox2.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 114
	$System_Drawing_Point.Y = 42
	$ConTextBox2.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 174
	$ConTextBox2.Size = $System_Drawing_Size
	$ConTextBox2.TabIndex = 1
    $ConTextBox2.UseSystemPasswordChar = $True
    $ConTextBox2.add_KeyUp(
        {
            if($_.KeyCode -eq 'Enter')
            {
                &$Handler_Textbox_Enter
            }
        }
    )

	$ConForm1.Controls.Add($ConTextBox2)

    # Username
    $ConTextBox1.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 114
	$System_Drawing_Point.Y = 16
	$ConTextBox1.Location = $System_Drawing_Point
	$ConTextBox1.Text = $UserDomainName
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 20
	$System_Drawing_Size.Width = 174
	$ConTextBox1.Size = $System_Drawing_Size
    $ConTextBox1.TabIndex = 0
    $ConTextBox1.add_KeyUp(
        {
            if($_.KeyCode -eq 'Enter')
            {
                &$Handler_Textbox_Enter
            }
        }
    )

	$ConForm1.Controls.Add($ConTextBox1)

	$ConLabel4.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 94
	$ConLabel4.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 83
	$ConLabel4.Size = $System_Drawing_Size
	$ConLabel4.Text = "Site"

	$ConForm1.Controls.Add($ConLabel4)

	$ConLabel3.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 71
	$ConLabel3.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 83
	$ConLabel3.Size = $System_Drawing_Size
	$ConLabel3.Text = "Server"

	$ConForm1.Controls.Add($ConLabel3)

	$ConLabel2.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 45
	$ConLabel2.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 83
	$ConLabel2.Size = $System_Drawing_Size
	$ConLabel2.Text = "Password"

	$ConForm1.Controls.Add($ConLabel2)

	$ConLabel1.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 19
	$ConLabel1.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 83
	$ConLabel1.Size = $System_Drawing_Size
	$ConLabel1.Text = "User Account"

	$ConForm1.Controls.Add($ConLabel1)


	$ConButton2.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 12
	$System_Drawing_Point.Y = 120
	$ConButton2.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 75
	$ConButton2.Size = $System_Drawing_Size
	$ConButton2.TabIndex = 4
	$ConButton2.Text = "Test"
	$ConButton2.UseVisualStyleBackColor = $True
	$ConButton2.add_Click($handler_button2_Click)

	$ConForm1.Controls.Add($ConButton2)


	$ConButton1.DataBindings.DefaultDataSourceUpdateMode = 0

	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 213
	$System_Drawing_Point.Y = 164
	$ConButton1.Location = $System_Drawing_Point
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 23
	$System_Drawing_Size.Width = 75
	$ConButton1.Size = $System_Drawing_Size
	$ConButton1.TabIndex = 6
	$ConButton1.Text = "Cancel"
	$ConButton1.UseVisualStyleBackColor = $True
	$ConButton1.add_Click($handler_button1_Click)

	$ConForm1.Controls.Add($ConButton1)

	#endregion Generated Form Code

	#Save the initial state of the form
	$ConInitialFormWindowState = $ConForm1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$ConForm1.add_Load($OnLoadForm_StateCorrection)
	#Show the Form
	$ConForm1.ShowDialog()| Out-Null
	return $ConnectionDetails
}

$ConnectionDetails = ConnectionDetails
if ($ConnectionDetails.Password)
{
    SCCMCollectionSearch $ConnectionDetails
}



