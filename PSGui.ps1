<#
.SYNOPSIS  
    GUI to asynchronously execute powershell jobs from commands
.DESCRIPTION
	PSGui is very much what the name implies.
    It makes executing and keeping track of all executed commands quite a bit easier,
    especially if the user does not have a lot of powershell knowledge.

    Select a command, fill in all required data and click execute.
    A double click at the finished (running or failed job) shows the output of the job.
    Right clicking on a not running job will delete it.
.EXAMPLE
    If you can read this in the GUI there is not much you'll have to do

    Else: Call the script without parameters via:
    PS C:\<path> .\PSGui.ps1
.NOTES:
    AUTHOR:
        Dominik Schmidt
    VERSION:
        1.3
    CHANGELOG:
        V1.0 (07.03.2017; DS): First version
        V1.1 (10.03.2017; DS): Shift-Clicking on a datagridview entry will pipe the output of the job to a Out-GridView (will do nothing, if job has no results)
        V1.2 (13.03.2017; DS): Moved some functions around, added some missing function descriptions.
        V1.3 (15.03.2017; DS): Boolean values as parameter now generate a listbox, too. (Shouldn't be that common. That what switch parameters are for. The newest configmr commandlets have removed this)
                               Pressing escape closes the main and results windows
    SETUP:
        1. Create a subfolder called "PSGui"
        2. Drop all module files with functions that you want to use in that subfolder
        3. Create a text file in that subfolder and call it "commands.txt" (see: $script:commandsfile)
        4. Fill that text file with one command name per line (can be commands from the modules or default commands like Test-Connection)
        5. ???
        6. Profit
    TODO:
        - What about commands with dependencies in other modules or modules that can only be run in a specified drive? (like the sccm commandlets) -> use #Requires: https://technet.microsoft.com/de-de/library/hh847765.aspx
        - Display default values as cue text instead of variable type (how to get them though???)
        - Scrollbar on tabpage that contains the command parameter controls (So if a command has loads of parameters the user would not have to resize the window)
#>


####################################
#region    SCRIPT VARIABLES        #
####################################

# Main variables
$script:assetfolder  = "$PSScriptRoot\PSGUIData"     # the path to all .psm1 files that should be imported
$script:commandsfile = "$assetfolder\commands.txt"   # the textfile with the list of all displayed commands
$script:jobresultrefreshratems = 500                 # less = jobs will be shown as finished faster, but needs more CPU to compute

# Visuals
$script:formname = "PSGui"                           # The title of the form
$script:formiconpath = "$assetfolder\ps.ico"         # The path to the form icon
$script:defaultjobname = "Unnamedjob_"               # The default name of the jobs that were not given a name
$script:initialtabheader = "Message of the day"      # The initial text in the gui tabcontrol
$script:initialtabtexts = @("A day for firm decisions!  Or is it?",      # One of those will be picked at random to display at startup. Use this as branding, MOTD, etc...
                            "A few minutes of grace before the madness begins again.",
                            "Are you making all this up as you go along?",
                            "Avoid reality at all costs.",
                            "Bank error in your favor.  Collect €200.",
                            "Things won't get any better so get used to it.",
                            "Beware of low-flying butterflies.",
                            "Caution: breathing may be hazardous to your health.",
                            "Caution: Keep out of reach of children.",
                            "Don't hate yourself in the morning - sleep till noon.",
                            "Don't let your mind wander - it's too little to be let out alone.",
                            "Don't read everything you believe.",
                            "Don't relax! It's only your tension that's holding you together.",
                            "Don't tell any big lies today. Small ones can be just as effective.",
                            "Don't you wish you had more energy... or less ambition?",
                            "Expect the worst, it's the least you can do.",
                            "Give thought to your reputation. Consider changing name and moving to a new town.",
                            "If you learn one useless thing every day, in a single year you'll learn 365 useless things.",
                            "If you stand on your head, you will get footprints in your hair.",
                            "If you think yesterday was a drag, wait till you see what happens tomorrow!",
                            "Is that really YOU that is reading this?",
                            "It may or may not be worthwhile, but it still has to be done.",
                            "It was all so different before everything changed.",
                            "It was all so the same before nothing changed.",
                            "It's lucky you're going so slowly, because you're going in the wrong direction.",
                            "Just because the message may never be received does not mean it is not worth sending.",
                            "Just to have it is enough.",
                            "Keep emotionally active.  Cater to your favorite neurosis.",
                            "Let me put it this way: today is going to be a learning experience.",
                            "Look afar and see the end from the beginning.",
                            "Love is in the offing. Be affectionate to one who adores you.",
                            "Never look up when dragons fly overhead.",
                            "Never reveal your best argument.",
                            "Perfect day for scrubbing the floor and other exciting things.",
                            "So this is it. We're going to die.",
                            "How unusual!",
                            "Someone whom you reject today, will reject you tomorrow.",
                            "Today is the first day of the rest of the mess.",
                            "Today is the first day of the rest of your life.",
                            "Today is the last day of your life so far.",
                            "Tomorrow will be cancelled due to lack of interest.",
                            "You are fighting for survival in your own special way.",
                            "You are only young once, but you can stay immature indefinitely.",
                            "You are taking yourself far too seriously.",
                            "You can create your own opportunities this week. Blackmail a senior executive.",
                            "You can rent this space for only 2€ a week.",
                            "You don't become a failure until you're satisfied with being one.",
                            "You have been selected for a secret mission.",
                            "You have the capacity to learn from mistakes.  You'll learn a lot today.",
                            "You should emulate your heroes, but don't carry it too far. Especially if they are dead.",
                            "You should go home.",
                            "You single-handedly fought your way into this hopeless mess.",
                            "You will be a winner today. Pick a fight with a four-year-old.",
                            "You will soon forget this.",
                            "You will wish you hadn't.",
                            "You work very hard. Don't try to think as well.",
                            "You worry too much about your job. Stop it. You are not paid enough to worry.",
                            "You'll be sorry...",
                            "You'll feel much better once you've given up hope.",
                            "You're working under a slight handicap. You happen to be human.",
                            "Your life would be very empty if you had nothing to regret.",
                            "Your motives for doing whatever good deed you may have in mind will be misinterpreted by somebody.",
                            "Your reasoning is excellent - it's only your basic assumptions that are wrong.",
                            "Your true value depends entirely on what you are compared with."
                            )

# General variables: caches and stuff
$script:cmdletnames = @()                            # List of all available commands. These names are fetched from $commandsfile and checked for availability after the import of all modules
[hashtable] $script:serializedcommandcache = @{}     # Everytime a command is selected in the gui the serialized command info will be searched for in this list.
                                                     # If it does not jet exist (first time clicking that command) it will be queried and cached in here for later. -> performance boost
                                                     # Format: @{ $commandname, SerializedCommandInfo }
[string] $script:selectedcommand = $PSCommandPath    # The name of the currently selected command in the commandslist. Until a command in the list is selected this is the path to this script to show help.
[hashtable] $script:controllist = @{}                # Here all the controls that were generated for all parametersets of the currently selected command will be stored.
                                                     # Format: @{ $parametersetname, @($control1, $control2,...) }
[hashtable] $script:controllistjobname = @{}         # The textboxes of each commands parameter sets
                                                     # Format: @{ $parametersetname, $textboxcontrol }
$script:serializedcommandinfo = ""                   # The serialized command info of the currently selected command. A little bit faster because it does not have to be searched in the hashtable on executing

# Job variables
$script:nextjobid   = 0    # Counter of executed jobs. Will be used if a job was not given a name to generate default job names with a counter. See: $script:defaultjobname
$script:currentjobs = @()  # Buffer of all current local ps jobs
$jobcommands        = @{}  # PSJobs don't store their parameters as property (only the command). For nice display how the command line was called, especially for not as powershell proficient users as example.
                           # Some stuff (like securestrings) should NEVER be displayed as parameter. Using this method that can be disguised, too.
                           # Format: @{ Jobid, parameters of job }


####################################
#endregion SCRIPT VARIABLES        #
####################################



####################################
#region    GENERATE MAIN FORM      #
####################################


<#
.SYNOPSIS
    Generates the main form that is presented to the user on program start
#>
Function Generate-MainForm {

    ###########################################################################################
    # Code in part generated by: SAPIEN Technologies PrimalForms (Community Edition) v1.0.8.0 #
    ###########################################################################################

    #----------------------------------------------
    #region IMPORT THE ASSEMBLIES
    #----------------------------------------------

    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

    #----------------------------------------------
    #endregion IMPORT THE ASSEMBLIES
    #----------------------------------------------


    #----------------------------------------------
    #region FORM OBJECTS
    #----------------------------------------------

    $f_main = New-Object System.Windows.Forms.Form
    $splitter1 = New-Object System.Windows.Forms.Splitter
    $panel1 = New-Object System.Windows.Forms.Panel
    $tc_parametersets = New-Object System.Windows.Forms.TabControl
    $tp_1 = New-Object System.Windows.Forms.TabPage
    $l_motd = New-Object System.Windows.Forms.Label
    $b_execute = New-Object System.Windows.Forms.Button
    $ll_cmdletname = New-Object System.Windows.Forms.LinkLabel
    $panel2 = New-Object System.Windows.Forms.Panel
    $lb_commands = New-Object System.Windows.Forms.ListBox
    $tb_search = New-Object System.Windows.Forms.CueTextBox
    $dgv_results = New-Object System.Windows.Forms.DataGridView
    $stb_bottom = New-Object System.Windows.Forms.StatusBar
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    #----------------------------------------------
    #endregion FORM OBJECTS
    #----------------------------------------------


    #----------------------------------------------
    #region FORM CODE
    #----------------------------------------------

    # MAIN FORM
    $f_main.Font = New-Object System.Drawing.Font("Lucida Console",9,0,3,1)
    $f_main.Text = $script:formname
    $f_main.Name = "f_main"
    $f_main.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($script:formiconpath)
    $f_main.AutoSizeMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 610
    $System_Drawing_Size.Height = 500
    $f_main.ClientSize = $System_Drawing_Size

    # SPLITTER BETWEEN COMMAND LIST AND TAB CONTROL
    $f_main.Controls.Add($splitter1)

    # PANEL FOR LINKLABEL AND TABCONTROL ON THE RIGHT
    $panel1.Dock = 5
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 600
    $System_Drawing_Size.Height = 400
    $panel1.Size = $System_Drawing_Size
    $panel1.AutoScroll = $true    # TODO: Does not seem to work. No scrollbar appears on resizing
    $f_main.Controls.Add($panel1)

    # TABCONTROL FOR PARAMETER SETS
    $tc_parametersets.TabIndex = 9
    $tc_parametersets.Dock = 5
    $tc_parametersets.Appearance = 2
    $tc_parametersets.Name = "tc_parametersets"
    $tc_parametersets.SelectedIndex = 0
    $panel1.Controls.Add($tc_parametersets)

    # INITIAL TAB PAGE
    $tp_1.AutoScroll = $True
    $tp_1.Text = $script:initialtabheader
    $tc_parametersets.Controls.Add($tp_1)

    # INITIAL CENTERED TEXT LABEL ON PROGRAM START
    $l_motd.Dock = 5
    $l_motd.TextAlign = 32
    $l_motd.Text = $Script:initialtabtexts[$(Get-Random -Maximum (0..($Script:initialtabtexts.Count-1)))] # random message from array
    $l_motd.Name = "l_motd"
    $tp_1.Controls.Add($l_motd)

    # LINKLABEL THAT SHOWS PARAMETER NAME
    $ll_cmdletname.Dock = 1
    $ll_cmdletname.TabIndex = 2
    $ll_cmdletname.Text = "Main Help"
    $ll_cmdletname.Name = "ll_cmdletname"
    $System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding # TODO: WHAT'S UP WITH THE LABEL THAT WE NEED A PADDING HERE?!
    $System_Windows_Forms_Padding.Bottom = 0
    $System_Windows_Forms_Padding.Left = 3
    $System_Windows_Forms_Padding.Right = 3
    $System_Windows_Forms_Padding.Top = 3
    $ll_cmdletname.Margin = $System_Windows_Forms_Padding
    $System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
    $System_Windows_Forms_Padding.Bottom = 0
    $System_Windows_Forms_Padding.Left = 5
    $System_Windows_Forms_Padding.Right = 0
    $System_Windows_Forms_Padding.Top = 5
    $ll_cmdletname.Padding = $System_Windows_Forms_Padding
    $panel1.Controls.Add($ll_cmdletname)

    # PANEL ON THE LEFT FOR COMMAND LISTBOX AND COMMAND SEARCHBOX
    $panel2.Dock = 3
    $f_main.Controls.Add($panel2)

    # LISTBOX WITH ALL AVAILABLE COMMANDS
    $lb_commands.Dock = 5
    $lb_commands.Items.Add("If you can read this something is wrooooong") | Out-Null
    $lb_commands.ItemHeight = 12
    $lb_commands.Name = "lb_commands"
    $lb_commands.TabIndex = 1
    $panel2.Controls.Add($lb_commands)

    # TEXT BOX ABOVE COMMAND LIST FOR SEARCHING
    $tb_search.Dock = 1
    $tb_search.Name = "tb_search"
    $tb_search.TabIndex = 0
    $tb_search.Cue = "Filter cmdlets..."
    $panel2.Controls.Add($tb_search)

    $s_horizontal = New-Object System.Windows.Forms.Splitter
    $s_horizontal.Dock = 2
    $f_main.Controls.Add($s_horizontal)

    # THE CLUSTERFUCK THAT IS THE DATA GRID VIEW AND IT'S DEFAULT ITEMS
    $System_Windows_Forms_DataGridViewTextBoxColumn_73 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $System_Windows_Forms_DataGridViewTextBoxColumn_73.DataPropertyName = "State"
    $System_Windows_Forms_DataGridViewTextBoxColumn_73.HeaderText = "State"
    $System_Windows_Forms_DataGridViewTextBoxColumn_73.FillWeight = 1
    $System_Windows_Forms_DataGridViewTextBoxColumn_73.AutoSizeMode = 16
    $dgv_results.Columns.Add($System_Windows_Forms_DataGridViewTextBoxColumn_73) | Out-Null

    $System_Windows_Forms_DataGridViewTextBoxColumn_74 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $System_Windows_Forms_DataGridViewTextBoxColumn_74.DataPropertyName = "StartTime"
    $System_Windows_Forms_DataGridViewTextBoxColumn_74.HeaderText = "StartTime"
    $System_Windows_Forms_DataGridViewTextBoxColumn_74.FillWeight = 2
    $System_Windows_Forms_DataGridViewTextBoxColumn_74.AutoSizeMode = 16
    $dgv_results.Columns.Add($System_Windows_Forms_DataGridViewTextBoxColumn_74) | Out-Null

    $System_Windows_Forms_DataGridViewTextBoxColumn_75 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $System_Windows_Forms_DataGridViewTextBoxColumn_75.DataPropertyName = "Name"
    $System_Windows_Forms_DataGridViewTextBoxColumn_75.HeaderText = "Name"
    $System_Windows_Forms_DataGridViewTextBoxColumn_75.FillWeight = 3
    $System_Windows_Forms_DataGridViewTextBoxColumn_75.AutoSizeMode = 16
    $dgv_results.Columns.Add($System_Windows_Forms_DataGridViewTextBoxColumn_75) | Out-Null

    $System_Windows_Forms_DataGridViewTextBoxColumn_76 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $System_Windows_Forms_DataGridViewTextBoxColumn_76.DataPropertyName = "Command"
    $System_Windows_Forms_DataGridViewTextBoxColumn_76.HeaderText = "Command"
    $System_Windows_Forms_DataGridViewTextBoxColumn_76.FillWeight = 4
    $System_Windows_Forms_DataGridViewTextBoxColumn_76.AutoSizeMode = 16
    $dgv_results.Columns.Add($System_Windows_Forms_DataGridViewTextBoxColumn_76) | Out-Null

    $System_Windows_Forms_DataGridViewCellStyle_76 = New-Object System.Windows.Forms.DataGridViewCellStyle
    $System_Windows_Forms_DataGridViewCellStyle_76.BackColor = [System.Drawing.Color]::FromArgb(255,0,36,82)
    $dgv_results.AlternatingRowsDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_76

    $dgv_results.ShowEditingIcon = $False
    $dgv_results.AllowUserToAddRows = $False
    $dgv_results.AutoGenerateColumns = $false
    $dgv_results.Name = "dgv_results"
    $dgv_results.AllowUserToOrderColumns = $True
    $dgv_results.SelectionMode = 1  
    $dgv_results.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
    $dgv_results.BackgroundColor = [System.Drawing.Color]::FromArgb(255,0,36,82)
    $dgv_results.Dock = 2
    $dgv_results.AllowUserToResizeRows = $false
    $dgv_results.ColumnHeadersHeightSizeMode = 1
    $dgv_results.EditMode = 4
    $dgv_results.StandardTab = $True
    $dgv_results.GridColor = [System.Drawing.Color]::FromArgb(255,0,36,82)
    $dgv_results.RowHeadersVisible = $False
    $dgv_results.TabIndex = 0
    $dgv_results.RowHeadersBorderStyle = 4
    $dgv_results.MultiSelect = $True
    $dgv_results.ColumnHeadersHeight = 18
    $System_Windows_Forms_DataGridViewCellStyle_78 = New-Object System.Windows.Forms.DataGridViewCellStyle
    $System_Windows_Forms_DataGridViewCellStyle_78.BackColor = [System.Drawing.Color]::FromArgb(255,0,36,82)
    $dgv_results.RowsDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_78
    $dgv_results.CellBorderStyle = 4
    $f_main.Controls.Add($dgv_results)

    # STATUS BAR AT THE BOTTOM
    $stb_bottom.Name = "stb_bottom"
    $stb_bottom.Text = "Total: 0 jobs | Running: 0 - Completed: 0 - Failed: 0"
    $f_main.Controls.Add($stb_bottom)

    #----------------------------------------------
    #endregion FORM CODE
    #----------------------------------------------


    #----------------------------------------------
    #region INITIALISATION
    #----------------------------------------------

    # Get all available commands and stuff
    Initialize-Data
    # Show all available commands in listbox
    Apply-CMDLetListFilter -Filter ""
    # Register all eventhandlers (clicking etc..)
    Register-EventHandlers

    # Set the initial selection to the search textbox
    $tb_search.Select()

    # REGISTER THE AUTOUPDATE TIMER THAT DISPLAYS CURRENT JOB STATUS
    # Has to be a [System.Timers.Timer] and no winforms timer, since it can't synchronize with the gui and thus would only fire on gui update
    $checktimer = New-Object -TypeName System.Timers.Timer
    $checktimer.Interval = $script:jobresultrefreshratems
    $checktimer.SynchronizingObject = $f_main
    $checktimer.Start()

    # Everytime the timer ticks update the joblist
    $checktimer.add_Elapsed({ 
        $checktimer.Stop()
        Write-Debug "checktimer tick."
        Refresh-JobList
        $checktimer.Start()
    })

    #----------------------------------------------
    #endregion INITIALISATION
    #----------------------------------------------

    # Correct the initial state of the form to prevent the .Net maximized form issue
    $OnLoadForm_StateCorrection = {
	    $f_main.WindowState = $InitialFormWindowState
    }

    # Save the initial state of the form
    $InitialFormWindowState = $f_main.WindowState

    # Init the OnLoad event to correct the initial state of the form
    $f_main.add_Load($OnLoadForm_StateCorrection)

    # SHOW THE FORM
    $f_main.ShowDialog() | Out-Null
}


<#
.SYNOPSIS
    Show the output of a job in a gui with separate list for output, warnings and errors
#>
Function Generate-JobResultsWindow {

    Param (
        # The job to display information from
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        $Job
    )

    # Inner function with the sole purpose of displaying the data
    Function Show-JobResultsForm {

        Param (
            [Parameter(Mandatory=$true)]  [string] $JobName,
            [Parameter(Mandatory=$false)] [PSCustomObject[]] $JobOutput,
            [Parameter(Mandatory=$false)] [System.Management.Automation.WarningRecord[]] $JobWarnings,
            [Parameter(Mandatory=$false)] [System.Management.Automation.ErrorRecord[]]   $JobErrors
        )

        if ($JobOutput) {
            $DisplayOutput = $JobOutput | Out-String
        } else {
            $DisplayOutput = "<None>"
        }


        ########################################################################
        # Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.8.0
        # Generated On: 23.02.2017 15:04
        # Generated By: SchmidtD045D
        ########################################################################

        #region Import the Assemblies
        #[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
        #[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
        #endregion

        #region Generated Form Objects
        $f_jobresults = New-Object System.Windows.Forms.Form
        $tb_output = New-Object System.Windows.Forms.RichTextBox
        $l_warnings = New-Object System.Windows.Forms.Label
        $splitter1 = New-Object System.Windows.Forms.Splitter
        $tb_warnings = New-Object System.Windows.Forms.RichTextBox
        $l_errors = New-Object System.Windows.Forms.Label
        $splitter2 = New-Object System.Windows.Forms.Splitter
        $tb_errors = New-Object System.Windows.Forms.RichTextBox
        $l_output = New-Object System.Windows.Forms.Label
        $sb_summary = New-Object System.Windows.Forms.StatusBar
        $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
        #endregion Generated Form Objects

        #region Generated Form Code
        $f_jobresults.Font = New-Object System.Drawing.Font("Lucida Console",9,0,3,1)
        $f_jobresults.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $f_jobresults.Text = "JobResults: $JobName"
        $f_jobresults.Name = "f_jobresults"
        $f_jobresults.DataBindings.DefaultDataSourceUpdateMode = 0
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 539
        $System_Drawing_Size.Height = 621
        $f_jobresults.ClientSize = $System_Drawing_Size

        $tb_output.Dock = 5
        $tb_output.Multiline = $True
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 539
        $System_Drawing_Size.Height = 404
        $tb_output.Size = $System_Drawing_Size
        $tb_output.DataBindings.DefaultDataSourceUpdateMode = 0
        $tb_output.ReadOnly = $True
        $tb_output.Name = "tb_output"
        $tb_output.BackColor = [System.Drawing.Color]::FromArgb(255,0,36,82)
        $tb_output.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = 0
        $System_Drawing_Point.Y = 23
        $tb_output.Location = $System_Drawing_Point
        $tb_output.TabIndex = 4
        $tb_output.Text = $DisplayOutput

        $f_jobresults.Controls.Add($tb_output)

        if ($JobWarnings) {

            $l_warnings.TabIndex = 1
            $l_warnings.Dock = 2
            $l_warnings.TextAlign = 16
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 23
            $l_warnings.Size = $System_Drawing_Size
            $l_warnings.Text = "Warnings"
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 427
            $l_warnings.Location = $System_Drawing_Point
            $l_warnings.DataBindings.DefaultDataSourceUpdateMode = 0
            $l_warnings.Name = "l_warnings"

            $f_jobresults.Controls.Add($l_warnings)

            $splitter1.Dock = 2
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 3
            $splitter1.Size = $System_Drawing_Size
            $splitter1.TabIndex = 5
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 450
            $splitter1.Location = $System_Drawing_Point
            $splitter1.DataBindings.DefaultDataSourceUpdateMode = 0
            $splitter1.TabStop = $False
            $splitter1.Name = "splitter1"

            $f_jobresults.Controls.Add($splitter1)

            $tb_warnings.Dock = 2
            $tb_warnings.Multiline = $True
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 60
            $tb_warnings.Size = $System_Drawing_Size
            $tb_warnings.DataBindings.DefaultDataSourceUpdateMode = 0
            $tb_warnings.ReadOnly = $True
            $tb_warnings.Name = "tb_warnings"
            $tb_warnings.BackColor = [System.Drawing.Color]::FromArgb(255,162,90,0)   #(255,255,142,0)
            $tb_warnings.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 453
            $tb_warnings.Location = $System_Drawing_Point
            $tb_warnings.TabIndex = 4
            $tb_warnings.Text = $JobWarnings

            $f_jobresults.Controls.Add($tb_warnings)

        }

        if ($JobErrors) {

            $l_errors.TabIndex = 3
            $l_errors.Dock = 2
            $l_errors.TextAlign = 16
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 23
            $l_errors.Size = $System_Drawing_Size
            $l_errors.Text = "Errors"
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 513
            $l_errors.Location = $System_Drawing_Point
            $l_errors.DataBindings.DefaultDataSourceUpdateMode = 0
            $l_errors.Name = "l_errors"

            $f_jobresults.Controls.Add($l_errors)

            $splitter2.Dock = 2
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 3
            $splitter2.Size = $System_Drawing_Size
            $splitter2.TabIndex = 6
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 536
            $splitter2.Location = $System_Drawing_Point
            $splitter2.DataBindings.DefaultDataSourceUpdateMode = 0
            $splitter2.TabStop = $False
            $splitter2.Name = "splitter2"

            $f_jobresults.Controls.Add($splitter2)

            $tb_errors.Dock = 2
            $tb_errors.Multiline = $True
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 539
            $System_Drawing_Size.Height = 60
            $tb_errors.Size = $System_Drawing_Size
            $tb_errors.DataBindings.DefaultDataSourceUpdateMode = 0
            $tb_errors.ReadOnly = $True
            $tb_errors.Name = "tb_errors"
            $tb_errors.BackColor = [System.Drawing.Color]::FromArgb(255,107,4,16)
            $tb_errors.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
            $System_Drawing_Point = New-Object System.Drawing.Point
            $System_Drawing_Point.X = 0
            $System_Drawing_Point.Y = 539
            $tb_errors.Location = $System_Drawing_Point
            $tb_errors.TabIndex = 2
            $tb_errors.Text = $JobErrors

            $f_jobresults.Controls.Add($tb_errors)

        }

        $l_output.TabIndex = 0
        $l_output.Dock = 1
        $l_output.TextAlign = 16
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 539
        $System_Drawing_Size.Height = 23
        $l_output.Size = $System_Drawing_Size
        $l_output.Text = "Output:"
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = 0
        $System_Drawing_Point.Y = 0
        $l_output.Location = $System_Drawing_Point
        $l_output.DataBindings.DefaultDataSourceUpdateMode = 0
        $l_output.Name = "l_output"

        $f_jobresults.Controls.Add($l_output)

        $sb_summary.Name = "sb_summary"
        $sb_summary.Text = "Results: $($JobOutput.Count) - Warnings: $($JobWarnings.Count) - Errors: $($JobErrors.Count)"
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 539
        $System_Drawing_Size.Height = 22
        $sb_summary.Size = $System_Drawing_Size
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = 0
        $System_Drawing_Point.Y = 599
        $sb_summary.Location = $System_Drawing_Point
        $sb_summary.DataBindings.DefaultDataSourceUpdateMode = 0
        $sb_summary.TabIndex = 0

        $f_jobresults.Controls.Add($sb_summary)

        #endregion Generated Form Code

        # Pressing escape closes the window
        $f_jobresults.KeyPreview = $true
        $f_jobresults.Add_KeyDown({

            if ($_.KeyCode -eq "Escape") {
                $f_jobresults.Close()
            }
        })

        #Save the initial state of the form
        $InitialFormWindowState = $f_jobresults.WindowState
        #Init the OnLoad event to correct the initial state of the form
        $f_jobresults.add_Load($OnLoadForm_StateCorrection)
        #Show the Form
        $f_jobresults.ShowDialog()| Out-Null

    }

    $jobname = $Job.Name
    $joboutput = Receive-Job -Job $Job -Keep -WarningVariable jobwarnings -ErrorVariable joberrors -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

    #Call the Function
    Show-JobResultsForm -JobName $jobname -JobOutput $joboutput -JobWarnings $jobwarnings -JobErrors $joberrors

}


<#
.SYNOPSIS
    Checks to see if a key or keys are currently pressed.
.DESCRIPTION
    Checks to see if a key or keys are currently pressed. If all specified keys are pressed then will return true, but if 
    any of the specified keys are not pressed, false will be returned.
.PARAMETER Keys
    Specifies the key(s) to check for. These must be of type "System.Windows.Forms.Keys"
.EXAMPLE
    PS> Test-KeyPress -Keys ControlKey

    Check to see if the Ctrl key is pressed
.EXAMPLE
    PS> Test-KeyPress -Keys ControlKey,Shift

    Test if Ctrl and Shift are pressed simultaneously (a chord)
.LINK
    Uses the Windows API method GetAsyncKeyState to test for keypresses
    http://www.pinvoke.net/default.aspx/user32.GetAsyncKeyState

    The above method accepts values of type "system.windows.forms.keys"
    https://msdn.microsoft.com/en-us/library/system.windows.forms.keys(v=vs.110).aspx
.INPUTS
    System.Windows.Forms.Keys
.OUTPUTS
    System.Boolean
#>
function Test-KeyPress {
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Keys[]] $Keys
    )
    
    # load assembly if not yet loaded
    # use the User32 API to define a keypress datatype
    $signature = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
'@
    $api = Add-Type -MemberDefinition $Signature -Name 'Keypress' -Namespace Keytest -PassThru -ErrorAction SilentlyContinue

    # test if each key in the collection is pressed
    $result = foreach ($key in $Keys) {
        [bool]($api::GetAsyncKeyState($key) -eq -32767)
    }
    
    # if all are pressed, return true, if any are not pressed, return false
    $result -notcontains $false
}


####################################
#endregion GENERATE MAIN FORM      #
####################################
 


####################################
#region    FORM EVENT FUNCTIONS    #
####################################


<#
.SYNOPSIS
    Maps controls for <<parametersetname : controls for parameterset>>
#>
Function Add-ToControlList([string] $ParameterSetName, [System.Windows.Forms.Control] $Control) {
    if ($script:controllist.ContainsKey($ParameterSetName)) {
        $script:controllist[$ParameterSetName] += $Control
    } else {
        $script:controllist.Add($ParameterSetName, @($Control))
    }
}


<#
.SYNOPSIS
    Registers gui eventhandlers
.DESCRIPTION
    - searchbox to filter command listbox
    - selection of command from listbox to update gui with tabs for parametersets and input objects for each parameter
    - Link for the command name label to oppen command help
    - doublecklick on a datagridview entry to open job results
    - right click on datagridview to delete not running job
#>
Function Register-EventHandlers {

    # Pressing escape closes the window
    $f_main.KeyPreview = $true
    $f_main.Add_KeyDown({

        if ($_.KeyCode -eq "Escape") {
            $f_main.Close()
        }
    })

    # typing in the search textbox: filter visible commands
    $tb_search.Add_TextChanged({ 
        Apply-CMDLetListFilter -Filter $($tb_search.Text)
    })

    # if a command is selected in the listbox: display command stuff
    $lb_commands.Add_SelectedValueChanged({
        Select-CommandFromList $lb_commands.SelectedItem
    })
    # focus the command listbox on mouse enter so it can be scrolled without kaving to click it beforehand. (Dumb Winforms...)
    $lb_commands.Add_MouseEnter({
        $this.Focus()
    })

    # MOUSE SUPPORT
    # if the label with the command name is clicked: open command help
    $ll_cmdletname.Add_Click({ 
        Write-Verbose "Opening help for command $script:selectedcommand."
        Get-Help -Name $script:selectedcommand -ShowWindow
    })

    # Double click in a datagridview entry opens result window
    # Shift click opens the result in a gridview instead!
    $dgv_results.Add_CellDoubleClick({

        # get job
        $job = $currentjobs[$_.RowIndex]

        if (Test-KeyPress -Keys ShiftKey) {
            
            Receive-Job -Job $job -Keep | Out-GridView -Title "Result of job $($job.Name)"
        } else {
            Generate-JobResultsWindow -Job $job
        }
    })

    # right click deletes non running jobs
    $dgv_results.Add_CellMouseClick({
        if ($_.Button -eq "Right") {
            $selectedjob = $currentjobs[$_.RowIndex]
            Write-Verbose "Right clicked job `"$($selectedjob.Name)`""

            #if ($selectedjob.State -ne "Running") {
                $selectedjob | Remove-Job -Force
                Refresh-JobList
            #}
        }
    })

    # KEYBOARD SUPPORT:
    # Pressing enter/space when a datagridview entry is selected: also open results window
    # Delete removes athe selected job
    # this is done by overwriting the default keypress event (default: next row on enter)
    $dgv_results.Add_KeyDown({
        
        if ($_.KeyData -match "Return" -or $_.KeyData -match "Space") {
            # disable next row
            $_.SuppressKeyPress = $true

            # get job
            $job = $currentjobs[$dgv_results.SelectedRows.Index]

            # and open result window
            if (Test-KeyPress -Keys ShiftKey) {
                Receive-Job -Job $job -Keep | Out-GridView -Title "Result of job $($job.Name)"
            } else {
                Generate-JobResultsWindow -Job $job
            }
        # On entf: remove job
        } elseif ($_.KeyData -eq "Delete") {
            # get job
            $job = $currentjobs[$dgv_results.SelectedRows.Index]

            $job  | Remove-Job -Force
            Refresh-JobList
        }
    })

}


<#
.SYNOPSIS
    Filters the list of commands to only those matching $Filter parameter.
    Used for textbox to filter listbox opf commands.
#>
Function Apply-CMDLetListFilter ($Filter) {
    $lb_commands.Items.Clear()
    $script:cmdletnames | Where-Object { $_ -match $Filter } | Sort-Object | ForEach-Object {
        $lb_commands.Items.Add($_) | Out-Null
    }
}


<#
.SYNOPSIS
    Sets the gui to show the given command
.DESCRIPTION
    - Change linklabel text
    - Serializes command information
    - Calls the function generating the gui
#>
Function Select-CommandFromList ($CommandName) {
    # Store the selected command name in variable
    $script:selectedcommand = $CommandName

    # Update the linklabel text to commandname (the link itself is always dependent on $selectedcommand. No need to update)
    $ll_cmdletname.Text = $CommandName
    
    # Is the command info already stored in the serialisation cache?
    if ($script:serializedcommandcache.ContainsKey($CommandName)) {
        # Yes: Use stored info. A lot faster!
        $script:serializedcommandinfo = $script:serializedcommandcache[$CommandName]
    } else {
        # No: Serialize info of command and store it for later
        $script:serializedcommandinfo = Get-SerializedCommandInfo -Commands $CommandName -NoWellKnownParameters
        $script:serializedcommandcache.Add($CommandName, $script:serializedcommandinfo)
    }

    # Call function to generate gui
    Generate-GuiForCommand -SerializedCommandInfo $script:serializedcommandinfo
}



<#
.SYNOPSIS
    Generates the layout matching for everything the selected command has
.DESCRIPTION
    - A tab for every parameterset
    - Checkboxes for switches
    - Dropdowns for enums
    - Textboxes for strings, ints, doubles etc... (everything else essentially)
#>
Function Generate-GuiForCommand() {

    Param ($SerializedCommandInfo)

    # Clear stored controls
    $script:controllist.Clear() 
    $script:controllistjobname.Clear()

    # Remove all currently existing tabpages
    #Write-Host $tc_parametersets.SelectedIndex
    #foreach ($tab in $tc_parametersets.TabPages) {
    #    $tc_parametersets.TabPages.Remove($tab)
    #}
    $tc_parametersets.TabPages.Clear()
    #$tc_parametersets.
    #Write-Host $tc_parametersets.SelectedIndex

    # Don't refresh the form for every tiny change. Only At the end
    $tc_parametersets.SuspendLayout()

    # for each parameterset: create a tab page
    # and fill this tabpage with controls matching the input of each parameter type
    $currenttab = 0
    foreach ($parameterset in $($SerializedCommandInfo.ParameterSets | Sort-Object IsDefault)) {


        # Get parametersetname and all parameters
        $parametersetname = $parameterset.Name
        $parametersetparams = $parameterset.Parameters

        # Add tabpage matching parameter set name
        if ($parametersetname -eq "__AllParameterSets") {
            $tabname = "Default"
        } else {
            $tabname = $parametersetname
        }
        $tc_parametersets.TabPages.Add($tabname)
        $workingtabpage = $tc_parametersets.TabPages[$currenttab]
        $workingtabpage.AutoScroll = $True

        # Create two panels: one for all labels on the left and one for all input forms on the right
        $p_commandparamcontrols = New-Object System.Windows.Forms.Panel
        $p_commandparamnames = New-Object System.Windows.Forms.Panel
        $s_commandparams = New-Object System.Windows.Forms.Splitter
        

        $p_commandparamcontrols.Dock = 5 # Dock right
        $workingtabpage.Controls.Add($p_commandparamcontrols)

        $s_commandparams.Name = "s_execute"
        $workingtabpage.Controls.Add($s_commandparams)

        $p_commandparamnames.Dock = 3 # Dock left
        $workingtabpage.Controls.Add($p_commandparamnames)
        # Finished creating panels

        $i = 0
        # Reverse the order of the params -> When using docking style the last one will be docked on top. In that case the last one will be the downmost one
        # Includes really, really weird sorting stuff, because vvv this vvv keeps the array reversed even after reassignment. 
        $orderedparams = $(($parametersetparams | ForEach-Object { [System.Tuple]::Create($_, $i); $i++ } | Sort-Object -Property Item2 -Descending).Item1) | Sort-Object -Property @{Expression=”IsMandatory”;Ascending=$true},@{Expression=”Position”;Ascending=$true}

        # Loop trough each parameter of the current parameter set
        $paramid = 0
        foreach ($param in $($orderedparams)) {

            # Get parameter data
            $paramname = $param.Name
            $parammandatory = $param.IsMandatory
            $paramtype = $param.ParameterType
            $paramvalidvalues = $param.ValidParamSetValues
            $paramisenum = $paramtype.IsEnum

            # Add a * to the param display name if it is a mandatory parameter
            if ($parammandatory) {
                $paramdesc = "$paramname*"
            } else {
                $paramdesc = $paramname
            }

            # What control should be generated for the parameter depending on type
            if ($paramtype.FullName -eq "System.Management.Automation.SwitchParameter" ) {
                # switch -> checkbox
                $paramcontrol = "checkbox"
            } elseif ($paramvalidvalues.Count -gt 0) {
                # list of valid values -> dropdown
                $paramcontrol = "dropdown"
            } elseif ($paramisenum) {
                # enum -> dropdown
                $paramcontrol = "enumdropdown"
            } elseif ($paramtype.FullName -eq "System.Boolean") {
                # boolean -> dropdown
                $paramcontrol = "booleandropdown"
            } elseif ($paramtype.FullName -eq "System.Security.SecureString") {
                # enum -> dropdown
                $paramcontrol = "textboxsecure"
            } else {
                # else -> textbox
                $paramcontrol = "textbox"
            }
            Write-Debug "Param: $paramname needs a $paramcontrol"

            # Create the controls

            # Label with name on the left panel
            $l_name = New-Object System.Windows.Forms.Label

            $l_name.AutoEllipsis = $true
            $System_Drawing_Size = New-Object System.Drawing.Size
            $System_Drawing_Size.Width = 190
            $System_Drawing_Size.Height = 20
            $l_name.Size = $System_Drawing_Size
            $l_name.Dock = 1
            $l_name.Text = $paramdesc
            $p_commandparamnames.Controls.Add($l_name)


            # main control
            switch ($paramcontrol) {
                'checkbox' {
                    $checkBox1 = New-Object System.Windows.Forms.CheckBox

                    $checkBox1.UseVisualStyleBackColor = $True
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 150
                    $System_Drawing_Size.Height = 20
                    $checkBox1.Size = $System_Drawing_Size
                    $checkBox1.Text = ""
                    $checkbox1.Dock = 1
                    $checkbox1.Tag = $paramname   # store parameter name in tag
                    $checkbox1.TabIndex = 10 + $parametersetparams.Count - $paramid

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $checkBox1
                    $p_commandparamcontrols.Controls.Add($checkBox1)
                }
                'dropdown' {
                    $comboBox1 = New-Object System.Windows.Forms.ComboBox

                    $comboBox1.FormattingEnabled = $True
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 200
                    $System_Drawing_Size.Height = 20
                    $comboBox1.Size = $System_Drawing_Size
                    $comboBox1.Tag = $paramname
                    $comboBox1.Dock = 1
                    $comboBox1.TabIndex = 10 + $parametersetparams.Count - $paramid

                    # Is the parameter mandatory? In the case iot isn't add an empty object to the list of choices. "not set"
                    if (-not $parammandatory) {
                        $comboBox1.Items.Add("")
                    }

                    # Add all valid values to the combobox
                    foreach ($paramvalidvalue in $paramvalidvalues) {
                        $comboBox1.Items.Add($paramvalidvalue)
                    }
                    $comboBox1.SelectedIndex = 0

                    # remove red color from missing mandatory
                    <#if ($parammandatory) {
                        $comboBox1.Add_SelectedIndexChanged({
                            if ($this.BackColor -ne [System.Drawing.Color]::FromName("Window")) {
                                $this.BackColor = [System.Drawing.Color]::FromName("Window")
                            }
                        })
                    }#>

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $comboBox1
                    $p_commandparamcontrols.Controls.Add($comboBox1)
                }
                'enumdropdown' {
                    $cb = New-Object System.Windows.Forms.ComboBox

                    $cb.FormattingEnabled = $True
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 200
                    $System_Drawing_Size.Height = 20
                    $cb.Size = $System_Drawing_Size
                    $cb.Name = "dd_$currenttab`_$paramname"
                    $cb.Tag = $paramname
                    $cb.Dock = 1
                    $cb.TabIndex = 10 + $parametersetparams.Count - $paramid

                    # Is the parameter mandatory? In the case iot isn't add an empty object to the list of choices. "not set"
                    if (-not $parammandatory) {
                        $cb.Items.Add("")
                    }

                    # Add all valid values to the combobox
                    foreach ($enumvalue in $paramtype.EnumValues) {
                        $cb.Items.Add($enumvalue)
                    }
                    $cb.SelectedIndex = 0

                    # remove red color from missing mandatory
                    if ($parammandatory) {
                        $cb.Add_SelectedIndexChanged({
                            if ($this.BackColor -ne [System.Drawing.Color]::FromName("Window")) {
                                $this.BackColor = [System.Drawing.Color]::FromName("Window")
                            }
                        })
                    }

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $cb
                    $p_commandparamcontrols.Controls.Add($cb)
                }
                'booleandropdown' {
                    $cb = New-Object System.Windows.Forms.ComboBox

                    $cb.FormattingEnabled = $True
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 200
                    $System_Drawing_Size.Height = 20
                    $cb.Size = $System_Drawing_Size
                    $cb.Name = "dd_$currenttab`_$paramname"
                    $cb.Tag = $paramname
                    $cb.Dock = 1
                    $cb.TabIndex = 10 + $parametersetparams.Count - $paramid

                    # Is the parameter mandatory? In the case iot isn't add an empty object to the list of choices. "not set"
                    if (-not $parammandatory) {
                        $cb.Items.Add("")
                    }

                    # Add all valid values to the combobox
                    $cb.Items.Add('$true')
                    $cb.Items.Add('$false')

                    $cb.SelectedIndex = 0

                    # remove red color from missing mandatory on selection change
                    if ($parammandatory) {
                        $cb.Add_SelectedIndexChanged({
                            if ($this.BackColor -ne [System.Drawing.Color]::FromName("Window")) {
                                $this.BackColor = [System.Drawing.Color]::FromName("Window")
                            }
                        })
                    }

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $cb
                    $p_commandparamcontrols.Controls.Add($cb)
                }
                'textboxsecure'  {
                    $tb = New-Object System.Windows.Forms.CueTextBox
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 200
                    $System_Drawing_Size.Height = 20
                    $tb.Size = $System_Drawing_Size
                    $tb.Dock = 1
                    $tb.Tag = $paramname
                    $tb.Cue = $paramtype.FullName
                    $tb.TabIndex = 10 + $parametersetparams.Count - $paramid
                    $tb.PasswordChar = "*"

                    # remove red color from missing mandatory
                    if ($parammandatory) {
                        $tb.Add_TextChanged({
                            if ($this.BackColor -ne [System.Drawing.Color]::FromName("Window")) {
                                $this.BackColor = [System.Drawing.Color]::FromName("Window")
                            }
                        })
                    }

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $tb
                    $p_commandparamcontrols.Controls.Add($tb)
                }
                'textbox'  {
                    $tb = New-Object System.Windows.Forms.CueTextBox
                    $System_Drawing_Size = New-Object System.Drawing.Size
                    $System_Drawing_Size.Width = 200
                    $System_Drawing_Size.Height = 20
                    $tb.Size = $System_Drawing_Size
                    $tb.Dock = 1
                    $tb.Tag = $paramname
                    $tb.Cue = $paramtype.FullName
                    $tb.TabIndex = 10 + $parametersetparams.Count - $paramid

                    # remove red color from missing mandatory
                    if ($parammandatory) {
                        $tb.Add_TextChanged({
                            if ($this.BackColor -ne [System.Drawing.Color]::FromName("Window")) {
                                $this.BackColor = [System.Drawing.Color]::FromName("Window")
                            }
                        })
                    }

                    # Add the control to list of stored controls for parameter set and add it to parent control
                    Add-ToControlList -ParameterSetName $parametersetname -Control $tb
                    $p_commandparamcontrols.Controls.Add($tb)
                }
            }
            $paramid++
            # Finished iteration for current parameter
        }

        # New objects
        $p_execute = New-Object System.Windows.Forms.Panel
        $tb_executename = New-Object System.Windows.Forms.CueTextBox
        $s_execute = New-Object System.Windows.Forms.Splitter
        $b_execute = New-Object System.Windows.Forms.Button

        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 100
        $System_Drawing_Size.Height = 20
        $p_execute.Size = $System_Drawing_Size
        $p_execute.Dock = 2
        $workingtabpage.Controls.Add($p_execute)

        $s_execute.Name = "s_execute"
        $p_execute.Controls.Add($s_execute)

        # Generate "execute" button for each parameter set
        $b_execute.TabIndex = 999
        $b_execute.Dock = 5
        $b_execute.Name = "b_execute_$parametersetname"
        $b_execute.UseVisualStyleBackColor = $True
        $b_execute.Text = "Execute>>"
        $b_execute.Tag = $parametersetname
        
        # Add click event to button to execute command with parameterset of the button
        $b_execute.add_Click( {
            Start-Command -ParameterSetName $this.Tag
        } )

        # Add button to parent control
        $p_execute.Controls.Add($b_execute)

        # 
        $tb_executename.Dock = 3
        $tb_executename.TabIndex = 998
        $tb_executename.Cue = "JobName"
        $p_execute.Controls.Add($tb_executename)

        # save the control of the control that holds the optional name for the job
        $script:controllistjobname.Add($parametersetname, $tb_executename)

        $currenttab++
        # Finished iteration for current parameter set
    }

    # Set the focus back to command listbox
    $lb_commands.Focus()

    # Don't refresh the form for every tiny change. Only At the end
    $tc_parametersets.ResumeLayout()

}


####################################
#endregion FORM EVENT FUNCTIONS    #
####################################



####################################
#region    HELP FUNCTIONS          #
####################################


<#
.SYNOPSIS
    - Imports all cmdlets in $assetfolder so they can be used in the gui
    - Gets all commands that should be displayed from $commandsfile
    - Refreshes list of all active jobs
#>
Function Initialize-Data {

    # import cmdlets (old version without links)
    # $(Get-ChildItem -Path $script:assetfolder -Filter "*.psm1").FullName | Import-Module -WarningAction SilentlyContinue

    # imports cmdlets
    $assetitems = $(Get-ChildItem -Path $script:assetfolder).FullName

    $modulefiles = $assetitems | Where-Object { $_ -match ".psm1" }
    $links = $assetitems | Where-Object { $_ -match ".lnk" }

    # get link file destinations
    if ($links) {
        $modulefiles += $(Get-ShortCutData -FileInfo $links).Target
    }

    # import all the modules
    Import-Module $modulefiles -WarningAction SilentlyContinue

    # Get all commands from file and check if they are available. Add those to the script list
    $cmdletnamesfromfile = Get-Content -Path $script:commandsfile
    $script:cmdletnames = $cmdletnamesfromfile | Foreach-Object { 
                                                        if ($(Get-Command -Name $_ -ErrorAction SilentlyContinue)) { 
                                                            $_ 
                                                        } else {
                                                            Write-Warning -Message "Command $_ is not available and thus will not be available for selection." 
                                                        }
                                                    }

    # Show all currently running jobs
    Refresh-JobList
}


<#
.SYNOPSIS
    Returns a shortcut object from a FileInfo
.DESCRIPTION
    The resulting Object contains:
    - Full Path to the shrtcut
    - The name of the shortcut
    - The path of the shortcut
    - The shortcut target

    Used to create shortcuts to psm1/psd1 files, that are by license not allowd to be moved.
    What a terrible, terrible hack. But it works, so ¯\_(ツ)_/¯
.EXAMPLE
    PS> Get-ChildItem -Path . -Filter "*.lnk" | Get-ShortcutData

    Returns shortcut information of all Shortcuts in the current folder
#>
Function Get-ShortcutData {

    [CmdletBinding(DefaultParameterSetName='FromPath',
                   PositionalBinding=$true)]
    [OutputType([String])]

    Param (
        # File Info Objects
        [Parameter(ParameterSetName='FromFileInfo',
                   Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateScript({$_.Name -match ".lnk" })] # is link?
        [System.IO.FileInfo[]] $FileInfo,

        # Filepath
        [Parameter(ParameterSetName='FromPath',
                   Mandatory=$false, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [Alias("PSPath", "FolderPath")]
        [string] $Path = $PWD
    )

    Begin {
        Write-Debug "Starting $($MyInvocation.Mycommand)."
    }

    Process {

        if ($PSCmdlet.ParameterSetName -eq "FromPath") {
            $FileInfo = Get-ChildItem -Path $Path -Filter "*.lnk"
        }

        foreach ($fi in $FileInfo) {
            $Shortcut = Get-ChildItem -Path $fi -Recurse -Include *.lnk
            $Shell = New-Object -ComObject WScript.Shell

            $Properties = @{
                ShortcutName = $Shortcut.Name;
                ShortcutFull = $Shortcut.FullName;
                ShortcutPath = $shortcut.DirectoryName
                Target = $Shell.CreateShortcut($Shortcut).targetpath
            }
            New-Object PSObject -Property $Properties
        }

        [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
    }

    End {
        Write-Debug "Finished $($MyInvocation.Mycommand)."
    }
}


<#
.SYNOPSIS
    Registers a textbox, that is able to show a cue banner (text shadow) using C# code.
#>
Function Register-CueTextBoxObject {

    $Assembly = @("System", "System.ComponentModel","System.Windows.Forms","System.Runtime.InteropServices")

    $Source =  @"
using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace System.Windows.Forms { 
public class CueTextBox : TextBox {
[Localizable(true)]
public string Cue {
    get { return mCue; }
    set { mCue = value; updateCue(); }
}

private void updateCue() {
    if (this.IsHandleCreated && mCue != null) {
        SendMessage(this.Handle, 0x1501, (IntPtr)1, mCue);
    }
}
protected override void OnHandleCreated(EventArgs e) {
    base.OnHandleCreated(e);
    updateCue();
}
        private string mCue;

        // PInvoke
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, string lp);
    }
}
"@

    Add-Type -TypeDefinition $Source -Language CSharp -ReferencedAssemblies $Assembly
}


####################################
#endregion HELP FUNCTIONS          #
####################################



####################################
#region    JOBMANAGER              #
####################################


<#
.SYNOPSIS
    - Retrieves the values from all controls of the parameterset
    - Builds a scriptblock out of it
    - And starts the managed job

    If missing mandatory parameters:
    - mark the control of all the missing parameters red
    - return false
#>
Function Start-Command ([string] $ParameterSetName) {

    # Get all controls for the currently selected parameter set
    $parameterobjects = $script:controllist[$ParameterSetName]
    
    $parameters = @{} # $parametername: $parametervalue
    $secureparameters = @{} # $parametername: $parametervalue
    $switches   = @() # $switchname

    # Get data from all gui objects for the parameter set and store their values in $parameters and $switches, if not empty/set
    foreach ($parameterobject in $parameterobjects) {

        if ($parameterobject -is [System.Windows.Forms.CheckBox]) {
            # is a checkbox -> add to switches
            if ($parameterobject.Checked) {
                $switches += $parameterobject.Tag
            }
        } else {
            # dropdown, textbox, ...
            # Is the element set by the user or left empty?
            if ($parameterobject.Text) {
                # If it is a textbox with passwordchar the value should be converted to a secure string
                if ($parameterobject -is [System.Windows.Forms.CueTextBox] -and $parameterobject.PasswordChar) {
                    # read text and convert to secure string
                    $secureparameters.Add($parameterobject.Tag, $parameterobject.Text)
                } else {
                    # normal dropdown, password or something
                    # add as normal string
                    $parameters.Add($parameterobject.Tag, $parameterobject.Text)
                }
            }
        }   
    }

    # Are all mandatory parameters set?
    $mandatoryparameters = $(($script:serializedcommandinfo.Parametersets | 
                                    Where-Object { $_.Name -eq $ParameterSetName }).Parameters | 
                                    Where-Object {
                                        $_.IsMandatory -eq $true -and $_.ParameterType.FullName -ne "System.Management.Automation.SwitchParameter"
                                    }).Name
    $missingparameters = $mandatoryparameters | Where-Object { $_ -notin $parameters.Keys -and $_ -notin $secureparameters.Keys }

    # if there are missing mandatory parameters set the color of the control to red and return false
    if ($missingparameters) {
        # color missing parameters red
        foreach ($missingparameter in $missingparameters) {
            $script:controllist[$ParameterSetName] | 
            Where-Object { $_.Tag -eq $missingparameter } | 
            Select-Object -First 1 |
            ForEach-Object { $_.BackColor = [System.Drawing.Color]::FromArgb(255,255,179,179) }
        }

        # cancel execution now if there are missing mandatory parameters
        return $false
    }

    # Create the final command line from all filled out parameters and switches
    $commandline = "$selectedcommand"
    $displayedcommandline = "$selectedcommand" # in case of secure strings the password would be shown as "$($parameterobject.Text) | ConvertTo-SecureString -AsPlainText -Force)"!!!
    foreach ($parameter in $parameters.GetEnumerator()) {
        $commandline += " -$($parameter.Name) $($parameter.Value)"
        $displayedcommandline += " -$($parameter.Name) $($parameter.Value)"
    }
    foreach ($secureparameter in $secureparameters.GetEnumerator()) {
        $commandline += " -$($secureparameter.Name) `$(`'$($secureparameter.Value)`' | ConvertTo-SecureString -AsPlainText -Force)"
        $displayedcommandline += " -$($secureparameter.Name) [SecureString]"
    }
    foreach ($switch in $switches) {
        $commandline += " -$switch"
        $displayedcommandline += " -$switch"
    }

    # Cast the command string to a scriptblock
    $scriptblock = [ScriptBlock]::Create($commandline)

    # Get the text of the textbox where the user can enter a job name
    $userdefinedjobname = $($script:controllistjobname[$ParameterSetName]).Text

    if ($userdefinedjobname) {
        # Start an asyncronous job using the scriptblock
        Start-ManagedJob -CommandLine $scriptblock -DisplayedCommandLine $displayedcommandline -JobName $userdefinedjobname
    } else {
        # Start an asyncronous job using the scriptblock
        Start-ManagedJob -CommandLine $scriptblock -DisplayedCommandLine $displayedcommandline
    }
}


<#
.SYNOPSIS
    - starts a job with given command line
    - registers job finishing event
    - updates job name
    - refreshes the datagridview
#>
Function Start-ManagedJob {

    [CmdletBinding(PositionalBinding=$true)]

    Param (
        # Command line, that the job should execute
        [Parameter(Mandatory=$true,  Position=0)]
        [ValidateNotNullOrEmpty()]
        [ScriptBlock] $CommandLine,

        # Command line, that the job should execute
        [Parameter(Mandatory=$false,  Position=1)]
        [ValidateNotNullOrEmpty()]
        [string] $JobName,

        # Command line, that the job should execute
        [Parameter(Mandatory=$false,  Position=2)]
        [ValidateNotNullOrEmpty()]
        [string] $DisplayedCommandLine
    )

    Write-Verbose "Starting job with command: `"$CommandLine`""

    # get the jobs module location
    # TODO: what about importing all modules in the assets folder? If a script has dependencies, for example
    $commandmodulelocation = $(Get-Command -Name $selectedcommand).Module.Path

    # If jobname is not set: use dynamically generated one
    if (-not $JobName) {
        $JobName = "$script:defaultjobname$script:nextjobid"
    }

    # Increase job id
    $script:nextjobid++

    # Start the job and store it as object
    $job = Start-Job -Name $JobName -ScriptBlock {

                                                # import required module if not null ("Get-Help" would be null, for example)
                                                $ConfirmPreference = "none" # nobody would be able to confirm a command inside a job...
                                                if ($args[1]) {
                                                    Import-Module $args[1] -WarningAction SilentlyContinue 3>$null # supress ALL warnings in every way
                                                }

                                                # make the command a scriptblock... AGAIN. TODO: WHY???
                                                $scriptblock = [ScriptBlock]::Create($args[0])
                                                #$scriptblock.Invoke()
                                                Invoke-Command -ScriptBlock $scriptblock -ErrorAction Stop # TODO: Stop sets the job to "completed" instead of "failed" anyways on ErrorAction Stop :(
                                            } -ArgumentList $CommandLine, $commandmodulelocation

    # Store the jobs command line in the hashtable
    if ($DisplayedCommandLine) {
        $jobcommands.Add($job.Id, $DisplayedCommandLine)
    } else {
        $jobcommands.Add($job.Id, $CommandLine)
    }

    # Update visuals
    Refresh-JobList
}


<#
.SYNOPSIS
    - Displays all currently running jobs in the datagridview
    - Updates form footer with job counts
#>
Function Refresh-JobList {

    [CmdletBinding()]
    Param ()

    # Get all currently running local os jobs
    $script:currentjobs = Get-Job | Where-Object { $_.PSJobTypeName -eq "BackgroundJob" } | Sort-Object -Property Id -Descending

    # Format data from jobs and add stored info from variables (will be a pscustomobject)
    $jobinfo = $currentjobs | ForEach-Object { 
            $j = New-Object -TypeName PSObject
            $j | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
            $j | Add-Member -MemberType NoteProperty -Name "State" -Value $_.State
            $j | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $_.PSBeginTime
            $j | Add-Member -MemberType NoteProperty -Name "Command"   -Value $(if ($jobcommands.Keys   -contains $_.Id) { $jobcommands[$_.Id]   } else { "unknown" })
            $j
    }

    # Calculate counts for footer
    $jobstotal     = @($jobinfo).Count
    $jobscompleted = @($jobinfo | Where-Object { $_.State -eq "Completed" }).Count
    $jobsrunning   = @($jobinfo | Where-Object { $_.State -eq "Running" }).Count
    $jobsfailed    = @($jobinfo | Where-Object { $_.State -eq "Failed" }).Count

    # Update form footer
    $stb_bottom.Text = "Total: $jobstotal jobs | Completed: $jobscompleted - Running: $jobsrunning - Failed: $jobsfailed"

    # Update the data in the datagridview
    $jobarray = New-Object System.Collections.ArrayList
    if ($jobinfo -is [System.Management.Automation.PSCustomObject]) {
        # a single job exist
        $jobarray.Add($jobinfo) | Out-Null

        $dgv_results.DataSource = $jobarray
        $f_main.refresh()
    } elseif ($jobinfo.Count -gt 0) {
        # multiple jobs exist
        $jobarray.AddRange($jobinfo) | Out-Null
        $jobstotal = $jobinfo.Count

        $dgv_results.DataSource = $jobarray
        $f_main.refresh()
    } else {
        # no jobs exist at all
        $dgv_results.DataSource = $null
        $f_main.refresh()
    }
}


####################################
#endregion JOBMANAGER              #
####################################



# MAIN

# Register the custom CueTextBox from c# assembly, if not yet loaded
try {
    $test = New-Object System.Windows.Forms.CueTextBox
    Remove-Variable test
} catch {
    Register-CueTextBoxObject
}

# Call the main function: Displaying the main form
# Inside this the form wil be initialized, eventhandlers in the gui registered, modules imported, ...
Generate-MainForm
