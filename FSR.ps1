################################################################################
#                            File System Reporter                              #
#                     Written By: MSgt Anthony Brechtel                        #
#                                                                              #
################################################################################
#####Global Variables###########################################################
################################################################################
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location $dir
Add-Type -AssemblyName 'System.Windows.Forms'
Add-Type -AssemblyName 'System.Drawing'
Add-Type -AssemblyName 'PresentationFramework'
[System.Windows.Forms.Application]::EnableVisualStyles();
################################################################################
clear-host
$version="2.2"
$script:prompt_return = "Null";
$script:excel_report = "Null" 
$loading = New-Object System.Windows.Forms.Label
$loading.Font = New-Object System.Drawing.Font("Copperplate Gothic Bold",10,[System.Drawing.FontStyle]::Regular)
################################################################################
######Main######################################################################
function main
{
    ##################################################################################
    ###########Main Form
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'Fixed3D'
    $form.BackColor = "#434343"
    $form.MaximizeBox = $false
    $form.Icon = $icon
    $Form.SizeGripStyle = "Hide"
    $form.Size='500,230'
    $form.Text = "File System Reporter"
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.TopMost = $false;
    
    ##################################################################################
    ###########Title Main
    $title1 = New-Object System.Windows.Forms.Label   
    $title1.Font = New-Object System.Drawing.Font("Copperplate Gothic Bold",21,[System.Drawing.FontStyle]::Regular)
    $title1.Text="File System Reporter  "
    $title1.TextAlign = 'MiddleCenter'
    $title1.Width=$form.Width
    $title1.height = 35
    $title1.Top = 3
    $title1.ForeColor = "white"
    $title1.Location    = New-Object System.Drawing.Size((($form.width / 2) - ($form.width / 2)),15)
    $form.Controls.Add($title1)

    ##################################################################################
    ###########Title Written By
    $title2 = New-Object System.Windows.Forms.Label
    $title2.Font = New-Object System.Drawing.Font("Copperplate Gothic",7.5,[System.Drawing.FontStyle]::Regular)
    $title2.Text="Written by: Anthony Brechtel`nVer $version"
    $title2.TextAlign = 'MiddleCenter'
    $title2.ForeColor = "darkgray"
    $title2.Width=$form.Width
    $title2.Height=40
    $title2.Location    = New-Object System.Drawing.Size((($form.width / 2) - ($form.width / 2)),($title1.location.y + 30))
    $form.Controls.Add($title2)

    ##################################################################################
    ###########Browse Button
    $target_box = New-Object System.Windows.Forms.TextBox
    $browse_button = New-Object System.Windows.Forms.Button
    $browse_button.Location= New-Object System.Drawing.Size(15,($title2.location.y + 40))
    $browse_button.BackColor = "#606060"
    $browse_button.ForeColor = "White"
    $browse_button.Width=70
    $browse_button.Height=25
    $browse_button.Text='Browse'
    $browse_button.Add_Click(
    {    
		    $script:prompt_return = prompt_for_folder
            if(($prompt_return -ne $Null) -and ($prompt_return -ne "") -and ((Test-Path $prompt_return) -eq $True))
            {
                write-host $prompt_return
                $target_box.Text="$prompt_return"
            }
    })
    $form.Controls.Add($browse_button)
  
    ##################################################################################
    ###########Target Box  
    $target_box.Location = New-Object System.Drawing.Point(($browse_button.location.x + $browse_button.width + 3),($browse_button.location.y + 3))
    $target_box.width = 385
    $target_box.Height = 40
    $target_box.Text = "Browse or Enter a file path"
    $target_box.Add_Click({
        if($target_box.Text -eq "Browse or Enter a file path")
        {
            $target_box.Text = ""
        }
    })
    $target_box.Add_TextChanged({
    
        [string]$script:prompt_return = $target_box.text
        if(($script:prompt_return -ne $null) -and ($script:prompt_return -ne ""))
        {
            if(Test-Path $target_box.text)
            {
                $form.Controls.Add($scan_target_button)
            }
            else
            {
                $form.Controls.Remove($scan_target_button)
            }
        }
        else
        {
            $form.Controls.Remove($scan_target_button)
        }
    })
    $form.Controls.Add($target_box)

    ##################################################################################
    ###########Scan Button
    $scan_target_button = New-Object System.Windows.Forms.Button
    $scan_target_button.Width=150
    $scan_target_button.BackColor = "#606060"
    $scan_target_button.ForeColor = "White"
    $scan_target_button.Location    = New-Object System.Drawing.Size((($form.width / 2) - ($scan_target_button.width / 2)),($target_box.location.y + 30))   
    $scan_target_button.Text='Scan Target'
    $scan_target_button.Add_Click({
        if(!(test-path "$dir\Results"))
        {
            New-Item -ItemType Directory "$dir\Results"
        }
        $form.Controls.Remove($report_button)
        $scan_target_button.Enabled = $false
        scan_target $script:prompt_return
        $scan_target_button.Enabled = $true
        $form.Controls.Add($report_button)
    })
    
    ##################################################################################
    ###########Report Button
    $report_button = New-Object System.Windows.Forms.Button
    $report_button.Width=150
    $report_button.BackColor = "#606060"
    $report_button.ForeColor = "White"
    $report_button.Location    = New-Object System.Drawing.Size((($form.width / 2) - ($report_button.width / 2)),($scan_target_button.location.y + 25))   
    $report_button.Text='View Report'
    $report_button.Add_Click({
        Invoke-Item "$script:excel_report"
    })


    $form.ShowDialog()    
}
################################################################################
######Prompt for Folder#########################################################
function prompt_for_folder()
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
################################################################################
######Scan Target###############################################################
function scan_target ($target)
{
    ##################################################################################
    ###########Build Output File
    $output = build_output_file_name $target

    ##################################################################################
    ###########Scan Job
    $scan_folder = {
        
        ##################################################################################
        ###########Job Variables
        $dir = $using:dir
        $output = $using:output
        $target = $using:target
        Write-Output "Preparing Scan..."
        $display_timer = Get-Date

        ##################################################################################
        ###########Prepare File
        add-content "$dir\Results\$output" "Path,File,Directory,File,Type,Size GB,Size MB,Size KB,Last Modified Date,Last Accessed Date,Creation Date,Full Path"

        ##################################################################################
        ###########Start Scanning
        $files_folders_count = ( Get-ChildItem -literalPath "$target" -Recurse -File -ErrorAction SilentlyContinue | Measure-Object ).Count;
        Write-Output "Scanning... $files_folders_count Files"
        $writer = new-object system.IO.StreamWriter("$dir\Results\$output",$true)
        $file_counter = 0;
        Get-ChildItem -LiteralPath "$target" -Recurse -File -ErrorAction SilentlyContinue | where {! $_.PSIsContainer} | sort Name | ForEach-Object {
            $name      = $_.Name
            $file_counter++;

            $duration = (Get-Date) - $display_timer
            if(($duration.TotalSeconds -gt 2))
            {    
                $display_timer = Get-Date
                [int]$percent = (($file_counter / $files_folders_count) * 100)
                Write-Output "Working on: $file_counter / $files_folders_count Files ($percent%)"
            }

            $status = "Scanning $file_counter of $files_folders_count"
            $directory = $_.Directory       
            $extention = $_.Extension
            $sizeGB      = $_.Length /1Gb
            $sizeGB = [math]::Round($sizeGB,2)
            $sizeMB      = $_.Length /1Mb
            $sizeMB = [math]::Round($sizeMB,2)
            $sizeKB      = $_.Length /1Kb
            $sizeKB = [math]::Round($sizeKB,2)
            $modified  = $_.LastWriteTime
            $accessed  = $_.LastAccessTime
            $created   = $_.CreationTime
            $full_path = $_.FullName


            ###Fix Formula Dashes
        
            $dir_link = "=HYPERLINK(`"`"$directory`"`",`"`"Path`"`")";
            $file_link = "=HYPERLINK(`"`"$directory\$name`"`",`"`"File`"`")";

            if($name -match "^-")
            {
                $write_line = "`"$dir_link`",`"$file_link`",`"$directory`",=`"$name`",$extention,$sizeGB,$sizeMB,$sizeKB,$modified,$accessed,$created,`"$full_path`"";
            }
            else
            {
                $write_line = "`"$dir_link`",`"$file_link`",`"$directory`",`"$name`",$extention,$sizeGB,$sizeMB,$sizeKB,$modified,$accessed,$created,`"$full_path`"";
            }
        
            $writer.write("$write_line`r`n");

        }
        $writer.Close()
    }


    ##################################################################################
    ###########Start Job and Display Updates

    $job = Start-Job -ScriptBlock  $scan_folder
    $loading.TextAlign = 'MiddleCenter'
    $loading.Width=$form.Width
    $loading.top = $form.height - 60
    $loading.Height = 20;
    $loading.BackColor = "Red"
    $loading.Left = (($form.width / 2) - ($form.width / 2))
    $form.Controls.Add($loading)
    $status_count = 0;
    Do {[System.Windows.Forms.Application]::DoEvents() 
        [string]$status = $job.ChildJobs.Output | Select-Object -Last 1    
        if($status -ne $loading.Text)
        {
            $loading.Text = $status
        } 
    } Until ($job.State -eq "Completed")

    ##################################################################################
    ###########Finish Job
    Remove-Job $job
    $loading.Text = "Building Report"
    csv_to_xlsx $output

    $form.Controls.remove($script:loading)  
}
################################################################################
######Loading Start#############################################################
function loading_start ($status)
{
    $loading.Text= $status
    $loading.TextAlign = 'MiddleCenter'
    $loading.Width=$form.Width
    $loading.top = $form.height - 60
    $loading.Height = 20;
    #$title1.ForeColor = "white"
    $loading.BackColor = "Red"
    $loading.Left = (($form.width / 2) - ($form.width / 2))
    $form.Controls.Add($loading)
    
}
################################################################################
######Loading Stop##############################################################
function loading_stop($loading)
{
    #$script:loading.Text=""
    $form.Controls.remove($script:loading)
    $form.refresh();
}
################################################################################
################Build Output File###############################################
function build_output_file_name($target)
{
    
    #$target_name = [System.IO.Path]::GetFileNameWithoutExtension($target)
    $target_name = $target
    $target_name = $target_name.replace(':\',")");
    $target_name = $target_name.replace('/',")");
    $target_name = $target_name.replace('\',")");
    $target_name = $target_name.replace('--',"");
    
    $date = Get-Date -Format G
    [regex]$pattern = " "
    $date = $pattern.replace($date, " @ ", 1);
    $date = $date.replace('/',"-");
    $date = $date.replace(':',".");

    $output = "$target_name      ($date)" + ".csv";
    return $output
}
################################################################################
######CSV to XLSX###############################################################
function csv_to_xlsx($output)
{
    ### Set input and output path
    $inputCSV = "$dir\Results\$output"
    $output2 = [io.path]::GetFileNameWithoutExtension($inputCSV)
    $outputXLSX = "$dir\Results\$output2.xlsx"

    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open("$inputCSV")
    $worksheet = $workbook.worksheets.item(1) 
    $objExcel.Visible=$false
    $objExcel.DisplayAlerts = $False


    ### Make it pretty
    $worksheet.UsedRange.Columns.Autofit();
    
    $worksheet.Columns.item("A").NumberFormat = "@"
    $worksheet.Columns.item("B").NumberFormat = "@"
    $worksheet.Columns.item("F").NumberFormat = "0"
    $worksheet.Columns.item("G").NumberFormat = "0"
    $worksheet.Columns.item("H").NumberFormat = "0"
    $headerRange = $worksheet.Range("a1","L1")
    $headerRange.AutoFilter() | Out-Null
    $headerRange.Interior.ColorIndex =48
    $headerRange.Font.Bold=$True
    $row_count = $worksheet.UsedRange.Rows.Count
    #$objRange = $worksheet.Range("C2:C$row_count")  
    #[void] $objRange.Sort($objRange) 

    $empty_Var = [System.Type]::Missing
    $sort_col = $worksheet.Range("C1:C$row_count")
    $worksheet.UsedRange.Sort($sort_col,1,$empty_Var,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1)

    $borderrange = $worksheet.Range(“A1","L$row_count")
    $borderrange.Borders.Color = 0
    $borderrange.Borders.Weight = 2

    $workbook.SaveAs($outputXLSX,51)
    $objExcel.Quit()

    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()


    $script:excel_report = $outputXLSX;
    if(Test-Path -literalpath $outputXLSX)
    {
        Remove-Item "$dir\Results\$output"
    }
}
################################################################################
######Show Console##############################################################
function Show-Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }
}
################################################################################
Show-Console -Hide
main