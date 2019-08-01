# Citrix Profile Analysis Script
# Written by Joshua Woleben
# 2/25/2019

# Purpose: Read through directory structure of profiles and determine user's full name and supervisor

Param([string]$hostlist)

# Import required modules
Import-Module ActiveDirectory -Force -ErrorAction SilentlyContinue



# Set up Excel Object
$excel_output = New-Object -ComObject Excel.Application
$excel_output.Visible = $true

# Create workbook
$workbook = $excel_output.WorkBooks.Add()

# Add and name worksheets
$workbook.Worksheets.Add()
$workbook.Worksheets.Add()
$profile_list_worksheet = $workbook.Worksheets.Item(1)
$host_profile_totals_worksheet = $workbook.Worksheets.Item(2)
$profile_list_worksheet.Name = "Total Profile List"
$host_profile_totals_worksheet.Name = "Host Profile Sizes"

# Set Header items
$profile_list_worksheet.Cells.Item(1,1) = "User ID"
$profile_list_worksheet.Cells.Item(1,2) = "User Full Name"
$profile_list_worksheet.Cells.Item(1,3) = "Manager Full Name"
$profile_list_worksheet.Cells.Item(1,4) = "Department"
$profile_list_worksheet.Cells.Item(1,5) = "Profile size (MB)"
$profile_list_worksheet.Cells.Item(1,6) = "Hostname"

$host_profile_totals_worksheet.Cells.Item(1,1) = "Hostname"
$host_profile_totals_worksheet.Cells.Item(1,2) = "Total Profile Size (MB)"

# Format Header font
$cell_format_range = $profile_list_worksheet.UsedRange

$cell_format_range.Interior.ColorIndex = 19
$cell_format_range.Font.ColorIndex = 11
$cell_format_range.Font.Bold = $true
$cell_format_range.EntireColumn.AutoFit()

$cell_format_range2 = $host_profile_totals_worksheet.UsedRange

$cell_format_range2.Interior.ColorIndex = 19
$cell_format_range2.Font.ColorIndex = 11
$cell_format_range2.Font.Bold = $true
$cell_format_range2.EntireColumn.AutoFit()

$row_counter = 2

# Construct file path
$base_path = "C:\Temp\"
$log_time = [datetime]::Now
$log_stamp = $log_time.ToString('yyyyMMdd-hhmmss')
$filename = ($base_path + "-ProfileSizing-" + $log_stamp + ".xlsx")

# Read file
$host_list = Get-Content -Path $hostlist

$second_counter = 2
ForEach ($hostname in $host_list) {
# Set variables
$profile_path="\\$hostname\C`$\Users\"

# Get user list from directory
$user_list = Get-ChildItem $profile_path | Select-Object -ExpandProperty Name

$profile_total = 0
# Query AD for each user's full name and manager's full name
ForEach ($user in $user_list) {
    $profile_list_worksheet.Cells.Item($row_counter,1) = $user
    if ($user -ne $null) {
        $full_name = (Get-ADUser $user -ErrorAction SilentlyContinue| Select -ExpandProperty givenName -ErrorAction SilentlyContinue) + " " + (Get-ADUser $user -ErrorAction SilentlyContinue | Select -ExpandProperty surname -ErrorAction SilentlyContinue)
        $profile_list_worksheet.Cells.Item($row_counter,2) = $full_name
        $manager = Get-ADUser $user -Properties Manager -ErrorAction SilentlyContinue | Select -ExpandProperty Manager -ErrorAction SilentlyContinue
        if ($manager -ne $null) {
            $manager_full_name = (Get-ADUser $manager -ErrorAction SilentlyContinue | Select -ExpandProperty givenName -ErrorAction SilentlyContinue) + " " + (Get-ADUser $manager -ErrorAction SilentlyContinue| Select -ExpandProperty surname -ErrorAction SilentlyContinue)
            $profile_list_worksheet.Cells.Item($row_counter,3) = $manager_full_name
        }
        $department = Get-ADUser $user -Properties Department -ErrorAction SilentlyContinue | Select -ExpandProperty Department -ErrorAction SilentlyContinue
        $profile_list_worksheet.Cells.Item($row_counter,4) = $department
        $profile_size = "{0:N2}" -f ((Get-ChildItem ($profile_path + "\" + $user) -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB)
        $profile_list_worksheet.Cells.Item($row_counter,5) = $profile_size
        $profile_total = $profile_total + $profile_size
        $profile_list_worksheet.Cells.Item($row_counter,6) = $hostname
       Write-Output ($user + ", " + $full_name + ", " + $manager_full_name + ", " + $department + ", " + $profile_size)
    }
    $row_counter = $row_counter + 1
}

$host_profile_totals_worksheet.Cells.Item($second_counter,2) = $profile_total
$host_profile_totals_worksheet.Cells.Item($second_counter,1) = $hostname
$second_counter = $second_counter + 1

}
$cell_format_range.EntireColumn.AutoFit()

$sort_range = $profile_list_worksheet.UsedRange

$column_range = $sort_range.Range("E1")

[void] $sort_range.Sort($column_range,2,$null,$null,1,$null,1,1)

$profile_list_worksheet.SaveAs($filename)