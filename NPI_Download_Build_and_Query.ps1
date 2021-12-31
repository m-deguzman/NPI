
Clear-Host
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0
Get-Date
## ---------- Prompt to Query Loop ------------------------------------------
$About = @"

This script can be used to get a list of MD's and NPI Codes
using data from CMS (Centers for Medicare and Medicaid).

It will build a local SQLite database using open source tools.
Once the NPI SQLite DB is built, you can query for code and specialty.
Query result is saved in csv format and imported to Excel.

"@
Write-Host $About -ForegroundColor White -BackgroundColor Blue
PAUSE

if(-not $PSScriptRoot) {Write-Host "Script Root Dir: $PSScriptRoot is not defined" -ForegroundColor White -BackgroundColor Red; break} 
Set-Location $PSScriptRoot

## ---- sqlite utility ------------------
$sqlite_uri = "https://sqlite.org/2021/sqlite-tools-win32-x86-3360000.zip"
$sqlite3 = Join-Path -Path $PSScriptRoot -ChildPath "sqlite3.exe"

## ---- sqlite core dll ------------------
$version = '1.0.112'
$sqlite_core_uri = "https://www.nuget.org/api/v2/package/System.Data.SQLite.Core/$version"
$sqlite_core  =  Join-Path -Path $PSScriptRoot -ChildPath "System.Data.SQLite.dll"
$sqlite_interop  =  Join-Path -Path $PSScriptRoot -ChildPath "SQLite.Interop.dll"


## ------ cms npi code file ---------
$npi_uri = "https://data.cms.gov/provider-data/sites/default/files/resources/69a75aa9d3dc1aed6b881725cf0ddc12_1639689642/DAC_NationalDownloadableFile.csv"
$npi_db_name = "NPI_dac." + (Get-Date -Format "yyyyMMdd") + ".SQLite.db"
$npi_db_full_path = Join-Path -Path $PSScriptRoot -ChildPath $npi_db_name
$npi_file_csv = "DAC_NationalDownloadableFile.csv"
$npi_file_csv_path = Join-Path -Path $PSScriptRoot -ChildPath $npi_file_csv


## --------- Intro -------------------
Get-Date
Set-Location $PSScriptRoot
Write-Host ""

If ( (-not(Test-Path $sqlite3)) -OR (-not(Test-Path $npi_db_full_path)) -OR (-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {
Write-Host "This tool will perform the following:"
Write-Host "- Download, unzip and extract required files from the following:"
If (-not(Test-Path $sqlite3)) {Write-Host $sqlite_uri -ForegroundColor White -BackgroundColor Black}
If ((-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {Write-Host $sqlite_core_uri -ForegroundColor White -BackgroundColor Black}
If (-not(Test-Path $npi_db_full_path)) {
Write-Host $npi_uri -ForegroundColor White -BackgroundColor Black
Write-Host "-Build CMS NPI Code Database File."
}
Write-Host "-Verify table count from CMS NPI Code DB."
Write-Host "Press Ctrl+C to cancel." -ForegroundColor White -BackgroundColor Red
Write-Host "Press Enter to proceed." -ForegroundColor White -BackgroundColor Green
PAUSE
}
else {
Write-Host "Required $sqlite3 found."
Write-Host "Required $npi_db_full_path found."
Write-Host "Required $sqlite_core found."
Write-Host "Required $sqlite_interop."
}

## --------- Download sqlite -------------------
If (-not (Test-Path $sqlite3) ) {
Get-Date
$sqlite_uri = $sqlite_uri.Trim()
$sqlite_zip = Join-Path -Path $PSScriptRoot -ChildPath "sqlite.zip"
$sqlite_dir = Join-Path -Path $PSScriptRoot -ChildPath "sqlite"
$sqlite_exe = Join-Path -Path $sqlite_dir -ChildPath "sqlite-tools-win32-x86-3360000\sqlite3.exe"
Write-Host "Downloading file: $sqlite_uri." -ForegroundColor White -BackgroundColor Black
Write-Host "Destination: " $sqlite_zip -ForegroundColor Black -BackgroundColor White
Write-Host "Please wait..."
    try
    {   
    $Response = Invoke-WebRequest -Uri $sqlite_uri -OutFile $sqlite_zip 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Succesful:   $sqlite_zip"  -ForegroundColor White -BackgroundColor Green 

        If (-not (Test-Path $sqlite_zip) ) {
        Write-Host "$sqlite_zip <---- Required File not found." -ForegroundColor White -BackgroundColor Red 
        PAUSE
        BREAK
        }
        else {  
        If (-not (Test-Path $sqlite_dir )) {New-Item -ItemType directory $sqlite_dir | Out-Null}
        Expand-Archive -Path $sqlite_zip -DestinationPath $sqlite_dir -Force 
        copy-item $sqlite_exe $PSScriptRoot -Force 
        remove-item $sqlite_zip -recurse 
        remove-item $sqlite_dir -recurse 
        }      
    }
    catch
    {
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download:  $sqlite_uri" -ForegroundColor White -BackgroundColor Red
    break
    }
    #$StatusCode

}

## --------- Download DLL -------------------
If ((-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {
$file = "system.data.sqlite.core.$version"
$sqlite_core_dll = $file + "/lib/netstandard2.0/System.Data.SQLite.dll"
$sqlite_interop_dll = $file + "/runtimes/win-x64/native/netstandard2.0/SQLite.Interop.dll"
$temp_download_dir =  Join-Path -Path $PSScriptRoot -ChildPath "temp_download"

If (-not (Test-Path $temp_download_dir) ) {New-Item -ItemType directory $temp_download_dir | Out-Null}
Set-Location $temp_download_dir

$dl = @{
	uri = $sqlite_core_uri
	outfile = "$file.zip"
}

Write-Host "Downloading file: $sqlite_core_uri." -ForegroundColor White -BackgroundColor Black
Write-Host "Destination: $temp_download_dir" -ForegroundColor Black -BackgroundColor White
try
{
    $Response = Invoke-WebRequest @dl 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Succesful $sqlite_core_uri " -ForegroundColor White -BackgroundColor Green  
}
catch
{
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download:  $sqlite_core_uri " -ForegroundColor White -BackgroundColor Red
    break
}
#$StatusCode

If (-not (Test-Path "$file.zip") ) {
Write-Host "$file zip <---- Required File not found." -ForegroundColor White -BackgroundColor Red 
PAUSE
BREAK
} else {
If (Test-Path "$file.zip") {
Expand-Archive $dl.outfile -Force
copy-item $sqlite_core_dll $PSScriptRoot -Force
copy-item $sqlite_interop_dll $PSScriptRoot -Force
               }
    }
Set-Location $PSScriptRoot
If (Test-Path $file) {remove-item $file -recurse}
If (Test-Path $temp_download_dir) {remove-item $temp_download_dir -recurse}

}


## --------- Download npi -------------------
If (-not (Test-Path $npi_db_full_path)) {

    If (-not (Test-Path $npi_file_csv_path)) {
    $activity_timer = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "Downloading file: $npi_uri." -ForegroundColor White -BackgroundColor Black
    Write-Host "Destination: " $npi_file_csv_path -ForegroundColor Black -BackgroundColor White
    Write-Host "Download may take 3-5 minutes. Please wait..."
    $csv_size = (Invoke-WebRequest $npi_uri -Method Head).Headers.'Content-Length'
    Write-Host "Size of File to download:" $csv_size
    try
    {
    $Response = Invoke-WebRequest -Uri $npi_uri -OutFile $npi_file_csv_path 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Succesful $npi_file_zip_path " -ForegroundColor White -BackgroundColor Green }
    catch
    {
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download: " $npi_uri
    PAUSE
    break
    }
    Write-Host $StatusCode
    $activity_timer.Stop()
    Write-Host "Download Time (HH:MM:SS.ms) - " $activity_timer.Elapsed -ForegroundColor White -BackgroundColor Red
}

    
    If (-not (Test-Path $npi_file_csv_path) ) {
       Write-Host "$npi_file_csv_path <---- Required Date File not found." -ForegroundColor White -BackgroundColor Red 
       PAUSE
       BREAK
       } 
   else {
## --------- Build npi.db -------------------
      Write-Host "$npi_file_csv_path <---- downloaded file found." -ForegroundColor White -BackgroundColor Green
  
      Set-Location $PSScriptRoot
      $sqlite_exe = Join-Path -Path $PSScriptRoot -ChildPath "sqlite3.exe"

      If (-not (Test-Path $sqlite_exe) ){Write-Host "$sqlite_exe <---- Exe File not found." -ForegroundColor White -BackgroundColor Red;PAUSE;BREAK}

## --------- npi DB Build Parameters -------------------
$activity_timer.Start()
$build_parameters = @"
.mode csv
.import DAC_NationalDownloadableFile.csv npi_code
create index idx on npi_code(npi);
create index ldx on npi_code(" lst_nm");
"@

Get-Date
Write-host $build_parameters -ForegroundColor White -BackgroundColor Black
Write-host "NPI Database Build will take 3-5 minutes."
Write-Host "$npi_db_full_path DB Build in progress. Please wait....."
$npi_db_full_path = $npi_db_full_path -replace "\\", "/"
$build_parameters | .\sqlite3 $npi_db_full_path

$activity_timer.Stop()
Write-Host "NPI DB Build Time (HH:MM:SS.ms) - " $activity_timer.Elapsed -ForegroundColor White -BackgroundColor Red

If (-not (Test-Path $npi_db_full_path) ){Write-Host "$npi_db_full_path <---- Exe File not found." -ForegroundColor White -BackgroundColor Red;BREAK}   
Write-Host "$npi_db_full_path build completed." -ForegroundColor White -BackgroundColor Blue
Remove-item $npi_file_csv_path -Force
      }
}
##--------------------End of Build ----------------------------------------



##-------------------- Row Count Verification ----------------------------------------
Write-Host ""
Write-Host "Verifying row count."  -ForegroundColor White -BackgroundColor Blue
Set-Location $PSScriptRoot

$timestamp = Get-Date -Format "dddd_yyyy_MM_dd_HH_mm_ss"
$output = "row_count_" + $timestamp + ".sql"
$output = Join-Path -Path $PSScriptRoot -ChildPath $output
$output = $output -replace "\\", "/"

$query_parameters = @"
.output $output
WITH RECURSIVE
tbl(name) AS (Select name FROM sqlite_master WHERE type IN ("table"))
SELECT  ' select " No. of rows from ' || name || ' table -  ", printf ("%,d",count (*))  from ' || name || ';' FROM tbl;
.output
.mode tab
.header off
.echo off
.read $output
"@

$query_parameters | .\sqlite3 $npi_db_full_path
Remove-Item $output

Add-Type -Path $sqlite_core
$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$npi_db_full_path"
$con.Open()
$sql = $con.CreateCommand()

$timestamp = Get-Date -Format "dddd_yyyy_MM_dd_HH_mm_ss"
$query_result_csv = "NPI_Query_Result_" + $timestamp + ".csv"
$query_result_xl = "NPI_Query_Result_" + $timestamp + ".xlsx"
$query_result_csv = Join-Path -Path $PSScriptRoot -ChildPath $query_result_csv
$query_result_xl = Join-Path -Path $PSScriptRoot -ChildPath $query_result_xl

$timer = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "$npi_db_full_path is ready to receive your query."  -ForegroundColor White -BackgroundColor Magenta


## ---------- Prompt to Query Loop ------------------------------------------
$qry_timer = [System.Diagnostics.Stopwatch]::StartNew()
$searchString = ""
Do
{ 
$qry_timer.Stop()
Write-host ""
Write-host "Enter provider info to search (e.g. cardiology)." -ForegroundColor White -BackgroundColor Magenta
	
$searchString = Read-Host 
    
if ((-not ($searchString)) -OR ($searchString -eq "exit")  -OR ($searchString.Length -lt 3)) {Write-Host "..." -BackgroundColor Magenta; [console]::beep(1000,100); BREAK}
$searchString = $searchString -replace " ", "%"

Write-host "Please wait while we search for: " $searchString

$sql.CommandText =  @"
select  "NPI", " lst_nm", " frst_nm", " mid_nm", " Cred" , " st", " pri_spec"
from  npi_code
where " st" like '%$searchString%' or " pri_spec" like '%$searchString%' or " lst_nm" like '%$searchString%' 
order by " lst_nm",  " frst_nm";	
"@

    $qry_timer.Start()
    Write-Host "SQLite Query: " $sql.CommandText
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
	$data = New-Object System.Data.DataSet
	[void]$adapter.Fill($data)
	

	If ($data.Tables.Rows) {
        [console]::beep(400,500)
        Write-Host "Retrieving result. Please wait...."
        #$data.tables.Rows | Out-Gridview -Title "$searchString"
        $Records  = $data.tables.Rows.Count 
        $Records  = '{0:N0}' -f $Records
        Write-Host $Records " - records retrieved." -ForegroundColor White -BackgroundColor Green
        $qry_timer.Stop()
        Write-Host "Retrive Time (HH:MM:SS.ms) - " $qry_timer.Elapsed -ForegroundColor White -BackgroundColor Red
        $data.tables.Rows | Export-Csv -Path $query_result_csv -Append
		Write-Host "Result saved to $query_result_csv"
		Write-host "Press ENTER to end search." -ForegroundColor Black -BackgroundColor Yellow		
  
	}
	Else {
        [console]::beep(200,1000)
		Write-Host "NO entry found for:  $searchString" -ForegroundColor Red -BackgroundColor Black
	}


} 
While ($searchString)
## ---------- Prompt to Query Loop ------------------------------------------

$sql.Dispose()
$con.Close()



If (Test-Path $query_result_csv){
    
Write-Host "Result file: $query_result_csv found." -ForegroundColor White -BackgroundColor Magenta

$inputCSV = $query_result_csv
$outputXLSX = $query_result_xl
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)
$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $Excel.Application.International(5)
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.Refresh() | Out-Null
$query.Delete() 
$workbook.SaveAs($outputXLSX,51) 

    If (-not(Test-Path $query_result_xl)) {
          $excel.Quit()     
          Write-Host "Result file: $query_result_csv can be imported in Excel." -ForegroundColor White -BackgroundColor Magenta      
          }  
    else  {
            Remove-Item $query_result_csv
            Write-Host "Prepping $query_result_xl. Please wait..." -ForegroundColor White -BackgroundColor Magenta	
            $worksheet.Cells.Item(1,1).EntireRow.Delete() | Out-Null
            $UsedRange = $worksheet.UsedRange
            $RowCount = $UsedRange.Rows.count
            $RowCount = '{0:N0}' -f $RowCount
            $ColCount = $UsedRange.Columns.count
            Write-Host "Total Rows Imported - " $RowCount -ForegroundColor White -BackgroundColor Green
            #Write-Host "Total Columns - "  $ColCount -ForegroundColor White -BackgroundColor Green

            Write-Host "Setting split pane."
            $workbook.Application.ActiveWindow.SplitColumn = 1
            $workbook.Application.ActiveWindow.SplitRow = 1
            $workbook.Application.ActiveWindow.FreezePanes = $true

            Write-Host "Setting font and highlight."
            for($i = 1; $i -lt $ColCount+1; $i++){ 
            $worksheet.Cells.Item(1,$i).Font.Bold = $True
            $worksheet.Cells.Item(1,$i).Interior.ColorIndex = 6
            #$worksheet.Cells.Item(1,$i).columnwidth = 30  
            }	
 
            Write-Host "Displaying spreadsheet."
            $workbook.Save()
            $excel.visible = $true
 
          }  
  
}

$timer.Stop()
Write-Host "Session Time (HH:MM:SS.ms) - " $timer.Elapsed -ForegroundColor White -BackgroundColor Red
Get-Date
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

