using namespace Microsoft.Office.Interop.Excel
using namespace	System.Reflection


# Function: merge_run_csv
#
function merge_run_csv
{
  $xl = new-object -comobject Excel.Application
  $xl.visible = $false
  

  $merge_keys	= "capacitor_code", "ra_offset_gain", "comp_offset", "latch_offset",
		  "comp_offset_injected_errors", "latch_offset_injected_errors",
		  "ra_offset_injected_errors", "ra_gain_injected_errors",
		  "capacitor_mismatch_injected_errors", "im3_dither_capacitor_mismatch_injected_errors",
		  "regaccess"


  $csv_count = @{}

  ls -filter *csv | foreach {
    $split = $_ -split '__'


    $crfrundir = split-path $split[0] -leaf
    $testrun   = $split[1]
    $csvtype   = ($split[2] -split '\.')[0]
    $savename  = join-path $PWD $($crfrundir, $testrun -join '__')
  
    if (!$csv_count.$crfrundir) {$csv_count.$crfrundir = @{}}  
    if (!$csv_count.$crfrundir.$testrun) {$csv_count.$crfrundir.$testrun = @{}}  

    $wb = $xl.workbooks.open($_.fullname)
    $ws = $wb.worksheets[1]

    $csv_count.$crfrundir.$testrun.$csvtype = @{ws=$ws; wb=$wb}

    write-output  "Count $($csv_count.$crfrundir.$testrun.count)  crfrundir<$crfrundir>  testrun<$testrun>  csvtype<$csvtype>"

    $table=$ws.listobjects.add([XlListObjectSourceType]::xlSrcRange, $ws.usedrange, $null, [XlYesNoGuess]::xlYes, $null, "TableStyleMedium15")
    $table.ShowTableStyleRowStripes = $false
    $table.range.rowheight = 15

  
    switch ($csvtype)
    {
      "fft"  {
        $sheet = "FFT"

	$table.HeaderRowRange | where value2 -notmatch 'ENOB|SNR|SFDR|FINR|Hz' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -match 'DESIGN FREQS').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['DESIGN FREQS (#:+/- Hz/Amp)'].range.WrapText = $true
	$table.ListColumns['DESIGN FREQS (#:+/- Hz/Amp)'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null

	# Conditional Formatting
	$table.ListColumns("STATUS").range  | 
	foreach {$_.FormatConditions.add([XlFormatConditionType]::xlCellValue, [XlFormatConditionOperator]::xlEqual, "FAIL")} | 
	foreach {$_.font.color = -16383844; $_.interior.color = 13551615}

	$table.ListColumns("ENOB").range  | 
	foreach {if ($_.row -ne 1) {$_.FormatConditions.add([XlFormatConditionType]::xlCellValue, [XlFormatConditionOperator]::xlLess, "11.7")}} | 
	foreach {$_.font.color = -16383843; $_.interior.color = 13551615}

	$table.ListColumns("ENOB").range  | 
	foreach {if ($_.row -ne 1) {$_.FormatConditions.add([XlFormatConditionType]::xlCellValue, [XlFormatConditionOperator]::xlGreater, "11.7")}} | 
	foreach {$_.font.color = -16752384; $_.interior.color = 13561798}

	$table.ListColumns("DESIGN TONE COUNT (+/-)").range |foreach {if ($_.row -ne 1 -and $_.value2 -ne $_.cells.offset(0, -1).value2) {$_.font.color = -16383844; $_.interior.color = 13551615}}

	}
	
      "capacitor_code" {
        $sheet = "CAPACITOR CODE"

	$table.HeaderRowRange | where value2 -ne 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "ra_offset_gain" {
        $sheet = "RESIDUE RA OFFSET GAIN"

	$table.HeaderRowRange | where value2 -notin 'STEP SIZE', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "comp_offset" {
        $sheet = "RESIDUE COMP OFFSET"

	$table.HeaderRowRange | where value2 -notin 'IDEAL', 'ACTUAL', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	# $table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
       }
  
      "latch_offset"  {
        $sheet = "RESIDUE LATCH OFFSET"

	$table.HeaderRowRange | where value2 -notin 'IDEAL', 'ACTUAL', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	# $table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "comp_offset_injected_errors" {
        $sheet = "INJ ERR COMP OFFSET"

	$table.HeaderRowRange | where value2 -notin 'INT ERROR NAME', 'FLOAT ERROR NAME', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "latch_offset_injected_errors"  {
        $sheet = "INJ ERR LATCH OFFSET"

	$table.HeaderRowRange | where value2 -notin 'INT ERROR NAME', 'FLOAT ERROR NAME', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "ra_offset_injected_errors" {
        $sheet = "INJ ERR RA OFFSET"

	$table.HeaderRowRange | where value2 -ne 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "ra_gain_injected_errors"	{
        $sheet = "INJ ERR RA GAIN"

	$table.HeaderRowRange | where value2 -ne 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "capacitor_mismatch_injected_errors"  {
        $sheet = "INJ ERR CAPACITOR MISMATCH"

	$table.HeaderRowRange | where value2 -notin 'INT ERROR NAME', 'FLOAT ERROR NAME', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null
      }
  
      "im3_dither_capacitor_mismatch_injected_errors" {
        $sheet = "INJ ERR IM3 DITHER CAP MISMATCH"

	$table.HeaderRowRange | where value2 -notin 'INT ERROR NAME', 'FLOAT ERROR NAME', 'PATH' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	$table.ListColumns['PATH'].range.WrapText = $true
	$table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | where name -ne 'PATH' | foreach {$_.range.columns.autofit()} | out-null

      }

      "regaccess" {
        $sheet = "REG ACCESS"

	$table.HeaderRowRange | where value2 -in 'TIMESTAMP', 'SCOPE', 'SLICE' | foreach {$_.entirecolumn.Horizontalalignment = [Constants]::xlcenter}
	$table.HeaderRowRange.entirerow.Horizontalalignment = [Constants]::xlleft
	# ($table.HeaderRowRange | where value2 -eq 'PATH').horizontalalignment = [Constants]::xlcenter
	# $table.ListColumns['PATH'].range.WrapText = $true
	# $table.ListColumns['PATH'].range.ColumnWidth = 75
	$table.ListColumns | foreach {$_.range.columns.autofit()} | out-null
      }
    }
  
    
    $xl.activewindow.Splitrow = 1
    $xl.activewindow.FreezePanes = $true

    $ws.name = $sheet

    if ($csv_count.$crfrundir.$testrun.count -eq 12) {
      "Got 12 CSV files for CRFRUNDIR:$crfrundir TESTRUN:$testrun, now merging them.."

      $merge_keys | foreach {$csv_count.$crfrundir.$testrun.$_.ws.move([Missing]::Value, $csv_count.$crfrundir.$testrun.fft.wb.activesheet)}

      # Activating the FFT tab
      # 
      # This method is equivalent to clicking the tab at the bottom of the sheet
      #
      $csv_count.$crfrundir.$testrun.fft.ws.activate()

      "Saving the resulting Workbook as $savename.xlsx .."
      $csv_count.$crfrundir.$testrun.fft.wb.saveas($savename,			      # Filename
                                                   [XlFileFormat]::xlWorkbookDefault, # FileFormat
                                                   [Missing]::Value,		      # Password
                                                   [Missing]::Value,    	      # WriteResPassword
                                                   $false,			      # ReadOnlyRecommended
                                                   $false,            		      # CreateBackup
                                                   [Missing]::Value,    	      # AccessMode
                                                   [Missing]::Value,    	      # ConflictResolution
                                                   [Missing]::Value,    	      # AddToMru
                                                   [Missing]::Value,    	      # TextCodepage
                                                   [Missing]::Value,    	      # TextVisualLayout
                                                   [Missing]::Value     	      # Local
						  )
      
      "Done."
      $csv_count.$crfrundir.$testrun.fft.wb.close($false, [Missing]::Value, $false)

      ""
    }
  } 
  

  $xl.quit()

} # merge_run_csv



# Function: tc_status_info
#
function tc_status_info
{
  $status    = @{0="FAILED"; 1="PASS"; 2="PASS"; 3="FAILED"}
  $changelist= 0
  $seed      = 0
  $rundir    = ""

  $xl = new-object -comobject Excel.Application
  $xl.SheetsInNewWorkbook = 1
  $xl.visible = $true

  $wb = $xl.workbooks.add()
  $ws = $wb.activesheet

  $ws.range("A1") = "STATUS"
  $ws.range("A5") = "CHANGELIST"
  $ws.range("A6") = "SEED"
  $ws.range("A7") = "RUNDIR"
                
  $ws.range("B1") = 0
  $ws.range("B2") = 1
  $ws.range("B3") = 2
  $ws.range("B4") = 3

  $ws.columns("B").autofit()
  $ws.columns("B").horizontalalignment = [Constants]::xlcenter
  $ws.columns("B").columnwidth = 3

  $ws.range("A1:A4").merge()
  $ws.range("A1").verticalalignment = [Constants]::xlcenter
  $ws.range("B5:C5").merge()
  $ws.range("B6:C6").merge()
  $ws.range("B7:C7").merge()

  $ws.range("A1:A7").indentlevel = 1
  $ws.columns("A").autofit()

  $ws.range("A1:C7").borders.color = 1 
  $ws.range("A1:C7").borders([XlBordersIndex]::xlEdgeBottom).Weight = [XlBorderWeight]::xlMedium
  $ws.range("A4:C4").borders([XlBordersIndex]::xlEdgeBottom).Weight = [XlBorderWeight]::xlMedium

}


# Function: fg_cal_passfail
#
function fg_cal_passfail ($csvfile)
{
  $xl = new-object -comobject Excel.Application
  $xl.SheetsInNewWorkbook = 1
  $xl.visible = $true
  
  $wb = $xl.workbooks.open($(resolve-path $csvfile))
  $ws = $wb.worksheets[1]

  $ws.range($ws.range("G1"), $ws.range("G1").offset(1).end([XlDirection]::xltoright).offset(-1)).select()
  $xl.selection.merge()
  $xl.selection.value2 = "STEPS"
  $xl.selection.horizontalalignment = [Constants]::xlcenter
  


  $ws.range($ws.range("E2"), $ws.range("E2").offset(0, 1).end([XlDirection]::xldown)).select()
  $xl.Selection.FormatConditions.add([XlFormatConditionType]::xlCellValue, [XlFormatConditionOperator]::xlEqual, "FAIL") | foreach {$_.font.color = -16383844; $_.interior.color = 13551615}
  $xl.Selection.FormatConditions.add([XlFormatConditionType]::xlCellValue, [XlFormatConditionOperator]::xlEqual, "PASS") | foreach {$_.font.color = -16752384; $_.interior.color = 13561798}

  $ws.usedrange.borders.color = 1 
  $ws.usedrange.borders([XlBordersIndex]::xlEdgeBottom).Weight = [XlBorderWeight]::xlMedium
  $ws.usedrange.columns.autofit()

  $ws.range($ws.range("A1"), $ws.range("A1").offset(1, 1).end([XlDirection]::xltoright).offset(-1)).select()
  $xl.selection.font.bold = $true
  $xl.selection.borders([XlBordersIndex]::xlEdgeBottom).Weight = [XlBorderWeight]::xlMedium


  vertical_cell_merge $ws.range("B2") -left_index
  vertical_cell_merge $ws.range("C2")
  #$LastRow = $ws.range("B2").end([XlDirection]::xlDown).row;
  #$s  = $ws.range("B2")
  #$e  = $s
  #$id =	1
  #$s.offset(0, -1).value2 = $id;

  #while ($e.row -lt $LastRow) {
  #  $e = $e.offset(1)

  #  if ($e.value2 -eq $s.value2) {
  #    $e.value2 = ""
  #    $ws.range($s, $e).merge()
  #    $ws.range($s.offset(0, -1), $e.offset(0, -1)).merge()
  #  }
  #  else {
  #    $s = $e
  #    $e.offset(0, -1).value2 = ++$id;
  #  }
  #}


  $ws.range($ws.range("D2"), $ws.range("D2").end([XlDirection]::xlDown)).offset(0, -1).indentlevel = 1

  $ws.columns("A:C").verticalalignment = [Constants]::xlcenter
  $xl.union($ws.columns("A"), $ws.range("B1"), $ws.columns("C:F")).horizontalalignment = [Constants]::xlcenter

  $ws.usedrange.columns.autofit()

  $ws.columns('E:F').columnwidth = 8
  $ws.range("A1").select()

}

# Function: vertical_cell_merge
#
function vertical_cell_merge ($rng, [switch]$left_index)
{
  $w = $rng.worksheet;
  
  $LastRow = $rng.end([XlDirection]::xlDown).row;
  $s  = $rng
  $e  = $s
  $id =	1

  if ($left_index) { $s.offset(0, -1).value2 = $id }

  while ($e.row -lt $LastRow) {
    $e = $e.offset(1)

    if ($e.value2 -eq $s.value2) {
      $e.value2 = ""
      $ws.range($s, $e).merge()

      if ($left_index) { $ws.range($s.offset(0, -1), $e.offset(0, -1)).merge() }
    }
    else {
      $s = $e

      if ($left_index) { $e.offset(0, -1).value2 = ++$id }
    }
  }


}


export-modulemember  -function merge_run_csv,tc_status_info,fg_cal_passfail
