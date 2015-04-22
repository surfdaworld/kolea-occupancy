#Include AHKCSV.ahk

global checkouts     := "checkouts.csv"
global checkins      := "checkins.txt"
global occupiedlist  := "occupied.csv"
global shutoffs      := "shutoffs.txt"
global inputcsv      := "input.csv"
global BillingTotals := "billing.csv"
global DateStart     :=
global DateEnd       :=
global DaysInRange   :=
global DateArray     := []
global DateRange     :=
global UnitArray     := []
global Occupancy     := [] ; master occupancy array, UnitNumber:Date(occupied)
global Billable      := 1  ; store billable variable for use as Occupancy[] array dimension
global Occupied      := 2  ; store occupied variable for use as Occupancy[] array dimension
global CleanupPrompt :=    ; decide whether to prompt user to save/delete temp files (used in GUI checkbox)

global ColNumParty   := 8  ; column number of number of people in party
global ColArrDate    := 2  ; column number of arrival date
global ColDepDate    := 4  ; column number of departure date
global ColOwnRes     := 11 ; column number of Owner Res.
global ColUnit       := 1  ; column number of Unit

FileDelete, %OccTotals%
FileDelete, %inputcsv%
FileDelete, %BillingTotals%
FileDelete, %checkouts%
FileDelete, %checkins%
FileDelete, %occupiedlist%
FileDelete, %shutoffs%

goto, MakeChoice

3ButtonGenerateCodeList:
Gui, 3:show, hide
GetCSV()
GetInOutOcc()
GetShutoffs()
TestExist()
Cleanup(CleanupPrompt)

ExitApp

3ButtonMonthlyBillingReport:
Gui, 3:show, hide
GetCSV()
Gosub GetDateRange
CreateArrays()
ProcessOccupancy()
GetTotals()
Notify()
Cleanup(CleanupPrompt)

ExitApp

;=================================================================
;Set up GUI elements
;=================================================================
GetDateRange:
	Gui, 2:Add, MonthCal,multi x2 y0 w230 h170 vDateRange, 
	Gui, 2:Add, Button, x7 y170 w110 h30 , &Submit
	Gui, 2:Add, Button, x127 y170 w100 h30 , &Cancel
	Gui, 2:Show, w237 h206, Choose Dates
	WinWaitClose, Choose Dates
2ButtonSubmit:
	Gui, 2:Submit
return

2GuiClose:
2ButtonCancel:
	ExitApp
	
MakeChoice:
Gui, 3:Add, Button, x12 y10 w120 h30 , Generate Code List
Gui, 3:Add, Button, x142 y10 w130 h30 , Monthly Billing Report
Gui, 3:Add, CheckBox, x72 y50 w130 h30 vCleanupPrompt, Keep temporary files
; Generated using SmartGUI Creator for SciTE
Gui, 3:Show, w280 h94, Kolea HOA Access Control
WinWaitClose, Kolea HOA Access Control
return

3GuiClose:
ExitApp
	
;=================================================================
;Get filename of input xls file, and convert it to CSV
;=================================================================
GetCSV()
{
	FileSelectFile, inputxls, 3
	if inputxls =
	{
		ExitApp
	}
	inputcsvtemp = %A_ScriptDir%\%inputcsv%
	oExcel := ComObjCreate("Excel.Application")
	oBook := oExcel.Workbooks.Open(inputxls)
	oBook.SaveAs(inputcsvtemp, 6)
	oBook.Close(False)
	oExel.Quit
}

;=================================================================
;Parse inputcsv for checkins, checkouts, and occupied units, and
;dump unit list to checkins/checkouts/occupied txt files
;=================================================================
GetInOutOcc()
{
	CSV_Load(inputcsv,1)
    FileLength := CSV_TotalRows(1)
	Loop, %FileLength%
		{
			CheckoutDay := CSV_ReadCell(1, A_Index, ColDepDate)
            Unit := CSV_ReadCell(1, A_Index, ColUnit)
			CheckinDay := CSV_ReadCell(1, A_Index, ColArrDate)
			
            FormatTime, DateToday,, M/d/yyyy
			DateTodaytemp := convert(DateToday)
			CheckinDaytemp :=convert(CheckinDay)
			CheckoutDaytemp := convert(CheckoutDay)
            EnvSub, CheckoutDaytemp, DateTodaytemp, days
			EnvSub, CheckinDaytemp, DateTodaytemp, days
			If CheckoutDaytemp = -1
			{
              FileAppend, %Unit%`n, %checkouts%
			}
			
			If CheckinDaytemp = 1
			{
              FileAppend, %Unit%`n, %checkins%
			}
            
            If (CheckinDaytemp < 1 AND CheckoutDaytemp > -1)
            {
              FileAppend, %Unit%`n, %occupiedlist%
            }
		}
}

;=================================================================
;Parse checkout list with occupancy and checkin list, to generate final list
;of codes to shut off
;=================================================================
GetShutoffs()
{
	CSV_Load(checkouts,2)
	checkoutsLength := CSV_TotalRows(2)
	
	Loop, %checkoutsLength%
	{
		UnitCheckout := CSV_ReadCell(2, A_Index, ColUnit)
		OccupiedFlag := 0
		
		OccupiedFlag += UnitSearch(occupied, 3, UnitCheckout)
		OccupiedFlag += UnitSearch(checkins, 4, UnitCheckout)

		if OccupiedFlag = 0
		{
			FileAppend, %UnitCheckout%`n, %shutoffs%
		}
	}
}

;=================================================================
;Search FileName.csv for a specific unit number
;=================================================================
UnitSearch(FileName, FileNumber, UnitNumber)
{
	CSV_Load(FileName,FileNumber)
	FileLength := CSV_TotalRows(FileNumber)
	UnitFound := 0
	Loop, %FileLength%
	{
		UnitOccupied := CSV_ReadCell(FileNumber, A_Index, ColUnit)
		if UnitNumber = %UnitOccupied%
		{
			UnitFound := 1
		}
	}
	return, UnitFound
}

;=================================================================
;Convert m/d/yyyy to AHK timestamp format -TJ
;=================================================================
Convert(str) {
	StringSplit, Dates, str, "/"
	Month := Pad(Dates1)
	Day := Pad(Dates2)
	Year := Dates3
	NewDate = %Year%%Month%%Day%
	Return, NewDate
}

;=================================================================
;Add a leading '0' if str is only 1 digit long. Used to pad
;months and days for timestamp conversion
;=================================================================
Pad(str1) {
	StringLen, strLength, str1
	if strLength = 1
	{
		str2 = 0%str1%
	}
	else
	{
		str2 = %str1%
	}
	Return, str2
}

;=================================================================
;Notify the user if output is missing
;=================================================================
TestExist() {
	
	IfExist, %checkins%
	{
		Run, Notepad.exe "%checkins%"
	}
	else
	{
		MsgBox No checkins today!
	}

	IfExist, %shutoffs%
	{
		Run, Notepad.exe "%shutoffs%"
	}
	else
	{
		MsgBox No codes to shut off today!
	}
}

;=================================================================
;Ask user whether to delete temp files
;=================================================================
Cleanup(CleanupPrompt)
{
	If CleanupPrompt != 1
	{
		FileDelete, %checkouts%
		;FileDelete, %checkins%
		FileDelete, %occupiedlist%
		;FileDelete, %shutoffs%
		FileDelete, %inputcsv%
	}
}

;=================================================================
;Create reference arrays for processing
;=================================================================
CreateArrays()
{
	StringLeft, DateStart, DateRange, 8
	StringRight, DateEnd, DateRange, 8
	;MsgBox %DateRange%
	;Calculate DaysInRange
	DaysInRange := DateEnd
	EnvSub, DaysInRange, DateStart, days
	DaysInRange += 1
	
	;Create date reference array containing all dates in range
	DateTemp := DateStart
	;MsgBox Days in range: %DaysInRange%
	loop, %DaysInRange%
	{
		DateArray.Insert(DateTemp)
		EnvAdd, DateTemp, 1, days
	}
	
	Units = 01A,01B,01C,01D,01E,01F,02A,02B,02C,02D,02E,02F,03A,03B,03C,03D,03E,03F,04A,04B,04C,04D,04E,04F,05A,05B,05C,05D,05E,05F,06A,06B,06C,06D,06E,06F,07A,07B,07C,07D,07E,07F,08A,08B,08C,08D,08E,08F,09A,09B,09C,09D,09E,09F,09G,09H,09J,09K,09L,09M,10A,10B,10C,10D,10E,10F,11A,11B,11C,11D,11E,11F,11G,11H,11J,11K,11L,11M,12A,12B,12C,12D,12E,12F,13A,13B,13C,13D,13E,13F,14A,14B,14C,14D,14E,14F,14G,14H,14J,14K,14L,14M,15A,15B,15C,15D,15E,15F,15G,15H,15J,15K,15L,15M,16A,16B,16C,16D,16E,16F,16G,16H,16J,16K,16L,16M,Lot#01,Lot#02,Lot#03,Lot#04,Lot#05,Lot#06,Lot#07,Lot#08,Lot#09,Lot#10,Lot#11,Lot#12,Lot#13,Lot#14,Lot#15,Lot#16,Lot#17
	
	UnitArray := StrSplit(Units, ",")
		
		
}

;=================================================================
;Parse through inputcsv and populate Occupancy[]
;=================================================================
ProcessOccupancy()
{
	CSV_Load(inputcsv,1)
	FileLength := CSV_TotalRows(1)
	loop, 143 ;loop through all units
	{
		UnitIndex := A_Index
		CurrentUnit := UnitArray[A_Index]
		loop, %FileLength%
		{
			FileUnit := CSV_ReadCell(1, A_Index, ColUnit)
			CheckoutDay := CSV_ReadCell(1, A_Index, ColDepDate)
			CheckinDay := CSV_ReadCell(1, A_Index, ColArrDate)
			IsOwner := CSV_ReadCell(1, A_Index, ColOwnRes)
			StringLen, IsOwner, IsOwner
			
			CheckinDay := convert(CheckinDay)
			CheckoutDay := convert(CheckoutDay)
			
			If FileUnit = %CurrentUnit%
			{
				loop, %DaysInRange%
				{
					DayIndex := A_Index
					CheckinDayTemp := CheckinDay
					CheckoutDayTemp := CheckoutDay
					CurrentDay := DateArray[A_Index]
					StringLeft, CurrentDay, CurrentDay, 8
					EnvSub, CheckinDayTemp, CurrentDay, days
					EnvSub, CheckoutDayTemp, CurrentDay, days
					If CheckoutDayTemp > -1
					{
						If CheckinDayTemp < 1
						{
							Occupancy[UnitIndex, DayIndex, Occupied] := 1
							TestOcc := Occupancy[UnitIndex, DayIndex, Occupied]
							If IsOwner < 5
							{
								Occupancy[UnitIndex, DayIndex, Billable] := 1
							}
						}
					}
				}
			}
		}
	}
}

;=================================================================
;Get SUM of billable and non-billable occupied nights per unit
;=================================================================
GetTotals()
{
	FormatTime, DateStartTemp, %DateStart%, M/d/yyyy
	FormatTime, DateEndTemp, %DateEnd%, M/d/yyyy
	FileAppend,
	(
	Reporting Period: %DateStartTemp% to %DateEndTemp%
	Unit, Total Billable Days, Total Non-Billable Days
	
	), %BillingTotals%
	
	loop, 143
	{
		UnitIndex := A_Index
		BillableSum := 0
		NonBillableSum := 0
		OccupiedSum := 0
		
		loop, %DaysInRange%
		{
			DayIndex := A_Index
			BillableSum += Occupancy[UnitIndex, DayIndex, Billable]
			OccupiedSum += Occupancy[UnitIndex, DayIndex, Occupied]
		}
		NonBillableSum := OccupiedSum-BillableSum
		UnitTemp := UnitArray[UnitIndex]
		FileAppend,
		(
		%UnitTemp%, %BillableSum%, %NonBillableSum%
		
		), %BillingTotals%
	}		
}

;=================================================================
;Notify user of completion, and report file location
;=================================================================
Notify()
{
	MsgBox Finished processing.`nReport saved at: %A_ScriptDir%\billing.csv
	Run, %BillingTotals%
	SetTitleMatchMode, 2
	WinActivate, Excel
}
