#Include AHKCSV.ahk

global checkouts       := "checkouts.csv"
global checkins        := "checkins.txt"
global occupiedlist    := "occupied.csv"
global shutoffs        := "shutoffs.txt"
global inputcsv        := "input.csv"
global BillingTotals   := "billing.csv"
global inputxls        :=                ; input filename for excel file downloaded
global DateStart       :=
global DateEnd         :=
global DaysInRange     :=
global DateArray       := []
global DateRange       :=
global UnitArray       := []
global Occupancy       := []             ; master occupancy array, UnitNumber:Date(occupied)
global Billable        := 1              ; store billable variable for use as Occupancy[] array dimension
global Occupied        := 2              ; store occupied variable for use as Occupancy[] array dimension
global CleanupPrompt   :=                ; decide whether to prompt user to save/delete temp files (used in Gui1 checkbox)
global VersionNum      := "1.7"          ; Set version number for display
global FileContents    :=                ; variable for displaying occupancy data in Gui4
global TextWindow      :=                ; Gui4control variable for edit field to display file contents
global WindowTitle     := Occupancy Data ; Variable to store window name for Gui4

global ColNumParty   := 8  ; column number of number of people in party
global ColArrDate    := 2  ; column number of arrival date
global ColDepDate    := 4  ; column number of departure date
global ColOwnRes     := 11 ; column number of Owner Res.
global ColUnit       := 1  ; column number of Unit

;Delete any temp files from last run
FileDelete, %OccTotals%
FileDelete, %inputcsv%
FileDelete, %BillingTotals%
FileDelete, %checkouts%
FileDelete, %checkins%
FileDelete, %occupiedlist%
FileDelete, %shutoffs%

goto, MakeChoice ; Display Gui4 to prompt user for Daily Code List, or Monthly Billing Report

;=================================================================
;Generate Daily Code/Card Programming List
;=================================================================
3ButtonGenerateCodeList:
Gui, 3:Submit
Gui, 3:show, hide
GetCSV()
GetInOutOcc()
GetShutoffs()
DisplayData()
Cleanup(CleanupPrompt)
ExitApp

;=================================================================
;Generate monthly billing report
;=================================================================
3ButtonMonthlyBillingReport:
Gui, 3:Submit
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
;Prompt for billing report date range
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

;=================================================================
;Initial GUI displayed - prompt user for billing report/code list
;=================================================================
MakeChoice:
Gui, 3:Add, Button, x12 y10 w120 h30 , Generate Code List
Gui, 3:Add, Button, x142 y10 w130 h30 , Monthly Billing Report
Gui, 3:Add, CheckBox, x72 y50 w130 h30 vCleanupPrompt, Keep temporary files
Gui, 3:Add, Text, x240 y75 w30 h20 , v.%VersionNum%
Gui, 3:Show, w280 h94, Kolea HOA Access Control
WinWaitClose, Kolea HOA Access Control
return

3GuiClose:
ExitApp
	
;=================================================================
;Prompt for input xls file, and convert it to CSV
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
			IsOwner := CSV_ReadCell(1, A_Index, ColOwnRes)
			
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
              FileAppend,
				(
				%Unit%, %IsOwner%`n
				), %checkins%
			}
            
            If (CheckinDaytemp < 1 AND CheckoutDaytemp > -1)
            {
              FileAppend, %Unit%`n, %occupiedlist%
            }
		}
	CleanOccupancies(occupiedlist)
}

;=================================================================
;Clean up OccupiedList file by removing extraneous header lines,
;and filtering out duplicate unit numbers
;=================================================================
CleanOccupancies(occupiedlist)
{
	FileDelete, occtemp.txt ; delete temp file from previous run, if it exists
	CSV_Load(occupiedlist,11)
    FileLength := CSV_TotalRows(11)
	Loop, %FileLength%
	{
		UnitTemp := CSV_ReadCell(11, A_Index, 1)
		If (A_Index>3) ; ignore first three (junk) lines of OccupiedList header
		{
		FileAppend, %UnitTemp%`n, occtemp.txt
		}
	}
	FileDelete, %occupiedlist%
	FileMove, occtemp.txt, %occupiedlist% ; replace old OccupiedList with occtemp.txt
	FileDelete, occtemp.txt ; delete temp file
	
	FileRead, FileContents, %occupiedlist%
	
	Sort, FileContents
    FileDelete, %occupiedlist%
    FileAppend, %FileContents%, %occupiedlist%
    FileContents =  ; Free the memory.
}

;=================================================================
;Parse checkout list with occupancy and checkin list, to generate
;code list to shut off
;=================================================================
GetShutoffs()
{
	CSV_Load(checkouts,2)
	checkoutsLength := CSV_TotalRows(2)
	
	Loop, %checkoutsLength%
	{
		UnitCheckout := CSV_ReadCell(2, A_Index, ColUnit)
		OccupiedFlag := 0
		
		OccupiedFlag += UnitSearch(occupiedlist, 3, UnitCheckout)
		OccupiedFlag += UnitSearch(checkins, 4, UnitCheckout)

		if OccupiedFlag = 0
		{
			FileAppend, %UnitCheckout%`n, %shutoffs%
		}
	}
}

;=================================================================
;Search FileName.csv for a specific unit number, return 1 if found
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
;Convert m/d/yyyy to AHK timestamp format
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
;Notify the user if no checkins or checkouts
;=================================================================
DisplayData() {
	
	FileRead, FileContents, %checkins%
	CurrentFile = Checkins
	WindowTitle = Checkins
	Gui, 4:font, , Courier New
	Gui, 4:add, edit, +readonly vscroll x5 y5 w590 h370 vTextWindow, %FileContents%
	Gui, 4:Add, Button, x100 y380 w110 h30 , &Checkins
	Gui, 4:Add, Button, x250 y380 w110 h30 , &Shutoffs
	Gui, 4:Add, Button, x400 y380 w110 h30 , &Occupied
	Gui, 4:show,w600 h415,%WindowTitle%
	Gui, 4:+LastFound
	WinWaitClose
	return
	
	4ButtonCheckins:
	FileRead, FileContents, %checkins%
	CurrentFile = Checkins
	GuiControl,, TextWindow, %FileContents%
	WindowTitle = Checkins
	Gui, 4:show,w600 h415,%WindowTitle%
	return
	
	4ButtonShutoffs:
	FileRead, FileContents, %shutoffs%
	CurrentFile = Checkouts
	GuiControl,, TextWindow, %FileContents%
	WindowTitle = Shutoffs
	Gui, 4:show,w600 h415,%WindowTitle%
	return
	
	4ButtonOccupied:
	FileRead, FileContents, %occupiedlist%
	CurrentFile = Occupied Units
	GuiControl,, TextWindow, %FileContents%
	WindowTitle = Occupied Units
	Gui, 4:show,w600 h415,%WindowTitle%
	return
	
	4GuiClose:
	Gui, 4:Destroy
	return
}

;=================================================================
;Delete all temp files if user has not chosen to keep temp files
;=================================================================
Cleanup(CleanupPrompt)
{
	If CleanupPrompt != 1
	{
		FileDelete, %checkouts%
		FileDelete, %shutoffs%
		FileDelete, %checkins%
		FileDelete, %occupiedlist%
		FileDelete, %inputcsv%
		FileDelete, %inputxls%
	}
}

;=================================================================
;Create reference array of all dates in chosen range, and unit
;number master reference array.
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
;Parse through inputcsv and populate Occupancy[] array with all
;occupied units & dates, for monthly billing report
;=================================================================
ProcessOccupancy()
{
	CSV_Load(inputcsv,1)
	FileLength := CSV_TotalRows(1)
	loop, 143 ;loop through all units
	{
		UnitIndex := A_Index
		CurrentUnit := UnitArray[A_Index]
		loop, %FileLength% ; loop through file line by line
		{
			FileUnit := CSV_ReadCell(1, A_Index, ColUnit) ; store unit number of current line
			CheckoutDay := CSV_ReadCell(1, A_Index, ColDepDate) ; store checkout date of current line
			CheckinDay := CSV_ReadCell(1, A_Index, ColArrDate) ; store checkin date of current line
			IsOwner := CSV_ReadCell(1, A_Index, ColOwnRes) ; is this reservation for an owner?
			StringLen, IsOwner, IsOwner ; convert IsOwner to the length of itself, for billable testing later
			
			CheckinDay := convert(CheckinDay) ; convert CheckinDay to ahk timestamp format
			CheckoutDay := convert(CheckoutDay) ; convert CheckoutDay to ahk timestamp format
			CheckoutDay += -1, days ; subtract one day from checkout date (causes report to generate based on number of NIGHTS, rather than DAYS.)
			; As instructed per Dennis on 4/28/15 -TJ
			
			If FileUnit = %CurrentUnit%
			{
				loop, %DaysInRange% ; loop through all dates in user-selected range
				{
					DayIndex := A_Index
					CheckinDayTemp := CheckinDay
					CurrentDay := DateArray[A_Index] ; current day in loop (daysinrange)
					CheckoutDayTemp := CheckoutDay
					StringLeft, CurrentDay, CurrentDay, 8 ; crop time portion off of CurrentDay timestamp
					EnvSub, CheckinDayTemp, CurrentDay, days ; measure number of days from CheckinDayTemp to CurrentDay
					EnvSub, CheckoutDayTemp, CurrentDay, days ; measure number of days from CheckoutDayTemp to CurrentDay
					If CheckoutDayTemp > -1 ; if Checkout Day is after yesterday(CurrentDay)
					{
						If CheckinDayTemp < 1 ; if Checkin Day is before tomorrow(CurrentDay)
						{
							Occupancy[UnitIndex, DayIndex, Occupied] := 1 ; flag unit as occupied on this date in the Occupancy[] array
							If IsOwner < 5 ; Test length of IsOwner. Will only be < 5 if it is NOT an owner. Length will be = 5 if it is an owner.
							{
								Occupancy[UnitIndex, DayIndex, Billable] := 1 ; If NOT an owner, flag this day/unit as billable.
							}
						}
					}
				}
			}
		}
	}
}

;=================================================================
;Get sum of billable and non-billable (owner-occuped)
;nights per unit
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

;=================================================================
;Attempt to program all cards/codes by directly manipulating
;the AlarmLock mbd database
;=================================================================
ProgramCodes()
{
	;Populate card/code starting array (Unit, Cards, Code)
	;Shut off all cards/codes
	;Turn on appropriate cards
	;Turn on appropriate codes
}
