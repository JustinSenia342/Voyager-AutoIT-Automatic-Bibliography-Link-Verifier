#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>


;Cases in which this program Will not work:
;If there is a "" in the title being searched for as "" is used as a boolean instead of being searched for implicitly
;If it is a document that has protection on it
;If it is a document that is composed of scanned images, because they cannot be searched like text with ctrl F
;If the link takes you to a file system, or search directory
;If there is no 856 cell value in the bib document
;If there is a date entered after an zxcv, done so you can go back and recheck to make sure it's correct and up to date

;PROGRAMS YOU NEED INSTALLED FOR THIS PROGRAM TO WORK
;Windows 7 (Programmed with all service packs)
;Adobe reader (Programmed with version XI)
;Firefox 20.0.1
;Microsoft Excel (Progammed with Excel2011)
;Microsoft Word (Programmed with Word2011)
;Notepad (ver6.1 Service pack 1 Windows 7)
;Voyager (ver 8.2.0 Build 1422)

;PROGRAM SETTINGS REQUIRED TO WORK
;Firefox set to be default browser
;Firefoxset to always ask what to do when you click on a link to a pdf (and have it's default set to adobe reader)

;Heres what you need to do for the setup
;open adobe pdf reader, make it the default for pdf'S
;open voyager, search for one file, and save said file without making changes in order to preserve settings in prompt boxes
;open masterlist.docx in word
;populate word doc with proper numbers
;open masterexcel.xlsx
;open firefox, make sure it's your default and that it is set to ask about opening pdf files (default to adobe reader)
;open TEXTCHECK.TXT

;Opt ( "WinTitleMatchMode", 2 ) ; changes title mode to partial match


;********************Functions START*******************************
;hotkey to exit the loop, checks to see if "pause/break" has been pressed, exits if it has
$closed = True
HotKeySet("{PAUSE}", "CloseHotKey")
Func CloseHotKey()
    $closed = Not $closed
    While $closed

    Exit(0)

    WEnd
    ToolTip("")
EndFunc
;********************Functions END********************************

;**************Callibration Variables START*************************
;open the "AutoIT Window Info(x64 or x86)".exe and use it to determine X & Y Coordinates on your windows. however if your config and monitor resolution is like mine this whole step should be unecessary.
;Calibration Instructions: Please Open and then maximize Firefox, adobe reader, microsoft excel. also open/maximize Voyager, open a bibliography and the hypertext window by hitting Ctrl and the "K" key


;Calibration Instructions: Please move your mouse over the first cell in the first row of your excel window and then enter in those X and Y coordinates below
		Global $CalibrationXExcelFirstCell = 155
		Global $CalibrationYExcelFirstCell = 236

;Calibration Instructions: Please move your mouse over to the right side of the first line in your properly formatted word document and then enter in those X and Y coordinates below
		Global $CalibrationXWordFirstLine = 597
		Global $CalibrationYWordFirstLine = 264

;Calibration Instructions: Please move your mouse over the first cell in the first row of your voyager bib document and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerFirstCell = 60
		Global $CalibrationYVoyagerFirstCell = 308

;Calibration Instructions: please move your mouse over the Voyager document(not the program) close button in the upper right hand corner of the screen and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerCloseDocument = 1270
		Global $CalibrationYVoyagerCloseDocument = 35

;Calibration Instructions: enter in a junk character anywhere in the bib document, click the close button on the voyager document (not the program) and then mouse over the dialogue box button that reads "yes" in the popup that says "are you sure you want to close? there have been changes"  and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerAcceptClosing = 627
		Global $CalibrationYVoyagerAcceptClosing = 572

;Calibration Instructions: make the hypertext link window in voyager active, then mouse over the top-most link available and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerHyperlink = 446
		Global $CalibrationYVoyagerHyperlink = 407

;Calibration Instructions: open up a document in adobe reader, hit ctrl+F to open the search box, then searh for a long nonsense word (one that it won't find in the document), when it doesnt find it, mouse over the "ok" dialogue button and then enter in those X and Y coordinates below
		Global $CalibrationXAdobeReaderCloseDialogue = 827
		Global $CalibrationYAdobeReaderCloseDialogue = 543

;Calibration Instructions: click on the hypertext link window in voyager to make it active and then mouse over the hypertext window close button and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerHyperlinkClose = 641
		Global $CalibrationYVoyagerHyperlinkClose = 639

;Calibration Instructions: Go to firefox, mouse over the first internet tab  and then enter in those X and Y coordinates below
		Global $CalibrationXFirefoxTab = 67
		Global $CalibrationYFirefoxTab = 15

;Calibration Instructions: Now right click on the firefox tab you were hovering over, then mouse over the menu option "close all other tabs" and then enter in those X and Y coordinates below
		Global $CalibrationXFirefoxTabClose = 146
		Global $CalibrationYFirefoxTabClose = 174

;Calibration Instructions: Mouse over the "save to database" button in voyager and then enter in those X and Y coordinates below
		Global $CalibrationXVoyagerSaveToDatabase = 355
		Global $CalibrationYVoyagerSaveToDatabase = 72

;Calibration Instructions: open up a document in adobe reader, then mouse over the adobe reader document (not the program) close button and then enter in those X and Y coordinates below
		Global $CalibrationXAdobeReaderCloseDocument = 1268
		Global $CalibrationYAdobeReaderCloseDocument = 31
;**************Callibration Variables END***************************

;*********Declaring and Initializing variables START****************
Global $UserNameStore  ;temporarily stores username entered by user
Global $UserPassStore  ;temporarily stores password entered by user
Global $DelimTag = "zxcv" ;used for checking if bib links have already been checked
Global $AlreadyChekCount = 0  ;counter for various loops
Global $AlreadyChekStrArray[10]  ;array that stores data found in 856 row configurations to determine if the links have already been checked
Global $AlreadyChekBoolArray[10]  ;array that stores true/false values based on if the delimiter tag was found in any of the $AlreadyChekStrArray elements
Global $Int = 0  ;variable that stores the row numbers to check and make sure the program found the right row
Global $Int245 = 0  ;variable that is checked against a value of 245 to see if the right row is selected
Global $IntN856 = 0  ;variable that acts as a counter for the number of 856 rows in the bib document being searched
Global $String_StripAll = 8 ;variable used to determine the type of string stripping
Global $TagCheck = "A" ;holds the initial whole title copied from voyager (with delimiters and all)
Global $DelimiterPos = 0 ;used to store delimiter position
Global $StringLength = 0 ;used to store length of string being checked
Global $StringCut = 0 ;used to determine how much to cut off of the end of the first copied title string
Global $StringCut2 = 0 ;used to determine how much to cut off of the end of the second copied title String
Global $DocCheck = "A" ;stores what is found in the target document/default value if not found
Global $EndDoc = "Whee" ;used to determine when to stop the program, checks if masterlist is empty
Global $sStringd = "A" ;stored 1st copied title, used to format string while keeping original data intact
Global $sString = "A" ;stored string that was found in the target document, used to format string while keeping original data intact
Global $sString2 = "A" ;stored 2nd copied title, used to format string while keeping original data intact
Global $ErrorWinExists = 0 ;if the bibliography is not found, this is set to 1 and loop resets
Global $YCoordLocation = $CalibrationYVoyagerHyperlink ;y coordinate of hyperlinks in voyager, modified in program to click multiple links
Global $LinkNumberRepetitions = 0 ;drives number of time loop is performed, based off of $IntN856
Global $CorrectLinkArray[10] ;array used to determine if links were indeed correct or not
Global $ArrayForLoop = 0 ;used to drive array reset loop at beginning of program
Global $ZXCVAdjuster = 0 ;used to drive the element correctLinkArray is referring to, when assigning value to CorrectLinkArray
Global $ZXCVLoop = 0 ;used to drive the element correctLinkArray is referring to when using the value of correctLinkArray to enter delimiter values
Global $ExcelLoop = 0 ;used to drive the element correctLinkArray is referring to when entering Y/N values in Excel
Global $IsChecked = False ;used to determine if bib link has already been checked or not
Global $BibNum = 0 ;used to store bib # copied from word doc
Global $QuoteCheck = False ;used to check if there is a "Quote Error"
Global $dlBoxExists = 0 ;used to check if firefox couldn't find the string, this drives a closing function

Global $ArIn = 0
Global $ArCount = 0
;*********Declaring and Initializing variables END****************


;*********Graphical User Interface START**************************
#Region ### START Koda GUI section ### Form=c:\users\owner.myworkstation\desktop\gui\thegoodone.kxf
$Form1_1 = GUICreate("Form1", 699, 414, -1, -1, BitOR($WS_POPUP,$DS_MODALFRAME))
$Pic1 = GUICtrlCreatePic("C:\JSeniaAutoBib\Graphics\programGUIwreminder.jpg", 0, 0, 698, 413, 0)
GUICtrlSetState(-1, $GUI_DISABLE)
$BtStart = GUICtrlCreateButton("BtStart", 464, 216, 219, 57, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\StartUP.bmp", -1)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetCursor (-1, 0)
$BtOpenFiles = GUICtrlCreateButton("BtOpenFiles", 464, 144, 219, 57, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\buttons.bmp", -1)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetCursor (-1, 0)
$BtClose = GUICtrlCreateButton("Close", 624, 16, 59, 57, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\exit.bmp", -1)
GUICtrlSetCursor (-1, 0)
$InputUser = GUICtrlCreateInput("", 16, 200, 193, 21)
GUICtrlSetCursor (-1, 5)
$InputPass = GUICtrlCreateInput("", 16, 280, 193, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_PASSWORD))
GUICtrlSetCursor (-1, 5)
$BtSetup = GUICtrlCreateButton("BtSetup", 232, 216, 219, 57, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\SetupUP.bmp", -1)
GUICtrlSetCursor (-1, 0)
$BtInformation = GUICtrlCreateButton("BtInformation", 232, 144, 219, 57, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\InformationUP.bmp", -1)
GUICtrlSetCursor (-1, 0)
$BtSignIn = GUICtrlCreateButton("BtSignIn", 16, 320, 195, 41, $BS_BITMAP)
GUICtrlSetImage(-1, "C:\JSeniaAutoBib\Graphics\SignInUP.bmp", -1)
GUICtrlSetCursor (-1, 0)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
;*********Graphical User Interface END****************************

;*********Main Program Body/ GUI Switch cases START***************
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

			  ;CASE opens Information text file
			  Case $BtInformation
				  Run("cmd /c ""C:\JSeniaAutoBib\ProgramInformation.txt""","",@SW_HIDE)




			  ;CASE opens all necessary programs, logs into voyager, maximizes everything, puts cursor in proper places
			  Case $BtOpenFiles
				  ;QUICK OPEN FILES
				  Run("cmd /c ""C:\Program Files (x86)\Mozilla Firefox\firefox.exe""","",@SW_HIDE)
				  WinWaitActive("Eastern Michigan University - Mozilla Firefox")
				  Run("cmd /c ""C:\JSeniaAutoBib\MasterDocs\MasterExcel.xlsx""","",@SW_HIDE)
				  WinWaitActive("Microsoft Excel - MasterExcel")
				  Run("cmd /c ""C:\JSeniaAutoBib\MasterDocs\MasterList.docx""","",@SW_HIDE)
				  WinWaitActive("MasterList - Microsoft Word")
				  Run("cmd /c ""C:\JSeniaAutoBib\MasterDocs\TEXTCHECK.txt""","",@SW_HIDE)
				  WinWaitActive("TEXTCHECK - Notepad")
				  Run("cmd /c ""C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe""","",@SW_HIDE)
				  WinWaitActive("Adobe Reader")
				  Run("cmd /c ""C:\Voyager\Catalog.exe""","",@SW_HIDE)
				  WinWaitActive("Voyager Cataloging")
				  Send($UserNameStore)
				  Send("{TAB}")
				  Send($UserPassStore)
				  Send("{ENTER}")
				  Sleep (1000)


				  ;Prepping Windows (Maximizing, selecting proper boxes, etc etc)
				  WinActivate("Eastern Michigan University - Mozilla Firefox")
				  WinWaitActive("Eastern Michigan University - Mozilla Firefox")
				  Sleep (500)
				  Send("!")
				  Sleep (100)
				  Send("{SPACE}")
				  Sleep (100)
				  Send("x")
				  Sleep (500)
				  WinActivate("Voyager Cataloging")
				  WinWaitActive("Voyager Cataloging")
				  Sleep (500)
				  Send("!")
				  Sleep (100)
				  Send("{SPACE}")
				  Sleep (100)
				  Send("x")
				  Sleep (500)
				  WinActivate("MasterList - Microsoft Word")
				  WinWaitActive("MasterList - Microsoft Word")
				  Sleep (500)
				  Send("!")
				  Sleep (100)
				  Send("{SPACE}")
				  Sleep (100)
				  Send("x")
				  Sleep (500)
				  WinActivate("Microsoft Excel - MasterExcel")
				  WinWaitActive("Microsoft Excel - MasterExcel")
				  Sleep (500)
				  Send("!")
				  Sleep (100)
				  Send("{SPACE}")
				  Sleep (100)
				  Send("x")
				  Sleep (500)
				  MouseMove($CalibrationXExcelFirstCell, $CalibrationYExcelFirstCell)
				  MouseClick("left")
				  Send("{Left}");






			  ;CASE opens setup pdf that has setup information on it
			  Case $BtSetup
					  Run("cmd /c ""C:\JSeniaAutoBib\ProgramInformation\JSeniaAutoBibSetup.pdf""","",@SW_HIDE)




			  ;CASE initiates main program loop that does all work
			  Case $BtStart

				  ;setting proper windows to associated handles so they can be called later even though the window name may have changed
				  Global $handleWinPDF = WinGetHandle("Adobe Reader")
				  Global $handleWinTXT = WinGetHandle("Eastern Michigan University - Mozilla Firefox")

				  Do
					  ;resets variables for loop re-use
					  $Int = 0
					  $Int245 = 0
					  $IntN856 = 0
					  $TagCheck = "A"
					  $DelimiterPos = 0
					  $StringLength = 0
					  $StringCut = 0
					  $StringCut2 = 0
					  $DocCheck = "A"
					  $EndDoc = "Whee"
					  $ErrorWinExists = 0
					  $YCoordLocation = $CalibrationYVoyagerHyperlink
					  $linkNumberRepetitions = 0
					  $ArrayForLoop = 0
					  $ZXCVAdjuster = 0
					  $ZXCVLoop = 0
					  $ExcelLoop = 0
					  $AlreadyChekCount = 0
					  $IsChecked = False
					  $BibNum = 0
					  $QuoteCheck = False
					  $dlBoxExists = 0

					  ;resets all arrays to default values for loop re-use
					  Do
						 $CorrectLinkArray[$ArrayForLoop] = 0
						 $AlreadyChekStrArray[$ArrayForLoop] = ""
						 $AlreadyChekBoolArray[$ArrayForLoop] = Null
						 $ArrayForLoop = $ArrayForLoop + 1
						 Until $ArrayForLoop = 10

					  ;opens bib search in voyager
					  WinActivate("Voyager Cataloging")
					  WinWaitActive("Voyager Cataloging")
					  Sleep(1000)
					  Send("!r")
					  Send("i")
					  Send("b")
					  WinWaitActive("Retrieve a Record")

					  ;cuts bib number from MasterList.docx
					  WinActivate("MasterList - Microsoft Word")
					  WinWaitActive("MasterList - Microsoft Word")
					  MouseMove($CalibrationXWordFirstLine, $CalibrationYWordFirstLine)
					  MouseClick("left")
					  send("^+{LEFT}")
					  send("^x")
					  Sleep(500)

					  ;Ends doc if the word END is copied off of the Master word doc (END signals end of document)
					  $EndDoc = ClipGet()
					  If $EndDoc = ("END") Then
						 Exit 0
						 EndIf


					  ;pastes bibliography id in excel 1st column
					  WinActivate("Microsoft Excel - MasterExcel")
					  WinWaitActive("Microsoft Excel - MasterExcel")
					  $BibNum = ClipGet()
					  Send($BibNum)
					  Sleep(500)

					  ;reactivates voyager, pastes bliography number then searches for bib
					  WinActivate("Voyager Cataloging")
					  Sleep(500)
					  WinWaitActive("Retrieve a Record")
					  Send("^v")
					  Send("{ENTER}")


					  ;if bib isn't found enter not found in voyager, then moves on to next iteration of loop
					  ;this is done by checking to see if the error prompt pops up after the search is initiated
					  Sleep(500)
					  $ErrorWinExists = WinExists("Voyager Cataloging", "not found or unable to display it" )
					  If $ErrorWinExists = 1 Then
						 Send("{ENTER}")
						 WinActivate("Microsoft Excel - MasterExcel")
						 WinWaitActive("Microsoft Excel - MasterExcel")
						 Send("{RIGHT}")
						 Send("N")
						 Send("{DOWN}")
						 Send("{LEFT}")
						 Sleep(500)


					  Else

						 ;initiates856 counter: parses data in first column cells and counts the number of 856 rows
						 ;(done cy copying cell contents and comparing against predefined integer)
						 WinActivate("Voyager Cataloging")
						 WinWaitActive("Voyager Cataloging")
						 MouseMove($CalibrationXVoyagerFirstCell, $CalibrationYVoyagerFirstCell)
						 Sleep(300)
						 MouseClick("left")
						 Do
							Send("{DOWN}{ENTER}^+{LEFT}")
							Send("^c")
							$Int = ClipGet()
							If $Int = 856 Then $IntN856 = $IntN856 + 1
						 Until $Int = 962


						 ;initiates already checked method: parses first column of cells until it finds the cell containing 856
						 ;then it moves the cursor over to the column with the actual hyperlink/potential delimiter+tag, copies what is found
						 ;into an array, for each hyperlink that is available in voyager (number if iterations based on number of 856 rows found
						 ;in previous method)
						 WinActivate("Voyager Cataloging")
						 MouseMove($CalibrationXVoyagerFirstCell, $CalibrationYVoyagerFirstCell)
						 Sleep(300)
						 MouseClick("left")
						 Do
							Send("{DOWN}{ENTER}^+{LEFT}")
							Send("^c")
							$Int = ClipGet()
						 Until $Int = 856

					     Send("{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}")

						 Do
							Send("{ENTER}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}")
							Send("^c")
							$AlreadyChekStrArray[$AlreadyChekCount] = ClipGet()
							Send("{DOWN}")
							$AlreadyChekCount = $AlreadyChekCount +1
						 Until ($AlreadyChekCount = $IntN856)

						 WinActivate("Microsoft Excel - MasterExcel")
						 WinWaitActive("Microsoft Excel - MasterExcel")

						 ;initiaties already done checker: parses strings inside of array to check and see if any of the 856 lines copied
						 ;were already checked by doing a StringInString to find if the delimiter values were already entered
						 $AlreadyChekCount = 0
						 Do
							If StringInStr ($AlreadyChekStrArray[$AlreadyChekCount], $DelimTag, 0) > 2 Then
							   $IsChecked = True
							   $AlreadyChekBoolArray[$AlreadyChekCount] = True
							ElseIf StringInStr ($AlreadyChekStrArray[$AlreadyChekCount], $DelimTag, 0) = 0 Then
							   $AlreadyChekBoolArray[$AlreadyChekCount] = False
							EndIf
							$AlreadyChekCount = $AlreadyChekCount +1
						 Until ($AlreadyChekCount = $IntN856)


						 $AlreadyChekCount = 0

						 ;if string is already tagged, the results are entered in excel, the voyager document is closed
						 ;and the loop moves on to the next link to check, otherwise if not checked, continues on in the
						 ;link checking process
						 If $IsChecked = True    Then
							Send ("{RIGHT}Y{RIGHT}Y{RIGHT}"& $IntN856 & "{RIGHT}")
							Do
							   If $AlreadyChekBoolArray[$AlreadyChekCount] = True Then
								  Send ("AC{RIGHT}")
							   ElseIf $AlreadyChekBoolArray[$AlreadyChekCount] = False Then
								  Send ("NC{RIGHT}")
							   ElseIf $AlreadyChekBoolArray[$AlreadyChekCount] = Null Then
								  Send ("{RIGHT}")
							   EndIf
							   $AlreadyChekCount = $AlreadyChekCount + 1
							Until ($AlreadyChekCount = 10)
							Send ("{DOWN}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}")
							Sleep(1000)

							WinActivate("Voyager Cataloging")
							WinWaitActive("Voyager Cataloging")
							Sleep(500)
							MouseMove($CalibrationXVoyagerCloseDocument, $CalibrationYVoyagerCloseDocument)
							MouseClick("LEFT")
							MouseMove($CalibrationXVoyagerAcceptClosing, $CalibrationYVoyagerAcceptClosing)
							MouseClick("LEFT")



						 ElseIf $IsChecked = False    Then
							Send("{RIGHT}Y{RIGHT}N")



							;enters number of links found in excel spreadsheet
							WinActivate("Microsoft Excel - MasterExcel")
							WinWaitActive("Microsoft Excel - MasterExcel")
							Send("{RIGHT}"& $IntN856 &"{RIGHT}")




							;finds 245 by parsing up through the first column of cells until it finds the cell containing 245
							;(done cy copying cell contents and comparing against predefined integer)
							WinActivate("Voyager Cataloging")
							WinWaitActive("Voyager Cataloging")
							Send("{UP}{LEFT}{LEFT}{LEFT}")
							Do
							   Send("{UP}{ENTER}^+{LEFT}")
							   Send("^c")
							   $Int245 = ClipGet()
						    Until $Int245 = 245

							;selects and copies the title contents from the corresponding 245 row's field
							Send("{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{ENTER}")
							Send("+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}+^{LEFT}")
							Send("^c")
							$TagCheck = ClipGet()

						    ;assigns double dagger delimiter to variable in order to bypass AutoIT's inablilty to use the proper ASCII code
							$Delimiter = StringMid($TagCheck, 1, 1)

							;trims off a and souble dagger delimiter
							$TagCheck = StringTrimLeft($TagCheck, 3)
							$sString2 = $Tagcheck ;used to check for alternative spacing at end check

							;checks to find next iteration of delimiter
							$DelimiterPos = StringInStr ($TagCheck, $Delimiter)

							;gets string length
							$StringLength = StringLen ($TagCheck)

							;determines cut lengths (for two most common instances of bib formatting)
							$StringCut = ($StringLength - $DelimiterPos +4)
							$StringCut2 = ($StringLength - $DelimiterPos +2)

							;trims off everything after 2nd delimiter
							$TagCheck = StringTrimRight($TagCheck, $StringCut)
							$sString2 = StringTrimRight($sString2, $StringCut2)

							;setting variable to number of repetitions for loop counter
							$LinkNumberRepetitions = $IntN856

						    ;checks to see if any quotation marks were found, if they were, it documents that for later because
							;quotation marks act as booleans in a ctrl F search, which messes up how this program functions.
							If StringInStr($TagCheck, '"') > 0 Then
							  $QuoteCheck = True
							EndIf

							;initiates loop that repeats job until all links have been checked
							WinActivate("Voyager Cataloging")
							WinWaitActive("Voyager Cataloging")
							send("^k")


						    ;opens hyperlink window and clicks on links (must have links set to open with Firefox so the proper prompt window will open)
							Do
							   ;sets variables to default for loop re-use
							   $DocCheck = "WInniFrEDLiKESBannaNASJRJDFKR"
							   ClipPut ("Default")
							   $IPosition = 0
							   $dlBoxExists = 0

							   ;actually clicks on hyperlinks
							   MouseMove($CalibrationXVoyagerHyperlink, $YCoordLocation)
							   Sleep(300)
							   MouseClick("left")
							   Sleep(7000)

							   ;checks to see if either a pdf or text doc and handles accordingly (deals with variable download times)
							   ;if prompt comes up, it's a pdf that needs to be downloaded which can take time, this waits for the pdf
							   ;to download as well as gives it enough time to load
							   $dlBoxExists = WinExists("[CLASS:MozillaDialogClass]")
							   If $dlBoxExists = 1 Then
								  Sleep(200)
								  Send("{ENTER}")
								  WinWaitActive($handleWinPDF)
								  Sleep(6000)
							   EndIf


							   ;sometimes there is a comma in voyager title where there isn't in the actual document, this trims the comma
							   ;off of the copied string before it is entered into the pdf to be used to search.
							   $TagCheck = StringRegExpReplace($TagCheck, ",", "")
							   $sString2 = StringRegExpReplace($sString2, ",", "")
							   ;opens ctrl f search function, and types the string that was previously saved and trimmed
							   Sleep(500)
							   send("^f")
							   Sleep(100)
							   Send($TagCheck)
							   Sleep(100)

							   ;If hyperlink document was a pdf that was opened in adobe reader, searches the document, deals with "not found" errors
							   ;exits the search feature (which still keeps what was found highlighted), copies what was found to a varaible
							   If WinActive($handleWinTXT) = 0 Then
								  Sleep(300)
								  Send("{ENTER}")
								  Sleep(5000)
								  If WinActive("Adobe Reader", "Reader has finished searching the document. No matches were found.") Then
									 Sleep(300)
									 MouseMove($CalibrationXAdobeReaderCloseDialogue, $CalibrationYAdobeReaderCloseDialogue)
									 Sleep(300)
									 MouseClick("Left")
								  Else
									 Send("{ESCAPE}")
									 Sleep(1000)
									 Send("^c")
									 Sleep(300)
									 $DocCheck = ClipGet()
								  EndIf

							   ;Else the hyperlink was a text document and was opened up in firefox, searches the document, deals with "not found" errors
							   ;exits the search feature (which still keeps what was found highlighted), copies what was found to a varaible
							   ElseIf WinActive($handleWinPDF) = 0 Then
								  Sleep(600)
								  Send("{ESCAPE}")
								  Sleep(300)
								  Send("^c")
								  Sleep(1000)
								  $DocCheck = ClipGet()
								  ;Send("{ENTER}")
							   EndIf


							   WinActivate("MasterList - Microsoft Word")
							   WinWaitActive("MasterList - Microsoft Word")

							   ;setting strings to second variable for trimming
							   $sStringd = StringLower($TagCheck)
							   $sString = StringLower($DocCheck)

							   ;program retypes variables in notepad to make it into a standard uniform font so it can then be more accurately compared
							   WinActivate("TEXTCHECK - Notepad")
							   WinWaitActive("TEXTCHECK - Notepad")
							   Send("^a")
							   Sleep(200)
							   Send("{BACKSPACE}")

							   Send($sStringd)
							   Send("^a")
							   Sleep(200)
							   Send("^c")
							   Sleep(200)
							   $sStringd = ClipGet()
							   Send("{BACKSPACE}")

							   Send($sString)
							   Send("^a")
							   Sleep(200)
							   Send("^c")
							   Sleep(200)
							   $sString = ClipGet()
							   Send("{BACKSPACE}")

							   Send($sString2)
							   Send("^a")
							   Sleep(200)
							   Send("^c")
							   Sleep(200)
							   $sString2 = ClipGet()
							   Send("{BACKSPACE}")


							   ;trimming excess spaces and characters
							   $sStringd = StringRegExpReplace($sStringd, " |'|’|‘|–|\/|—", "")
							   $sString = StringRegExpReplace($sString, " |'|’|‘|–|\/|—", "")
							   $sString2 = StringRegExpReplace($sString2, " |'|’|‘|–|\/|—", "")

							   $sStringd = StringRegExpReplace($sStringd, "\(|\)", "")
							   $sString = StringRegExpReplace($sString, "\(|\)", "")
							   $sString2 = StringRegExpReplace($sString2, "\(|\)", "")

							   $sStringd = StringRegExpReplace($sStringd, '\-|\-|\:|\!|\@|\#|\$|\%|\^|\&|\*|\_|\+|\=|\;|"|\?|\/|\>|\.|\,|\<|\`|\~|\\' & @LF & @TAB, '')
							   $sString = StringRegExpReplace($sString, '\-|\-|\:|\!|\@|\#|\$|\%|\^|\&|\*|\_|\+|\=|\;|"|\?|\/|\>|\.|\,|\<|\`|\~|\\' & @LF & @TAB, '')
							   $sString2 = StringRegExpReplace($sString2, '\-|\-|\:|\!|\@|\#|\$|\%|\^|\&|\*|\_|\+|\=|\;|"|\?|\/|\>|\.|\,|\<|\`|\~|\\' & @LF & @TAB, '')

							   $sStringd = StringStripWS($sStringd, $String_StripAll)
							   $sString = StringStripWS($sString, $String_StripAll)
							   $sString2 = StringStripWS($sString2, $String_StripAll)



							   ;compares the results to see if they are equal
							   If($sStringd = $sString Or $sString = $sString2) Then
								 $IPosition = 1
							   EndIf



							   ;puts whether link was correct or not in an array

							   ;(Correct Link)
							   If $IPosition > 0 Then
								  $CorrectLinkArray[$ZXCVAdjuster] = 1
								  $ZXCVAdjuster = $ZXCVAdjuster + 1
								  WinActivate("Voyager Cataloging")
								  WinWaitActive("Voyager Cataloging")
								  WinActivate("Hypertext links")
								  WinWaitActive("Hypertext links")
								  Sleep(500)

							   ;(Incorrect Link)
							   ElseIf $IPosition = 0 Then
								  $CorrectLinkArray[$ZXCVAdjuster] = 2
								  $ZXCVAdjuster = $ZXCVAdjuster + 1
								  WinActivate("Voyager Cataloging")
								  WinWaitActive("Voyager Cataloging")
								  WinActivate("Hypertext links")
								  WinWaitActive("Hypertext links")
								  Sleep(500)

							   EndIf

								  ;modifies clicking coordinate variable and loop rep variable
								  $YCoordLocation = $YCoordLocation + 18
								  $LinkNumberRepetitions = $LinkNumberRepetitions - 1

						    Until ($LinkNumberRepetitions = 0)

						    ;closes hyperlink window
							WinActivate("Hypertext links")
							WinWaitActive("Hypertext links")
							MouseMove($CalibrationXVoyagerHyperlinkClose, $CalibrationYVoyagerHyperlinkClose)
							Sleep(300)
							MouseClick("Left")

							;finds 856 by parsing cells until it finds the value 856 in the first column
						    MouseMove($CalibrationXVoyagerFirstCell, $CalibrationYVoyagerFirstCell)
						    MouseClick("left")
						    Do
							   Send("{DOWN}{ENTER}^+{LEFT}")
							   Send("^c")
							   $Int = ClipGet()
						    Until $Int = 856

							;moves cursor to hyperlink location in document
							Send("{UP}{DOWN}{RIGHT}{RIGHT}{RIGHT}")

							;populates 856 fields with delimiter tags if applicable, skips if not applicable

						    Do
							   If $CorrectLinkArray[$ZXCVLoop] = 1 Then
								  Send ("{ENTER}{SPACE}{F9}xzxcv{DOWN}")
								  $ZXCVLoop = $ZXCVLoop + 1
								  ElseIf $CorrectLinkArray[$ZXCVLoop] = 0 Then
								  Send ("{DOWN}")
								  $ZXCVLoop = $ZXCVLoop + 1
								  ElseIf $CorrectLinkArray[$ZXCVLoop] = 2 Then
								  Send ("{DOWN}")
								  $ZXCVLoop = $ZXCVLoop + 1
							   EndIf
						    Until $ZXCVLoop = 10


							   ;clicks out of current open bib after saving
						    MouseMove($CalibrationXVoyagerSaveToDatabase, $CalibrationYVoyagerSaveToDatabase)
						    Sleep(300)
						    MouseClick("LEFT")
						    Sleep(1000)
						    MouseMove($CalibrationXVoyagerCloseDocument, $CalibrationYVoyagerCloseDocument)
						    Sleep(300)
						    MouseClick("LEFT")

							;populates excel spreadsheet with proper YES/NO/QERROR output
							WinActivate("Microsoft Excel - MasterExcel")
							WinWaitActive("Microsoft Excel - MasterExcel")

						    Do
							   If $QuoteCheck = True Then
								  Send ("QError")
							   ElseIf $CorrectLinkArray[$ExcelLoop] = 1 Then
								  Send("Y")
							   ElseIf $CorrectLinkArray[$ExcelLoop] = 2 Then
								  Send("N")
							   EndIf
							   $ExcelLoop = $ExcelLoop + 1
							   Send("{Right}")

						    Until $ExcelLoop = 10

							; handles returning to next line in excel spreadsheet
							Send ("{DOWN}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}")
							Sleep(1000)

						    ;Closes all firefox tabs except the first
							WinActivate($handleWinTXT)
							Sleep(1500)
							MouseMove($CalibrationXFirefoxTab, $CalibrationYFirefoxTab)
							MouseClick("Right")
							Sleep(300)
							MouseMove($CalibrationXFirefoxTabClose, $CalibrationYFirefoxTabClose)
							MouseClick("Left")
							Sleep(300)
							Send("{ENTER}")
							Sleep(800)

							WinActivate($handleWinPDF)
							Sleep(1500)
							;just in case document took too long to search, this is a safeguard to prevent the program from failing to close out of documents
							If WinActive("Adobe Reader", "Reader has finished searching the document. No matches were found.") Then
							   Sleep(300)
							   MouseMove($CalibrationXAdobeReaderCloseDialogue, $CalibrationYAdobeReaderCloseDialogue)
							   Sleep(300)
							   MouseClick("Left")
							   Sleep(1000)
						    EndIf
							MouseMove($CalibrationXAdobeReaderCloseDocument, $CalibrationYAdobeReaderCloseDocument)
							MouseClick("Left")
							Sleep(400)
							MouseClick("Left")
							Sleep(400)
							MouseClick("Left")
							Sleep(400)


						 EndIf
					  EndIf
					  Until $EndDoc = "No"




			  ;CASE Lets users enter their name and password, for Voyager Login, and enables the main program buttons
			  ;only after they do so
			  Case $BtSignIn
				  MsgBox($MB_SYSTEMMODAL, "GUI Event", "Thank you for Entering your UserName and Password, You may now use the program.")
				  $UserNameStore = GUICtrlRead($InputUser)
				  $UserPassStore = GUICtrlRead($InputPass)
				  GUICtrlSetState($BtStart, $GUI_ENABLE)
				  GUICtrlSetState($BtOpenFiles, $GUI_ENABLE)




			  ;close button for the program
			  Case $BtClose
				  Exit(0)




			  ;user input username
			  Case $InputUser

			  ;user input password
			  Case $InputPass


	EndSwitch
 WEnd
;*********Main Program Body/ GUI Switch cases END*****************

GUIDelete($Form1_1)