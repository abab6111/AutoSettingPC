#include <MsgBoxConstants.au3> ;視窗測試宣告
#include <Excel.au3> ; 開啟excel
#include <Process.au3> ;Run CMD

#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;#AutoIt3Wrapper_UseX64=n ; If target application is running as 32 bit code
;#AutoIt3Wrapper_UseX64=y ; If target application is running as 64 bit code

#include "C:\install\UI Automation\Spy tool\Includes\UIA_Constants.au3" ; Can be copied from UIASpy Includes folder
#include "C:\install\UI Automation\Spy tool\Includes\UIA_Functions.au3" ; Can be copied from UIASpy Includes folder
#include "C:\install\UI Automation\Zip.au3" ;解壓縮用套件
;#include "C:\UI Automation\Spy tool\Includes\CUIAutomation2-a.au3"
;#include "C:\UI Automation\Spy tool\Includes\UIA_Functions.au3"
;#include "UIA_SafeArray.au3" ; Can be copied from UIASpy Includes folder
;#include "UIA_Variant.au3" ; Can be copied from UIASpy Includes folder
;MsgBox($MB_SYSTEMMODAL, "Title", "This message box will timeout after 10 seconds or select the OK button.", 10)


#RequireAdmin

;Framework()
;AutoShutdown()



computer()
sleep(500)
close_power()
sleep(500)
move_anydesk()
sleep(500)
framework()
sleep(500)
photoviewer()
sleep(500)
host()
sleep(1000)
adobe()
sleep(1000)
zip()
sleep(1000)
adobe_pdf()
sleep(1000)
chrome_install()
sleep(1000)
chrome_bookmark()
sleep(500)
officeCMD()
sleep(500)
excel_ACT()
sleep(1000)
preset_program()
sleep(1000)
close_update()
sleep(1000)
hicos()
sleep(1000)
nt64()


;~ sleep(500)
;~ ;reboot()
;excel_open()






func computer() ;桌面顯示本機
	Run(@ComSpec & " /c " & 'C:\install\Cmd\Computer', "", @SW_HIDE)
	WinWaitActive("桌面圖示設定")
	Send("!m")
	sleep(500)
	Send("{enter}")
	ConsoleWrite( "本機 ok" & @CRLF )
EndFunc

func move_anydesk() ;把Anydesk和Office移到桌面
	Run(@ComSpec & " /c " & 'C:\install\Cmd\MoveAnydesk.cmd', "", @SW_HIDE)
	ConsoleWrite( "移到桌面OK" & @CRLF )
EndFunc

func close_power() ;關閉螢幕休眠
	Run(@ComSpec & " /c " & 'C:\install\Cmd\ClosePower.bat', "", @SW_HIDE)
	ConsoleWrite( "螢幕休眠關閉OK" & @CRLF )
EndFunc


func officeCMD() ;安裝office 2013
	;Run(@ComSpec & " /k " & 'C:\install\office xlm\install_office.bat', "", @SW_HIDE)
	;/k 參數表示“執行字符串指定的命令但保留”，若改? /c 則表示“執行字符串指定的命令然後終斷”。對此比較直觀的解釋是 /k 將在執行完命令後保留命令提示窗口，而 /c 則將在執行完命令之後關閉命令提示窗口。
	Run("C:\install\Office\Office2013\install_office.bat")
	WinWait("C:\Windows\system32\cmd.exe", "" , 5 )
    Local $num = 0
	;--------------------------------
	;等待office安裝完成
	While 1
		If WinExists("選取 C:\Windows\system32\cmd.exe") Then
		ConsoleWrite( "選取" & @CRLF )
		ElseIf   WinExists("C:\Windows\system32\cmd.exe") Then
		ConsoleWrite( "沒選取" & @CRLF )
		ElseIf   WinExists("Microsoft Office Professional Plus 2013") Then
		ConsoleWrite( "Microsoft Office Professional Plus 2013" & @CRLF )
		Else
		;跳出loop
		ConsoleWrite( "跳出loop" & @CRLF )
		ExitLoop
	EndIf
	Sleep(100)
	WEnd
	;------------------------------------
EndFunc

func excel_ACT()
	; Constants
	$xlforce   = True            ; True = force new instance, False = use existing
	$xlkeybrd  = True            ; True = allow keys,           False = block keys
	$xlscreen  = True            ; True = live screen,          False = suppressed
	$xlvisible = True            ; True = visible,                False = invisible

	; Variables
	$book  = 0                    ; The workbook
	$excel = 0                    ; The instance of Excel
	$range = 0                    ; Range object within a sheet

	; Create an instance of Excel and attach this script to it
	; Use the constants defined above to control visibility of the Excel window
	$excel = _Excel_Open($xlvisible, False, $xlvisible, $xlkeybrd, $xlforce)
	If ($excel = 0) Then ConsoleWrite( "--- Error ---" & @CRLF )


	; Create a new workbook with one worksheet within the opened instance
	$book = _Excel_BookNew($excel, 1)
	If ($book = 0) Then HandleError("Workbook creation failed ")
	WinWaitActive("Microsoft Office 啟動精靈")
	Send("!n")
	sleep(2000)
	WinClose("Microsoft Office 啟動精靈")
	WinWaitClose("Microsoft Office 啟動精靈")
	sleep(1000)
	WinClose("活頁簿1 - Excel")
	WinWaitClose("活頁簿1 - Excel")
EndFunc


func AutoShutdown()
	Send("#q")
	sleep(1000)
	Send("工作排程器{enter}")
	WinWaitActive("工作排程器")
	sleep(3000)
	Send("!a")
	sleep("1000")
	Send("m")
	sleep("1000")
	Send("自動關機{enter}")
	sleep("1000")
	Send("!d")
	sleep(1000)
	Send("C:\install")
	sleep(1000)
	Send("!o")
	WinWaitActive("建立工作")
	Send("{enter}")
	sleep(1000)
	Send("!{F4}")
	ConsoleWrite( "AutoShutdown OK" & @CRLF )
EndFunc


func framework()
	Run("C:\install\dotnetfx35.exe")
	WinWaitActive("Windows 功能")
	sleep(3000)
	ControlClick("Windows 功能" , "下載並安裝此功能" , "[CLASS:Button; INSTANCE:4]")
	WinWaitActive("Windows 功能" , "您可能需要重新啟動")
	ControlClick("Windows 功能" , "關閉" , "[CLASS:Button; INSTANCE:2]")
	ConsoleWrite( "framework OK" & @CRLF )
EndFunc


func hicos()
   Run(@ComSpec & " /c " & 'C:\install\Cmd\HicosDownload.cmd', "", @SW_MAXIMIZE)
;~    ConsoleWrite( "下載中" & @CRLF )
;~    WinWaitActive("C:\WINDOWS\system32\cmd.exe")
   While 1
	  If fileExists("C:\install\HiCOS_Client.zip") Then
		 ;跳出loop
		 ConsoleWrite( "下載完成" & @CRLF )
		 ExitLoop
	  Else
		 sleep(3000)
		 ConsoleWrite( "-----下載中-----" & @CRLF )
	  EndIf
   WEnd
   ;----------解壓縮------------
   Dim $ZipHicos,$Destpath
   $ZipHicos = "C:\install\HiCOS_Client.zip"
   $Destpath = "C:\install"
   _Zip_UnzipAll($ZipHicos, $Destpath , 0)
   While 1
	  If fileExists("C:\install\HiCOS_Client.exe") Then
		 ;跳出loop
		 ConsoleWrite( "跳出loop" & @CRLF )
		 ExitLoop
	  Else
		 sleep(1000)
	  EndIf
   WEnd
   ConsoleWrite("Extract OK" & @CRLF)
   ;----------解壓縮------------
   Run("C:\install\HiCOS_Client.exe")
   WinWaitActive("HiCOS Client")
   Local $hWnd = WinWait("HiCOS Client", "", 3)
   ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:3]")
   Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
   WinWaitActive("HiCOS Client","安裝成功")
   ControlClick($hWnd, "離開", "[CLASS:Button; INSTANCE:14]")
   ConsoleWrite( "Hicos OK" & @CRLF )
EndFunc

func host()
	Run("Notepad.exe")
	WinWaitActive("未命名 - 記事本")
	Send("!f") ;點選檔案
	sleep(1000)
	Send("o")  ;開啟舊檔
	sleep(1000)
	Send("hosts")	;host
	sleep(1000)
	Send("!d")	 ;focus 搜尋列
	sleep(1000)
	Send("C:\Windows\System32\drivers\etc") ;path
	sleep(1000)
	Send("{enter}")
	sleep(1000)
	Send("!o") ;開啟
	sleep(1000)

	;----------- 輸入host網址 ------------
	WinWaitActive("hosts - 記事本")
	send("^{end}")
	$hWnd = WinGetHandle("[ACTIVE]");
	$ret = DllCall("user32.dll", "long", "LoadKeyboardLayout", "str", "08040804", "int", 1 + 0)
	DllCall("user32.dll", "ptr", "SendMessage", "hwnd", $hWnd, "int", 0x50, "int", 1, "int", $ret[0])
	Send("172.16.3.122 tycgcloud.tycg.gov.tw{enter}")
	Send("172.20.1.37 property.tycg.gov.tw{enter}")
	Send("172.20.1.140 odis.tycg.gov.tw{enter}")
	Send("172.20.1.140 edms.tycg.gov.tw{enter}")
	Send("172.20.1.140 odiswebedit.tycg.gov.tw{enter}")
	Send("20.189.79.72 time.windows.com{enter}")
	Send("132.163.97.2 time.nist.gov{enter}")
	Send("10.0.0.11 hylib.typl.gov.tw{enter}")
	Send("10.0.0.14 hylibcir.typl.gov.tw{enter}")
	sleep(1000)
	Send("#{space}")
	Send("^s")
	sleep(1000)
	Send("!{F4}")
	ConsoleWrite( "host OK" & @CRLF )
EndFunc

func close_update()
	Run("C:\install\stopupdates10setup")
	WinWaitActive("選擇安裝語言")
	ControlClick("選擇安裝語言", "確定", "[CLASS:TNewButton; INSTANCE:1]")
	sleep(1000)
	Send("!n") ;下一步
	sleep(1000)
	send("!a") ;我接受
	sleep(500)
	send("!n") ;下一步
	sleep("500")
	send("!n")
	sleep("500")
	send("!n")
	sleep("500")
	send("!n")
	sleep("500")
	send("!i")  ;安裝
	sleep(1000)
	WinWaitActive("StopUpdates10 安裝程式","安裝完成")
	Send("!f")
	;WinWaitActive("StopUpdates10 - Download - Google Chrome")
	; Wait 10 seconds for the Notepad window to appear.
    Local $hWnd = WinWait("StopUpdates10 - Download", "", 10)
    ; Activate the Notepad window using the handle returned by WinWait.
    WinActivate($hWnd)
	sleep(1000)
	Send("!{F4}")


	WinWaitActive("停止更新 Windows 10 -")
	;ControlClick("停止更新 Windows 10", "停止 Windows Updates!", "[CLASS:TGreatisButton; INSTANCE:5]")
	Send("{enter}")
	WinWaitActive("那麼，下一步該怎麼做?")
	Send("!{F4}")
	Send("!{F4}")
	ConsoleWrite( "更新關閉 OK" & @CRLF )


	#cs
    ; Create UI Automation object
    Local $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtag_IUIAutomation )
    If Not IsObj( $oUIAutomation ) Then Return ConsoleWrite( "$oUIAutomation ERR" & @CRLF )
    ConsoleWrite( "$oUIAutomation OK" & @CRLF )

	; Get Desktop element
    Local $pDesktop, $oDesktop
    $oUIAutomation.GetRootElement( $pDesktop )
    $oDesktop = ObjCreateInterface( $pDesktop, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
    If Not IsObj( $oDesktop ) Then Return ConsoleWrite( "$oDesktop ERR" & @CRLF )
    ConsoleWrite( "$oDesktop OK" & @CRLF )

	; --- UpdateClose window ---

    ConsoleWrite( "--- UpdateClose window ---" & @CRLF )

    Local $pCondition ; Note that $UIA_ClassNamePropertyId maybe ia a CASE SENSITIVE condition
    $oUIAutomation.CreatePropertyCondition( $UIA_ClassNamePropertyId, "TGreatisButton", $pCondition )
    If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
    ConsoleWrite( "$pCondition OK" & @CRLF )

    Local $pUPC, $oUPC
    $oDesktop.FindFirst( $TreeScope_Descendants, $pCondition, $pUPC )
    $oUPC = ObjCreateInterface( $pUPC, $sIID_IUIAutomationElement , $dtag_IUIAutomationElement )
    If Not IsObj( $oUPC ) Then Return ConsoleWrite( "$oUPC ERR" & @CRLF )
    ConsoleWrite( "$oUPC OK" & @CRLF )

	; --- 停止windows update ---

	ConsoleWrite( "--- windows update ---" & @CRLF )

	Local $pCondition1
	$oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_CheckBoxControlTypeId, $pCondition1 )
	If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
	ConsoleWrite( "$pCondition1 OK" & @CRLF )


	Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
	$oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "停止 Windows Updates!", $pCondition2 ) ; File <<<<<<<<<<<<<<<<<<<<
	If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
	ConsoleWrite( "$pCondition2 OK" & @CRLF )

	; And condition
	$oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
	If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
	ConsoleWrite( "$pCondition OK" & @CRLF )

	Local $pCancel, $oCancel
	$oUPC.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
	$oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
	If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oUPC ERR" & @CRLF )
	ConsoleWrite( "$oUPC OK" & @CRLF )

	Local $pInvoke, $oInvoke
	$oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
	$oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
	If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
	ConsoleWrite( "$oInvoke OK" & @CRLF )
	$oInvoke.Invoke()
	Sleep( 100 )
	#ce
;~ 	WinWaitActive("那麼，下一步該怎麼做?")
;~ 	Send("!{F4}")
;~ 	Send("!{F4}")
;~ 	ConsoleWrite( "更新關閉 OK" & @CRLF )

EndFunc

func photoviewer()
	Run(@ComSpec & " /c " & 'C:\install\Reg\PhotoViewer.reg', "", @SW_HIDE)
	WinWaitActive("登錄編輯程式")
	Local $hWnd = WinWait("登錄編輯程式", "", 5) ; Wait 5 seconds for the installed window to appear.
	ControlClick($hWnd, "是", "[CLASS:Button; INSTANCE:1]")
	WinWaitActive("登錄編輯程式","確定")
	Local $hWnd = WinWait("登錄編輯程式", "", 5) ; Wait 5 seconds for the installed window to appear.
	ControlClick($hWnd, "確定", "[CLASS:Button; INSTANCE:1]")
	ConsoleWrite( "PhotoViewer OK" & @CRLF )
EndFunc

func preset_program()
	Send("#i") ;開設定
	WinWaitActive("設定")
	sleep(2000)
	Send("#{UP}")
	sleep(2000)
	Send("預設應用程式") ;點應用程式
    Sleep(1000)
	Send("{enter}")
    sleep(2000)
	Send("{enter}") ; 點預設應用程式
	sleep(2000)
	Send("{tab}{tab}{tab}{enter}") ;相片檢視器
	sleep(3000)
	Send("{tab}{enter}")
	sleep(6000)
	Send("{tab}{tab}{enter}") ; 瀏覽器
	sleep(2000)
	Send("{tab}{enter}")
	sleep(2000)
	Send("{tab}{enter}")
	sleep(2000)

	Send("!{F4}")
	ConsoleWrite( "Preset OK" & @CRLF )
EndFunc

func office()
   Run("C:\install\office 2013 64 繁體中文\install_office.bat")
   #cs
   ;Run("cmd")
   ;WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
   ;Send("cd C:\install\office 2013 64 繁體中文{enter}")
   ;sleep(500)
   ;Send("start Setup.exe{enter}")
   Send("#e")
   WinWaitActive("檔案總管")
   Send("!d")
   sleep(500)
   Send("C:\install\office 2013 64 繁體中文\setup.exe")
   sleep(500)
   Send("{enter}{enter}")
   #ce
   WinWaitActive("Microsoft Office Professional Plus 2013")
   Send("a")
   sleep(50)
   Send("!c")
   Send("!i")

     ; Create UI Automation object
    Local $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtag_IUIAutomation )
    If Not IsObj( $oUIAutomation ) Then Return ConsoleWrite( "$oUIAutomation ERR" & @CRLF )
    ConsoleWrite( "$oUIAutomation OK" & @CRLF )

    ; Get Desktop element
    Local $pDesktop, $oDesktop
    $oUIAutomation.GetRootElement( $pDesktop )
    $oDesktop = ObjCreateInterface( $pDesktop, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
    If Not IsObj( $oDesktop ) Then Return ConsoleWrite( "$oDesktop ERR" & @CRLF )
    ConsoleWrite( "$oDesktop OK" & @CRLF )

	; --- office window ---

    ConsoleWrite( "--- office window ---" & @CRLF )

    Local $pCondition ; Note that $UIA_ClassNamePropertyId maybe ia a CASE SENSITIVE condition
    $oUIAutomation.CreatePropertyCondition( $UIA_ClassNamePropertyId, "NetUIButton", $pCondition )
    If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
    ConsoleWrite( "$pCondition OK" & @CRLF )

    Local $pOffice, $oOffice
    $oDesktop.FindFirst( $TreeScope_Descendants, $pCondition, $pOffice )
    $oOffice = ObjCreateInterface( $pOffice, $sIID_IUIAutomationElement , $dtag_IUIAutomationElement )
   ; While we haven't found the next button
   While Not IsObj($oOffice)
	  ; Search for it
	  $oDesktop.FindFirst( $TreeScope_Descendants, $pCondition, $pOffice )
	  ; Attempt to create the object
	  $oOffice = ObjCreateInterface( $pOffice, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
	  ; If we found it (it's an object)
	  If IsObj( $oOffice ) Then
		 ConsoleWrite( "$oOffice OK" & @CRLF)
	  Else
		 ; Wait a little before retrying so we don't kill the old computers
		 Sleep(500)

		 ; Optionally, count the number of times this happens...
		 ;~ $iCount += 1
		 ; If it loops too many times, return an error
		 ;~ If $iCount >= 10 Then Return SetError(1, 0, False)
	  EndIf
   WEnd
   Send("!c")
   sleep(1000)
   ConsoleWrite( "office OK" & @CRLF )
EndFunc

func reboot()
	Shutdown(2)
EndFunc

func excel_open()
   ; Constants
$xlforce   = True            ; True = force new instance, False = use existing
$xlkeybrd  = True            ; True = allow keys,           False = block keys
$xlscreen  = True            ; True = live screen,          False = suppressed
$xlvisible = True            ; True = visible,                False = invisible

; Variables
$book  = 0                    ; The workbook
$excel = 0                    ; The instance of Excel
$range = 0                    ; Range object within a sheet

; Create an instance of Excel and attach this script to it
; Use the constants defined above to control visibility of the Excel window
$excel = _Excel_Open($xlvisible, False, $xlvisible, $xlkeybrd, $xlforce)
If ($excel = 0) Then ConsoleWrite( "--- Error ---" & @CRLF )


; Create a new workbook with one worksheet within the opened instance
$book = _Excel_BookNew($excel, 1)
If ($book = 0) Then HandleError("Workbook creation failed ")

   Send("!f")
   sleep(1000)
   Send("dy2")
   sleep(3000)
   ;Send("{tab}{tab}{enter}")
   ;Send("Y"); Enter Licence
   Send("YC7DK")
   Send("!i")
   sleep(1000)
   WinClose("活頁簿1 - Excel")
   WinWaitClose("活頁簿1 - Excel")
EndFunc

func chrome_install()
	Run("C:\install\ChromeSetup")
	ConsoleWrite( "開始裝chrome" & @CRLF )
	$iTimeout = TimerInit()
	While 1
		If TimerDiff($iTimeout) >= 600000 Then
			ConsoleWrite( "7-zip Error" & @CRLF ) ;"5min timeout before either window came up!"
			Exit
		ElseIf WinExists("歡迎使用 Chrome - Google Chrome") Then
			; Hanle it
			Send("!{F4}")
			ExitLoop
		ElseIf WinExists("Google Chrome 安裝程式" , "安裝完成") Then
			; Handle it
			Local $hWnd = WinWait("Google Chrome 安裝程式", "", 5) ; Wait 5 seconds for the installed window to appear.
			ControlClick($hWnd, "關", "[CLASS:Button; INSTANCE:2]")
			ExitLoop
		EndIf

		Sleep(100)
	WEnd

	;WinWaitActive("歡迎使用 Chrome - Google Chrome")
	;Send("!{F4}")
	;WinWaitActive("新分頁 - Google Chrome")
	;WinClose("新分頁 - Google Chrome")
	ConsoleWrite( "Chrome OK" & @CRLF )
EndFunc

func chrome_bookmark()
   ;	匯入書籤
   Run("C:\Program Files\Google\Chrome\Application\chrome.exe")
  ;WinWaitActive("歡迎使用 Chrome - Google Chrome")
   sleep(5000)
   Send("^t")
   sleep(1000)
   Opt("WinTitleMatchMode", 1) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
   WinWaitActive("新分頁")
   ;Send("{LSHIFT}") ;切換輸入法
   Send("chrome://settings/importData{enter}") ;import bookmark
   Sleep(3000)
   ;open file manager
   Send("{Tab}")
   Sleep(500)
   Send("{Enter}")
   sleep(500)
   Send("{Down}")
   sleep(500)
   Send("{Down}")
   sleep(500)
   Send("{Down}")
   sleep(500)
   Send("{Down}")
   sleep(500)
   Send("{enter}")
   sleep(500)
   Send("{TAB}{TAB}{TAB}{enter}")
   sleep(1000)
   Send("{enter}")
   sleep(1000)
   ;input data
   Send("書籤{enter}")
   Send("!d")
   sleep(1000)
   ;跳出視窗
   Send("C:\install\{enter}{enter}{enter}")
   Sleep(1000)
   Send("!o")
   sleep(1000)
   Send("!{F4}")
   ConsoleWrite( "bookmark OK" & @CRLF )
EndFunc

func adobe()
   ;	安裝Adobe PDF Reader
   Run("C:\install\AcroRdrDC1900820071_zh_TW.exe")
   WinWaitActive("Adobe Acrobat Reader DC (Continuous) - 設定")
   Send("!i")
   WinWaitActive("Adobe Acrobat Reader DC (Continuous) - 設定","安裝程式已完成")
   Send("!f")
   ConsoleWrite( "Adobe OK" & @CRLF )
EndFunc

func adobe_pdf()
	; 設定pdf預設開啟
	Run("C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32")
	; Run(@ComSpec & " /c " & 'C:\install\Cmd\RunAdobe', "", @SW_HIDE)
	WinWaitActive("Adobe Acrobat Reader DC - 用於個人電腦上的分發授權合約" , "" , 5)
	ConsoleWrite("分發授權合約視窗跳出" & @CRLF )
	sleep(1000)
	ControlClick("Adobe Acrobat Reader DC - 用於個人電腦上的分發授權合約", "接受", "[CLASS:Button; INSTANCE:2]")
	WinWaitActive("Acrobat Reader" , "將 Adobe Acrobat Reader 設定")
	ConsoleWrite("分發授權合約接受" & @CRLF )
	Send("!y")
	sleep(6000)
	Send("{TAB}{Enter}")
	WinWaitActive("Click on 'Change' to select default PDF handler - 內容")
	Send("!c")
	sleep(1000)
	Send("{TAB}{TAB}{TAB}{Down}{Enter}")
	sleep(1000)
	ControlClick("Click on 'Change' to select default PDF handler - 內容", "確定", "[CLASS:Button; INSTANCE:5]")
	sleep(3000)
	Send("!{F4}")
	sleep(3000)
	Send("!{F4}")
	ConsoleWrite("PDF OK" & @CRLF )
EndFunc

func zip()
	;	安裝7-zip
	Run("C:\install\7z1604-x64.exe")
	WinWaitActive("7-Zip 16.04 (x64) Setup")
	Send("!i")
	Winwait("7-Zip 16.04 (x64) Setup", "You must restart your system to complete the installation", 5 )
	If WinExists("7-Zip 16.04 (x64) Setup" , "You must restart your system to complete the installation") Then
		Send("!n")
		ConsoleWrite( "Restart Close" & @CRLF )
	Else
		ConsoleWrite( "沒顯示" & @CRLF )
	EndIf

	#cs
	$iTimeout = TimerInit()
	While 1
		If TimerDiff($iTimeout) >= 60000 Then
			ConsoleWrite( "7-zip Error" & @CRLF ) ;"30sec timeout before either window came up!"
			Exit
		ElseIf WinExists("7-Zip 16.04 (x64) Setup", "You must restart your system to complete the installation") Then
			; Hanle it
			WinWaitActive("7-Zip 16.04 (x64) Setup","You must restart your system to complete the installation")
			Send("!n")
			ExitLoop
		ElseIf WinExists("7-Zip 16.04 (x64) Setup" , "7-Zip 16.04 (x64) is installed") Then
			; Handle it
			ConsoleWrite( "7-zip OK 2" & @CRLF )
			ExitLoop
		EndIf
		Sleep(100)
	WEnd
	#ce

	;Send("!n")
	;WinWaitActive("7-Zip 16.04 (x64) Setup","7-Zip 16.04 (x64) is installed")
	sleep(1000)
	WinClose("7-Zip 16.04 (x64) Setup")
	WinWaitClose("7-Zip 16.04 (x64) Setup")
	ConsoleWrite( "7-zip OK" & @CRLF )
EndFunc

func nt64()

    Run("C:\install\eis_nt64")
	WinWaitActive("ESET Internet Security")
	Sleep( 5000 )

    ; Create UI Automation object
    Local $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtag_IUIAutomation )
    If Not IsObj( $oUIAutomation ) Then Return ConsoleWrite( "$oUIAutomation ERR" & @CRLF )
    ConsoleWrite( "$oUIAutomation OK" & @CRLF )

    ; Get Desktop element
    Local $pDesktop, $oDesktop
    $oUIAutomation.GetRootElement( $pDesktop )
    $oDesktop = ObjCreateInterface( $pDesktop, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
    If Not IsObj( $oDesktop ) Then Return ConsoleWrite( "$oDesktop ERR" & @CRLF )
    ConsoleWrite( "$oDesktop OK" & @CRLF )

	; --- Nt64 window ---

    ConsoleWrite( "--- Nt64 window ---" & @CRLF )

    Local $pCondition ; Note that $UIA_ClassNamePropertyId maybe ia a CASE SENSITIVE condition
    $oUIAutomation.CreatePropertyCondition( $UIA_ClassNamePropertyId, "#32770", $pCondition )
    If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
    ConsoleWrite( "$pCondition OK" & @CRLF )

    Local $pNt64, $oNt64
    $oDesktop.FindFirst( $TreeScope_Descendants, $pCondition, $pNt64 )
    $oNt64 = ObjCreateInterface( $pNt64, $sIID_IUIAutomationElement , $dtag_IUIAutomationElement )
    If Not IsObj( $oNt64 ) Then Return ConsoleWrite( "$oNotepad ERR" & @CRLF )
    ConsoleWrite( "$oNotepad OK" & @CRLF )

   ; --- 取消安裝較新版本 ---

   ConsoleWrite( "--- Cancel ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId, $UIA_CheckBoxControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "安裝較新版本", $pCondition2 ) ; File <<<<<<<<<<<<<<<<<<<<
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )

  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 100 )

   ; --- 繼續 ---
   ConsoleWrite( "--- 繼續 ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "繼續", $pCondition2 ) ; File <<<<<<<<<<<<<<<<<<<<
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )

  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 6000 )

   ; --- 我接受 ---
   ConsoleWrite( "--- 我接受 ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "我接受", $pCondition2 ) ;
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  ConsoleWrite( $oCancel & @CRLF )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )


  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 7000 )

  ; --- 使用購買的授權金鑰 ---
  ConsoleWrite( "--- 使用購買的授權金鑰 ---" & @CRLF )
  Send("{TAB}{TAB}{TAB}{enter}")
  ConsoleWrite( "Enter OK" & @CRLF )
  sleep(7000)

  ;

  ;輸入金鑰  BV4E-XFD3-NXW6-CWBJ-8BS2
  ConsoleWrite( "--- 輸入金鑰 ---" & @CRLF )
  Send("{TAB}{TAB}{TAB}")
  sleep(500)
  Send("BV4E-XFD3-NXW6-CWBJ-8BS2")
  sleep(500)



  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "繼續", $pCondition2 ) ;
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  ConsoleWrite( $oCancel & @CRLF )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )


  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 7000 )

   ; --- 繼續 ---
   ConsoleWrite( "--- 繼續 ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "繼續", $pCondition2 ) ;
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  ConsoleWrite( $oCancel & @CRLF )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )


  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 7000 )

  ; --- 啟用 ESET ---
  ConsoleWrite( "--- 啟用ESET ---" & @CRLF )
  Send("{TAB}{TAB}{TAB}{enter}")



   ; --- 啟用潛在不需要應用程式偵測 ---
   ConsoleWrite( "--- 啟用潛在不需要應用程式偵測 ---" & @CRLF )
   Send("{TAB}{TAB}{TAB}{enter}")
   sleep(500)
   Send("{TAB}{TAB}{enter}")
   sleep(500)


   ; --- 是 我想加入計畫 ---
   ConsoleWrite( "--- 是 我想加入計畫 ---" & @CRLF )
   Send("{TAB}{TAB}{TAB}{TAB}{enter}")

   ; --- 安裝 ---
   ConsoleWrite( "--- 安裝 ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "安裝", $pCondition2 ) ;
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )


  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
  Sleep( 15000 )


   ; --- 完成 ---
   ConsoleWrite( "--- 完成 ---" & @CRLF )

  Local $pCondition1
  $oUIAutomation.CreatePropertyCondition( $UIA_ControlTypePropertyId,  $UIA_ButtonControlTypeId, $pCondition1 )
  If Not $pCondition1 Then Return ConsoleWrite( "$pCondition1 ERR" & @CRLF )
  ConsoleWrite( "$pCondition1 OK" & @CRLF )


  Local $pCondition2 ; $UIA_NamePropertyId is LOCALIZED and maybe CASE SENSITIVE
  $oUIAutomation.CreatePropertyCondition( $UIA_NamePropertyId, "完成", $pCondition2 ) ;
  If Not $pCondition2 Then Return ConsoleWrite( "$pCondition2 ERR" & @CRLF )
  ConsoleWrite( "$pCondition2 OK" & @CRLF )

  ; And condition
  $oUIAutomation.CreateAndCondition( $pCondition1, $pCondition2, $pCondition )
  If Not $pCondition Then Return ConsoleWrite( "$pCondition ERR" & @CRLF )
  ConsoleWrite( "$pCondition OK" & @CRLF )

  Local $pCancel, $oCancel
  $oNt64.FindFirst( $TreeScope_Descendants, $pCondition, $pCancel )
  $oCancel = ObjCreateInterface( $pCancel, $sIID_IUIAutomationElement, $dtag_IUIAutomationElement )
  ConsoleWrite( $oCancel & @CRLF )
  If Not IsObj( $oCancel ) Then Return ConsoleWrite( "$oCancel ERR" & @CRLF )
  ConsoleWrite( "$oCancel OK" & @CRLF )


  Local $pInvoke, $oInvoke
  $oCancel.GetCurrentPattern( $UIA_InvokePatternId, $pInvoke )
  $oInvoke = ObjCreateInterface( $pInvoke, $sIID_IUIAutomationInvokePattern, $dtag_IUIAutomationInvokePattern )
  If Not IsObj( $oInvoke ) Then Return ConsoleWrite( "$oInvoke ERR" & @CRLF )
  ConsoleWrite( "$oInvoke OK" & @CRLF )
  $oInvoke.Invoke()
;~   Sleep( 1000 )



WinWaitActive("設定其他 ESET 安全性工具 - ESET Internet Security")
WinClose("設定其他 ESET 安全性工具 - ESET Internet Security")
WinWaitClose("設定其他 ESET 安全性工具 - ESET Internet Security")
sleep(100)

WinWaitActive("ESET Internet Security")
WinClose("ESET Internet Security")
WinWaitClose("ESET Internet Security")
sleep(100)

ConsoleWrite( "nt64 OK" & @CRLF )
EndFunc


