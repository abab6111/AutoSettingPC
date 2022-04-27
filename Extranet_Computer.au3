#include <MsgBoxConstants.au3> ;視窗測試宣告
#include <Excel.au3> ; 開啟excel
#include <Process.au3> ;Run CMD

#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;#AutoIt3Wrapper_UseX64=n ; If target application is running as 32 bit code
;#AutoIt3Wrapper_UseX64=y ; If target application is running as 64 bit code

#include "C:\install\UI Automation\Spy tool\Includes\UIA_Constants.au3" ; Can be copied from UIASpy Includes folder
#include "C:\install\UI Automation\Spy tool\Includes\UIA_Functions.au3" ; Can be copied from UIASpy Includes folder
;#include "C:\UI Automation\Spy tool\Includes\CUIAutomation2-a.au3"
;#include "C:\UI Automation\Spy tool\Includes\UIA_Functions.au3"
;#include "UIA_SafeArray.au3" ; Can be copied from UIASpy Includes folder
;#include "UIA_Variant.au3" ; Can be copied from UIASpy Includes folder
;MsgBox($MB_SYSTEMMODAL, "Title", "This message box will timeout after 10 seconds or select the OK button.", 10)


#RequireAdmin

computer()
sleep(500)
Framework()
sleep(500)
AutoShutdown()
sleep(500)
close_power()
sleep(500)
move_shortcut()
sleep(500)
;~ move_anydesk()
photoviewer()
sleep(500)
host_search()
sleep(500)
close_update()
sleep(500)
chrome_install()
sleep(1000)
preset_program()
sleep(1000)
input_method()
sleep(1000)
live_tiles()



func computer() ;桌面顯示本機
	Run(@ComSpec & " /c " & 'C:\install\Cmd\Computer', "", @SW_HIDE)
	WinWaitActive("桌面圖示設定")
	;Send("!m") ;本機消失
	sleep(500)
	Send("!r") ;垃圾桶消失
	Sleep(500)
	Send("{enter}")
	ConsoleWrite( "本機 ok" & @CRLF )
EndFunc

func live_tiles()
	Run(@ComSpec & " /c " & 'C:\install\Reg\LockedStartLayout.reg', "", @SW_HIDE)
	WinWaitActive("登錄編輯程式")
	Local $hWnd = WinWait("登錄編輯程式", "", 5) ; Wait 5 seconds for the installed window to appear.
	ControlClick($hWnd, "是", "[CLASS:Button; INSTANCE:1]")
	WinWaitActive("登錄編輯程式","確定")
	Local $hWnd = WinWait("登錄編輯程式", "", 5) ; Wait 5 seconds for the installed window to appear.
	ControlClick($hWnd, "確定", "[CLASS:Button; INSTANCE:1]")
	ConsoleWrite( "動態磚 OK" & @CRLF )
EndFunc

func move_anydesk()
	Run(@ComSpec & " /c " & 'C:\install\Cmd\MoveAnydesk.cmd', "", @SW_HIDE)
EndFunc

func move_shortcut()
	Run(@ComSpec & " /c " & 'C:\install\Cmd\Move_ShortCut.cmd', "", @SW_HIDE)
EndFunc

func close_power()
	Run(@ComSpec & " /c " & 'C:\install\Cmd\ClosePower.bat', "", @SW_HIDE)
EndFunc

func host_search()
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
	Send("112.121.96.30  www.books.com.tw{enter}")
	Send("52.84.248.61   im2.book.com.tw{enter}")
	Send("211.72.248.247  cdn.kingstone.com.tw{enter}")
	Send("113.196.56.91     www.sanmin.com.tw{enter}")
	Send("54.230.215.53  static.findbook.tw{enter}")
	Send("113.196.250.30  addons.books.com.tw{enter}")
	Send("175.41.12.68  image.anobii.com{enter}")
	Send("194.244.14.33  beta.anobii.com{enter}")
	Send("61.220.250.181  www.koobe.com.tw{enter}")
	Send("14.0.52.66   static.anobii.com{enter}")
	Send("54.230.214.207  ecx.images-amazon.comm{enter}")
	Send("173.192.67.85  fb1.anobii.com{enter}")
	Send("199.184.255.169  contentcafe2.btol.com{enter}")
	Send("113.196.250.30  im1.book.com.tw{enter}")
	Send("168.63.131.14  img.kingstone.com.tw{enter}")
	Send("210.61.96.91 hylib.typl.gov.tw{enter}")
	Send("20.189.79.72 time.windows.com{enter}")
	Send("132.163.97.2 time.nist.gov{enter}")
	Send("210.241.76.206 webpac.typl.gov.tw{enter}")
	sleep(1000)
	Send("#{space}")
	Send("^s")
	sleep(1000)
	Send("!{F4}")
	ConsoleWrite( "host OK" & @CRLF )
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
	Send("C:\install{enter}")
	sleep(1000)
	Send("!o")
	WinWaitActive("建立工作")
	Send("{enter}")
	sleep(1000)
	Send("!{F4}")
	ConsoleWrite( "AutoShutdown OK" & @CRLF )
EndFunc

func input_method()  ;輸入法設定
	Send("#i") ;開設定
	WinWaitActive("設定")
	sleep(2000)
	Send("#{UP}")
	sleep(2000)
	Send("語言") ;點應用程式
    Sleep(1000)
	Send("{enter}")
    sleep(2000)
	Send("{enter}") ; 點選編輯語言與鍵盤選項
	sleep(2000)
	Send("{enter}") ;等待中文
	sleep(3000)
	Send("{tab}{enter}") ;進入選項
	sleep(3000)
	For $i=1 To 3
		Send("{tab}")
		sleep(1000)
	Next
	For $i=1 To 8
		Send("{enter}")
		sleep(1000)
	Next
	Sleep(1000)
	Send("!{F4}")
	ConsoleWrite( "Input Method OK" & @CRLF )
EndFunc


func framework()
	Run("C:\install\dotnetfx35.exe")
	WinWaitActive("Windows 功能")
	ControlClick("Windows 功能" , "下載並安裝此功能" , "[CLASS:Button; INSTANCE:4]")
	WinWaitActive("Windows 功能" , "您可能需要重新啟動")
	ControlClick("Windows 功能" , "關閉" , "[CLASS:Button; INSTANCE:2]")
	ConsoleWrite( "framework OK" & @CRLF )
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
	Send("172.20.1.140 odiswebedit.tycg.gov.tw{enter}")
	Send("20.189.79.72 time.windows.com{enter}")
	Send("132.163.97.2 time.nist.gov{enter}")
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
	WinWaitActive("那麼，下一步該怎麼做?")
	Send("!{F4}")
	Send("!{F4}")
	ConsoleWrite( "更新關閉 OK" & @CRLF )

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


func reboot()
	Shutdown(2)
EndFunc


func chrome_install()
	Run("C:\install\ChromeSetup")
	local $num = 0
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
		$num = $num + 1
		ConsoleWrite( $num & @CRLF )
		Sleep(100)
	WEnd

	;WinWaitActive("歡迎使用 Chrome - Google Chrome")
	;Send("!{F4}")
	;WinWaitActive("新分頁 - Google Chrome")
	;WinClose("新分頁 - Google Chrome")
	ConsoleWrite( "Chrome OK" & @CRLF )
EndFunc






