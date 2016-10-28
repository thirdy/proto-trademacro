; https://github.com/thirdy/trademacro
; based on Pete's Price Macro https://github.com/trackpete/exiletools-price-macro/blob/master/poe_price_macro.ahk
; as well as POE Item Info https://github.com/aRTy42/POE-ItemInfo

; Trade Macro
; Version: 1.0.0 (2016/09/23)
global versionNr := "1.0.0"
global repoName := "thirdy"
global projectName := "trademacro"
; Tested with AutoHotkey112401_Install, Unicode 64-bit
; Windows 7
;
; Written by /u/ProFalseIdol on reddit, ManicCompression in game and /u/Eruyome87
;
; CONFIGURATION NOTE: You must configure the LeagueName properly. Otherwise it will default to
; "Standard" - press ^F and search for LeagueName and you will find where to set it.
;
; USAGE NOTE: This requires your Path of Exile to be in Windowed (or Windowed Full Screen) 
; to work properly, otherwise the popup will just show in the background and you won't
; see it. Also, you *must* use the AHK from http://ahkscript.org NOT NOT autohotkey.com!
;
; WINDOWS 10 NOTE: You may need to run this .ahk file as an Administrator in order for the popups
; to show properly.
;
; AUTHOR'S NOTE: I'm not an AHK programmer, I learned everything on by Google only. Certainly this code will look sloppy to experienced AHK programmers, if you have any
; advice or would like to re-write it, please feel free and let me know. 
;
; ===================================================
; Change Log
; 0.1 (2016/09/23): Re-write to use pure AHK, previous version used Java, project called "longan"
;

; Register a function to be called on exit:
OnExit("ExitFunc")

; == Startup Options ===========================================
#SingleInstance force
#NoEnv 
#Persistent ; Stay open in background
SendMode Input 
StringCaseSense, On ; Match strings with case.
Menu, tray, Tip, Path of Exile Trade Macro

; https://github.com/cocobelgica/AutoHotkey-JSON
#Include, lib/JSON.ahk

If (A_AhkVersion <= "1.1.22")
{
    msgbox, You need AutoHotkey v1.1.22 or later to run this script. `n`nPlease go to http://ahkscript.org/download and download a recent version.
    exit
}

; Windows system tray icon
IfExist, %A_ScriptDir%\icon.ico
    Menu, Tray, Icon, %A_ScriptDir%\icon.ico
else
{
    UrlDownloadToFile, https://raw.githubusercontent.com/thirdy/trademacro/master/icon.ico, %A_ScriptDir%\icon.ico
    Menu, Tray, Icon, %A_ScriptDir%\icon.ico
}

; == Variables and Options and Stuff ===========================
FileRemoveDir, %A_WorkingDir%\tempFiles, 1
FileCreateDir, %A_WorkingDir%\tempFiles

; *******************************************************************
; *******************************************************************
;                      SET LEAGUENAME IN CONFIG FILE!!
; *******************************************************************
; *******************************************************************
; League Name must be specified in config file, otherwise the search defaults to Standard League
global tempFilesDirectory = "tempFiles\"
global LeagueJSONFile := tempFilesDirectory . "leagues.json"
global Leagues := FunctionGETLeagues()
global iniFilePath := "config.ini"
global tempLeagueIsRunning := FunctionCheckIfTempLeagueIsRunning()
global defaultLeague := ( (tempLeagueIsRunning > 0) ? "tmpstandard" : "standard" )
global IniLeagueName := FunctionReadValueFromIni("SearchLeague", defaultLeague, "Search")
global LeagueName := Leagues[IniLeagueName]
global MouseMoveThreshold := FunctionReadValueFromIni("MouseMoveThreshold", 40)
global CacheExpireAge := FunctionReadValueFromIni("Expire", 0, "Cache")
global fontSize := FunctionReadValueFromIni("FontSize", "9", "Misc")
global Debug := FunctionReadValueFromIni("Debug", 0, "Debug")
global ShowItems := FunctionReadValueFromIni("ShowItems", "15", "Misc")
global ReadFromClipboardKey := FunctionReadValueFromIni("PriceCheckHotKey", "^q", "Hotkeys")
global CustomInputSearchKey := FunctionReadValueFromIni("CustomInputSearchHotKey", "^i", "Hotkeys")
global ConfigPredefSearchesKey := FunctionReadValueFromIni("ConfigPredefSearchesHotKey", "^o", "Hotkeys")
global PredefSearch01Key := FunctionReadValueFromIni("PredefinedSearch01HotKey", "F9", "Hotkeys")
global PredefSearch02Key := FunctionReadValueFromIni("PredefinedSearch02HotKey", "F10", "Hotkeys")
global PredefSearch03Key := FunctionReadValueFromIni("PredefinedSearch03HotKey", "F11", "Hotkeys")
global PredefSearch04Key := FunctionReadValueFromIni("PredefinedSearch04HotKey", "F12", "Hotkeys")
global RepeatPredefSearchModifier := FunctionReadValueFromIni("RepeatPredefinedSearchModifier", "^", "Hotkeys")
global PredefinedSearch01Url := FunctionReadValueFromIni("PredefinedSearch01Url", "", "Search")
global PredefinedSearch02Url := FunctionReadValueFromIni("PredefinedSearch02Url", "", "Search")
global PredefinedSearch03Url := FunctionReadValueFromIni("PredefinedSearch03Url", "", "Search")
FunctionGetLatestRelease()
; There are multiple hotkeys to run this script now, defaults set as follows:
; ^p (CTRL-p) - Sends the item information to my server, where a price check is performed. Levels and quality will be automatically processed.
; ^i (CTRL-i) - Pulls up an interactive search box that goes away after 30s or when you hit enter/ok
;
; To modify these, you will need to modify the config file
; see http://www.autohotkey.com/docs/Hotkeys.htm for hotkey options
AssignHotkey(ReadFromClipboardKey, "ReadFromClipboard")
AssignHotkey(CustomInputSearchKey, "CustomInputSearch")
AssignHotkey(ConfigPredefSearchesKey, "ConfigPredefSearches")
AssignHotkey(PredefSearch01Key, "PredefSearch01")
AssignHotkey(RepeatPredefSearchModifier . PredefSearch01Key, "RepeatPredefSearch01")
AssignHotkey(PredefSearch02Key, "PredefSearch02")
AssignHotkey(RepeatPredefSearchModifier . PredefSearch02Key, "RepeatPredefSearch02")
AssignHotkey(PredefSearch03Key, "PredefSearch03")
AssignHotkey(RepeatPredefSearchModifier . PredefSearch03Key, "RepeatPredefSearch03")
AssignHotkey(PredefSearch04Key, "PredefSearch04")
AssignHotkey(RepeatPredefSearchModifier . PredefSearch04Key, "RepeatPredefSearch04")

CoordMode, Mouse, Screen
CoordMode, ToolTip, Screen

Fonts.Init(fontSize, fontSize)

; Price check w/ auto filters
ReadFromClipboard:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
  FunctionReadItemFromClipboard()
}
return

; Custom Input String Search
CustomInputSearch:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
  Global X
  Global Y
  MouseGetPos, X, Y	
  InputBox,ItemName,Price Check,Item Name,,250,100,X-160,Y - 250,,30,
  if ItemName {
	

	PostData := "league=" . LeagueName . "&type=&base=&name=" . ItemName . "&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&q_min=&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="

    FunctionPostItemData(PostData)
  }
}
return

; == Function PriceCache ==================================
getPriceCacheAge(itemname)
{
	filepath := "cache\price\" . itemname
	if FileExist(filepath)
	{
		FileGetTime, modificationTime, %filepath%
		var1 = %modificationTime% ; replace file date variable with var1
		var2 = %A_Now% ; replace current time variable with var2
		EnvSub, var2, %var1%, Minutes
		if var2 < CacheExpireAge
			return var2
		else
			return -1
	}
	else
	{
		return -1
	}
}

saveCache(itemname, data)
{
	filepath := "cache\price\" . itemname
	if FileExist(filepath)
		FileDelete %filepath%
	FileAppend, %data%, %filepath%
}

dispCacheData(itemname, cacheAge)
{
	filepath := "cache\price\" . itemname
	FileRead, data, %filepath%
	FunctionShowToolTipPriceInfo(data . "`n(" . cacheAge . " min ago)")
}

; == Function Stuff =======================================

FunctionPostItemData(Payload)
{  
  temporaryContent = Submitting...
  FunctionShowToolTipPriceInfo(temporaryContent)
  
  ; Send the PostData and parse the response
  html := FunctionDoPostRequest(Payload)
  result := FunctionParseHtml(html, Payload)
  
  ;FileDelete, tempFiles\result.txt
  ;FileAppend, %result%, tempFiles\result.txt
  FunctionShowToolTipPriceInfo(result . "`n(0 min ago)")
  
  return result
}

; This is for the tooltip, so it shows it and starts a timer that watches mouse movement.
; I imagine there's a better way of doing this. The crazy long name is to avoid
; overlap with other scripts in case people try to combine these into one big script.

FunctionShowToolTipPriceInfo(responsecontent)
{
    ; Get position of mouse cursor
    Sleep, 2
	Global X
    Global Y
    MouseGetPos, X, Y	
    Global fontSize
    size := "s" . fontSize
	;ToolTipFont(size, "Lucida Console")
    ToolTip, %responsecontent%, X - 135, Y + 30
    Fonts.SetFixedFont()
    ToolTip, %responsecontent%, X - 135, Y + 30
    SetTimer, SubWatchCursorPrice, 100     

}

; == Utility Functions ==================

GetAhkExeFilename(Default_="AutoHotkey.exe")
{
    AhkExeFilename := Default_
    If (A_AhkPath)
    {
        StringSplit, AhkPathParts, A_AhkPath, \
        Loop, % AhkPathParts0
        {
            IfInString, AhkPathParts%A_Index%, .exe
            {
                AhkExeFilename := AhkPathParts%A_Index%
                Break
            }
        }
    }
    return AhkExeFilename
}

; == Tooltip Custom Font Function ==================
class Fonts {
    
    Init(FontSizeFixed, FontSizeUI) 
    {
        this.FontSizeFixed := FontSizeFixed
        this.FontSizeUI := FontSizeUI
        this.FixedFont := this.CreateFixedFont(FontSizeFixed)
        this.UIFont := this.CreateUIFont(FontSizeUI)
    }
    
    CreateFixedFont(FontSize_)
    {
        Options :=
        If (!(FontSize_ == "")) 
        {
            Options = s%FontSize_%
        }
        Gui Font, %Options%, Courier New
        Gui Font, %Options%, Consolas
        Gui Add, Text, HwndHidden, 
        SendMessage, 0x31,,,, ahk_id %Hidden%
        return ErrorLevel
    }

    CreateUIFont(FontSize_)
    {
        Options :=
        If (!(FontSize_ == "")) 
        {
            Options = s%FontSize_%
        }
        Gui Font, %Options%, Tahoma
        Gui Font, %Options%, Segoe UI
        Gui Add, Text, HwndHidden, 
        SendMessage, 0x31,,,, ahk_id %Hidden%
        return ErrorLevel
    }
    
    Set(NewFont)
    {
        AhkExe := GetAhkExeFilename()
        SendMessage, 0x30, NewFont, 1,, ahk_class tooltips_class32 ahk_exe %AhkExe%
        ; Development versions of AHK
        SendMessage, 0x30, NewFont, 1,, ahk_class tooltips_class32 ahk_exe AutoHotkeyA32.exe
        SendMessage, 0x30, NewFont, 1,, ahk_class tooltips_class32 ahk_exe AutoHotkeyU32.exe
        SendMessage, 0x30, NewFont, 1,, ahk_class tooltips_class32 ahk_exe AutoHotkeyU64.exe
    }
    
    SetFixedFont(FontSize_=-1)
    {
        If (FontSize_ == -1)
        {
            FontSize_ := this.FontSizeFixed
        }
        Else
        {
            this.FontSizeFixed := FontSize_
            this.FixedFont := this.CreateFixedFont(FontSize_)
        }
        this.Set(this.FixedFont)
    }

    SetUIFont(FontSize_=-1)
    {
        If (FontSize_ == -1)
        {
            FontSize_ := this.FontSizeUI
        }
        Else
        {
            this.FontSizeUI := FontSize_
            this.UIFont := this.CreateUIFont(FontSize_)
        }
        this.Set(this.UIFont)
    }
    
    GetFixedFont()
    {
        return this.FixedFont
    }
    
    GetUIFont()
    {
        return this.UIFont
    }
}

; == The Goods =====================================

; Watches the mouse cursor to get rid of the tooltip after too much movement
SubWatchCursorPrice:
  MouseGetPos, CurrX, CurrY
  MouseMoved := (CurrX - X)**2 + (CurrY - Y)**2 > MouseMoveThreshold**2
  If (MouseMoved)
  {
    SetTimer, SubWatchCursorPrice, Off
    ToolTip
  }
return


FunctionReadItemFromClipboard() {
  ; Only does anything if POE is the window with focus
  IfWinActive, Path of Exile ahk_class Direct3DWindowClass
  {
	; Send a ^C to copy the item information to the clipboard
	; Note: This will trigger any Item Info/etc. script that monitors the clipboard
	Send ^c
	; Wait 250ms - without this the item information doesn't get to the clipboard in time
	Sleep 250
	; Get what's on the clipboard
	ClipBoardData = %clipboard%
	; Split the clipboard data into strings to make sure it looks like a properly
	; formatted item, looking for the Rarity: tag in the first line. Just in case
	; something weird got copied to the clipboard.
	StringSplit, data, ClipBoardData, `n, `r
		
	; Strip out extra CR chars
	StringReplace RawItemData, ClipBoardData, `r, , A

	; If the first line on the clipboard has Rarity: it is probably some item
	; information from POE, so we'll send it to my server to process. Otherwise
	; we just don't do anything at all.
	IfInString, data1, Rarity:
	{
		; TODO, write a better code for all this
		
		ItemName := data2 . " " . data3
		IfInString, data3, ---
		{
			ItemName := data2
		}
		
		cacheAge := getPriceCacheAge(ItemName)
		if (CacheExpireAge != 0 && cacheAge != -1 && cacheAge < CacheExpireAge)
		{
			dispCacheData(ItemName, cacheAge)
		}
		else
		{
			; If item was linked from chat, there's this extra string we need to eliminate
			StringReplace ItemName, ItemName, <<set:MS>><<set:M>><<set:S>>, , A
			
			QualityParam := "q_min="
			IfInString, data1, Rarity: Gem
			{
				IfInString, RawItemData, Quality: 
				{
					QualityParam .= StrX(RawItemData, "Quality: +", 1,10, "%",1,1 )
				}
			}
		
			Payload := "league=" . LeagueName . "&type=&base=&name=" . ItemName . "&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&" . QualityParam . "&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="
			
			resdata := FunctionPostItemData(Payload)
			saveCache(ItemName, resdata)
		}
	}
  }  
}

StrPutVar(Str, ByRef Var, Enc = "")
{
	Len := StrPut(Str, Enc) * (Enc = "UTF-16" || Enc = "CP1200" ? 2 : 1)
	VarSetCapacity(Var, Len, 0)
	Return, StrPut(Str, &Var, Enc)
}

FunctionDoPostRequest(payload)
{
	;FileDelete, tempFiles\payload.txt
    ;FileAppend, %payload%, tempFiles\payload.txt
    
    ; TODO: split this function, HTTP POST and Html parsing should be separate
    ; Reference in making POST requests - http://stackoverflow.com/questions/158633/how-can-i-send-an-http-post-request-to-a-server-from-excel-using-vba
    HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ;HttpObj := ComObjCreate("MSXML2.ServerXMLHTTP") 
    ; We use this instead of WinHTTP to support gzip and deflate - http://microsoft.public.winhttp.narkive.com/NDkh5vEw/get-request-for-xml-gzip-file-winhttp-wont-uncompress-automagically
    HttpObj.Open("POST","http://poe.trade/search")
    HttpObj.SetRequestHeader("Host","poe.trade")
    HttpObj.SetRequestHeader("Connection","keep-alive")
    HttpObj.SetRequestHeader("Content-Length",StrLen(payload))
    HttpObj.SetRequestHeader("Cache-Control","max-age=0")
    HttpObj.SetRequestHeader("Origin","http://poe.trade")
    HttpObj.SetRequestHeader("Upgrade-Insecure-Requests","1")
    HttpObj.SetRequestHeader("User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36")
    HttpObj.SetRequestHeader("Content-type","application/x-www-form-urlencoded")
    HttpObj.SetRequestHeader("Accept","text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
    HttpObj.SetRequestHeader("Referer","http://poe.trade/")
    HttpObj.SetRequestHeader("Accept-Encoding","gzip;q=0,deflate;q=0") ; disables compression
    ;HttpObj.SetRequestHeader("Accept-Encoding","gzip, deflate")
    HttpObj.SetRequestHeader("Accept-Language","en-US,en;q=0.8")

    HttpObj.Send(payload)
    HttpObj.WaitForResponse()

    ;MsgBox % HttpObj.StatusText . HttpObj.GetAllResponseHeaders()
    ;MsgBox % HttpObj.ResponseText
    ; Dear GGG, it would be nice if you can provide an API like http://pathofexile.com/trade/search?name=Veil+of+the+night&links=4
    ; Pete's indexer is open sourced here - https://github.com/trackpete/exiletools-indexer you can use this to provide this api
    html := HttpObj.ResponseText
    ;FileRead, html, %tempFilesDirectory%Test1.txt
    ;FileDelete, %tempFilesDirectory%html.htm
    ;FileAppend, %html%, %tempFilesDirectory%html.htm
    
    Return, html
}

FunctionGetLatestRelease() {
    HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    url := "https://github.com/" . repoName . "/" . projectName . "/releases"
    tagsUrl = url . "/tags" 
    HttpObj.Open("GET",url)
    HttpObj.SetRequestHeader("Content-type","application/html")
    HttpObj.Send("")
    HttpObj.WaitForResponse()
    
    html := HttpObj.ResponseText
    tag := StrX( html,  "<span class=""tag-name",N,0,  "</span>", 1,0, N )
    RegExMatch(tag, "i)>(.*)<", match)
    tag := match1
    
    RegExMatch(tag, "(\d).(\d).(\d)(.*)", latestVersion)
    RegExMatch(versionNr, "(\d).(\d).(\d)(.*)", currentVersion)
    
    Loop, 3 {
        If (latestVersion%A_Index% > currentVersion%A_Index%) {
            Gui, UpdateAvailable: New
    
            Gui, Add, Text, cBlack w230, Current version: %versionNr%
            Gui, Add, Text, cGreen w230, New version available, grab it here:
            Gui, Add, Link, cBlue, <a href="%url%">(%projectName% %tag% on Github)</a>
            Gui, Show
            break
        }
    }
}

FunctionParseHtml(html, payload)
{
    ; Target HTML Looks like the ff:
    ;<tbody id="item-container-97" class="item" data-seller="Jobo" data-sellerid="458008" data-buyout="15 chaos" data-ign="Lolipop_Slave" data-league="Essence" data-name="Tabula Rasa Simple Robe" data-tab="This is a buff" data-x="10" data-y="9"> <tr class="first-line">
    ; TODO: grab more data like corruption found inside <tbody>
    
    ItemName := StrX( payload,  "&name=",  1,6, "&",1,1 )
    StringReplace, ItemName, ItemName, +, %A_SPACE%, All
    Quality  := StrX( payload,  "&q_min=", 1,7, "&",1,1 )
    
    ; TODO Refactor this
    IfInString, Quality, q_max
    {
        Quality = 
    }
    Text := ItemName " " Quality "`n ---------- `n"

    ; Text .= StrX( html,  "<tbody id=""item-container-0",          N,0, "<tr class=""first-line"">",1,28, N )

    NoOfItemsToShow := ShowItems
    While A_Index < NoOfItemsToShow
          Item        := StrX( html,  "<tbody id=""item-container-" . %A_Index%,  N,0,  "<tr class=""first-line"">", 1,23, N )
        , AccountName := StrX( Item,  "data-seller=""",                           1,13, """"  ,                      1,1,  T )
        , Buyout      := StrX( Item,  "data-buyout=""",                           T,13, """"  ,                      1,1,  T )
        , IGN         := StrX( Item,  "data-ign=""",                              T,10, """"  ,                      1,1     )
        ;, Text .= StrPad(IGN, 30) StrPad(AccountName, 30) StrPad(Buyout,30) "`n"
        ;, Text .= StrPad(IGN,20) StrPad(Buyout,20,"left") "`n"
        , Text .= StrPad(IGN,20) StrPad(Buyout,20,"left")"`n"
    
    Return, Text
}

; Taken from Poe-Item-Info
; Pads a string with a multiple of PadChar to become a wanted total length.
; Note that Side is the side that is padded not the anchored side.
; Meaning, if you pad right side, the text will move left. If Side was an 
; anchor instead, the text would move right if anchored right.
StrPad(String, Length, Side="right", PadChar=" ")
{
    StringLen, Len, String
    AddLen := Length-Len
    If (AddLen <= 0)
    {
        return String
    }
    Pad := StrMult(PadChar, AddLen)
    If (Side == "right")
    {
        Result := String . Pad
    }
    Else
    {
        Result := Pad . String
    }
    return Result
}
StrMult(Char, Times)
{
    Result =
    Loop, %Times%
    {
        Result := Result . Char
    }
    return Result
}

; ------------------------------------------------------------------------------------------------------------------ ;
; StrX function for parsing html, see simple example usage at https://gist.github.com/thirdy/9cac93ec7fd947971721c7bdde079f94
; ------------------------------------------------------------------------------------------------------------------ ;

; Cleanup StrX function and Google Example from https://autohotkey.com/board/topic/47368-strx-auto-parser-for-xml-html
; By SKAN

;1 ) H = HayStack. The "Source Text"
;2 ) BS = BeginStr. Pass a String that will result at the left extreme of Resultant String
;3 ) BO = BeginOffset. 
; Number of Characters to omit from the left extreme of "Source Text" while searching for BeginStr
; Pass a 0 to search in reverse ( from right-to-left ) in "Source Text"
; If you intend to call StrX() from a Loop, pass the same variable used as 8th Parameter, which will simplify the parsing process.
;4 ) BT = BeginTrim. 
; Number of characters to trim on the left extreme of Resultant String
; Pass the String length of BeginStr if you want to omit it from Resultant String
; Pass a Negative value if you want to expand the left extreme of Resultant String
;5 ) ES = EndStr. Pass a String that will result at the right extreme of Resultant String
;6 ) EO = EndOffset. 
; Can be only True or False. 
; If False, EndStr will be searched from the end of Source Text. 
; If True, search will be conducted from the search result offset of BeginStr or from offset 1 whichever is applicable.
;7 ) ET = EndTrim. 
; Number of characters to trim on the right extreme of Resultant String
; Pass the String length of EndStr if you want to omit it from Resultant String
; Pass a Negative value if you want to expand the right extreme of Resultant String
;8 ) NextOffset : A name of ByRef Variable that will be updated by StrX() with the current offset, You may pass the same variable as Parameter 3, to simplify data parsing in a loop

StrX(H,  BS="",BO=0,BT=1,   ES="",EO=0,ET=1,  ByRef N="" ) 
{ 
        Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
            <0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
            +(0-ET)):(X+P)):(X)))-P)
}
; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
; | by Skan | 19-Nov-2009

; From https://redd.it/53ml1x
;*RButton::Send {shift Down}{RButton down}
;*RButton Up::Send {shift Up}{RButton up}

PredefSearch01:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch01()
}
return

RepeatPredefSearch01:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch01(true)
}
return

FunctionPredefSearch01(reset = false) {
    static LineNumber = 3
    if (reset) 
        LineNumber = 3
        
    Payload := "league=" . LeagueName . "&type=Gem&base=&name=Added+Chaos+Damage&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&q_min=&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="
    
    filename := tempFilesDirectory . "PredefSearch01-result.txt"
    FunctionDoMacroSearch(Payload, LineNumber, filename, reset)
    LineNumber += 1
}

PredefSearch02:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch02()
}
return

RepeatPredefSearch02:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch02(true)
}
return

FunctionPredefSearch02(reset = false) {
    static LineNumber = 3
    if (reset) 
        LineNumber = 3
        
    Payload := "league=" . LeagueName . "&type=Jewel&base=&name=&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=%23%25+increased+maximum+Life&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&mod_name=%28pseudo%29+%28total%29+%2B%23%25+to+Cold+Resistance&mod_min=&mod_max=&mod_name=%28pseudo%29+%28total%29+%2B%23%25+to+Lightning+Resistance&mod_min=&mod_max=&mod_name=%28pseudo%29+%28total%29+%2B%23%25+to+Fire+Resistance&mod_min=&mod_max=&group_type=Count&group_min=2&group_max=&group_count=3&q_min=&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="
    
    filename := tempFilesDirectory . "PredefSearch02-result.txt"
    FunctionDoMacroSearch(Payload, LineNumber, filename, reset)
    LineNumber += 1
}

PredefSearch03:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch03()
}
return

RepeatPredefSearch03:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch03(true)
}
return

FunctionPredefSearch03(reset = false) {
    static LineNumber = 3
    if (reset) 
        LineNumber = 3
        
    Payload := "league=" . LeagueName . "&type=&base=&name=Tabula+Rasa+Simple+Robe&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&q_min=&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=0&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="
    
    filename := tempFilesDirectory . "PredefSearch03-result.txt"
    FunctionDoMacroSearch(Payload, LineNumber, filename, reset)
    LineNumber += 1
}

PredefSearch04:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch04()
}
return

RepeatPredefSearch04:
IfWinActive, Path of Exile ahk_class Direct3DWindowClass 
{
    FunctionPredefSearch04(true)
}
return

FunctionPredefSearch04(reset = false) {
    static ItemName = "Lifesprig Driftwood Wand"
    static LineNumber = 3
    if (reset) {
        LineNumber = 3
        InputBox,ItemName,Search By Item Name,Item Name,,250,100,X-160,Y - 250,,30,
    }   
        
    Payload := "league=" . LeagueName . "&type=&base=&name=" . ItemName . "&dmg_min=&dmg_max=&aps_min=&aps_max=&crit_min=&crit_max=&dps_min=&dps_max=&edps_min=&edps_max=&pdps_min=&pdps_max=&armour_min=&armour_max=&evasion_min=&evasion_max=&shield_min=&shield_max=&block_min=&block_max=&sockets_min=&sockets_max=&link_min=&link_max=&sockets_r=&sockets_g=&sockets_b=&sockets_w=&linked_r=&linked_g=&linked_b=&linked_w=&rlevel_min=&rlevel_max=&rstr_min=&rstr_max=&rdex_min=&rdex_max=&rint_min=&rint_max=&mod_name=&mod_min=&mod_max=&group_type=And&group_min=&group_max=&group_count=1&q_min=&q_max=&level_min=&level_max=&ilvl_min=&ilvl_max=&rarity=&seller=&thread=&identified=&corrupted=&online=x&buyout=x&altart=&capquality=x&buyout_min=&buyout_max=&buyout_currency=&crafted=&enchanted="
    
    filename := tempFilesDirectory . "PredefSearch04-result.txt"
    FunctionDoMacroSearch(Payload, LineNumber, filename, reset)
    LineNumber += 1
}

FunctionDoMacroSearch(Payload, LineNumber, filename, reset)
{   
    if (reset)
        FunctionDoFreshMacroSearch(Payload, filename)
    
    IfNotExist, %A_ScriptDir%\%filename%
        FunctionDoFreshMacroSearch(Payload, filename)
    
    FileReadLine, line, %A_ScriptDir%\%filename%, %LineNumber%
    FileReadLine, itemName, %A_ScriptDir%\%filename%, 1

    if (!line)
        result := "No more items found in " filename
    else {
        result := FunctionToWTB(itemName, line)
        clipboard = %result%
    }
    
    FunctionShowToolTipPriceInfo(result)
    
}

FunctionDoFreshMacroSearch(Payload, filename)
{
    FunctionShowToolTipPriceInfo("Running search...")
    html := FunctionDoPostRequest(Payload)
    result := FunctionParseHtml(html, Payload)
    FileDelete, %filename%
    FileAppend, %result%, %filename%
}

FunctionToWTB(itemName, line)
{
    RegExMatch(line, "([^\s]+)\s+(.+)", SubPat)
    ign    := SubPat1
    buyout := Trim(SubPat2)
    msg    := "@" ign " I would like to buy your " Trim(itemName) " listed for " buyout " in " LeagueName
    Return msg
}

; ------------------ GET LEAGUES ------------------ 
FunctionGETLeagues(){
    JSON := FunctionGetLeaguesJSON()    
    FileRead, JSONFile, %tempFilesDirectory%leagues.json  
    ; too dumb to parse the file to JSON Object, skipping this tstep
    ;parsedJSON 	:= JSON.Load(JSONFile)	
        
    ; Loop over league info and get league names    
    leagues := []
	Loop, Parse, JSONFile, `n, `r
	{					
        If RegExMatch(A_LoopField,"iOm)id *: *""(.*)""",leagueNames) {
            If (RegExMatch(leagueNames[1], "i)^Standard$")) {
                leagues["standard"] := leagueNames[1]
            }
            Else If (RegExMatch(leagueNames[1], "i)^Hardcore$")) {
                leagues["hardcore"] := leagueNames[1]
            }
            Else If InStr(leagueNames[1], "Hardcore", false) {
                leagues["tmphardcore"] := leagueNames[1]
            }
            Else {
                leagues["tmpstandard"] := leagueNames[1]
            }
        }        
	}
    
	Return leagues
}

FunctionGetLeaguesJSON(){
    HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    HttpObj.Open("GET","http://api.pathofexile.com/leagues?type=main&compact=1")
    HttpObj.SetRequestHeader("Content-type","application/json")
    HttpObj.Send("")
    HttpObj.WaitForResponse()
    
    ; Trying to format the string as JSON
    json := "{""results"":" . HttpObj.ResponseText . "}"
    json := RegExReplace(HttpObj.ResponseText, ",", ",`r`n`" A_Tab) 
    json := RegExReplace(json, "{", "{`r`n`" A_Tab)
    json := RegExReplace(json, "}", "`r`n`}")    
    json := RegExReplace(json, "},", A_Tab "},")
    json := RegExReplace(json, "\[", "[`r`n`" A_Tab)
    json := RegExReplace(json, "\]", "`r`n`]")
    json := RegExReplace(json, "m)}$", A_Tab "}")
    json := RegExReplace(json, """(.*)"":", A_Tab "$1 : ")
    
    ;MsgBox % json
    FileDelete, %LeagueJSONFile%
    FileAppend, %json%, %LeagueJSONFile%
    
    Return, json
}

; ------------------ CHECK IF A TEMP-LEAGUE IS ACTIVE ------------------ 
FunctionCheckIfTempLeagueIsRunning() {
    tempLeagueDates := FunctionGetTempLeagueDates()
    
    UTCTimestamp := GetTimestampUTC()
    UTCFormatStr := "yyyy-MM-dd'T'HH:mm:ss'Z'"
    FormatTime, TimeStr, %UTCTimestamp%, %UTCFormatStr%
    
    timeDiffStart := DateParse(TimeStr) - DateParse(tempLeagueDates["start"])
    timeDiffEnd := DateParse(TimeStr) - DateParse(tempLeagueDates["end"])
    
    If (timeDiffStart > 0 && timeDiffEnd < 0) {
        ; Current datetime is between temp league start and end date
        defaultLeague := "tmpstandard"
        Return 1
    }
    Else {
        defaultLeague := "standard"
        Return 0
    }
}

GetTimestampUTC() { ; http://msdn.microsoft.com/en-us/library/ms724390
   VarSetCapacity(ST, 16, 0) ; SYSTEMTIME structure
   DllCall("Kernel32.dll\GetSystemTime", "Ptr", &ST)
   Return NumGet(ST, 0, "UShort")                        ; year   : 4 digits until 10000
        . SubStr("0" . NumGet(ST,  2, "UShort"), -1)     ; month  : 2 digits forced
        . SubStr("0" . NumGet(ST,  6, "UShort"), -1)     ; day    : 2 digits forced
        . SubStr("0" . NumGet(ST,  8, "UShort"), -1)     ; hour   : 2 digits forced
        . SubStr("0" . NumGet(ST, 10, "UShort"), -1)     ; minute : 2 digits forced
        . SubStr("0" . NumGet(ST, 12, "UShort"), -1)     ; second : 2 digits forced
}

DateParse(str) {
    ; Parse ISO 8601 Formatted Date/Time to YYYYMMDDHH24MISS timestamp
    str := RegExReplace(str, "i)-|T|:|Z")
    Return str
}

FunctionGetTempLeagueDates(){
    JSON := FunctionGetLeaguesJSON()    
    FileRead, JSONFile, %tempFilesDirectory%leagues.json  
    ; too dumb to parse the file to JSON Object, skipping this step
    ;parsedJSON 	:= JSON.Load(JSONFile)	
     
    ; complicated way to find start and end dates of temp leagues since JSON.load is not working 
    foundStart := 
    foundEnd := 
    lastOpenBracket := 0
    lastCloseBracket := 0
    tempLeagueDates := []
    
	Loop, Parse, JSONFile, `n, `r
	{					
        If (InStr(A_LoopField, "{", false)) {
            lastOpenBracket := A_Index
        }
        Else If (InStr(A_LoopField, "}", false)) {
            lastCloseBracket := A_Index
        }        
        
        ; Find startAt and remember line number
        If RegExMatch(A_LoopField,"iOm)startAt *: *""(.*)""",dates) {
            If (StrLen(dates[1]) > 0)  {
                foundStart := A_index
                start := dates[1]
            }
        }            
        Else If RegExMatch(A_LoopField,"iOm)endAt *: *""(.*)""",dates) {
            If (!RegExMatch(dates[1], "i)null")) {
                foundEnd := A_Index
                end := dates[1]
            }       
        }
        
        If (foundStart > lastCloseBracket && foundEnd > lastCloseBracket) {
            tempLeagueDates["start"] := start
            tempLeagueDates["end"] := end
            Return tempLeagueDates
        }          
    }
}

; ------------------ ASSIGN HOTKEY AND HANDLE ERRORS ------------------ 
AssignHotkey(Key, Label){
    Hotkey, %Key%, %Label%, UseErrorLevel
    if (ErrorLevel)	{
		if (errorlevel = 1)
			str := str . "`nASCII '" . Key . "' - 1) The Label parameter specifies a nonexistent label name."
		else if (errorlevel = 2)
			str := str . "`nASCII '" . Key . "' - 2) The KeyName parameter specifies one or more keys that are either not recognized or not supported by the current keyboard layout/language."
		else if (errorlevel = 3)
			str := str . "`nASCII '" . Key . "' - 3) Unsupported prefix key. For example, using the mouse wheel as a prefix in a hotkey such as WheelDown & Enter is not supported."
		else if (errorlevel = 4)
			str := str . "`nASCII '" . Key . "' - 4) The KeyName parameter is not suitable for use with the AltTab or ShiftAltTab actions. A combination of two keys is required. For example: RControl & RShift::AltTab."
		else if (errorlevel = 5)
			str := str . "`nASCII '" . Key . "' - 5) The command attempted to modify a nonexistent hotkey."
		else if (errorlevel = 6)
			str := str . "`nASCII '" . Key . "' - 6) The command attempted to modify a nonexistent variant of an existing hotkey. To solve this, use Hotkey IfWin to set the criteria to match those of the hotkey to be modified."
		else if (errorlevel = 50)
			str := str . "`nASCII '" . Key . "' - 50) Windows 95/98/Me: The command completed successfully but the operating system refused to activate the hotkey. This is usually caused by the hotkey being "" ASCII " . int . " - in use"" by some other script or application (or the OS itself). This occurs only on Windows 95/98/Me because on other operating systems, the program will resort to the keyboard hook to override the refusal."
		else if (errorlevel = 51)
			str := str . "`nASCII '" . Key . "' - 51) Windows 95/98/Me: The command completed successfully but the hotkey is not supported on Windows 95/98/Me. For example, mouse hotkeys and prefix hotkeys such as a & b are not supported."
		else if (errorlevel = 98)
			str := str . "`nASCII '" . Key . "' - 98) Creating this hotkey would exceed the 1000-hotkey-per-script limit (however, each hotkey can have an unlimited number of variants, and there is no limit to the number of hotstrings)."
		else if (errorlevel = 99)
			str := str . "`nASCII '" . Key . "' - 99) Out of memory. This is very rare and usually happens only when the operating system has become unstable."
        
        MsgBox, %str%
	}
}

; ------------------ READ ALL OTHER INI VALUES ------------------ 
SubroutineReadIniValues:
	MouseMoveThreshold := FunctionReadValueFromIni("MouseMoveThreshold", 40)
    Debug := FunctionReadValueFromIni("Debug", 0, "Debug")
    ReadFromClipboardKey := FunctionReadValueFromIni("PriceCheckHotKey", "^q", "Hotkeys")
    CustomInputSearchKey := FunctionReadValueFromIni("CustomInputSearchHotKey", "^i", "Hotkeys")
    ConfigPredefSearchesKey := FunctionReadValueFromIni("ConfigPredefSearchesHotKey", "^o", "Hotkeys")
    RepeatPredefSearchModifier := FunctionReadValueFromIni("RepeatPredefinedSearchModifier", "^", "Hotkeys")
    ShowItems := FunctionReadValueFromIni("ShowItems", "15", "Misc")
    PredefSearch01Key := FunctionReadValueFromIni("PredefinedSearch01HotKey", "F9", "Hotkeys")
    PredefSearch02Key := FunctionReadValueFromIni("PredefinedSearch02HotKey", "F10", "Hotkeys")
    PredefSearch03Key := FunctionReadValueFromIni("PredefinedSearch03HotKey", "F11", "Hotkeys")
    PredefSearch04Key := FunctionReadValueFromIni("PredefinedSearch04HotKey", "F12", "Hotkeys")
    PredefinedSearch01Url := FunctionReadValueFromIni("PredefinedSearch01Url", "", "Search")
    PredefinedSearch02Url := FunctionReadValueFromIni("PredefinedSearch02Url", "", "Search")
    PredefinedSearch03Url := FunctionReadValueFromIni("PredefinedSearch03Url", "", "Search")
return

; ------------------ READ INI AND CHECK IF VARIABLES ARE SET ------------------ 
FunctionReadValueFromIni(IniKey, DefaultValue = "", Section = "Misc"){
	IniRead, OutputVar, %iniFilePath%, %Section%, %IniKey%
    
	If (!OutputVar | RegExMatch(OutputVar, "^ERROR$")) { 
		OutputVar := DefaultValue
        ; Somehow reading some ini-values is not working with IniRead
        ; Fallback for these cases via FileReadLine        
        Loop {
            FileReadLine, line, %iniFilePath%, %A_Index%
            If ErrorLevel
            break
            If InStr(line, IniKey, false) {
                RegExMatch(line, "= *(.*)", value)
                If (StrLen(value1) = 0) {
                    OutputVar := DefaultValue
                    
                }
                Else {
                    OutputVar := value1
                }                
            }
        }
    }
    
    
    
	Return OutputVar
}

; ------------------ WRITE TO INI ------------------
FunctionWriteValueToIni(IniKey,NewValue,IniSection){
	IniWrite, %NewValue%, %iniFilePath%, %IniSection%, %IniKey%
	If ErrorLevel
		s := "Writing to config failed."
	Else
		s := "Config updated."
	
	Gosub, SubroutineReadIniValues
}

; ------------------ Create and handle Gui Window to configure predefined searches ------------------
ConfigPredefSearches:
    GoSub, SubroutineCreatePredefGui
return

SubroutineCreatePredefGui:
    Gui, SearchConfig: New
    
    Gui, Add, Text, cGreen w230, Add poe.trade search URLs as predefined searches. Keep in mind to select the right league since these urls don't use the league specified in the config.ini file.
    Gui, Add, Link, cBlue, <a href="http://poe.trade">(visit poe.trade)</a>
    
    Gui, Add, Text, cBlack, Predefined Search Nr. 1 (%PredefSearch01Key%):
    Gui, Add, Edit, r1 vPredefinedSearch01Url w230, %PredefinedSearch01Url%
    Gui, Add, Text, cBlack, Predefined Search Nr. 2 (%PredefSearch02Key%):
    Gui, Add, Edit, r1 vPredefinedSearch02Url w230, %PredefinedSearch02Url%
    Gui, Add, Text, cBlack, Predefined Search Nr. 3 (%PredefSearch03Key%): 
    Gui, Add, Edit, r1 vPredefinedSearch03Url w230, %PredefinedSearch03Url%
    
    Gui, Add, Button, Default gSubroutineSavePredefSearches, &Save
    Gui, Add, Button, gSubroutineCancelPredefSearches, Cancel
    Gui, Show
return

SubroutineSavePredefSearches:
    Gui, Submit
    Gui, Destroy
    FunctionSavePredefSearchesToIni()
return

SubroutineCancelPredefSearches:
    Gui, Destroy
return

FunctionSavePredefSearchesToIni(){
    ; FunctionWriteValueToIni only writes the first value, no idea why. Therefore we do it wihtout it for now.
    IniWrite, %PredefinedSearch01Url%, %iniFilePath%, Search, PredefinedSearch01Url
    IniWrite, %PredefinedSearch02Url%, %iniFilePath%, Search, PredefinedSearch02Url
    IniWrite, %PredefinedSearch03Url%, %iniFilePath%, Search, PredefinedSearch03Url
    Gosub, SubroutineReadIniValues
}

; ------------------ Handle OnExit ------------------
ExitFunc(ExitReason, ExitCode)
{
    FileRemoveDir, %A_WorkingDir%\tempFiles, 1
    ExitApp
}