#Requires AutoHotkey v2.0

;To do:
;[Ref1] - using regex here is probably not necessary?
;[Ref2] - header is not used anywhere, do i keep it?
;[Ref3] - is basically double checking the condition of while loop, but I like to initialize the current position at the start of the loop. Is not worthy of rewriting
;[Ref4] - dirty way of refreshing GUI

PrototypeExpiration := "20250723" ;this is to prompt me, that if I'm still using this tool, I should consider improving it.
If DateDiff(PrototypeExpiration, A_Now, "days") < 0 {
	MsgBox "Prototype validity expired"
	ExitApp
}

global numberOfLinks
numberOfLinks := 0
global CurrentLink
CurrentLink := 0

global striptRunning
scriptRunning := false
global acquiringLinks
acquiringLinks := false

global url
url := "not set"

GuiCreate

GrabWebsiteTestfile(filepath) { ;this function is for testing stuff and is currently unused during runtime
	HtmlText := FileRead(filepath)
	LenText := StrLen(HtmlText)
	StartingPoint := InStr(HtmlText, "divProductDetailSpecification", 1) ;Makes the following Regex less intensive
	RegexPos := RegExMatch(HtmlText, "<ul data-vendor-name=`"[^`"]+`" ", &SubPat, StartingPoint) ;Start of table
	EndOfString := InStr(HtmlText, "</ul>",, RegexPos)
	HtmlText := SubStr(HtmlText,RegexPos,EndOfString-RegexPos)

	StartingPoint := InStr(HtmlText, "<li",1)
	Header := SubStr(HtmlText, 1, StartingPoint-1)
	Table := SubStr(HtmlText, StartingPoint)
	position := 1
	Str := ""
	While position > 0 {
		position := RegexMatch(Table,"label_header`">([^<>]+)<", &SubPatN,position)
		if position = 0
			break
		Str .= String(SubPatN.1) . ">"
		nextheader := RegexMatch(Table,"label_header`">([^<>]+)<",,position+1)
		if nextheader = 0
			nextheader := StrLen(Table)
		Substring := SubStr(Table, position+SubPatN.Len(0), nextheader-(position+SubPatN.Len(0)))
		subposition := 1
		while subposition > 0 {
			subposition := RegexMatch(Substring,">([^<>]+)<", &SubPatN,subposition)
			if subposition = 0
				break
			Str .= String(SubPatN.1) . ">"
			subposition := subposition + 1
		}
		Str := RTrim(Str, ">") ;remove the last thingie
		Str := Str . "<"
		position := nextheader - 1
	}
	Str := RTrim(Str, "<") ; remove the last pipe (#)
		return Str
}

GrabWebsiteData(hyperlink) {
	WebObj := ComObject("WinHttp.WinHttpRequest.5.1")
	WebObj.Open("GET", hyperlink)
	WebObj.Send()
	HtmlText := WebObj.ResponseText
	StartingPoint := InStr(HtmlText, "divProductDetailSpecification", 1) ;Defines start of table, makes following regex less intensive
	RegexPos := RegExMatch(HtmlText, "<ul data-vendor-name=`"[^`"]+`" ", &SubPat, StartingPoint) ;Finds the start of table, could be made simpler [Ref1]
	EndOfString := InStr(HtmlText, "</ul>",, RegexPos)
	HtmlText := SubStr(HtmlText,RegexPos,EndOfString-RegexPos)

	StartingPoint := InStr(HtmlText, "<li",1)
	Header := SubStr(HtmlText, 1, StartingPoint-1) ;[Ref2]
	Table := SubStr(HtmlText, StartingPoint)
	position := 1
	Str := ""
	While position > 0 {
		position := RegexMatch(Table,"label_header`">([^<>]+)<", &SubPatN,position)
		if position = 0
			break	;[Ref3]
		foundHeader := String(SubPatN.1)
		Str .= foundHeader . ">" ;I'm using ">" and "<" to pack multiple results into a single string, because it's the one thing that you are guaranted not to find between html tags :D
		nextheader := RegexMatch(Table,"label_header`">([^<>]+)<",,position+1)
		if nextheader = 0
			nextheader := StrLen(Table) ;next header not found, sets the thing to run until the next of time
		Substring := SubStr(Table, position+SubPatN.Len(0), nextheader-(position+SubPatN.Len(0)))
		subposition := 1

		while subposition > 0 {
			subposition := RegexMatch(Substring,">([^<>]+)<", &SubPatN,subposition)
			if subposition = 0
				break
			Str .= String(SubPatN.1) . ">"
			subposition := subposition + 1
		}
		Str := RTrim(Str, ">") ;remove the last thingie
		Str := Str . "<"
		if (foundHeader = "Price") {
			subposition := 1
			pattern := 'a href="(http[^"]+)"'
			subposition := RegexMatch(Substring,pattern, &SubPatN,subposition)
			if (subposition > 0) {
				Str .= "Direct link>" . String(SubPatN.1) . "<"
			}
		}	
		position := nextheader - 1
	}
	Str := RTrim(Str, "<") ; remove the last pipe (#)
		return Str
}

*F1::
{

	if (url = "not set") {
		MsgBox "Url not set"
		Exit
	}
	global scriptRunning
	scriptRunning := true
	global acquiringLinks
	acquiringLinks := true

	 GuiDestroy ;[Ref4]
	 GuiCreate

	WebObj := ComObject("WinHttp.WinHttpRequest.5.1")
	WebObj.Open("GET", url)
	WebObj.Send()
	teststring := WebObj.ResponseText
	position := 1
	links := ""
	While position > 0 {
		position := InStr(teststring, "<div class=`"title`">",,position)
		if position = 0
			break
		position := RegexMatch(teststring,"a href=`"([^`"]+)`"", &SubPat,position)
		links .= "https://www.biocompare.com" . SubPat.1 . "<" 
		}
		links := RTrim(links, "<")
		MsgBox links
		acquiringLinks := false

		GuiDestroy
		GuiCreate
		 
		linkArray := StrSplit(links, "<")
		global numberOfLinks
		numberOfLinks := linkArray.Length
		global CurrentLink
		CurrentLink := 0
	For hyperlink in linkArray {
		sleep 5000 ;Please don't remove this. It makes the script slow to avoid DDOSing the biocompare website
		result := "Link>" . hyperlink . "<" .  GrabWebsiteData(hyperlink) . "`r`n"
		FileAppend result, "linkScrapperOutput.txt" ;also creates the file in the directory where the script runs
		CurrentLink++
		GuiDestroy
		GuiCreate
	}
	MsgBox "Finished"
	scriptRunning := false
	GuiDestroy
	GuiCreate
}

*F2:: ;set url
{
 Global url
 url := A_Clipboard
  if (inStr(url, "biocompare") = 0) {
	MsgBox "This works only for biocompare website"
	url := "not set"
	Exit
 }
 
 if (inStr(url, "vcmpv=true") = 0) {
	MsgBox "You need to swtich to product view. I will fix it for you."
	url .= "&vcmpv=true"
 }

 GuiDestroy
 GuiCreate
}

;-----GUI CODE-----
GuiCreate() {
	global MyGui
	MyGui := Gui("-Caption +AlwaysOnTop +Owner +LastFound")
	MyGui.SetFont("s14 w500 q4", "Times New Roman")
	if (scriptRunning = true and acquiringLinks = true) {
	MyGui.Add("Text", , "Acquiring links....")
	}
	if (scriptRunning = false and acquiringLinks = false) {
		MyGui.Add("Text", , "Esc - exit script")
		MyGui.Add("Text", , "F1 - start scraping the link")
		MyGui.Add("Text", , "F2 - load link from clipboard")
		if (url = "not set") {
			MyGui.SetFont("cRed s12 w500 q4", "Times New Roman")
		} else {
			MyGui.SetFont("cGreen s12 w500 q4", "Times New Roman")
		}
		MyGui.Add("Text", , "url:" url)
		MyGui.Show("x1000 y40 NoActivate")
	}
	else {
	MyGui.Add("Text", , CurrentLink . " / " . numberOfLinks)
	MyGui.Show("x1690 y40 NoActivate")
	}
}
GuiDestroy() {
	MyGui.Destroy()
}
;----END GUI-----

Esc::ExitApp