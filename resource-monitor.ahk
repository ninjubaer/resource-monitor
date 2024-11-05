/*************************************************************************
 * @author ninju | .ninju.
* @file ninjusResourceMonitor.ahk
* @date 03/15/2024
* @version 0.0.1
* @description shows cpu and ram usage, toggle alwaysOnTop with little white circle in captions
*************************************************************************/
#Requires AutoHotkey v2.1-alpha.14+
#SingleInstance Force
#MaxThreads 255
if (!pToken:=Gdip_Startup())
    msgbox "error"
OnMessage(0x0020, (*)=>1)
hCursorArrow := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32512, "Ptr")
hCursorLoading := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32514, "Ptr")
DllCall("SetCursor", "Ptr", hCursorArrow)
captionCol      := "0xFF181F2C"
backCol         := "0xFF1C2433"
cpuCol          := "0xfff35c65"
ramCol          := "0xFF34A9C3"
white           := "0xFFFFFFFF"
successCol      := "0xFF23CD85"
w:=140,h:=240,aot:=false
recentValues := Map("cpu",[],"ram",[])
main := Gui("+OwnDialogs +E0x80000 -Caption", "Ninju's Stat Monitor v2")
main.show()
main.add("text","x" w-10 " y2 w6 h6").OnEvent("Click", (*) => ExitApp())
main.add("Text","x" w-20 " y2 w6 h6").OnEvent("Click", tAot)
main.Add("text", "x10 y170 w120 h20").OnEvent("Click", (*)=>(DllCall("SetCursor", "ptr",hCursorLoading),mem:=GlobalMemoryStatusEx(1), FreeMemory(), updateGui(), SetTimer((*)=>MsgBox('Freed ' GlobalMemoryStatusEx(1) - mem "mb"),-5000), DllCall("SetCursor", "ptr", hCursorArrow)))
main.Add("text", "x10 y195 w120 h20").OnEvent("Click", (*)=>(DllCall("SetCursor", "ptr", hCursorLoading), am:=RemoveTempFiles(),updateGui(), DllCall("SetCursor", "ptr", hCursorArrow)))
main.add("text","x0 y0 w" w " h" h).OnEvent("Click", (*) => PostMessage(0xA1,2))
; ty to xspx
hbm := CreateDIBSection(w, h)
hdc := CreateCompatibleDC()
obm := SelectObject(hdc, hbm)
G := Gdip_GraphicsFromHDC(hdc)
Gdip_SetSmoothingMode(G, 2)
Gdip_SetInterpolationMode(G, 2)
UpdateLayeredWindow(main.Hwnd, hdc)
CPUSystemTimes() ;init
loop
    sleep(1000),updateGui(CPUSystemTimes(),GlobalMemoryStatusEx())
CPUSystemTimes(){
    static pIdleTime:=0, pKernelTime:=0, pUserTime:=0
    local s, pIdleTime2,pKernelTime2,pUserTime2
    if !(pIdleTime)
        return (DllCall("GetSystemTimes", "int64*", &pIdleTime , "int64*", &pKernelTime, "int64*", &pUserTime))
    DllCall("GetSystemTimes", "int64*", &pIdleTime2:=0 , "int64*", &pKernelTime2:=0, "int64*", &pUserTime2:=0)
    return ((s:=pKernelTime-pKernelTime2 + pUserTime - pUserTime2) - pIdleTime+pIdleTime2 ) * 100//s * !!(pIdleTime:=pIdleTime2,pKernelTime:=pKernelTime2,pUserTime:=pUserTime2)
}
GlobalMemoryStatusEx(returnMB?) {
    static _MEMORYSTATUSEX := Buffer(64, 0), _:=(NumPut("uint", 64, _MEMORYSTATUSEX))
    if (DllCall("Kernel32.dll\GlobalMemoryStatusEx", "Ptr", _MEMORYSTATUSEX)) {
		if returnMB ?? 0
			return NumGet(_MEMORYSTATUSEX, 16, "UInt64") // 1024**2
        return NumGet(_MEMORYSTATUSEX, 4, "uint")
	}
	return 0
}
FreeMemory() {
	for obj in ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process") {
		try {
			hProcess := DllCall("OpenProcess", "uint", 0x1F0FFF, "int", 0, "uint", obj.ProcessId, "ptr")
			DllCall("SetProcessWorkingSetSize", "ptr", hProcess, "int", -1, "int", -1, "int")
			DllCall("psapi\EmptyWorkingSet", "ptr", hProcess)
			DllCall("CloseHandle", "ptr", hProcess)
		}
	}
	return DllCall("psapi\EmptyWorkingSet", "ptr", -1)
}

/************************************************************************
 * @description SHFILEOP
 * @file SHFILEOP.ahk
 * @author ninju
 * @date 2024/05/27
 * @version 0.0.1
 ***********************************************************************/
Class SHFILEOP {
	static operations := { move: 0x1, copy: 0x2, delete: 0x3, rename: 0x4 }
	static fFlags := {
		FOF_MULTIDESTFILES: 0x1,
		FOF_CONFIRMMOUSE: 0x2,
		FOF_SILENT: 0x4,
		FOF_RENAMEONCOLLISION: 0x8,
		FOF_NOCONFIRMATION: 0x10,
		FOF_WANTMAPPINGHANDLE: 0x20,
		FOF_ALLOWUNDO: 0x40,
		FOF_FILESONLY: 0x80,
		FOF_SIMPLEPROGRESS: 0x100,
		FOF_NOCONFIRMMKDIR: 0x200,
		FOF_NOERRORUI: 0x400,
		FOF_NOCOPYSECURITYATTRIBS: 0x800,
		FOF_NORECURSION: 0x1000,
		FOF_NO_CONNECTED_ELEMENTS: 0x2000,
		FOF_WANTNUKEWARNING: 0x4000,
		FOF_NORECURSEREPARSE: 0x8000
	}
	static return := {
		2: "File not found",
		3: "Path not found",
		5: "Access denied",
		10: "Invalid handle",
		11: "Invalid parameter",
		12: "Disk full",
		15: "Invalid drive",
		16: "Sharing violation",
		17: "File exists",
		18: "Cannot create file",
		32: "Sharing buffer overflow",
		53: "Network path not found",
		67: "Network name not found",
		80: "File already exists",
		87: "Invalid parameter",
		1026: "Directory not empty",
		1392: "File or directory already exists",
		161: "Bad path name",
		206: "Path too long",
		995: "Operation cancelled"
	}
	Class _SHFILEOPSTRUCT {
		hwnd: uptr
		wFunc: u32
		pFrom: iptr
		pTo: iptr
		fFlags: u32
		fAnyOperationsAborted: i8
		hNameMappings: uptr
		lpszProgressTitle: iptr
	}
	static move(sourceDir, destDir, flags := this.fFlags.FOF_SILENT | this.fFlags.FOF_NOCONFIRMATION | this.fFlags.FOF_NOERRORUI, &errormsg:=false) {
		SHFILEOPSTRUCT := this._SHFILEOPSTRUCT()
		SHFILEOPSTRUCT.wFunc := this.operations.move
		SHFILEOPSTRUCT.pFrom := StrPtr(sourceDir)
		SHFILEOPSTRUCT.pTo := StrPtr(destDir)
		SHFILEOPSTRUCT.fFlags := flags
		r:= DllCall("shell32\SHFileOperationW", "Ptr", SHFILEOPSTRUCT)
		if this.return.HasProp(r)
			errormsg := this.return.%r%
		return r
	}
	static copy(sourceDir, destDir, flags := this.fFlags.FOF_SILENT | this.fFlags.FOF_NOCONFIRMATION | this.fFlags.FOF_NOERRORUI, &errormsg:=false) {
		SHFILEOPSTRUCT := this._SHFILEOPSTRUCT()
		SHFILEOPSTRUCT.wFunc := this.operations.copy
		SHFILEOPSTRUCT.pFrom := StrPtr(sourceDir)
		SHFILEOPSTRUCT.pTo := StrPtr(destDir)
		SHFILEOPSTRUCT.fFlags := flags
		r:= DllCall("shell32\SHFileOperationW", "Ptr", SHFILEOPSTRUCT)
		if this.return.HasProp(r)
			errormsg := this.return.%r%
		return r
	}
	static delete(sourceDir, flags := this.fFlags.FOF_SILENT | this.fFlags.FOF_NOCONFIRMATION | this.fFlags.FOF_NOERRORUI, &errormsg:=false) {
		SHFILEOPSTRUCT := this._SHFILEOPSTRUCT()
		SHFILEOPSTRUCT.wFunc := this.operations.delete
		SHFILEOPSTRUCT.pFrom := StrPtr(sourceDir)
		SHFILEOPSTRUCT.fFlags := flags
		r:= DllCall("shell32\SHFileOperationW", "Ptr", SHFILEOPSTRUCT)
		if this.return.HasProp(r)
			errormsg := this.return.%r%
		return r
	}
	static rename(sourceDir, destDir, flags := this.fFlags.FOF_SILENT | this.fFlags.FOF_NOCONFIRMATION | this.fFlags.FOF_NOERRORUI, &errormsg:=false) {
		SHFILEOPSTRUCT := this._SHFILEOPSTRUCT()
		SHFILEOPSTRUCT.wFunc := this.operations.rename
		SHFILEOPSTRUCT.pFrom := StrPtr(sourceDir)
		SHFILEOPSTRUCT.pTo := StrPtr(destDir)
		SHFILEOPSTRUCT.fFlags := flags
		r:= DllCall("shell32\SHFileOperationW", "Ptr", SHFILEOPSTRUCT)
		if this.return.HasProp(r)
			errormsg := this.return.%r%
		return r
	}

}

removeTempFiles() {
	loop files A_Temp "\*.*", "D"
		SHFILEOP.delete(A_LoopFileFullPath,SHFILEOP.fFlags.FOF_NOERRORUI |SHFILEOP.fFlags.FOF_NOCONFIRMATION)
}

tAOT(*){
    global
    aot := !aot
    WinSetAlwaysOnTop(aot,"ahk_id" main.hwnd)
    updateGui()
}
GetLastBootTime(){
	tot := A_TickCount//1000
	h:= Format("{:02}",tot//3600)
	m := Format("{:02}",(n:=mod(tot,3600))//60)
	s := Format("{:02}",mod(n,60))
	return h ":" m ":" s
}
updateGui(cpu?,ram?){
    global
    if (IsSet(cpu) && IsSet(ram)) {
    For k,v in recentValues
        if v.Length = 120
            v.RemoveAt(1)
    recentValues["cpu"].push(cpu)
    recentValues["ram"].push(ram)
    }
    Gdip_FillRectangle(G, pBrush := Gdip_BrushCreateSolid(captionCol),-1,-1,w+2,h+2), Gdip_DeleteBrush(pBrush)
    Gdip_FillRectangle(G,pBrush:=Gdip_BrushCreateSolid(backCol), -1,10,w+2,h-9),Gdip_DeleteBrush(pBrush)
    Gdip_FillEllipse(G,pBrush:=Gdip_BrushCreateSolid(cpuCol),w-10,2,6,6),Gdip_DeleteBrush(pBrush)
    Gdip_FillEllipse(G,pBrush:=Gdip_BrushCreateSolid(aot ? successCol : white),w-20,2,6,6),Gdip_DeleteBrush(pBrush)
    pPen := Gdip_CreatePen("0x22FFFFFF",0.5)
    pBrush := Gdip_BrushCreateSolid(captionCol)
    createGraphBox(10, 15)
    createGraphBox(10,70)
    createTextBox(10,120,'CPU: ' recentValues["cpu"][-1] "%",cpuCol)
    createTextBox(10,145,'RAM: ' recentValues["ram"][-1] "%",ramCol)
	createTextBox(10,170,'Free Memory', '0xFFFFFFFF')
	createTextBox(10, 195, 'Remove Temp', '0xFFFFFFFF')
    Gdip_DeletePen(pPen)
    pPenCpu := Gdip_CreatePen(cpuCol,1)
    pPenRam := Gdip_CreatePen(ramCol,1)
    For k,v in recentValues
        for i,j in v
            Gdip_DrawLine(G,pPen%k%,129-v.length +i, 5+ (k = "cpu" ? 1 : 2)*55, 129-v.length + i, 5+(k = "cpu" ? 1:2)*55 - 45/100*j)
    Gdip_DeletePen(pPenCpu),Gdip_DeletePen(pPenRam)
	Gdip_TextToGraphics(G,"Last Boot: " GetLastBootTime(),"x0 y" h-20 " Center vCenter c" (pBrush:=Gdip_BrushCreateSolid(white)),,w, 20)
    UpdateLayeredWindow(main.Hwnd,hdc,,,w,h)
}
createGraphBox(x,y){
    global
    Gdip_FillRectangle(G,pBrush, x,y,120,45)
    loop 7
        Gdip_DrawLine(G,pPen,x+(15*A_Index),y,x+15*A_Index,y+45)
    loop 2
        Gdip_DrawLine(G,pPen, x,y+A_Index*15,x+120,y+A_Index*15)
}
createTextBox(x,y,percent, col){
    Gdip_FillRectangle(G,pBrush := Gdip_BrushCreateSolid(captionCol),x,y,120,20), Gdip_DeleteBrush(pBrush)
    Gdip_TextToGraphics(G, percent,"x" x " y" y+2 " Center vCenter c" (pBrush:=Gdip_BrushCreateSolid(col)),,120,20), Gdip_DeleteBrush(pBrush)
}



;;;;;;;;;;;;;;;;;;;;;
; GDI LIBRARY
/**
 * @author builasz
 * @edited by xSPx
 * @url https://github.com/buliasz/AHKv2-Gdip
 */
Gdip_Startup()
{
	if (!DllCall("LoadLibrary", "str", "gdiplus", "UPtr")) {
		throw Error("Could not load GDI+ library")
	}

	si := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
	NumPut("UInt", 1, si)
	DllCall("gdiplus\GdiplusStartup", "UPtr*", &pToken:=0, "UPtr", si.Ptr, "UPtr", 0)
	if (!pToken) {
		throw Error("Gdiplus failed to start. Please ensure you have gdiplus on your system")
	}

	return pToken
}
GetDC(hwnd:=0) => DllCall("GetDC", "UPtr", hwnd)
Gdip_CreateHBITMAPFromBitmap(pBitmap, Background:=0xffffffff)
{
	DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "UPtr", pBitmap, "UPtr*", &hbm:=0, "Int", Background)
	return hbm
}
CreateDIBSection(w, h, hdc:="", bpp:=32, &ppvBits:=0,alt:=0)
{
	hdc2 := hdc ? hdc : GetDC()
	bi := Buffer(40, 0)

	NumPut("UInt", 40, "UInt", w, "UInt", h, "ushort", 1, "ushort", bpp, "UInt", 0, bi)

	hbm := DllCall("CreateDIBSection"
					, "UPtr", hdc2
					, "UPtr", bi.Ptr
					, "UInt", 0
					, "UPtr*", &ppvBits
					, "UPtr", 0
					, "UInt", 0, "UPtr")

	if (!hdc) {
		ReleaseDC(hdc2)
	}
	return (!alt ? hbm : bi)
}

ReleaseDC(hdc, hwnd:=0) => DllCall("ReleaseDC", "UPtr", hwnd, "UPtr", hdc)
Gdip_DisposeImage(pBitmap) => DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
CreateCompatibleDC(hdc:=0) => DllCall("CreateCompatibleDC", "UPtr", hdc)
SelectObject(hdc, hgdiobj) => DllCall("SelectObject", "UPtr", hdc, "UPtr", hgdiobj)
Gdip_GraphicsFromHDC(hdc)
{
	DllCall("gdiplus\GdipCreateFromHDC", "UPtr", hdc, "UPtr*", &pGraphics:=0)
	return pGraphics
}
Gdip_SetSmoothingMode(pGraphics, SmoothingMode) => DllCall("gdiplus\GdipSetSmoothingMode", "UPtr", pGraphics, "Int", SmoothingMode)
Gdip_SetInterpolationMode(pGraphics, InterpolationMode) => DllCall("gdiplus\GdipSetInterpolationMode", "UPtr", pGraphics, "Int", InterpolationMode)
Gdip_GraphicsClear(pGraphics, ARGB:=0x00ffffff) => DllCall("gdiplus\GdipGraphicsClear", "UPtr", pGraphics, "Int", ARGB)
Gdip_FillRectangle(pGraphics, pBrush, x, y, w, h)
{
	return DllCall("gdiplus\GdipFillRectangle"
					, "UPtr", pGraphics
					, "UPtr", pBrush
					, "Float", x
					, "Float", y
					, "Float", w
					, "Float", h)
}
Gdip_BrushCreateSolid(ARGB:=0xff000000)
{
	DllCall("gdiplus\GdipCreateSolidFill", "UInt", ARGB, "UPtr*", &pBrush:=0)
	return pBrush
}
Gdip_DeleteBrush(pBrush) => DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrush)
Gdip_CreatePen(ARGB, w)
{
	DllCall("gdiplus\GdipCreatePen1", "UInt", ARGB, "Float", w, "Int", 2, "UPtr*", &pPen:=0)
	return pPen
}
Gdip_DeletePen(pPen) => DllCall("gdiplus\GdipDeletePen", "UPtr", pPen)
Gdip_TextToGraphics(pGraphics, Text, Options, Font:="Arial", Width:="", Height:="", Measure:=0)
{
	IWidth := Width
	IHeight := Height
	PassBrush := 0
	Text := String(Text)


	pattern_opts := "i)"
	RegExMatch(Options, pattern_opts "X([\-\d\.]+)(p*)", &xpos:="")
	RegExMatch(Options, pattern_opts "Y([\-\d\.]+)(p*)", &ypos:="")
	RegExMatch(Options, pattern_opts "W([\-\d\.]+)(p*)", &Width:="")
	RegExMatch(Options, pattern_opts "H([\-\d\.]+)(p*)", &Height:="")
	RegExMatch(Options, pattern_opts "C(?!(entre|enter))([a-f\d]+)", &Colour:="")
	RegExMatch(Options, pattern_opts "Top|Up|Bottom|Down|vCentre|vCenter", &vPos:="")
	RegExMatch(Options, pattern_opts "NoWrap", &NoWrap:="")
	RegExMatch(Options, pattern_opts "R(\d)", &Rendering:="")
	RegExMatch(Options, pattern_opts "S(\d+)(p*)", &Size:="")

	if Colour && IsInteger(Colour[2]) && !Gdip_DeleteBrush(Gdip_CloneBrush(Colour[2])) {
		PassBrush := 1, pBrush := Colour[2]
	}

	if !(IWidth && IHeight) && ((xpos && xpos[2]) || (ypos && ypos[2]) || (Width && Width[2]) || (Height && Height[2]) || (Size && Size[2])) {
		return -1
	}

	Style := 0
	Styles := "Regular|Bold|Italic|BoldItalic|Underline|Strikeout"
	for eachStyle, valStyle in StrSplit( Styles, "|" ) {
		if RegExMatch(Options, "\b" valStyle)
			Style |= (valStyle != "StrikeOut") ? (A_Index-1) : 8
	}

	Align := 0
	Alignments := "Near|Left|Centre|Center|Far|Right"
	for eachAlignment, valAlignment in StrSplit( Alignments, "|" ) {
		if RegExMatch(Options, "\b" valAlignment) {
			Align |= A_Index*10//21	; 0|0|1|1|2|2
		}
	}

	xpos := (xpos && (xpos[1] != "")) ? xpos[2] ? IWidth*(xpos[1]/100) : xpos[1] : 0
	ypos := (ypos && (ypos[1] != "")) ? ypos[2] ? IHeight*(ypos[1]/100) : ypos[1] : 0
	Width := (Width && Width[1]) ? Width[2] ? IWidth*(Width[1]/100) : Width[1] : IWidth
	Height := (Height && Height[1]) ? Height[2] ? IHeight*(Height[1]/100) : Height[1] : IHeight

	if !PassBrush {
		Colour := "0x" (Colour && Colour[2] ? Colour[2] : "ff000000")
	}

	Rendering := (Rendering && (Rendering[1] >= 0) && (Rendering[1] <= 5)) ? Rendering[1] : 4
	Size := (Size && (Size[1] > 0)) ? Size[2] ? IHeight*(Size[1]/100) : Size[1] : 12

	hFamily := Gdip_FontFamilyCreate(Font)
	hFont := Gdip_FontCreate(hFamily, Size, Style)
	FormatStyle := NoWrap ? 0x4000 | 0x1000 : 0x4000
	hFormat := Gdip_StringFormatCreate(FormatStyle)
	pBrush := PassBrush ? pBrush : Gdip_BrushCreateSolid(Colour)

	if !(hFamily && hFont && hFormat && pBrush && pGraphics) {
		return !pGraphics ? -2 : !hFamily ? -3 : !hFont ? -4 : !hFormat ? -5 : !pBrush ? -6 : 0
	}

	CreateRectF(&RC:="", xpos, ypos, Width, Height)
	Gdip_SetStringFormatAlign(hFormat, Align)
	Gdip_SetTextRenderingHint(pGraphics, Rendering)
	ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, &RC)

	if vPos {
		ReturnRC := StrSplit(ReturnRC, "|")

		if (vPos[0] = "vCentre") || (vPos[0] = "vCenter")
			ypos += Floor(Height-ReturnRC[4])//2
		else if (vPos[0] = "Top") || (vPos[0] = "Up")
			ypos := 0
		else if (vPos[0] = "Bottom") || (vPos[0] = "Down")
			ypos := Height-ReturnRC[4]

		CreateRectF(&RC, xpos, ypos, Width, ReturnRC[4])
		ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, &RC)
	}

	if !Measure {
		Gdip_DrawString(pGraphics, Text, hFont, hFormat, pBrush, &RC)
	}

	if !PassBrush {
		Gdip_DeleteBrush(pBrush)
	}

	Gdip_DeleteStringFormat(hFormat)
	Gdip_DeleteFont(hFont)
	Gdip_DeleteFontFamily(hFamily)

	return ReturnRC
}
Gdip_SaveBitmapToFile(pBitmap, sOutput, Quality:=75)
{
	_p := 0

	SplitPath sOutput,,, &extension:=""
	if (!RegExMatch(extension, "^(?i:BMP|DIB|RLE|JPG|JPEG|JPE|JFIF|GIF|TIF|TIFF|PNG)$")) {
		return -1
	}
	extension := "." extension

	DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", &nCount:=0, "uint*", &nSize:=0)
	ci := Buffer(nSize)
	DllCall("gdiplus\GdipGetImageEncoders", "UInt", nCount, "UInt", nSize, "UPtr", ci.Ptr)
	if !(nCount && nSize) {
		return -2
	}

	loop nCount {
		address := NumGet(ci, (idx := (48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize, "UPtr")
		sString := StrGet(address, "UTF-16")
		if !InStr(sString, "*" extension)
			continue

		pCodec := ci.Ptr+idx
		break
	}

	if !pCodec {
		return -3
	}

	; from @iseahound ImagePut.select_codec
	if (quality ~= "^-?\d+$") and ("image/jpeg" = StrGet(NumGet(ci, idx+32+4*A_PtrSize, "ptr"), "UTF-16")) { ; MimeType
		; Use a separate buffer to store the quality as ValueTypeLong (4).
		v := Buffer(4), NumPut("uint", quality, v)

		; struct EncoderParameter - http://www.jose.it-berater.org/gdiplus/reference/structures/encoderparameter.htm
		; enum ValueType - https://docs.microsoft.com/en-us/dotnet/api/system.drawing.imaging.encoderparametervaluetype
		; clsid Image Encoder Constants - http://www.jose.it-berater.org/gdiplus/reference/constants/gdipimageencoderconstants.htm
		ep := Buffer(24+2*A_PtrSize)                  ; sizeof(EncoderParameter) = ptr + n*(28, 32)
			NumPut(  "uptr",     1, ep,            0)  ; Count
			DllCall("ole32\CLSIDFromString", "wstr", "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}", "ptr", ep.ptr+A_PtrSize, "hresult")
			NumPut(  "uint",     1, ep, 16+A_PtrSize)  ; Number of Values
			NumPut(  "uint",     4, ep, 20+A_PtrSize)  ; Type
			NumPut(   "ptr", v.ptr, ep, 24+A_PtrSize)  ; Value
	}

	_E := DllCall("gdiplus\GdipSaveImageToFile", "UPtr", pBitmap, "UPtr", StrPtr(sOutput), "UPtr", pCodec, "UInt", _p ? _p : 0)

	return _E ? -5 : 0
}
UpdateLayeredWindow(hwnd, hdc, x:="", y:="", w:="", h:="", Alpha:=255)
{
	if ((x != "") && (y != "")) {
		pt := Buffer(8)
		NumPut("UInt", x, "UInt", y, pt)
	}

	if (w = "") || (h = "") {
		WinGetRect(hwnd,,, &w, &h)
	}

	return DllCall("UpdateLayeredWindow"
		, "UPtr", hwnd
		, "UPtr", 0
		, "UPtr", ((x = "") && (y = "")) ? 0 : pt.Ptr
		, "Int64*", w|h<<32
		, "UPtr", hdc
		, "Int64*", 0
		, "UInt", 0
		, "UInt*", Alpha<<16|1<<24
		, "UInt", 2)
}
Gdip_CloneBrush(pBrush)
{
	DllCall("gdiplus\GdipCloneBrush", "UPtr", pBrush, "UPtr*", &pBrushClone:=0)
	return pBrushClone
}
Gdip_FontFamilyCreate(Font)
{
	DllCall("gdiplus\GdipCreateFontFamilyFromName"
					, "UPtr", StrPtr(Font)
					, "UInt", 0
					, "UPtr*", &hFamily:=0)

	return hFamily
}
Gdip_FontCreate(hFamily, Size, Style:=0)
{
	DllCall("gdiplus\GdipCreateFont", "UPtr", hFamily, "Float", Size, "Int", Style, "Int", 0, "UPtr*", &hFont:=0)
	return hFont
}
Gdip_StringFormatCreate(Format:=0, Lang:=0)
{
	DllCall("gdiplus\GdipCreateStringFormat", "Int", Format, "Int", Lang, "UPtr*", &hFormat:=0)
	return hFormat
}
CreateRectF(&RectF, x, y, w, h)
{
	RectF := Buffer(16)
	NumPut(
		"Float", x,
		"Float", y,
		"Float", w || 0,
		"Float", h || 0,
		RectF)
}
Gdip_DrawString(pGraphics, sString, hFont, hFormat, pBrush, &RectF) => DllCall("gdiplus\GdipDrawString", "UPtr", pGraphics, "UPtr", StrPtr(sString), "Int", -1, "UPtr", hFont, "UPtr", RectF.Ptr, "UPtr", hFormat, "UPtr", pBrush)
Gdip_DeleteStringFormat(hFormat) => DllCall("gdiplus\GdipDeleteStringFormat", "UPtr", hFormat)
Gdip_DeleteFont(hFont) => DllCall("gdiplus\GdipDeleteFont", "UPtr", hFont)
Gdip_DeleteFontFamily(hFamily) => DllCall("gdiplus\GdipDeleteFontFamily", "UPtr", hFamily)
Gdip_SetStringFormatAlign(hFormat, Align) => DllCall("gdiplus\GdipSetStringFormatAlign", "UPtr", hFormat, "Int", Align)
WinGetRect( hwnd, &x:="", &y:="", &w:="", &h:="" ) {
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	CreateRect(&winRect, 0, 0, 0, 0) ;is 16 on both 32 and 64
	;VarSetCapacity( winRect, 16, 0 )	; Alternative of above two lines
	DllCall( "GetWindowRect", "Ptr", hwnd, "Ptr", winRect )
	x := NumGet(winRect,  0, "UInt")
	y := NumGet(winRect,  4, "UInt")
	w := NumGet(winRect,  8, "UInt") - x
	h := NumGet(winRect, 12, "UInt") - y
}
CreateRect(&Rect, x, y, w, h)
{
	Rect := Buffer(16)
	NumPut("UInt", x, "UInt", y, "UInt", w, "UInt", h, Rect)
}
Gdip_SetTextRenderingHint(pGraphics, RenderingHint) => DllCall("gdiplus\GdipSetTextRenderingHint", "UPtr", pGraphics, "Int", RenderingHint)
Gdip_MeasureString(pGraphics, sString, hFont, hFormat, &RectF)
{
	RC := Buffer(16)
	DllCall("gdiplus\GdipMeasureString"
					, "UPtr", pGraphics
					, "UPtr", StrPtr(sString)
					, "Int", -1
					, "UPtr", hFont
					, "UPtr", RectF.Ptr
					, "UPtr", hFormat
					, "UPtr", RC.Ptr
					, "uint*", &Chars:=0
					, "uint*", &Lines:=0)

	return RC.Ptr ? NumGet(RC, 0, "Float") "|" NumGet(RC, 4, "Float") "|" NumGet(RC, 8, "Float") "|" NumGet(RC, 12, "Float") "|" Chars "|" Lines : 0
}
Gdip_DrawLine(pGraphics, pPen, x1, y1, x2, y2) =>DllCall("gdiplus\GdipDrawLine"
					, "UPtr", pGraphics
					, "UPtr", pPen
					, "Float", x1
					, "Float", y1
					, "Float", x2
					, "Float", y2)
Gdip_FillEllipse(pGraphics, pBrush, x, y, w, h) => DllCall("gdiplus\GdipFillEllipse", "UPtr", pGraphics, "UPtr", pBrush, "Float", x, "Float", y, "Float", w, "Float", h)
Gdip_BitmapFromHWND(hwnd)
{
	WinGetRect(hwnd,,, &Width, &Height)
	hbm := CreateDIBSection(Width, Height), hdc := CreateCompatibleDC(), obm := SelectObject(hdc, hbm)
	PrintWindow(hwnd, hdc)
	pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
	SelectObject(hdc, obm), DeleteObject(hbm), DeleteDC(hdc)
	return pBitmap
}
PrintWindow(hwnd, hdc, Flags:=0) => DllCall("PrintWindow", "UPtr", hwnd, "UPtr", hdc, "UInt", Flags)
Gdip_CreateBitmapFromHBITMAP(hBitmap, Palette:=0)
{
	DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "UPtr", hBitmap, "UPtr", Palette, "UPtr*", &pBitmap:=0)
	return pBitmap
}
DeleteObject(hObject) => DllCall("DeleteObject", "UPtr", hObject)
DeleteDC(hdc) => DllCall("DeleteDC", "UPtr", hdc)