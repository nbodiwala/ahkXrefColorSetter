#SingleInstance Force
F11:: Reload

;Layer Colors
F1::
{

;Get Selection
sel := Explorer_GetSelection()

for item in sel
{
WinName := "Autodesk AutoCAD 2017 - [" item.name "]"
Run % item.path
WinWait, %WinName%
WinActivate, %WinName%
WinWaitActive, %WinName%
	
ACAD := ComObjActive("AutoCAD.Application") ; GetObject

f = 0
Need_Center = false
Need_Dash = false

Loop
{
	try
	{
		d := acad.activedocument.Fullname
		break
	}
	catch
	{
		Sleep 25
	}
}

;Set seach Folder
If InStr(d, "HSC Tower")
	StringGetPos, i, d, _xref, 1
else
	StringGetPos, i, d, 03_xref, 1

StringLeft, d, d, i
i = Layer Settings.ini

;set by layer and linetype scale
acad.activedocument.sendcommand("_ai_selall`n")
acad.activedocument.sendcommand("setbylayer`ny`ny`n")
acad.activedocument.sendcommand("lts`n75`n")


;If drawing doesn't have center and dashed line type, add them
For LineTypes in acad.activedocument.linetypes
	{
		If (linetypes.name = "Center")
			Need_Center = true
			
		If (linetypes.name = "Dashed")
			Need_Dash = true
	}
	If Need_Center = false
		acad.activedocument.linetypes.load("Center","acad.lin")
	If Need_Dash = false
		acad.activedocument.linetypes.load("Dashed","acad.lin")

;Find Layers.ini file in project folder
Loop, Files, %d%\%i%, R
{
	If InStr(A_LoopFileName, i)
	{
		d = %A_LoopFileFullPath%
		f = 1
		break
	}
}
	If f = 0
		Msgbox Layer Settings Not Found

IniRead, 11, %d%, 11			;what layers need to be set to color 11
IniRead, 15, %d%, 15			; what layers need to be set to color 15
IniRead, 25, %d%, 25			; what layers need to be set to color 25
IniRead, 35, %d%, 35			; what layers need to be set to color 35
IniRead, 45, %d%, 45			; what layers need to be set to color 45
IniRead, 55, %d%, 55			; what layers need to be set to color 55
IniRead, 65, %d%, 65			; what layers need to be set to color 65
IniRead, 75, %d%, 75			; what layers need to be set to color 75
IniRead, 85, %d%, 85			; what layers need to be set to color 85
IniRead, 95, %d%, 95			; what layers need to be set to color 95
IniRead, 105, %d%, 105			; what layers need to be set to color 105
IniRead, 115, %d%, 115			; what layers need to be set to color 115
IniRead, 125, %d%, 125			; what layers need to be set to color 125
IniRead, 135, %d%, 135			; what layers need to be set to color 135
IniRead, 145, %d%, 145			; what layers need to be set to color 145
IniRead, 155, %d%, 155			; what layers need to be set to color 155
IniRead, 165, %d%, 165			; what layers need to be set to color 165
IniRead, 175, %d%, 175			; what layers need to be set to color 175
IniRead, 185, %d%, 185			; what layers need to be set to color 185
IniRead, 195, %d%, 195			; what layers need to be set to color 195
IniRead, 205, %d%, 205			; what layers need to be set to color 205
IniRead, 215, %d%, 215			; what layers need to be set to color 215
IniRead, 225, %d%, 225			; what layers need to be set to color 225
IniRead, 235, %d%, 235			; what layers need to be set to color 235
IniRead, 245, %d%, 245			; what layers need to be set to color 245
IniRead, 251, %d%, 251			; what layers need to be set to color 251
IniRead, Freeze, %d%, Freeze		;what layers need to be frozen


For layers in acad.activedocument.layers
{
	layers.Color := 251
	layers.linetype := "Continuous"
}
Loop, Parse, 11, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
		{
			layers.Color := 11
			layers.linetype := "Center"
		}
	}
}

Loop, Parse, Freeze, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.freeze := 1
	}
}
Loop, Parse, 15, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.Color := 15
	}
}
Loop, Parse, 25, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.Color := 25
	}
}
Loop, Parse, 35, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.Color := 35
	}
}
Loop, Parse, 45, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.Color := 45
	}
}
Loop, Parse, 55, `n
{
	For layers in acad.activedocument.layers
	{
		if instr(layers.name, A_LoopField)
			layers.Color := 55
	}
}
Loop, Parse, 65, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 65
	}
}
Loop, Parse, 75, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 75
	}
}
Loop, Parse, 85, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 85
	}
}
Loop, Parse, 95, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 95
	}
}
Loop, Parse, 105, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 105
	}
}
Loop, Parse, 115, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 115
	}
}
Loop, Parse, 125, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 125
	}
}
Loop, Parse, 135, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 135
	}
}
Loop, Parse, 145, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 145
	}
}
Loop, Parse, 155, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 155
	}
}
Loop, Parse, 165, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 165
	}
}
Loop, Parse, 175, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 175
	}
}
Loop, Parse, 185, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 185
	}
}
Loop, Parse, 195, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 195
	}
}
Loop, Parse, 205, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 205
	}
}
Loop, Parse, 215, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 215
	}
}
Loop, Parse, 225, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 225
	}
}
Loop, Parse, 235, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 235
	}
}
Loop, Parse, 245, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 245
	}
}
Loop, Parse, 251, `n
{
	For layers in acad.activedocument.layers
	{
		if (layers.name = A_LoopField)
			layers.Color := 251
	}
}

acad.activedocument.save
acad.activedocument.Close
}
return
}


Explorer_GetWindow(hwnd="")
{
	hwnd := hwnd ? hwnd : WinExist("A")
	WinGetClass class, ahk_id %hwnd%
	
	if (class="CabinetWClass" or class="ExploreWClass" or class="Progman")
		for window in ComObjCreate("Shell.Application").Windows
			if (window.hwnd==hwnd)
				return window
}
Explorer_GetShellFolderView(hwnd="")
{
	return Explorer_GetWindow(hwnd).Document
}


Explorer_GetPath(hwnd="")
{
	return Explorer_GetWindow(hwnd).LocationURL
}
Explorer_GetItems(hwnd="")
{
	return Explorer_GetShellFolderView(hwnd).Folder.Items
}
Explorer_GetSelection(hwnd="")
{
	return Explorer_GetShellFolderView(hwnd).SelectedItems
}