# vbs

## ini.vbs
```vbs
Set ini = new IniFile
ini.Open "test.ini"

' set/add values to ini
ini.SetValue "general", "key1", "value1"
ini.SetValue "general", "key1", "value12"

' retrieve value by key "a" from section "general"
WScript.Echo "Get value:" & vbNewLine & ini.GetValue("general", "a")
' here we list all sections,
' includes empty section name if there was no first section name
WScript.Echo "Sections:" & vbNewLine & Join(ini.GetSections(), ",")

' IF you need to get key/values by section name, do this:
' vLine(0=Section, 1=Key, 2=Value, 3=Line, 4=Comment, 5=DefineSectionBool)
For Each vLine In ini.GetLinesBySection("general")
	' index 1 is key name
	' index 2 is value
	If "" <> vLine(1) Then
		' display it like this: key=value
		MsgBox vLine(1) & "=" & vLine(2)
	End If
Next

' To save opened file:
' Write 0
' To save in new file:
' Write "newfile.ini"
ini.Write 0
```
