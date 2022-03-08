' License: Unlicense, or 0-BSD
' No copyright claims
Class CSharpCompiler
	Public Target, OutputFileName, Recurse, Unsafe, Optimize
	Public Main ' /main:<type>, /m:
	Public Platform ' /platform:x86, Itanium, x64, or anycpu
	Public Reference ' /r:System.dll,System.Net.dll
	Public Lib ' /lib:c:\csharp\libraries
	Public Icon ' /win32icon:<file>
	Public Win32Res ' /win32res:<file>
	Public Resource ' /resource:<resinfo>
	Public LinkResource ' /linkresource:<resinfo>
	Public Manifest ' /win32manifest:<file>
	Public NoManifest ' /nowin32manifest
	Public NoLogo ' /nologo
	Public WarnLevel ' /warn
	Public AdditionalArgs ' Any string
	Private WSS, FSO
	Sub Class_Initialize()
		Set WSS = CreateObject("WScript.Shell")
		Set FSO = CreateObject("Scripting.FileSystemObject")
		NoManifest = False
		NoLogo = True : Unsafe = True : Optimize = True
		AdditionalArgs = "" : WarnLevel = ""
		Main = "" :  Resource = "" : LinkResource = ""
		Lib = "" : Icon = "" : Win32Res = "" : Manifest = ""
		Reference = "" : Platform = "" : OutputFileName = ""
		Target = "exe"
		Recurse = "*.cs"
	End Sub
	Public Function GetCompilerResult(aOutputFileName, aTarget, aIcon)
		If IsNull(aTarget) Then aTarget = "exe"
		If Not IsNull(aIcon) Then Icon = aIcon
		Target = aTarget
		Dim vCmd: vCmd = GetCompilerCommandLine(aOutputFileName)
		GetCompilerResult = GetCSharpCompileResult(null, vCmd)
	End Function
	Function GetCompilerCommandLine(aOutputFileName)
		Dim vArgs: vArgs = ""
		If "" <> Main Then vArgs = vArgs & "/m:" & Main & " "
		If "" <> aOutputFileName Then
			vArgs = vArgs & "/out:" & aOutputFileName & " "
		ElseIf "" <> OutputFileName Then
			vArgs = vArgs & "/out:" & OutputFileName & " "
		End If
		If "" <> Target Then vArgs = vArgs & "/target:" & Target & " "
		If NoLogo Then vArgs = vArgs & "/nologo" & " "
		If Unsafe Then vArgs = vArgs & "/unsafe" & " "
		If "" <> WarnLevel Then vArgs = vArgs & "/warn:" & WarnLevel & " "
		If Optimize Then vArgs = vArgs & "/o" & " "
		If NoManifest Then vArgs = vArgs & "/nowin32manifest" & " "
		If "" <> Main Then vArgs = vArgs & "/main:" & Main & " "
		If "" <> Recurse Then vArgs = vArgs & "/recurse:" & Recurse & " "
		If "" <> Platform Then vArgs = vArgs & "/platform:" & Platform & " "
		If "" <> Win32Res Then vArgs = vArgs & "/win32res:" & Win32Res & " "
		If "" <> Resource Then vArgs = vArgs & "/resource:" & Resource & " "
		If "" <> Manifest Then vArgs = vArgs & "/win32manifest:" & Manifest & " "
		If "" <> Icon Then vArgs = vArgs & "/win32icon:" & Icon & " "
		If "" <> Lib Then vArgs = vArgs & "/lib:" & Lib & " "
		GetCompilerCommandLine = vArgs & " " & AdditionalArgs
	End Function
	Function GetCSharpCompileResult(aVersionType, aCommand)
		' aVersionType [0=v3.5, 1=v2.0.50727, 2=v4.0.30319]
		vCSC = GetCSharpCompilerPath(aVersionType)
		GetCSharpCompileResult = GetExecOutput(vCSC & " " & aCommand)
	End Function
	' Important
	Function Environ(aString) : Environ = WSS.ExpandEnvironmentStrings(aString) : End Function
	Function IO_FileExists(aFilePath) : IO_FileExists = FSO.FileExists(aFilePath) : End Function
	Function IO_DirExists(aDirectoryPath) : IO_DirExists = FSO.FolderExists(aDirectoryPath) : End Function
	' C#/CSharp compiler
	Function GetCSharpCompilerPath(aVersionType) : GetCSharpCompilerPath = GetDNFXPath(aVersionType) & "/csc.exe" : End Function
	Function GetDNFXPath(aVersionType)
		' aVersionType [0=v3.5, 1=v2.0.50727, 2=v4.0.30319]
		Dim vVersionStr: vVersionStr = "v3.5"
		If IsNull(aVersionType) Then aVersionType = 0
		If 2 = aVersionType Then vVersionStr = "v2.0.50727"
		If 1 = aVersionType Then vVersionStr = "v4.0.30319"
		Dim vPath: vPath = Environ("%windir%\Microsoft.NET\Framework")
		Dim vPath2: vPath2 = vPath & "64\" & vVersionStr
		GetDNFXPath = vPath & "\" & vVersionStr
		'Msgbox vPath2
		If IO_DirExists(vPath2) Then GetDNFXPath = vPath2
	End Function
	' Added:
	Function GetExecOutput(aCommand)
		Set vExecObj = WSS.Exec(aCommand)
		Dim vResult: vResult = ""
		Do While Not vExecObj.StdOut.AtEndOfStream
			vResult = vResult & vExecObj.StdOut.ReadLine()
			WScript.Sleep 1
		Loop
		GetExecOutput = vResult
	End Function
End Class
