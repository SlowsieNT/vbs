Function GetBitness()
	GetBitness = 16
	Dim vPARCH: vPARCH = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
	If 0 <> InStr(vPARCH, "86") Then GetBitness = 32
	If 0 <> InStr(vPARCH, "64") Then GetBitness = 64
	' Add unnecessary checks?
	If 0 <> InStr(vPARCH, "128") Then GetBitness = 128
	If 0 <> InStr(vPARCH, "256") Then GetBitness = 256
	If 0 <> InStr(vPARCH, "512") Then GetBitness = 512
End Function
