Author: Lovepreet Singh
Ducumenting the finding and research to Convert 32 bit Macro to 64 Bit
Date: 30-Nov-2023

----------------------------------------------------------------------
To Convert 32bit macro to 64bit macro.
----------------------------------------------------------------------
Pre-requsites:
Excel 64bit and other applications if any
eg. SAP 64bit
----------------------------------------------------------------------
- Convert "Long" datatype to "LongPtr"
- Convert all 32bit datatypes to 64 bit: As below
	- Some time we may find datatypes in 32bit macro/vba code which are like: "SIZE32","RECT32" etc.
	  Below is how we can craete custom Datatypes for 64Bit ->
----------------------------------------------------------------------
		            SAMPLE CODE
----------------------------------------------------------------------

#---------------Example 1----------------#
Public Type RGB32
        Red As LongPtr
        Green As LongPtr
        Blue As LongPtr
End Type

#---------------Example 2----------------#
Public Type RECT32
        Left As LongPtr
        Top As LongPtr
        Right As LongPtr
        Bottom As LongPtr
End Type

#---------------Example 3----------------#
Public Type SIZE32
        cx As LongPtr
        cy As LongPtr
End Type

#---------------Example 4----------------#
Public Type LOGFONT32
        lfHeight As LongPtr
        lfWidth As LongPtr
        lfEscapement As LongPtr
        lfOrientation As LongPtr
        lfWeight As LongPtr
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

----------------------------------------------------------------------
		       SCRIPTCONTROL for 64Bit
----------------------------------------------------------------------
ScriptControl is deprecated as part of 64 bit migrations, below is the WorkAround:

You can create ActiveX objects like ScriptControl, which available on 32-bit Office versions via mshta x86 host on 64-bit VBA version, here is the example (put the code in a standard VBA project module):
Option Explicit

Sub Test()
    
    Dim oSC As Object
    
    Set oSC = CreateObjectx86("ScriptControl") ' create ActiveX via x86 mshta host
    Debug.Print TypeName(oSC) ' ScriptControl
    ' do some stuff
    
    CreateObjectx86 Empty ' close mshta host window at the end
    
End Sub

Function CreateObjectx86(sProgID)
   
    Static oWnd As Object
    Dim bRunning As Boolean
    
    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If IsEmpty(sProgID) Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        If Not IsEmpty(sProgID) Then Set CreateObjectx86 = CreateObject(sProgID)
    #End If
    
End Function

Function CreateWindow()

    ' source https://github.com/lovepreetsingh-ai/32bit-to-64-bit-vba
    Dim sSignature, oShellWnd, oProc
    
    On Error Resume Next
    Do Until Len(sSignature) = 32
        sSignature = sSignature & Hex(Int(Rnd * 16))
    Loop
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop
    
End Function
----------------------------------------------------------------------
		        Declare PtrSafe functions
----------------------------------------------------------------------
From:
- Public Declare Function FunctionName

To:
- Public Declare PtrSafe Function FunctionName

----------------------------------------------------------------------
			      END OF FILE
----------------------------------------------------------------------
