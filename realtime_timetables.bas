Attribute VB_Name = "realtime_timetables"
Sub main()

Call request_realtime_bundle("U:\test_gtds_bundle_forest", "SMBSC002")

'SMBSC001: Busways Western Sydney
'SMBSC002: Interline Bus Services
'SMBSC003: Transit Systems
'SMBSC004: Hillsbus
'SMBSC005: Punchbowl Bus Company
'SMBSC006: State Transit Sydney
'SMBSC007: State Transit Sydney
'SMBSC008: State Transit Sydney
'SMBSC009: State Transit Sydney
'SMBSC010: Transdev NSW
'SMBSC012: Transdev NSW
'SMBSC013: Transdev NSW
'SMBSC014: Forest Coach LinesSMBSC002
'SMBSC015: Busabout
'OSMBSC001: Rover Coaches
'OSMBSC002: Hunter Valley Buses
'OSMBSC003: Port Stephens Coaches
'OSMBSC004: Hunter Valley Buses
'OSMBSC005: State Transit Newcastle
'OSMBSC006: Busways Central Coast
'OSMBSC007: Red Bus Service
'OSMBSC008: Blue Mountains Transit
'OSMBSC009: Premier Charters
'OSMBSC010: Premier Illawarra
'OSMBSC011: Coastal Liner
'OSMBSC012: Dions Bus Service

'buses
'ferries/sydneyferries
'lightrail/innerwest
'lightrail/newcastle
'nswtrains
'sydneytrains
'metro


End Sub
Sub request_realtime_bundle(ByVal folder_path As Variant, ByVal contract_name As Variant)

Dim api_url             As String
Dim oXH                 As MSXML2.XMLHTTP60
Dim oStream             As ADODB.Stream
Dim bodytxt             As String
Dim distanc_e, tim_e    As String

API_KEY = "IO3Bo6fumY4qOT1SKByqgraAGFA6cLX565E2"

Set oXH = New MSXML2.XMLHTTP60

api_url = "https://api.transport.nsw.gov.au/v1/gtfs/schedule/buses/" & contract_name

Debug.Print (api_url)
    
    With oXH
        .Open "get", api_url, False
        .setRequestHeader "Accept", "application/octet-stream"
        .setRequestHeader "Authorization", "apikey " & API_KEY
        .send
        Debug.Print (.statusText)
        Debug.Print (.Status)
        Debug.Print (.getAllResponseHeaders)
    End With

If oXH.Status = 200 Then
    Set oStream = New ADODB.Stream
    oStream.Open
    oStream.Type = 1
    oStream.Write oXH.responseBody
    oStream.SaveToFile folder_path & ".zip", 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If

Set oXH = Nothing

Call UnzipAFile(folder_path & ".zip", folder_path)

End Sub
Sub UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)

Dim fso As FileSystemObject
Set fso = New FileSystemObject

If Not fso.FileExists(zippedFileFullName) Then MsgBox "File does not exist" Else
If Not fso.FolderExists(unzipToPath) Then fso.CreateFolder (unzipToPath) Else

Dim ShellApp As Object
'Copy the files & folders from the zip into a folder
Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(unzipToPath).CopyHere ShellApp.Namespace(zippedFileFullName).Items

fso.DeleteFile (zippedFileFullName)

End Sub
Sub testere()
Debug.Print (get_file_name("U:\testing.txt"))
Debug.Print (get_folder_name("U:\testing.txt"))
Debug.Print (get_extension_name("U:\testing.txt"))
End Sub

Function get_file_name(ByVal full_path As String) As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
get_file_name = fso.GetBaseName(full_path)
Set fso = Nothing
End Function
Function get_folder_name(ByVal full_path As String) As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
get_folder_name = fso.GetParentFolderName(full_path)
Set fso = Nothing
End Function
Function get_extension_name(ByVal full_path As String) As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
get_extension_name = fso.GetExtensionName(full_path)
Set fso = Nothing
End Function

