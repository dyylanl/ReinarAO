Attribute VB_Name = "Module3"
Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
                                      "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
                                                                                               lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
                                                               lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
                                                               lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
                                                                                                                                  nFileSystemNameSize As Long) As Long

Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(strDrive, Temp1, _
                               Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function

