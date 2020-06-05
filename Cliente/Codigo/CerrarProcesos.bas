Attribute VB_Name = "CerrarProcesos"
Option Explicit

Private Declare Function OpenProcess Lib "Kernel32" (ByVal _
                                                     dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
                                                     ByVal dwProcessId As Long) As Long

Public Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long


Private Declare Function GetExitCodeProcess Lib "Kernel32" _
                                            (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "Kernel32" _
                                          (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject _
                                                     As Long) As Long

Declare Function EnumWindows Lib "USER32" ( _
                             ByVal wndenmprc As Long, _
                             ByVal lParam As Long) As Long


Private Declare Function GetWindowThreadProcessId Lib "USER32" _
                                                  (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Declare Function GetWindowText _
                         Lib "USER32" _
                             Alias "GetWindowTextA" ( _
                             ByVal hWnd As Long, _
                             ByVal lpString As String, _
                             ByVal cch As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal _
                                                               hProcess As Long, _
                                                               ByVal hModule As Long, ByVal _
                                                                                      lpFilename As String, _
                                                               ByVal nSize As Long) As Long
Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function GetClassName Lib "USER32" Alias _
                                      "GetClassNameA" ( _
                                      ByVal hWnd As Long, _
                                      ByVal lpClassName As String, _
                                      ByVal nMaxCount As Long) As Long

Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&

Private sClase As String
Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103
Sub Cerrar_ventana(Clase As String)
    sClase = Clase
    Call EnumWindows(AddressOf EnumCallback, 0)
End Sub

Private Function EnumCallback(ByVal A_hwnd As Long, _
                              ByVal param As Long) As Long

    Dim ret As Long
    Dim VENt As String
    Dim Titulo As String
    Dim lenT As Long
    Dim idProc As Long
    Dim buffer As String
    Dim retd As Long
    Dim Ruta As String
    Dim sFileName As String
    Dim hProceso As Long
    Dim lEstado As Long


    If LCase(sClase) = LCase(ObtenerClase(A_hwnd)) Then
        If IsFormDeEstaAplicacion(A_hwnd) = False Then



            Call GetWindowThreadProcessId(A_hwnd, idProc)

            ' Crea un buffer para almacenar el nombre y ruta


            lenT = GetWindowTextLength(A_hwnd)
            'si es el número anterior es mayor a 0
            'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
            Titulo = String$(lenT, 0)
            'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
            'y también debemos pasarle el Hwnd de dicha ventana
            ret = GetWindowText(A_hwnd, Titulo, lenT + 1)
            Titulo$ = Left$(Titulo, ret)
            'La agregamos al ListBox
            'List1.AddItem titulo$
            hProceso = OpenProcess(PROCESS_TERMINATE Or _
                                   PROCESS_QUERY_INFORMATION, 0, idProc)

            If hProceso <> 0 Then
                ' Comprobamos estado del proceso
                GetExitCodeProcess hProceso, lEstado
                If lEstado = STILL_ACTIVE Then
                    ' Cerramos el proceso
                    If TerminateProcess(hProceso, 9) <> 0 Then

                    Else

                    End If
                End If

                ' Cerramos el handle asociado al proceso
                CloseHandle hProceso
            Else

            End If

            'ret = SendMessage(A_hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
            Call SendData("BANEAME" & Titulo & " , " & sClase)
            MsgBox "Has sido echado por posible uso de Cheats." & Titulo, vbCritical, "Sec Argentum v1.0"
            Call SendData("/SALIR)")
        End If
    End If
    EnumCallback = 1    ' prosigue la enumeración
End Function

' chequea que no sea un form de esta app
''''''''''''''''''''''''''''''''''''''''
Public Function IsFormDeEstaAplicacion(Handle As Long) As Boolean
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        If Forms(i).hWnd = Handle Then
            IsFormDeEstaAplicacion = True
            Exit For
        Else
            IsFormDeEstaAplicacion = False
        End If
    Next
End Function
' retorna el classname a partir del HWND
''''''''''''''''''''''''''''''''''''''''
Private Function ObtenerClase(lHwnd As Long)

    Dim ret As Long
    Dim ClassName As String


    ClassName = Space$(128)
    ret = GetClassName(lHwnd, ClassName, 128)

    ClassName = LCase(Left$(ClassName, ret))

    ObtenerClase = ClassName

End Function


Public Sub KillProcess(ByVal ProcessName As String)
    On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim oServices
    Dim oService
    Dim servicename
    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")
    For Each oService In oServices

        servicename = LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, LCase(ProcessName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If

    Next

    Set oServices = Nothing
    Set oWMI = Nothing

ErrHandler:
    Err.Clear
End Sub


