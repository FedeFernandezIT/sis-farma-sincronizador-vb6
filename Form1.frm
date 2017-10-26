VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sync Farmatic"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Sincronizando..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "Menú"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "Mostrar"
      End
      Begin VB.Menu mSep 
         Caption         =   "-"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    
    With nid
     .cbSize = Len(nid)
     .hwnd = Me.hwnd
     .uId = vbNull
     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     .uCallBackMessage = WM_MOUSEMOVE
     .hIcon = Me.Icon
     .szTip = "Sync Farmatic" & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_ADD, nid
       
    ejecutarProcesos
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbCenterScreen
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbCenterScreen
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
    End Select
End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
    MatarProceso ("VBMTEjecutor.exe")
End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    Unload Me
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    
    Me.WindowState = vbCenterScreen
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub ejecutarProcesos()
    Dim E1 As VBMTEjecutor.Ejecutor
    Dim E2 As VBMTEjecutor.Ejecutor
    Dim E3 As VBMTEjecutor.Ejecutor
    Dim E4 As VBMTEjecutor.Ejecutor
    Dim E5 As VBMTEjecutor.Ejecutor
    Dim E6 As VBMTEjecutor.Ejecutor
    Dim E7 As VBMTEjecutor.Ejecutor
    Dim E8 As VBMTEjecutor.Ejecutor
    Dim E9 As VBMTEjecutor.Ejecutor
    Dim E10 As VBMTEjecutor.Ejecutor
    Dim E11 As VBMTEjecutor.Ejecutor
    Dim E12 As VBMTEjecutor.Ejecutor
    Dim E13 As VBMTEjecutor.Ejecutor
    Dim E14 As VBMTEjecutor.Ejecutor
    Dim E15 As VBMTEjecutor.Ejecutor
    Dim E16 As VBMTEjecutor.Ejecutor
    Dim E17 As VBMTEjecutor.Ejecutor
    Dim E18 As VBMTEjecutor.Ejecutor
    
    Set E1 = New VBMTEjecutor.Ejecutor
    Set E2 = New VBMTEjecutor.Ejecutor
    Set E3 = New VBMTEjecutor.Ejecutor
    Set E4 = New VBMTEjecutor.Ejecutor
    Set E5 = New VBMTEjecutor.Ejecutor
    Set E6 = New VBMTEjecutor.Ejecutor
    Set E7 = New VBMTEjecutor.Ejecutor
    Set E8 = New VBMTEjecutor.Ejecutor
    Set E9 = New VBMTEjecutor.Ejecutor
    Set E10 = New VBMTEjecutor.Ejecutor
    Set E11 = New VBMTEjecutor.Ejecutor
    Set E12 = New VBMTEjecutor.Ejecutor
    Set E13 = New VBMTEjecutor.Ejecutor
    Set E14 = New VBMTEjecutor.Ejecutor
    Set E15 = New VBMTEjecutor.Ejecutor
    Set E16 = New VBMTEjecutor.Ejecutor
    Set E17 = New VBMTEjecutor.Ejecutor
    Set E18 = New VBMTEjecutor.Ejecutor
    
    E1.Procesar ("PENDIENTE_PUNTOS")
    E2.Procesar ("PEDIDOS")
    E3.Procesar ("PRODUCTOS_CRITICOS")
    E4.Procesar ("ENCARGOS")
    E5.Procesar ("FAMILIAS")
    E6.Procesar ("CATEGORIASPS")
    E7.Procesar ("CLIENTES")
    E8.Procesar ("CONTROL_STOCK")
    E9.Procesar ("CONTROL_SIN_STOCK")
    E10.Procesar ("CONTROL_STOCK_FECHAS_ENTRADA")
    E11.Procesar ("CONTROL_STOCK_FECHAS_SALIDA")
    E12.Procesar ("LISTA_TIENDA")
    E13.Procesar ("LISTAS")
    E14.Procesar ("SINONIMOS")
    E15.Procesar ("ACTUALIZAR_PP")
    E16.Procesar ("ACTUALIZAR_PRODUCTOS_BORRADOS")
    E17.Procesar ("ACTUALIZAR_ENTREGAS_CLIENTES")
    E18.Procesar ("ACTUALIZAR_RECETAS_PENDIENTES")
        
    Set E1 = Nothing
    Set E2 = Nothing
    Set E3 = Nothing
    Set E4 = Nothing
    Set E5 = Nothing
    Set E6 = Nothing
    Set E7 = Nothing
    Set E8 = Nothing
    Set E9 = Nothing
    Set E10 = Nothing
    Set E11 = Nothing
    Set E12 = Nothing
    Set E13 = Nothing
    Set E14 = Nothing
    Set E15 = Nothing
    Set E16 = Nothing
    Set E17 = Nothing
    Set E18 = Nothing
End Sub

Private Function MatarProceso(StrNombreProceso As String) As Boolean
    MatarProceso = False
    Set ObjetoWMI = GetObject("winmgmts:")
    
    If IsNull(ObjetoWMI) = False Then
       Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
    
       For Each ProcesoACerrar In ListaProcesos
           If UCase(ProcesoACerrar.Name) = UCase(StrNombreProceso) Then
              ProcesoACerrar.Terminate (0)
              MatarProceso = True
           End If
           
           If MatarProceso Then
              Exit For
           End If
       Next
    End If
    
    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing
End Function
