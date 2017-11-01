VERSION 5.00
Begin VB.Form Proceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sisfarma"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer_Actualizar_Recetas_Pendientes '''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1800
      Top             =   2400
   End
   Begin VB.Timer Timer_Actualizar_Entregas_Clientes ''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer Timer_Actualizar_Productos_Borrados ''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   360
      Top             =   2400
   End
   Begin VB.Timer Timer_Actualizar_Pendiente_Puntos ''''''''''''''''' migrado
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer Timer_Sinonimos '''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   3240
      Top             =   1800
   End
   Begin VB.Timer Timer_Control_Sin_Stock_Inicial '''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.Timer Timer_Control_Stock_Fechas_Salida '''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Timer Timer_Pedidos '''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   3100
      Left            =   1080
      Top             =   1800
   End
   Begin VB.Timer Timer_Clientes_Huecos ''''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   63000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Timer Timer_Lista_Tienda '''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   3960
      Top             =   1080
   End
   Begin VB.Timer Timer_Categorias_PS ''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   64000
      Left            =   3240
      Top             =   1080
   End
   Begin VB.Timer Timer_Listas_Fechas ''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   62000
      Left            =   2520
      Top             =   1080
   End
   Begin VB.Timer Timer_Listas '''''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   61000
      Left            =   1800
      Top             =   1080
   End
   Begin VB.Timer Timer_Control_Stock_Inicial '''''''''''''''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   1080
   End
   Begin VB.Timer Timer_Familias ''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   360
      Top             =   1080
   End
   Begin VB.Timer Timer_Encargos ''''''''''''''''''' Migardo
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3960
      Top             =   360
   End
   Begin VB.Timer Timer_Productos_Criticos ''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   3240
      Top             =   360
   End
   Begin VB.Timer Timer_Control_Stock_Fechas_Entrada '''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2520
      Top             =   360
   End
   Begin VB.Timer Timer_Pendiente_Puntos '''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer Timer_Clientes ''''''''''''''''''''''''''''''''' Migrado
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   1080
      Top             =   360
   End
End
Attribute VB_Name = "Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim connMySql As ADODB.Connection
Dim connSqlServer As ADODB.Connection
Dim connSqlServerBP As ADODB.Connection

Dim serverLocal As String
Dim baseLocal As String
Dim serverRemoto As String
Dim baseRemoto As String
Dim usuarioLocal As String
Dim passLocal As String
Dim codListaTienda As String

Function SqlSafe(strInput As String) As String
    SqlSafe = Replace(strInput, "'", "''")
    SqlSafe = Replace(SqlSafe, """", """""")
End Function
    
Function StripString(MyStr As Variant) As Variant
    On Error GoTo StripStringError
    
    Dim strChar As String, strHoldString As String
    Dim i As Integer
    
    ' Exit if the passed value is null.
    If IsNull(MyStr) Then Exit Function
    
    ' Exit if the passed value is not a string.
    If VarType(MyStr) <> 8 Then Exit Function
    
    ' Check each value for invalid characters.
    For i = 1 To Len(MyStr)
        strChar = Mid$(MyStr, i, 1)
        Select Case strChar
            Case ".", "'", ",", "-", "\"
                ' Do nothing
            Case Else
                strHoldString = strHoldString & strChar
        End Select
    Next i
    
    ' Pass back corrected string.
    StripString = strHoldString
    
StripStringEnd:
    Exit Function
    
StripStringError:
    MsgBox Error$
    Resume StripStringEnd
End Function

Public Function CerosIzq(numero As String, TotalDigitos As Integer) As String
    If Len(numero) > TotalDigitos Then
        CerosIzq = numero
    Else
        CerosIzq = String(TotalDigitos - Len(numero), "0") & numero
    End If
End Function

Function quitarCaracterCadena(ByVal cadena As String, ByVal caracter) As String
    Dim i As Long
    Dim j As Long
    Dim cadTemporal As String
    Dim sCaracter$
    
    quitarCaracterCadena = ""
    If Not IsMissing(caracter) Then
        sCaracter = caracter
        cadTemporal = ""
        For i = 1 To Len(cadena)
            If InStr(sCaracter, Mid$(cadena, i, 1)) = 0 Then
                cadTemporal = cadTemporal & Mid$(cadena, i, 1)
            End If
        Next
        quitarCaracterCadena = cadTemporal
    End If
End Function

Public Sub leerFicherosConfiguracion()
ficheros:

    On Error GoTo errorFicheros
    
    Const FICHERO1 As String = "c:\server_local.txt"
    Open FICHERO1 For Binary As #1

    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
         
        serverLocal = buffer1$
        
        start& = start& + 20000
    Next
    
    buffer1$ = String$(part&, 0)
   
    Get #1, start&, buffer1$        'get the remaining bytes at the end
    
    serverLocal = buffer1$
  
    Close #1
           
    On Error GoTo errorFicheros
        
    Const FICHERO2 As String = "c:\base_local.txt"
    Open FICHERO2 For Binary As #1

    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
         
        baseLocal = buffer1$
        
        start& = start& + 20000
    Next
    
    buffer1$ = String$(part&, 0)
   
    Get #1, start&, buffer1$        'get the remaining bytes at the end
    
    baseLocal = buffer1$
  
    Close #1
    
    On Error GoTo errorFicheros
    
    Const FICHERO3 As String = "c:\server_remoto.txt"
    Open FICHERO3 For Binary As #1

    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
         
        serverRemoto = buffer1$
        
        start& = start& + 20000
    Next
    
    buffer1$ = String$(part&, 0)
   
    Get #1, start&, buffer1$        'get the remaining bytes at the end
    
    serverRemoto = buffer1$
  
    Close #1

    On Error GoTo errorFicheros

    Const FICHERO4 As String = "c:\base_remoto.txt"
    Open FICHERO4 For Binary As #1
    
    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
        
        baseRemoto = buffer1$
       
        start& = start& + 20000
    Next
        
    buffer1$ = String$(part&, 0)
 
    Get #1, start&, buffer1$        'get the remaining bytes at the end
  
    baseRemoto = buffer1$
  
    Close #1
    
    On Error GoTo errorFicheros
    
    Const FICHERO5 As String = "c:\usuario_local.txt"
    Open FICHERO5 For Binary As #1

    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
         
        usuarioLocal = buffer1$
        
        start& = start& + 20000
    Next
    
    buffer1$ = String$(part&, 0)
   
    Get #1, start&, buffer1$        'get the remaining bytes at the end
    
    usuarioLocal = buffer1$
  
    Close #1
    
    On Error GoTo errorFicheros
    
    Const FICHERO6 As String = "c:\pass_local.txt"
    Open FICHERO6 For Binary As #1

    whole& = LOF(1) \ 20000
    part& = LOF(1) Mod 20000
    buffer1$ = String$(20000, 0)
    start& = 1
    
    For X& = 1 To whole&
        Get #1, start&, buffer1$
         
        passLocal = buffer1$
        
        start& = start& + 20000
    Next
    
    buffer1$ = String$(part&, 0)
   
    Get #1, start&, buffer1$        'get the remaining bytes at the end
    
    passLocal = buffer1$
  
    Close #1
    
    codListaTienda = -1
    
    If Dir$("c:\cod_lista_tienda.txt") <> "" Then
        On Error GoTo errorFicheros
        
        Const FICHERO7 As String = "c:\cod_lista_tienda.txt"
        Open FICHERO7 For Binary As #1
    
        whole& = LOF(1) \ 20000
        part& = LOF(1) Mod 20000
        buffer1$ = String$(20000, 0)
        start& = 1
        
        For X& = 1 To whole&
            Get #1, start&, buffer1$
             
            codListaTienda = buffer1$
            
            start& = start& + 20000
        Next
        
        buffer1$ = String$(part&, 0)
       
        Get #1, start&, buffer1$        'get the remaining bytes at the end
        
        codListaTienda = buffer1$
      
        Close #1
    End If
    
final:
    GoTo fin
    
errorFicheros:
    MsgBox "Ha habido un error en la lectura de algún fichero de configuración. Compruebe que existen dichos ficheros de configuración", vbCritical
        
fin:
    
End Sub

Public Sub establecerConexionesBasesDatos()
conexion:

    On Error GoTo errorConexion

    Set connMySql = New ADODB.Connection
    connMySql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" + serverRemoto + ";Database=" + baseRemoto + "; User=fisiotes_fede;Password=tGLjuIUr9A;Option=3;"
    connMySql.ConnectionTimeout = 3600
    connMySql.CommandTimeout = 3600
    connMySql.Open
    
    On Error GoTo errorConexion
    
    Set connSqlServer = New ADODB.Connection
    connSqlServer.ConnectionString = "Provider=sqloledb;Data Source=" + serverLocal + ";Initial Catalog=" + baseLocal + ";User Id=" + usuarioLocal + ";Password=" + passLocal + ";"
    connSqlServer.ConnectionTimeout = 3600
    connSqlServer.CommandTimeout = 3600
    connSqlServer.Open
    
    On Error GoTo errorConexion

    Set connSqlServerBP = New ADODB.Connection
    connSqlServerBP.ConnectionString = "Provider=sqloledb;Data Source=" + serverLocal + ";Initial Catalog=Consejo;User Id=" + usuarioLocal + ";Password=" + passLocal + ";"
    connSqlServerBP.ConnectionTimeout = 3600
    connSqlServerBP.CommandTimeout = 3600
    connSqlServerBP.Open
    
final:
    GoTo fin
    
errorConexion:
    GoTo conexion

fin:

End Sub

''''''''''' MIGRACION HECHA ''''''''''''''''
Public Sub setCeroClientes()
    Dim sql As String
    
inicio:
    On Error GoTo errorActualizarClientesCero
    
    leerFicherosConfiguracion
    
    On Error GoTo errorActualizarClientesCero
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorActualizarClientesCero

    sql = "UPDATE clientes SET dni_tra = 0"
    
    connMySql.Execute (sql)
    
final:
    GoTo fin
    
errorActualizarClientesCero:
    GoTo inicio
    
fin:
    
End Sub

Public Sub Procesar(ByVal Id As Long)
    iId = Id
    Timer1.Enabled = True
End Sub

Public Sub procesarTimerClientes()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos

    Timer_Clientes.Enabled = True
End Sub

Public Sub procesarTimerClientesHuecos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos

    Timer_Clientes_Huecos.Enabled = True
End Sub

Public Sub procesarTimerPendientePuntos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Pendiente_Puntos.Enabled = True
End Sub

Public Sub procesarTimerControlStockInicial()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Control_Stock_Inicial.Enabled = True
End Sub

Public Sub procesarTimerControlSinStockInicial()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Control_Sin_Stock_Inicial.Enabled = True
End Sub

Public Sub procesarTimerControlStockFechasEntrada()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Control_Stock_Fechas_Entrada.Enabled = True
End Sub

Public Sub procesarTimerControlStockFechasSalida()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Control_Stock_Fechas_Salida.Enabled = True
End Sub

Public Sub procesarTimerProductosCriticos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Productos_Criticos.Enabled = True
End Sub

Public Sub procesarTimerPedidos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Pedidos.Enabled = True
End Sub

Public Sub procesarTimerEncargos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Encargos.Enabled = True
End Sub

Public Sub procesarTimerFamilias()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Familias.Enabled = True
End Sub

Public Sub procesarTimerCategorias()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Categorias_PS.Enabled = True
End Sub

Public Sub procesarTimerListas()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Listas.Enabled = True
End Sub

Public Sub procesarTimerListasFechas()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Listas_Fechas.Enabled = True
End Sub

Public Sub procesarTimerListaTienda()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Lista_Tienda.Enabled = True
End Sub

Public Sub procesarTimerSinonimos()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Sinonimos.Enabled = True
End Sub

Public Sub procesarTimerActualizarPP()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Actualizar_Pendiente_Puntos.Enabled = True
End Sub

Public Sub procesarTimerActualizarPB()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Actualizar_Productos_Borrados.Enabled = True
End Sub

Public Sub procesarTimerActualizarEntregasClientes()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Actualizar_Entregas_Clientes.Enabled = True
End Sub

Public Sub procesarTimerActualizarRecetasPendientes()
    leerFicherosConfiguracion
    'establecerConexionesBasesDatos
    
    Timer_Actualizar_Recetas_Pendientes.Enabled = True
End Sub

Private Sub Timer_Clientes_Timer()''''''''''''''''' Migrado ''''''''''''''''''''''''''''''
    Dim ultimoCliente As String
        
    Dim telefono As String
    Dim movil As String
    Dim email As String
    Dim direccion As String
    Dim nombre As String
    Dim tarjeta As String
    Dim trabajador As String
    
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim FieldExistsInRS1 As Boolean
    Dim existCampoSexo As Boolean
    Dim oField
    
    FieldExistsInRS = False
    FieldExistsInRS1 = False
    existCampoSexo = False
    
    On Error GoTo errorTimerClientes
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorTimerClientes
        
    '''''''''' MIGRACION HECHA ''''''''''''''''''''''''''''
    sql = "CREATE TABLE IF NOT EXISTS `clientes_huecos` (" & _
            "`hueco` varchar(255) DEFAULT NULL" & _
            ") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
            
    connMySql.Execute (sql)
    
    On Error GoTo errorTimerClientes
    
    
    '''''''''' MIGRACION HECHA ''''''''''''''''''''''''''''
    sql = "SELECT * from clientes LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "baja" Then
            FieldExistsInRS = True
        End If
        If oField.Name = "fechaAlta" Then
            FieldExistsInRS1 = True
        End If
    Next
    
    rs.Close
    
    '''''''''' MIGRACION HECHA ''''''''''''''''''''''''''''
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE clientes ADD `baja` tinyint(1) DEFAULT 0 AFTER dia_alta;"
        
        connMySql.Execute (sql)
    End If
    
    '''''''''' MIGRACION HECHA ''''''''''''''''''''''''''''
    If FieldExistsInRS1 = False Then
        sql = "ALTER TABLE clientes ADD `fechaAlta` datetime AFTER dia_alta;"
        
        connMySql.Execute (sql)
    End If
    
    On Error GoTo errorTimerClientes
    
    
    '''''''''' MIGRACION HECHA ''''''''''''''''''''''''''''
    sql = "SELECT TOP 1 * FROM ClienteAux"
        
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If LCase(oField.Name) = "sexo" Then
            existCampoSexo = True
        End If
    Next
    
    rs.Close
    
    'If enTimerClientes Then
        On Error GoTo errorTimerClientes
        
        'sql = "SELECT * FROM clientes ORDER BY dni DESC LIMIT 0,1"
        
        '' MIGRACION HECHA ''
        sql = "SELECT * FROM clientes WHERE dni_tra = 1"
        
        rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        If Format(Now, "HHmm") = "1500" Or Format(Now, "HHmm") = "2300" Then
            ultimoCliente = 0
        Else
            If rs.EOF Then
                ultimoCliente = 0
            Else
                ultimoCliente = rs!dni
                
                '' SE VA AL SEGUNDO ''
                rs.MoveNext
                
                If Not rs.EOF Then
                    ultimoCliente = rs!dni
                End If
            End If
        End If
        
        contadorHuecos = -1
        
        rs.Close
        
        On Error GoTo errorTimerClientes
        
        sql = "SELECT * FROM cliente WHERE Idcliente > " & ultimoCliente & " ORDER BY CAST(Idcliente AS DECIMAL(20)) ASC"
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
             
        Do While Not rs.EOF
            DoEvents
            
            If contadorHuecos = -1 Then
                contadorHuecos = CDbl(rs!idcliente)
            End If
            
            On Error GoTo errorTimerClientes
            
            sql = "SELECT * FROM clientes WHERE dni='" & rs!idcliente & "'"
            
            rs3.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
            
            On Error GoTo errorTimerClientes
            
            sql = "SELECT * FROM Destinatario WHERE fk_Cliente_1 ='" & rs!idcliente & "'"
            
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If Not rs2.EOF Then
                If IsNull(rs2!TlfMovil) Then
                    movil = ""
                Else
                 movil = Trim(rs2!TlfMovil)
                End If
                
                If IsNull(rs2!email) Then
                    email = ""
                Else
                 email = Trim(rs2!email)
                End If
            Else
                 movil = ""
                 email = ""
            End If
            
            rs2.Close
            
            On Error GoTo errorTimerClientes
            
            sql = "SELECT * FROM ClienteAux WHERE idCliente = '" & rs!idcliente & "'"
    
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If Not rs2.EOF Then
                If Not IsNull(rs2!fechaNac) Then
                    fechaNacimiento = Year(rs2!fechaNac) & CerosIzq(Month(rs2!fechaNac), 2) & CerosIzq(Day(rs2!fechaNac), 2)
                Else
                    fechaNacimiento = ""
                End If
                
                If existCampoSexo Then
                    If rs2!sexo = "V" Then
                        sexo = "Hombre"
                    Else
                        If rs2!sexo = "M" Then
                            sexo = "Mujer"
                        Else
                            If IsNull(rs2!sexo) Then
                                sexo = ""
                            End If
                        End If
                    End If
                End If
            Else
                fechaNacimiento = ""
            End If
            
            rs2.Close
            
            If IsNull(rs!FIS_NIF) Then
                baja = 0
            Else
                If Trim(rs!FIS_NIF) = "" Or Trim(rs!FIS_NIF) = "No" Or Trim(rs!FIS_NIF) = "N" Then
                    baja = 0
                Else
                    baja = 1
                End If
            End If
            
            If IsNull(rs!TIPOTARIFA) Then
                lopd = 0
            Else
                If Trim(rs!TIPOTARIFA) = "No" Or Trim(rs!XCLIE_IDCLIENTEFACT) = "Si" Then
                    lopd = 0
                Else
                    lopd = 1
                End If
            End If
            
            If sexo = "" Then
                If IsNull(rs!fis_nombre) Then
                    sexo = ""
                Else
                    sexo = rs!fis_nombre
                End If
            End If
            
            If IsNull(rs!FIS_PROVINCIA) Then
                fechaAlta = "NULL"
            Else
                If InStr(rs!FIS_PROVINCIA, ":") > 0 Then
                    fechaAux = Left(rs!FIS_PROVINCIA, (InStr(rs!FIS_PROVINCIA, ":") - 4))
                Else
                    If IsDate(rs!FIS_PROVINCIA) Then
                        fechaAux = rs!FIS_PROVINCIA
                    Else
                        fechaAux = ""
                    End If
                End If
                
                If fechaAux <> "" Then
                    fechaAlta = Format(fechaAux, "yyyy-MM-dd HH:mm:ss")
                Else
                    fechaAlta = "NULL"
                End If
            End If
            
            If IsNull(rs!FIS_FAX) Then
                tarjeta = ""
            Else
                tarjeta = rs!FIS_FAX
            End If
            
            If IsNull(rs!PER_TELEFONO) Then
                If IsNull(rs!FIS_TELEFONO) Then
                    telefono = ""
                Else
                    telefono = Trim(rs!FIS_TELEFONO)
                End If
            Else
                telefono = Trim(rs!PER_TELEFONO)
            End If
            
            If IsNull(rs!PER_DIRECCION) Then
                direccion = ""
            Else
                If Trim(rs!PER_DIRECCION) = "" Then
                    direccion = ""
                Else
                    direccion = Trim(rs!PER_DIRECCION)
                    
                    If Not IsNull(rs!PER_CODPOSTAL) And Trim(rs!PER_CODPOSTAL) <> "" Then
                        direccion = direccion & " - " & Trim(rs!PER_CODPOSTAL)
                    End If
                    
                    If Not IsNull(rs!PER_POBLACION) And Trim(rs!PER_POBLACION) <> "" Then
                        direccion = direccion & " - " & Trim(rs!PER_POBLACION)
                    End If
                    
                    If Not IsNull(rs!PER_PROVINCIA) And Trim(rs!PER_PROVINCIA) <> "" Then
                        direccion = direccion & " (" & Trim(rs!PER_PROVINCIA) & ")"
                    End If
                End If
            End If
            
            If IsNull(rs!PER_NOMBRE) Then
                nombre = ""
            Else
                If Trim(rs!PER_NOMBRE) = "" Then
                    nombre = ""
                Else
                    nombre = StripString(Trim(rs!PER_NOMBRE))
                End If
            End If
            
            id_vendedor = rs!XVend_IdVendedor
                
            On Error GoTo errorTimerClientes
            
            sql = "SELECT * FROM vendedor WHERE IdVendedor='" & id_vendedor & "'"

            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            If rs2.EOF Then
                trabajador = ""
            Else
                trabajador = rs2!nombre
            End If
            
            rs2.Close
            
            On Error GoTo errorTimerClientes
            
            sql = "SELECT ISNULL(SUM(cantidad), 0) AS puntos FROM HistoOferta " & _
                    "WHERE IdCliente = '" & rs!idcliente & "' AND TipoAcumulacion = 'P'"
                    
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs2.EOF Then
                puntos = 0
            Else
                puntos = Replace(rs2!puntos, ",", ".")
            End If
            
            rs2.Close
            
            If rs3.EOF Then
            
               On Error GoTo errorTimerClientes
            
               tipo = "cliente"
               
               sql = "UPDATE clientes SET dni_tra = 0"
               
               connMySql.Execute (sql)
                                          
               sql = "INSERT INTO clientes " & "(dni_tra,nombre_tra,tarjeta,dni,apellidos,telefono,direccion,movil,email,puntos,fecha_nacimiento,sexo,tipo,fechaAlta,baja,lopd) VALUES('1', '" & _
                                                    trabajador & "','" & _
                                                    tarjeta & "','" & _
                                                    rs!idcliente & "','" & _
                                                    StripString(nombre) & "','" & _
                                                    telefono & "','" & _
                                                    StripString(direccion) & "','" & _
                                                    movil & "','" & _
                                                    email & "','" & _
                                                    puntos & "','" & _
                                                    fechaNacimiento & "','" & _
                                                    sexo & "','" & _
                                                    tipo & "','" & _
                                                    fechaAlta & "','" & _
                                                    baja & "','" & _
                                                    lopd & "')"
                
                connMySql.Execute (sql)
            Else
                On Error GoTo errorTimerClientes
                
                sql = "UPDATE clientes SET dni_tra = 0"
               
                connMySql.Execute (sql)
               
                sql = "UPDATE clientes SET dni_tra = '1', nombre_tra = '" & trabajador & "', tarjeta = '" & tarjeta & "', apellidos = '" & StripString(nombre) & "', telefono = '" & telefono & "', " & _
                        "direccion = '" & StripString(direccion) & "', movil = '" & movil & "', email = '" & email & "', puntos = '" & puntos & "', " & _
                        "fecha_nacimiento = '" & fechaNacimiento & "', sexo = '" & sexo & "', fechaAlta = '" & fechaAlta & "', baja = '" & baja & "', lopd = '" & lopd & "' " & _
                        "WHERE dni = '" & rs!idcliente & "'"
                    
                connMySql.Execute (sql)
            End If
               
            On Error GoTo errorTimerClientes
            
            If CDbl(rs!idcliente) <> contadorHuecos Then
                For i = contadorHuecos To CDbl(rs!idcliente) - 1
                    On Error GoTo errorTimerClientes
                    
                    sql = "SELECT * FROM clientes_huecos WHERE hueco = " & i
                    
                    rs4.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs4.EOF Then
                        On Error GoTo errorTimerClientes
                        
                        sql = "INSERT INTO clientes_huecos (hueco) VALUES ('" & i & "')"
                        
                        connMySql.Execute (sql)
                    End If
                    
                    rs4.Close
                Next
                
                contadorHuecos = CDbl(rs!idcliente)
            End If
            
            contadorHuecos = contadorHuecos + 1
            
            rs.MoveNext
                    
            rs3.Close
        Loop
         
        rs.Close
    'End If
    
final:
    GoTo fin
    
errorTimerClientes:
    Sleep 1500
    
fin:
    procesarTimerClientesHuecos
    
End Sub

Private Sub Timer_Clientes_Huecos_Timer() ''''''''''' Migrado '''''''''''''''''''
    Dim telefono As String
    Dim movil As String
    Dim email As String
    Dim direccion As String
    Dim nombre As String
    Dim tarjeta As String
    
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Dim existCampoSexo As Boolean
    '''''''''''' EXISTE SEXO '''''''''''''''''''''
    existCampoSexo = False
    
    On Error GoTo errorTimerClientesHuecos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorTimerClientesHuecos
    
    sql = "SELECT TOP 1 * FROM ClienteAux"
        
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If LCase(oField.Name) = "sexo" Then
            existCampoSexo = True
        End If
    Next
    
    rs.Close
    
    On Error GoTo errorTimerClientesHuecos
	'''''''''''' huecos en orden ASC'''''''''''''''
    sql = "SELECT hueco FROM clientes_huecos ORDER BY hueco ASC"
    
    rs4.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    Do While Not rs4.EOF
        DoEvents
        
        On Error GoTo errorTimerClientesHuecos
    
        sql = "SELECT * FROM cliente WHERE Idcliente = " & rs4!hueco
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs.EOF Then
            On Error GoTo errorTimerClientesHuecos
            
            sql = "SELECT * FROM clientes WHERE dni='" & rs!idcliente & "'"
            
            rs3.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
            
            On Error GoTo errorTimerClientesHuecos
            
            sql = "SELECT * FROM Destinatario WHERE fk_Cliente_1 ='" & rs!idcliente & "'"
            
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            ''''''''''''''' SETEAMOS MOVIL Y EMAIL ''''''''''''''''''''''''''''
			If Not rs2.EOF Then
                If IsNull(rs2!TlfMovil) Then
                    movil = ""
                Else
                 movil = Trim(rs2!TlfMovil)
                End If
                
                If IsNull(rs2!email) Then
                    email = ""
                Else
                 email = Trim(rs2!email)
                End If
            Else
                 movil = ""
                 email = ""
            End If
            
            rs2.Close
            
            On Error GoTo errorTimerClientesHuecos
            
            sql = "SELECT * FROM ClienteAux WHERE idCliente = '" & rs!idcliente & "'"
    
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            ''''''''''' SETEAMOS FECHA NACIMIENTO Y SEXO
			If Not rs2.EOF Then
                If Not IsNull(rs2!fechaNac) Then
                    fechaNacimiento = Year(rs2!fechaNac) & CerosIzq(Month(rs2!fechaNac), 2) & CerosIzq(Day(rs2!fechaNac), 2)
                Else
                    fechaNacimiento = ""
                End If
                
                If existCampoSexo Then
                    If rs2!sexo = "V" Then
                        sexo = "Hombre"
                    Else
                        If rs2!sexo = "M" Then
                            sexo = "Mujer"
                        Else
                            If IsNull(rs2!sexo) Then
                                sexo = ""
                            End If
                        End If
                    End If
                End If
            Else
                fechaNacimiento = ""
            End If
            
            rs2.Close
            
            '''''''''''''' SETEAMOS BAJA ''''''''''''''''''
			If IsNull(rs!FIS_NIF) Then
                baja = 0
            Else
                If Trim(rs!FIS_NIF) = "" Or Trim(rs!FIS_NIF) = "No" Or Trim(rs!FIS_NIF) = "N" Then
                    baja = 0
                Else
                    baja = 1
                End If
            End If
            
            If IsNull(rs!TIPOTARIFA) Then
                lopd = 0
            Else
                If Trim(rs!TIPOTARIFA) = "No" Or Trim(rs!XCLIE_IDCLIENTEFACT) = "Si" Then
                    lopd = 0
                Else
                    lopd = 1
                End If
            End If
            
            If sexo = "" Then
                If IsNull(rs!fis_nombre) Then
                    sexo = ""
                Else
                    sexo = rs!fis_nombre
                End If
            End If
            
            '''''''' SETEAMOS FECHA_ALATA ''''''''''''''''
			If IsNull(rs!FIS_PROVINCIA) Then
                fechaAlta = "NULL"
            Else
                If InStr(rs!FIS_PROVINCIA, ":") > 0 Then
                    fechaAux = Left(rs!FIS_PROVINCIA, (InStr(rs!FIS_PROVINCIA, ":") - 4))
                Else
                    If IsDate(rs!FIS_PROVINCIA) Then
                        fechaAux = rs!FIS_PROVINCIA
                    Else
                        fechaAux = ""
                    End If
                End If
                
                If fechaAux <> "" Then
                    fechaAlta = Format(fechaAux, "yyyy-MM-dd HH:mm:ss")
                Else
                    fechaAlta = "NULL"
                End If
            End If
            
            
			'''' SETEAMOS _ TARJETA
			If IsNull(rs!FIS_FAX) Then
                tarjeta = ""
            Else
                tarjeta = rs!FIS_FAX
            End If
            
            '''' SETEAMOS TELEFONO
			If IsNull(rs!PER_TELEFONO) Then
                If IsNull(rs!FIS_TELEFONO) Then
                    telefono = ""
                Else
                    telefono = Trim(rs!FIS_TELEFONO)
                End If
            Else
                telefono = Trim(rs!PER_TELEFONO)
            End If
            
            If IsNull(rs!PER_DIRECCION) Then
                direccion = ""
            Else
                If Trim(rs!PER_DIRECCION) = "" Then
                    direccion = ""
                Else
                    direccion = Trim(rs!PER_DIRECCION)
                    
                    If Not IsNull(rs!PER_CODPOSTAL) And Trim(rs!PER_CODPOSTAL) <> "" Then
                        direccion = direccion & " - " & Trim(rs!PER_CODPOSTAL)
                    End If
                    
                    If Not IsNull(rs!PER_POBLACION) And Trim(rs!PER_POBLACION) <> "" Then
                        direccion = direccion & " - " & Trim(rs!PER_POBLACION)
                    End If
                    
                    If Not IsNull(rs!PER_PROVINCIA) And Trim(rs!PER_PROVINCIA) <> "" Then
                        direccion = direccion & " (" & Trim(rs!PER_PROVINCIA) & ")"
                    End If
                End If
            End If
            
            If IsNull(rs!PER_NOMBRE) Then
                nombre = ""
            Else
                If Trim(rs!PER_NOMBRE) = "" Then
                    nombre = ""
                Else
                    nombre = StripString(Trim(rs!PER_NOMBRE))
                End If
            End If
            
            id_vendedor = rs!XVend_IdVendedor
                
            On Error GoTo errorTimerClientesHuecos
            
            sql = "SELECT * FROM vendedor WHERE IdVendedor='" & id_vendedor & "'"

            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            If rs2.EOF Then
                trabajador = ""
            Else
                trabajador = rs2!nombre
            End If
            
            rs2.Close
            
            On Error GoTo errorTimerClientesHuecos
            
            sql = "SELECT ISNULL(SUM(cantidad), 0) AS puntos FROM HistoOferta " & _
                    "WHERE IdCliente = '" & rs!idcliente & "' AND TipoAcumulacion = 'P'"
                    
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs2.EOF Then
                puntos = 0
            Else
                puntos = Replace(rs2!puntos, ",", ".")
            End If
            
            rs2.Close
            
            If rs3.EOF Then
            
               On Error GoTo errorTimerClientesHuecos
            
               tipo = "cliente"
               
               sql = "INSERT INTO clientes " & "(nombre_tra,tarjeta,dni,apellidos,telefono,direccion,movil,email,puntos,fecha_nacimiento,sexo,tipo,fechaAlta,baja,lopd) VALUES('" & _
                                                    trabajador & "','" & _
                                                    tarjeta & "','" & _
                                                    rs!idcliente & "','" & _
                                                    StripString(nombre) & "','" & _
                                                    telefono & "','" & _
                                                    StripString(direccion) & "','" & _
                                                    movil & "','" & _
                                                    email & "','" & _
                                                    puntos & "','" & _
                                                    fechaNacimiento & "','" & _
                                                    sexo & "','" & _
                                                    tipo & "','" & _
                                                    fechaAlta & "','" & _
                                                    baja & "','" & _
                                                    lopd & "')"
                
                connMySql.Execute (sql)
            Else
                On Error GoTo errorTimerClientesHuecos
                
                ssql = "UPDATE clientes SET nombre_tra = '" & trabajador & "', tarjeta = '" & tarjeta & "', apellidos = '" & StripString(nombre) & "', telefono = '" & telefono & "', " & _
                        "direccion = '" & StripString(direccion) & "', movil = '" & movil & "', email = '" & email & "', puntos = '" & puntos & "', " & _
                        "fecha_nacimiento = '" & fechaNacimiento & "', sexo = '" & sexo & "', fechaAlta = '" & fechaAlta & "', baja = '" & baja & "', lopd = '" & lopd & "' " & _
                        "WHERE dni = '" & rs!idcliente & "'"
                    
                connMySql.Execute (sql)
            End If
            
            rs3.Close
            
            On Error GoTo errorTimerClientesHuecos
                
            connMySql.Execute ("DELETE FROM clientes_huecos WHERE hueco = '" & rs!idcliente & "'")
        End If
        
        rs4.MoveNext
        
        rs.Close
    Loop
     
    rs4.Close
    
final:
    GoTo fin
    
errorTimerClientesHuecos:
    Sleep 1500
    
fin:
        
End Sub

Private Sub Timer_Pendiente_Puntos_Timer() ''''''''''''''''''''' MIGRADO *******************************
    Dim puesto As String
    Dim numero As String
    Dim cargado As String
    Dim dni As String
    Dim numero_max As Double
    Dim familia As String
    Dim receta As String
    Dim codLaboratorio As String
    Dim nombreLaboratorio As String
    Dim fechaVenta As String
    Dim pcoste As String
    Dim precioMed As String
    Dim superfamilia As String
    Dim descuentoVenta As Boolean
    Dim dtoLinea As String
    Dim dtoVenta As String
    
    Dim sql As String
    
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim FieldExistsInRS1 As Boolean
    Dim FieldExistsInRS2 As Boolean
    Dim FieldExistsInRS3 As Boolean
    Dim existCampoSexo As Boolean
    Dim oField
    
    FieldExistsInRS = False
    FieldExistsInRS1 = False
    FieldExistsInRS2 = False
    FieldExistsInRS3 = False
    existCampoSexo = False
    
    On Error GoTo errorPendientePuntos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorPendientePuntos
        
    sql = "SELECT * from pendiente_puntos LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "dtoLinea" Then
            FieldExistsInRS = True
        End If
        If oField.Name = "proveedor" Then
            FieldExistsInRS1 = True
        End If
        If oField.Name = "redencion" Then
            FieldExistsInRS2 = True
        End If
        If oField.Name = "recetaPendiente" Then
            FieldExistsInRS3 = True
        End If
    Next
    
    rs.Close
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE pendiente_puntos ADD dtoLinea FLOAT DEFAULT 0;"
        
        connMySql.Execute (sql)
        
        sql = "ALTER TABLE pendiente_puntos ADD dtoVenta FLOAT DEFAULT 0;"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS1 = False Then
        sql = "ALTER TABLE pendiente_puntos ADD proveedor VARCHAR(255) DEFAULT NULL AFTER laboratorio;"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS2 = False Then
        sql = "ALTER TABLE pendiente_puntos ADD redencion FLOAT DEFAULT NULL;"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS3 = False Then
        sql = "ALTER TABLE pendiente_puntos ADD recetaPendiente CHAR(2) DEFAULT NULL;"
        
        connMySql.Execute (sql)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   
    On Error GoTo errorPendientePuntos
    
    sql = "SELECT TABLE_NAME AS tipo From information_schema.TABLES " & _
          "WHERE TABLE_SCHEMA = '" & baseRemoto & "' AND TABLE_NAME = 'entregas_clientes'"
    
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        On Error GoTo errorPendientePuntos
            
        sql = "CREATE TABLE IF NOT EXISTS `entregas_clientes` (" & _
              "`cod` bigint(255) NOT NULL AUTO_INCREMENT," & _
              "`idventa` bigint(255) NOT NULL," & _
              "`idnlinea` bigint(255) NOT NULL," & _
              "`codigo` varchar(255) NOT NULL," & _
              "`descripcion` varchar(255) NOT NULL," & _
              "`cantidad` int(255) NOT NULL," & _
              "`precio` decimal(20,2) NOT NULL," & _
              "`tipo` char(2) DEFAULT NULL," & _
              "`fecha` int(255) NOT NULL," & _
              "`dni` varchar(255) NOT NULL," & _
              "`hora` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP," & _
              "`puesto` varchar(100) NOT NULL," & _
              "`trabajador` varchar(100) NOT NULL," & _
              "`fechaEntrega` datetime DEFAULT NULL," & _
              "`pvp` float DEFAULT NULL," & _
              "PRIMARY KEY (`cod`)," & _
              "UNIQUE KEY `unico` (`idventa`,`idnlinea`)," & _
              "KEY `tx_codigo` (`codigo`)," & _
              "KEY `tx_fecha` (`fecha`)," & _
              "KEY `tx_fecha_entrega` (`fechaEntrega`)," & _
              "KEY `tx_venta` (`idventa`,`idnlinea`)" & _
            ") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
                    
        connMySql.Execute (sql)
    End If
    
    rs.Close
    '#####################################################################################
    On Error GoTo errorPendientePuntos
    
    sql = "SELECT TOP 1 * FROM ClienteAux"
        
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If LCase(oField.Name) = "sexo" Then
            existCampoSexo = True
        End If
    Next
    
    rs.Close
    '##################################################################################
    On Error GoTo errorPendientePuntos
    
    sql = "SELECT data_type AS tipo " & _
          "From information_schema.Columns " & _
          "WHERE TABLE_SCHEMA = '" & baseRemoto & "' AND TABLE_NAME = 'pendiente_puntos' AND COLUMN_NAME = 'tipoPago'"
          
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If Not rs.EOF Then
        If LCase(rs!tipo) <> "char" Then
            sql = "ALTER TABLE pendiente_puntos MODIFY COLUMN tipoPago CHAR(2);"
            
            connMySql.Execute (sql)
        End If
    Else
        sql = "ALTER TABLE pendiente_puntos ADD tipoPago CHAR(2) DEFAULT NULL AFTER precio;"
        
        connMySql.Execute (sql)
    End If
    
    rs.Close
    '################################################################################
    On Error GoTo errorPendientePuntos
    
    sql = "SELECT * FROM pendiente_puntos ORDER BY idventa DESC LIMIT 0,1"
       
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

    If rs.EOF Then
        venta = 1
    Else
        venta = rs!IdVenta
    End If

    rs.Close
	'###################################################################################
    On Error GoTo errorPendientePuntos
        
    sql = "SELECT * FROM venta WHERE ejercicio >= 2015 AND IdVenta >= " & venta & " ORDER BY IdVenta ASC"
    
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents
      
        puesto = rs!Maquina
        
        tipoPago = rs!TipoVenta
    
        fechaVenta = Format(rs!FechaHora, "yyyy-MM-dd HH:mm:ss")
        
        dni = LTrim(RTrim(StripString(rs!XClie_IdCliente)))
                
        'If (dni = "") Then
        '    dni = 0
        'End If
        
        If IsNull(dni) Then
            dni = ""
        End If
        
        If Len(dni) > 0 Then
            On Error GoTo errorPendientePuntos
            
            sql = "SELECT * FROM cliente WHERE Idcliente = '" & dni & "'"
            
            rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                 
            If Not rs5.EOF Then
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT * FROM clientes WHERE dni='" & dni & "'"
                
                rs3.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
                
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT * FROM Destinatario WHERE fk_Cliente_1 ='" & dni & "'"
                
                rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If Not rs2.EOF Then
                    If IsNull(rs2!TlfMovil) Then
                        movil = ""
                    Else
                     movil = Trim(rs2!TlfMovil)
                    End If
                    
                    If IsNull(rs2!email) Then
                        email = ""
                    Else
                     email = Trim(rs2!email)
                    End If
                Else
                     movil = ""
                     email = ""
                End If
                
                rs2.Close
                
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT * FROM ClienteAux WHERE idCliente = '" & dni & "'"
        
                rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If Not rs2.EOF Then
                    If Not IsNull(rs2!fechaNac) Then
                        fechaNacimiento = Year(rs2!fechaNac) & CerosIzq(Month(rs2!fechaNac), 2) & CerosIzq(Day(rs2!fechaNac), 2)
                    Else
                        fechaNacimiento = ""
                    End If
                    
                    If existCampoSexo Then
                        If rs2!sexo = "V" Then
                            sexo = "Hombre"
                        Else
                            If rs2!sexo = "M" Then
                                sexo = "Mujer"
                            Else
                                If IsNull(rs2!sexo) Then
                                    sexo = ""
                                End If
                            End If
                        End If
                    End If
                Else
                    fechaNacimiento = ""
                End If
                
                rs2.Close
                
                If IsNull(rs5!FIS_NIF) Then
                    baja = 0
                Else
                    If Trim(rs5!FIS_NIF) = "" Or Trim(rs5!FIS_NIF) = "No" Or Trim(rs5!FIS_NIF) = "N" Then
                        baja = 0
                    Else
                        baja = 1
                    End If
                End If
                
                If IsNull(rs5!TIPOTARIFA) Then
                    lopd = 0
                Else
                    If Trim(rs5!TIPOTARIFA) = "No" Or Trim(rs5!XCLIE_IDCLIENTEFACT) = "Si" Then
                        lopd = 0
                    Else
                        lopd = 1
                    End If
                End If
                
                If sexo = "" Then
                    If IsNull(rs5!fis_nombre) Then
                        sexo = ""
                    Else
                        sexo = rs5!fis_nombre
                    End If
                End If
                
                If IsNull(rs5!FIS_PROVINCIA) Then
                    fechaAlta = "NULL"
                Else
                    If InStr(rs5!FIS_PROVINCIA, ":") > 0 Then
                        fechaAux = Left(rs5!FIS_PROVINCIA, (InStr(rs5!FIS_PROVINCIA, ":") - 4))
                    Else
                        If IsDate(rs5!FIS_PROVINCIA) Then
                            fechaAux = rs5!FIS_PROVINCIA
                        Else
                            fechaAux = ""
                        End If
                    End If
                    
                    If fechaAux <> "" Then
                        fechaAlta = Format(fechaAux, "yyyy-MM-dd HH:mm:ss")
                    Else
                        fechaAlta = "NULL"
                    End If
                End If
                
                If IsNull(rs5!FIS_FAX) Then
                    tarjeta = ""
                Else
                    tarjeta = rs5!FIS_FAX
                End If
                
                If IsNull(rs5!PER_TELEFONO) Then
                    If IsNull(rs5!FIS_TELEFONO) Then
                        telefono = ""
                    Else
                        telefono = Trim(rs5!FIS_TELEFONO)
                    End If
                Else
                    telefono = Trim(rs5!PER_TELEFONO)
                End If
                
                If IsNull(rs5!PER_DIRECCION) Then
                    direccion = ""
                Else
                    If Trim(rs5!PER_DIRECCION) = "" Then
                        direccion = ""
                    Else
                        direccion = Trim(rs5!PER_DIRECCION)
                        
                        If Not IsNull(rs5!PER_CODPOSTAL) And Trim(rs5!PER_CODPOSTAL) <> "" Then
                            direccion = direccion & " - " & Trim(rs5!PER_CODPOSTAL)
                        End If
                        
                        If Not IsNull(rs5!PER_POBLACION) And Trim(rs5!PER_POBLACION) <> "" Then
                            direccion = direccion & " - " & Trim(rs5!PER_POBLACION)
                        End If
                        
                        If Not IsNull(rs5!PER_PROVINCIA) And Trim(rs5!PER_PROVINCIA) <> "" Then
                            direccion = direccion & " (" & Trim(rs5!PER_PROVINCIA) & ")"
                        End If
                    End If
                End If
                
                If IsNull(rs5!PER_NOMBRE) Then
                    nombre = ""
                Else
                    If Trim(rs5!PER_NOMBRE) = "" Then
                        nombre = ""
                    Else
                        nombre = StripString(Trim(rs5!PER_NOMBRE))
                    End If
                End If
                
                id_vendedor = rs5!XVend_IdVendedor
                    
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT * FROM vendedor WHERE IdVendedor='" & id_vendedor & "'"
    
                rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                If rs2.EOF Then
                    trabajador = ""
                Else
                    trabajador = rs2!nombre
                End If
                
                rs2.Close
                
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT ISNULL(SUM(cantidad), 0) AS puntos FROM HistoOferta " & _
                        "WHERE IdCliente = '" & dni & "' AND TipoAcumulacion = 'P'"
                        
                rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs2.EOF Then
                    puntos = 0
                Else
                    puntos = Replace(rs2!puntos, ",", ".")
                End If
                
                rs2.Close
                
                If rs3.EOF Then
                
                   On Error GoTo errorPendientePuntos
                
                   tipo = "cliente"
                   
                   sql = "INSERT INTO clientes " & "(nombre_tra,tarjeta,dni,apellidos,telefono,direccion,movil,email,puntos,fecha_nacimiento,sexo,tipo,fechaAlta,baja,lopd) VALUES('" & _
                                                        trabajador & "','" & _
                                                        tarjeta & "','" & _
                                                        dni & "','" & _
                                                        StripString(nombre) & "','" & _
                                                        telefono & "','" & _
                                                        StripString(direccion) & "','" & _
                                                        movil & "','" & _
                                                        email & "','" & _
                                                        puntos & "','" & _
                                                        fechaNacimiento & "','" & _
                                                        sexo & "','" & _
                                                        tipo & "','" & _
                                                        fechaAlta & "','" & _
                                                        baja & "','" & _
                                                        lopd & "')"
                    
                    connMySql.Execute (sql)
                Else
                    On Error GoTo errorPendientePuntos
                    
                    sql = "UPDATE clientes SET nombre_tra = '" & trabajador & "', tarjeta = '" & tarjeta & "', apellidos = '" & StripString(nombre) & "', telefono = '" & telefono & "', " & _
                            "direccion = '" & StripString(direccion) & "', movil = '" & movil & "', email = '" & email & "', puntos = '" & puntos & "', " & _
                            "fecha_nacimiento = '" & fechaNacimiento & "', sexo = '" & sexo & "', fechaAlta = '" & fechaAlta & "', baja = '" & baja & "', lopd = '" & lopd & "' " & _
                            "WHERE dni = '" & dni & "'"
                        
                    connMySql.Execute (sql)
                End If
                
                rs3.Close
            End If
            
            rs5.Close
            
        Else
            dni = 0 '###### no se usa
        End If
        '#######################################################################################################
        id_vendedor = rs!XVend_IdVendedor
            
        On Error GoTo errorPendientePuntos
        
        sql = "SELECT * FROM vendedor WHERE IdVendedor='" & id_vendedor & "'"

        rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs5.EOF Then
            trabajador = "NO"
        Else
            trabajador = rs5!nombre
        End If
            
        rs5.Close
        
        On Error GoTo errorPendientePuntos
        
        sql = "SELECT * FROM lineaventa WHERE IdVenta='" & rs!IdVenta & "'"

        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
        ArrayFecha = Split(rs!FechaHora, "/")

        numero_dia = ArrayFecha(0)
        numero_mes = ArrayFecha(1)
        numero_anio = ArrayFecha(2)
        
        numero_Fecha = Split(numero_anio, " ")
        numero_anio = numero_Fecha(0)

        'Fecha = (numero_anio * 10000) + (numero_mes * 100) + numero_dia
        Fecha = Format(rs!FechaHora, "yyyyMMdd")
        
        descuentoVenta = False
    
        On Error GoTo errorPendientePuntos
        
        Do While Not rs2.EOF
        
            recetaPendiente = rs2!recetaPendiente
    
            receta = rs2!TipoAportacion
            
            precioMed = Replace(rs2!pvp, ",", ".")
            
            dtoVenta = 0
            If Not descuentoVenta Then
                dtoVenta = Replace(rs!DescuentoOpera, ",", ".")
                descuentoVenta = True
            End If
            
            dtoLinea = Replace(rs2!DescuentoLinea, ",", ".")
            
            On Error GoTo errorPendientePuntos
            '####################################################################################################
            sql = "SELECT * FROM LineaVentaReden WHERE IdVenta = " & rs2!IdVenta & " AND IdNLinea = " & rs2!idnlinea

            rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            redencion = 0
            If Not rs3.EOF Then
                redencion = Replace(rs3!redencion, ",", ".")
            End If
            
            rs3.Close
			'##############################################################
            On Error GoTo errorPendientePuntos
            
            sql = "SELECT * FROM articu WHERE IdArticu='" & rs2!codigo & "'"

            rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            On Error GoTo errorPendientePuntos
            
            sql = "SELECT * FROM sinonimo WHERE IdArticu='" & rs2!codigo & "'"

            rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            If rs4.EOF Then
                cod_barras = "847000" + CerosIzq(rs2!codigo, 6)
            Else
                cod_barras = rs4!Sinonimo
            End If
            
            rs4.Close
    
            If rs3.EOF Then
                familia = ""
                codLaboratorio = ""
                pcoste = 0
                nombreLaboratorio = "<Sin Laboratorio>"
                proveedor = ""
            Else
                pcoste = Replace(rs3!puc, ",", ".")
                
                On Error GoTo errorPendientePuntos
                
                sql = "SELECT * FROM familia WHERE IdFamilia='" & rs3!XFam_IdFamilia & "'"

                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
         
                If rs4.EOF Then
                    familia = ""
                Else
                    familia = rs4!descripcion
                End If
    
                rs4.Close
      
                If IsNull(familia) Then
                    familia = "<Sin Clasificar>"
                Else
                    If Len(familia) = 0 Then
                        familia = "<Sin Clasificar>"
                    End If
                End If
            
                If IsNull(rs3!laboratorio) Then
                    codLaboratorio = ""
                Else
                    codLaboratorio = rs3!laboratorio
                End If
        
                If familia = "<Sin Clasificar>" Then
                    superfamilia = "<Sin Clasificar>"
                Else
                    On Error GoTo errorPendientePuntos
                    
                    sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                            " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(familia) & "'"
                    
                    rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs4.EOF Then
                        superfamilia = "<Sin Clasificar>"
                    Else
                        superfamilia = rs4!descripcion
                    End If
                    
                    rs4.Close
                End If
                
                On Error GoTo errorPendientePuntos
                '##################################################################################################
                sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs3!proveedorHabitual & "'"
               
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                If rs4.EOF Then
                    proveedor = ""
                Else
                    proveedor = rs4!fis_nombre
                End If
        
                rs4.Close
                
                If Len(Trim(codLaboratorio)) > 0 Then
                    On Error GoTo errorPendientePuntos
                    
                    sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(codLaboratorio) & "'"
                    
                    rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                        
                    If rs4.EOF Then
                        rs4.Close
                
                        On Error GoTo errorPendientePuntos
                    
                        sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(codLaboratorio) & "'"
                        
                        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                        
                        If rs4.EOF Then
                            nombreLaboratorio = "<Sin Laboratorio>"
                        Else
                            nombreLaboratorio = rs4!nombre
                        End If
                    Else
                        nombreLaboratorio = rs4!nombre
                    End If
                    
                    rs4.Close
                Else
                    nombreLaboratorio = "<Sin Laboratorio>"
                End If
            End If
			'################################################################################################################
            On Error GoTo errorPendientePuntos
            
            sql = "SELECT * FROM pendiente_puntos WHERE IdVenta='" & rs2!IdVenta & "' AND Idnlinea= '" & rs2!idnlinea & "'"
       
            rs4.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
            numero = Replace(rs2!importeneto, ",", ".")
                
            cargado = "no"
                
            If rs4.EOF Then
                On Error GoTo errorPendientePuntos
                
                sql = "INSERT INTO pendiente_puntos " & "(idventa,idnlinea,cod_barras,cod_nacional,descripcion,familia,cantidad,precio,tipoPago,fecha,dni,cargado,puesto,trabajador,cod_laboratorio,laboratorio,proveedor,receta,fechaVenta,superFamilia, pvp, puc, dtoLinea, dtoVenta, redencion, recetaPendiente) VALUES('" & _
                                 rs2!IdVenta & "','" & _
                                 rs2!idnlinea & "','" & _
                                 StripString(cod_barras) & "','" & _
                                 StripString(rs2!codigo) & "','" & _
                                 LTrim(RTrim(StripString(rs2!descripcion))) & "','" & _
                                 LTrim(RTrim(StripString(familia))) & "','" & _
                                 rs2!cantidad & "','" & _
                                 numero & "','" & _
                                 tipoPago & "','" & _
                                 Fecha & "','" & _
                                 dni & "','" & _
                                 cargado & "','" & _
                                 puesto & "','" & _
                                 trabajador & "','" & _
                                 LTrim(RTrim(StripString(codLaboratorio))) & "','" & _
                                 LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                                 LTrim(RTrim(StripString(proveedor))) & "','" & _
                                 receta & "','" & _
                                 fechaVenta & "','" & _
                                 LTrim(RTrim(StripString(superfamilia))) & "','" & _
                                 precioMed & "','" & _
                                 pcoste & "','" & _
                                 dtoLinea & "','" & dtoVenta & "','" & redencion & "', '" & recetaPendiente & "')"
                                 
                connMySql.Execute (sql)
            End If
        
            If (rs!IdVenta > venta) Then
                venta = rs!IdVenta - 3
            End If
        
            rs4.Close
        
            rs2.MoveNext
     
            rs3.Close
        Loop
   
        rs2.Close
        
        On Error GoTo errorPendientePuntos
        
        sql = "SELECT * FROM lineaventavirtual WHERE IdVenta='" & rs!IdVenta & "' AND (codigo = 'Pago' OR codigo = 'A Cuenta')"

        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorPendientePuntos
        
        Do While Not rs2.EOF
            On Error GoTo errorPendientePuntos
        
            sql = "SELECT * FROM entregas_clientes WHERE IdVenta='" & rs2!IdVenta & "' AND Idnlinea= '" & rs2!idnlinea & "'"
       
            rs4.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
            numero = Replace(rs2!importeneto, ",", ".")
            
            pvp = Replace(rs2!pvp, ",", ".")
                
            If rs4.EOF Then
                On Error GoTo errorPendientePuntos
                
                sql = "INSERT INTO entregas_clientes " & "(idventa,idnlinea,codigo,descripcion,cantidad,precio,tipo,fecha,dni,puesto,trabajador,fechaEntrega,pvp) VALUES('" & _
                                 rs2!IdVenta & "','" & _
                                 rs2!idnlinea & "','" & _
                                 StripString(rs2!codigo) & "','" & _
                                 StripString(rs2!descripcion) & "','" & _
                                 rs2!cantidad & "','" & _
                                 numero & "','" & _
                                 rs2!tipoLinea & "','" & _
                                 Fecha & "','" & _
                                 dni & "','" & _
                                 puesto & "','" & _
                                 trabajador & "','" & _
                                 fechaVenta & "','" & _
                                 pvp & "')"
                                 
                connMySql.Execute (sql)
            End If
        
            rs4.Close
        
            rs2.MoveNext
        Loop
        
        rs.MoveNext
        
        rs2.Close
    Loop
 
    rs.Close
 
final:
    GoTo fin
 
errorPendientePuntos:
    Sleep 1500
 
fin:

End Sub

Private Sub Timer_Control_Stock_Fechas_Entrada_Timer() ''''''''''''''''''''''' MIGRADO ''''''''''''''''''''''''''''''''''''
    Dim precio As String
    Dim pcoste As String
    Dim pvpsiva As String
    Dim stock As String
    Dim stockMinimo As String
    Dim stockMaximo As String
    Dim desc As String
    Dim laboratorio As String
    Dim nombreLaboratorio As String
    Dim descripcion As String
    Dim present As String
    Dim activo As String

    Dim sql As String

    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    On Error GoTo errorControlStock
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorControlStock
    
    sql = "SELECT * FROM configuracion WHERE campo = 'fechaActualizacionStockEntrada'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStock
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('fechaActualizacionStockEntrada', NULL)"
        
        connMySql.Execute (sql)
    Else
        If IsNull(rs!valor) Or Trim(rs!valor) = "" Then
            fechaActualizacionStock = DateAdd("d", -7, Date)
            
            fechaActualizacionStock = Format(fechaActualizacionStock, "yyyy-dd-MM")
        Else
            resultado = Abs(DateDiff("d", Now, Format(rs!valor, "yyyy-MM-dd")))
            
            If resultado > 7 Then
                fechaActualizacionStock = DateAdd("d", -7, Date)
            
                fechaActualizacionStock = Format(fechaActualizacionStock, "yyyy-dd-MM")
            Else
                fechaActualizacionStock = Format(rs!valor, "yyyy-dd-MM")
            End If
        End If
    End If
    
    rs.Close
    
    'sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
    '      "WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
    '      "AND ((FechaUltimaEntrada >= DATEADD(dd, -1, DATEDIFF(dd, 0, GETDATE())) AND FechaUltimaEntrada <= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()))) " & _
    '      "OR (FechaUltimaSalida >= DATEADD(mi, -10, GetDate()) AND FechaUltimaSalida <= GetDate()))"
    
    On Error GoTo errorControlStock
    
    sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
          "WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
          "AND FechaUltimaEntrada >= '" & fechaActualizacionStock & "' ORDER BY FechaUltimaEntrada ASC"
          
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents

        On Error GoTo errorControlStock
        
        sql = "UPDATE configuracion SET valor = '" & Format(rs!FechaUltimaEntrada, "yyyy-MM-dd") & "' WHERE campo = 'fechaActualizacionStockEntrada'"
        
        connMySql.Execute (sql)
        
        On Error GoTo errorControlStock
        
        sql = "select * from medicamentos where cod_nacional ='" & rs!IdArticu & "'"
    
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorControlStock
        
        sql = "select * from familia where IdFamilia='" & rs!XFam_IdFamilia & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not IsNull(rs!FechaUltimaEntrada) And rs!FechaUltimaEntrada <> "" Then
            fechaUltimaCompra = Format(rs!FechaUltimaEntrada, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaCompra = "NULL"
        End If
        
        If Not IsNull(rs!FechaUltimaSalida) And rs!FechaUltimaSalida <> "" Then
            fechaUltimaVenta = Format(rs!FechaUltimaSalida, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaVenta = "NULL"
        End If
    
        precio = Replace(rs!pvp, ",", ".")
        
        pcoste = Replace(rs!puc, ",", ".")
        
        pvpsiva = (rs!pvp * 100) / (rs!iva + 100)
        
        pvpsiva = Round(pvpsiva, 2)
        
        pvpsiva = Replace(pvpsiva, ",", ".")
        
        stock = rs!StockActual
        
        stockMinimo = rs!stockMinimo
        
        stockMaximo = rs!stockMaximo
        
        desc = LTrim(RTrim(StripString(rs!descripcion)))
        
        descripcion = ""

        If IsNull(rs3!descripcion) Or Len(rs3!descripcion) = 0 Then
            superfamilia = ""
        Else
            On Error GoTo errorControlStock
            
            sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                    " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(rs3!descripcion) & "'"
                    
            rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                superfamilia = ""
            Else
                superfamilia = rs4!descripcion
            End If
            
            rs4.Close
        End If
        
        On Error GoTo errorControlStock
            
        sql = "SELECT * FROM sinonimo WHERE IdArticu = '" & rs!IdArticu & "'"

        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            cod_barras = "847000" + CerosIzq(rs!IdArticu, 6)
        Else
            cod_barras = rs4!Sinonimo
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStock
                
        sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs!proveedorHabitual & "'"
       
        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            proveedor = ""
        Else
            proveedor = rs4!fis_nombre
        End If

        rs4.Close

        If Len(Trim(rs!laboratorio)) > 0 Then
            On Error GoTo errorControlStock
            
            sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(rs!laboratorio) & "'"
            
            rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                rs4.Close
                
                On Error GoTo errorControlStock
            
                sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(rs!laboratorio) & "'"
                
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs4.EOF Then
                    nombreLaboratorio = "<Sin Laboratorio>"
                Else
                    nombreLaboratorio = rs4!nombre
                End If
            Else
                nombreLaboratorio = rs4!nombre
            End If
            
            rs4.Close
        Else
            nombreLaboratorio = "<Sin Laboratorio>"
        End If
        
        On Error GoTo errorControlStock
        
        sql = "SELECT * FROM ESPEPARA WHERE CODIGO = '" & rs!IdArticu & "'"
        
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
        
        If rs4.EOF Then
            present = ""
        Else
            present = rs4!presentacion
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStock
        
        sql = "SELECT t.TEXTO FROM TEXTOS t INNER JOIN TEXTOSESPE te ON te.CODIGOTEXTO = t.CODIGOTEXTO " & _
               "WHERE te.CODIGOESPEPARA = '" & rs!IdArticu & "' ORDER BY te.CODIGOEPIGRAFE"
               
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
        If rs4.EOF Then
            descripcion = ""
        Else
            Do While Not rs4.EOF
                If descripcion = "" Then
                    descripcion = rs4!TEXTO
                Else
                    descripcion = descripcion & " <br> " & rs4!TEXTO
                End If
            
                rs4.MoveNext
            Loop
            
            If Len(descripcion) < 30000 Then
                descripcion = StripString(quitarCaracterCadena(Replace(descripcion, vbCrLf, "<br>"), Chr(0)))
            Else
                descripcion = ""
            End If
        End If
        
        rs4.Close
        
        If CBool(rs!baja) Then
            activo = 0
            baja = 1
        Else
            activo = 1
            baja = 0
        End If
        
        If Not IsNull(rs!fechaCaducidad) And rs!fechaCaducidad <> "" And rs!fechaCaducidad <> 0 Then
            fechaCaducidad = Format(rs!fechaCaducidad, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaCaducidad = ""
        End If
                       
        ''' rs medicamentos '''''''''''''''
		If rs2.EOF Then
            On Error GoTo errorControlStock
        
            sql = "INSERT INTO medicamentos " & "(cod_barras,cod_nacional,nombre,superFamilia,familia,precio,descripcion,laboratorio,nombre_laboratorio,proveedor,pvpSinIva,iva,stock,puc,stockMinimo,stockMaximo,presentacion,descripcionTienda,activoPrestashop,actualizadoPS,fechaCaducidad,fechaUltimaCompra,fechaUltimaVenta,baja) " & _
                    "VALUES('" & StripString(cod_barras) & "','" & _
                    StripString(rs!IdArticu) & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(superfamilia))) & "','" & _
                    LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                    precio & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(rs!laboratorio))) & "','" & _
                    LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                    LTrim(RTrim(StripString(proveedor))) & "','" & _
                    pvpsiva & "','" & _
                    rs!iva & "','" & _
                    stock & "','" & _
                    pcoste & "','" & _
                    stockMinimo & "','" & _
                    stockMaximo & "','" & _
                    LTrim(RTrim(StripString(present))) & "','" & _
                    LTrim(RTrim(descripcion)) & "', " & _
                    activo & ", 1, '" & fechaCaducidad & "', '" & fechaUltimaCompra & "', '" & fechaUltimaVenta & "'," & baja & ")"
         
            connMySql.Execute (sql)
        Else
            On Error GoTo errorControlStock
            
            sqlExtra = ""
            If (LTrim(RTrim(StripString(desc))) <> rs2!nombre _
                        Or precio <> rs2!precio Or LTrim(RTrim(StripString(rs!laboratorio))) <> rs2!laboratorio _
                        Or rs!iva <> rs2!iva Or stock <> rs2!stock _
                        Or LTrim(RTrim(StripString(present))) <> rs2!presentacion _
                        Or LTrim(RTrim(descripcion)) <> rs2!descripcion) Then
                sqlExtra = " cargadoPS = 0, actualizadoPS = 1, "
            End If
            
            sql = "UPDATE medicamentos SET cod_barras = '" & StripString(cod_barras) & "', nombre = '" & LTrim(RTrim(StripString(desc))) & "', superFamilia = '" & LTrim(RTrim(StripString(superfamilia))) & "', familia = '" & LTrim(RTrim(StripString(rs3!descripcion))) & "', " & _
                   "precio = '" & precio & "', descripcion = '" & LTrim(RTrim(StripString(desc))) & "', laboratorio = '" & LTrim(RTrim(StripString(rs!laboratorio))) & "', " & _
                   "nombre_laboratorio = '" & LTrim(RTrim(StripString(nombreLaboratorio))) & "', proveedor = '" & LTrim(RTrim(StripString(proveedor))) & "', " & _
                   "iva = '" & rs!iva & "', pvpSinIva = '" & pvpsiva & "', stock = " & stock & ", puc = '" & pcoste & "', stockMinimo = " & stockMinimo & ", " & _
                   "stockMaximo = " & stockMaximo & ", " & _
                   "presentacion = '" & LTrim(RTrim(StripString(present))) & "', descripcionTienda = '" & LTrim(RTrim(descripcion)) & "', " & sqlExtra & " activoPrestashop = " & activo & ", fechaCaducidad = '" & fechaCaducidad & "', " & _
                   "fechaUltimaCompra = '" & fechaUltimaCompra & "', fechaUltimaVenta = '" & fechaUltimaVenta & "', baja = " & baja & " " & _
                   " WHERE cod_nacional = '" & rs!IdArticu & "'"
                   
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        rs3.Close
     
        rs.MoveNext
    
    Loop

    rs.Close

final:
    GoTo fin
 
errorControlStock:
    Sleep 1500
 
fin:
    
End Sub

Private Sub Timer_Control_Stock_Fechas_Salida_Timer() '############ migracion ################
    Dim precio As String
    Dim pcoste As String
    Dim pvpsiva As String
    Dim stock As String
    Dim stockMinimo As String
    Dim stockMaximo As String
    Dim desc As String
    Dim laboratorio As String
    Dim nombreLaboratorio As String
    Dim descripcion As String
    Dim present As String
    Dim activo As String

    Dim sql As String

    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    On Error GoTo errorControlStock
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorControlStock
    
    sql = "SELECT * FROM configuracion WHERE campo = 'fechaActualizacionStockSalida'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStock
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('fechaActualizacionStockSalida', NULL)"
        
        connMySql.Execute (sql)
    Else
        If IsNull(rs!valor) Or Trim(rs!valor) = "" Then
            fechaActualizacionStock = DateAdd("d", -7, Date)
            
            fechaActualizacionStock = Format(fechaActualizacionStock, "yyyy-dd-MM")
        Else
            resultado = Abs(DateDiff("d", Now, Format(rs!valor, "yyyy-MM-dd")))
            
            If resultado > 7 Then
                fechaActualizacionStock = DateAdd("d", -7, Date)
            
                fechaActualizacionStock = Format(fechaActualizacionStock, "yyyy-dd-MM")
            Else
                fechaActualizacionStock = Format(rs!valor, "yyyy-dd-MM")
            End If
        End If
    End If
    
    rs.Close
    
    'sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
    '      "WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
    '      "AND ((FechaUltimaEntrada >= DATEADD(dd, -1, DATEDIFF(dd, 0, GETDATE())) AND FechaUltimaEntrada <= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()))) " & _
    '      "OR (FechaUltimaSalida >= DATEADD(mi, -10, GetDate()) AND FechaUltimaSalida <= GetDate()))"
    
    On Error GoTo errorControlStock
    
    sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
          "WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
          "AND FechaUltimaSalida >= '" & fechaActualizacionStock & "' ORDER BY FechaUltimaSalida ASC"
          
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents

        On Error GoTo errorControlStock
        
        sql = "UPDATE configuracion SET valor = '" & Format(rs!FechaUltimaSalida, "yyyy-MM-dd") & "' WHERE campo = 'fechaActualizacionStockSalida'"
        
        connMySql.Execute (sql)
        
        On Error GoTo errorControlStock
        
        sql = "select * from medicamentos where cod_nacional ='" & rs!IdArticu & "'"
    
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorControlStock
        
        sql = "select * from familia where IdFamilia='" & rs!XFam_IdFamilia & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not IsNull(rs!FechaUltimaEntrada) And rs!FechaUltimaEntrada <> "" Then
            fechaUltimaCompra = Format(rs!FechaUltimaEntrada, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaCompra = "NULL"
        End If
        
        If Not IsNull(rs!FechaUltimaSalida) And rs!FechaUltimaSalida <> "" Then
            fechaUltimaVenta = Format(rs!FechaUltimaSalida, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaVenta = "NULL"
        End If
    
        precio = Replace(rs!pvp, ",", ".")
        
        pcoste = Replace(rs!puc, ",", ".")
        
        pvpsiva = (rs!pvp * 100) / (rs!iva + 100)
        
        pvpsiva = Round(pvpsiva, 2)
        
        pvpsiva = Replace(pvpsiva, ",", ".")
        
        stock = rs!StockActual
        
        stockMinimo = rs!stockMinimo
        
        stockMaximo = rs!stockMaximo
        
        desc = LTrim(RTrim(StripString(rs!descripcion)))
        
        descripcion = ""

        If IsNull(rs3!descripcion) Or Len(rs3!descripcion) = 0 Then
            superfamilia = ""
        Else
            On Error GoTo errorControlStock
            
            sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                    " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(rs3!descripcion) & "'"
                    
            rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                superfamilia = ""
            Else
                superfamilia = rs4!descripcion
            End If
            
            rs4.Close
        End If
        
        On Error GoTo errorControlStock
            
        sql = "SELECT * FROM sinonimo WHERE IdArticu = '" & rs!IdArticu & "'"

        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            cod_barras = "847000" + CerosIzq(rs!IdArticu, 6)
        Else
            cod_barras = rs4!Sinonimo
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStock
                
        sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs!proveedorHabitual & "'"
       
        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            proveedor = ""
        Else
            proveedor = rs4!fis_nombre
        End If

        rs4.Close

        If Len(Trim(rs!laboratorio)) > 0 Then
            On Error GoTo errorControlStock
            
            sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(rs!laboratorio) & "'"
            
            rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                rs4.Close
                
                On Error GoTo errorControlStock
            
                sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(rs!laboratorio) & "'"
                
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs4.EOF Then
                    nombreLaboratorio = "<Sin Laboratorio>"
                Else
                    nombreLaboratorio = rs4!nombre
                End If
            Else
                nombreLaboratorio = rs4!nombre
            End If
            
            rs4.Close
        Else
            nombreLaboratorio = "<Sin Laboratorio>"
        End If
        
        On Error GoTo errorControlStock
        
        sql = "SELECT * FROM ESPEPARA WHERE CODIGO = '" & rs!IdArticu & "'"
        
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
        
        If rs4.EOF Then
            present = ""
        Else
            present = rs4!presentacion
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStock
        
        sql = "SELECT t.TEXTO FROM TEXTOS t INNER JOIN TEXTOSESPE te ON te.CODIGOTEXTO = t.CODIGOTEXTO " & _
               "WHERE te.CODIGOESPEPARA = '" & rs!IdArticu & "' ORDER BY te.CODIGOEPIGRAFE"
               
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
        If rs4.EOF Then
            descripcion = ""
        Else
            Do While Not rs4.EOF
                If descripcion = "" Then
                    descripcion = rs4!TEXTO
                Else
                    descripcion = descripcion & " <br> " & rs4!TEXTO
                End If
            
                rs4.MoveNext
            Loop
            
            If Len(descripcion) < 30000 Then
                descripcion = StripString(quitarCaracterCadena(Replace(descripcion, vbCrLf, "<br>"), Chr(0)))
            Else
                descripcion = ""
            End If
        End If
        
        rs4.Close
        
        If CBool(rs!baja) Then
            activo = 0
            baja = 1
        Else
            activo = 1
            baja = 0
        End If
        
        If Not IsNull(rs!fechaCaducidad) And rs!fechaCaducidad <> "" And rs!fechaCaducidad <> 0 Then
            fechaCaducidad = Format(rs!fechaCaducidad, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaCaducidad = ""
        End If
                       
        If rs2.EOF Then
            On Error GoTo errorControlStock
        
            sql = "INSERT INTO medicamentos " & "(cod_barras,cod_nacional,nombre,superFamilia,familia,precio,descripcion,laboratorio,nombre_laboratorio,proveedor,pvpSinIva,iva,stock,puc,stockMinimo,stockMaximo,presentacion,descripcionTienda,activoPrestashop,actualizadoPS,fechaCaducidad,fechaUltimaCompra,fechaUltimaVenta,baja) " & _
                    "VALUES('" & StripString(cod_barras) & "','" & _
                    StripString(rs!IdArticu) & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(superfamilia))) & "','" & _
                    LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                    precio & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(rs!laboratorio))) & "','" & _
                    LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                    LTrim(RTrim(StripString(proveedor))) & "','" & _
                    pvpsiva & "','" & _
                    rs!iva & "','" & _
                    stock & "','" & _
                    pcoste & "','" & _
                    stockMinimo & "','" & _
                    stockMaximo & "','" & _
                    LTrim(RTrim(StripString(present))) & "','" & _
                    LTrim(RTrim(descripcion)) & "', " & _
                    activo & ", 1, '" & fechaCaducidad & "', '" & fechaUltimaCompra & "', '" & fechaUltimaVenta & "'," & baja & ")"
         
            connMySql.Execute (sql)
        Else
            On Error GoTo errorControlStock
            
            sqlExtra = ""
            If (LTrim(RTrim(StripString(desc))) <> rs2!nombre _
                        Or precio <> rs2!precio Or LTrim(RTrim(StripString(rs!laboratorio))) <> rs2!laboratorio _
                        Or rs!iva <> rs2!iva Or stock <> rs2!stock _
                        Or LTrim(RTrim(StripString(present))) <> rs2!presentacion _
                        Or LTrim(RTrim(descripcion)) <> rs2!descripcion) Then
                sqlExtra = " cargadoPS = 0, actualizadoPS = 1, "
            End If
            
            sql = "UPDATE medicamentos SET cod_barras = '" & StripString(cod_barras) & "', nombre = '" & LTrim(RTrim(StripString(desc))) & "', superFamilia = '" & LTrim(RTrim(StripString(superfamilia))) & "', familia = '" & LTrim(RTrim(StripString(rs3!descripcion))) & "', " & _
                   "precio = '" & precio & "', descripcion = '" & LTrim(RTrim(StripString(desc))) & "', laboratorio = '" & LTrim(RTrim(StripString(rs!laboratorio))) & "', " & _
                   "nombre_laboratorio = '" & LTrim(RTrim(StripString(nombreLaboratorio))) & "', proveedor = '" & LTrim(RTrim(StripString(proveedor))) & "', " & _
                   "iva = '" & rs!iva & "', pvpSinIva = '" & pvpsiva & "', stock = " & stock & ", puc = '" & pcoste & "', stockMinimo = " & stockMinimo & ", " & _
                   "stockMaximo = " & stockMaximo & ", " & _
                   "presentacion = '" & LTrim(RTrim(StripString(present))) & "', descripcionTienda = '" & LTrim(RTrim(descripcion)) & "', " & sqlExtra & " activoPrestashop = " & activo & ", fechaCaducidad = '" & fechaCaducidad & "', " & _
                   "fechaUltimaCompra = '" & fechaUltimaCompra & "', fechaUltimaVenta = '" & fechaUltimaVenta & "', baja = " & baja & " " & _
                   " WHERE cod_nacional = '" & rs!IdArticu & "'"
                   
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        rs3.Close
     
        rs.MoveNext
    
    Loop

    rs.Close

final:
    GoTo fin
 
errorControlStock:
    Sleep 1500
 
fin:
    
End Sub


Private Sub Timer_Productos_Criticos_Timer() '############ migrado ##########################3333
    Dim familia As String
    Dim pcoste As String
    Dim precioMed As String
    Dim codLaboratorio As String
    Dim nombreLaboratorio As String
    Dim superfamilia As String
    Dim Fecha As String
    Dim FechaPedido As String

    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Set rs6 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim oField
    
    FieldExistsInRS = False
    
    On Error GoTo errorProductosCriticos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorProductosCriticos
        
    sql = "SELECT * from faltas LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "proveedor" Then
            FieldExistsInRS = True
        End If
    Next
    
    rs.Close
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE faltas ADD proveedor VARCHAR(255) DEFAULT NULL AFTER laboratorio;"
        
        connMySql.Execute (sql)
    End If
    
    On Error GoTo errorProductosCriticos
    
    sql = "select * from faltas order by idPedido Desc Limit 0,1"
   
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

    If rs.EOF Then
        FechaPedido = Format(Now, "yyyyMMdd")
        'FechaPedido = 20030101
    
        rs.Close
        
        On Error GoTo errorProductosCriticos
        
        sql = "SELECT * From pedido WHERE Fecha >= '" & FechaPedido & "' Order by IdPedido ASC"

        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    Else
        pedido = rs!IdPedido
        
        rs.Close
        
        On Error GoTo errorProductosCriticos
        
        sql = "SELECT * From pedido WHERE IdPedido >= " & pedido & " Order by IdPedido ASC"
    
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    End If

    Do While Not rs.EOF
        DoEvents
        
        FechaPedido = Format(rs!Hora, "yyyy-MM-dd HH:mm:ss")
     
        Fecha = Format(Now, "yyyy-MM-dd HH:mm:ss")
        
        On Error GoTo errorProductosCriticos
     
        sql = "select * from lineaPedido where IdPedido='" & rs!IdPedido & "'"

        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        Do While Not rs2.EOF
            If (Trim(rs2!XArt_IdArticu) <> "" And Not IsNull(rs2!XArt_IdArticu)) Then
                On Error GoTo errorProductosCriticos
                
                sql = "select * from articu where IdArticu='" & rs2!XArt_IdArticu & "'"
        
                rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs3.EOF Then
                    familia = ""
                    codLaboratorio = ""
                    pcoste = 0
                    precioMed = 0
                    nombreLaboratorio = "<Sin Laboratorio>"
                Else
                    pcoste = Replace(rs3!puc, ",", ".")
                    
                    On Error GoTo errorProductosCriticos
                    
                    sql = "select * from familia where IdFamilia='" & rs3!XFam_IdFamilia & "'"
            
                    rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                     
                    If rs4.EOF Then
                        familia = ""
                    Else
                        familia = rs4!descripcion
                    End If
            
                    rs4.Close
                    
                    If IsNull(familia) Then
                        familia = "<Sin Clasificar>"
                    Else
                        If Len(familia) = 0 Then
                            familia = "<Sin Clasificar>"
                        End If
                    End If
            
                    If IsNull(rs3!laboratorio) Then
                        codLaboratorio = ""
                    Else
                        codLaboratorio = rs3!laboratorio
                    End If
            
                    precioMed = Replace(rs3!pvp, ",", ".")
                    
                    If familia = "<Sin Clasificar>" Then
                        superfamilia = "<Sin Clasificar>"
                    Else
                        On Error GoTo errorProductosCriticos
                        
                        sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                                " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(familia) & "'"
                                
                        rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                        
                        If rs5.EOF Then
                            superfamilia = "<Sin Clasificar>"
                        Else
                            superfamilia = rs5!descripcion
                        End If
                        
                        rs5.Close
                    End If
                    
                    On Error GoTo errorProductosCriticos
                
                    sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs3!proveedorHabitual & "'"
                   
                    rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
                    If rs5.EOF Then
                        proveedor = ""
                    Else
                        proveedor = rs5!fis_nombre
                    End If
            
                    rs5.Close
                        
                    If Len(Trim(codLaboratorio)) > 0 Then
                        On Error GoTo errorProductosCriticos
                        
                        sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(codLaboratorio) & "'"
                        
                        rs5.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                        
                        If rs5.EOF Then
                            rs5.Close
                
                            On Error GoTo errorProductosCriticos
                        
                            sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(codLaboratorio) & "'"
                            
                            rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                            
                            If rs5.EOF Then
                                nombreLaboratorio = "<Sin Laboratorio>"
                            Else
                                nombreLaboratorio = rs5!nombre
                            End If
                        Else
                            nombreLaboratorio = rs5!nombre
                        End If
                        
                        rs5.Close
                    Else
                        nombreLaboratorio = "<Sin Laboratorio>"
                    End If
                End If
          '''''''''''''''''''''''''''''''
                On Error GoTo errorProductosCriticos
        
                sql = "select * from faltas where idPedido='" & rs2!IdPedido & "' AND idLinea= '" & rs2!IdLinea & "'"
        
                rs6.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
           
                'si se realiza un pedido y su stock es 0 se genera una falta
                If (rs3!StockActual = 0 And rs6.EOF) Then
                    On Error GoTo errorProductosCriticos
             
                    sql = "INSERT INTO faltas " & "(idPedido,idLinea,cod_nacional,descripcion,familia,superFamilia,cantidadPedida,fechaFalta,cod_laboratorio,laboratorio,proveedor,fechaPedido,pvp,puc) VALUES('" & _
                                      rs2!IdPedido & "','" & _
                                      rs2!IdLinea & "','" & _
                                      StripString(rs3!IdArticu) & "','" & _
                                      LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                                      LTrim(RTrim(StripString(familia))) & "','" & _
                                      LTrim(RTrim(StripString(superfamilia))) & "','" & _
                                      rs2!unidades & "','" & _
                                      Fecha & "','" & _
                                      LTrim(RTrim(StripString(codLaboratorio))) & "','" & _
                                      LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                                      LTrim(RTrim(StripString(proveedor))) & "','" & _
                                      FechaPedido & "','" & _
                                      precioMed & "','" & _
                                      pcoste & "')"
                                     
                     connMySql.Execute (sql)
                End If
         
                rs3.Close
                rs6.Close
            End If
    
            rs2.MoveNext
        Loop
   
        rs.MoveNext
        rs2.Close
    Loop

    rs.Close

final:
    GoTo fin
 
errorProductosCriticos:
    Sleep 1500
 
fin:

End Sub

Private Sub Timer_Encargos_Timer() '''''''''''' Migrado ################################333
    Dim sql As String
    Dim idEncargo As String
    Dim cod_nacional As String
    Dim cliente As String
    Dim nombre As String
    Dim familia As String
    Dim codLaboratorio As String
    Dim pcoste As String
    Dim precioMed As String
    Dim nombreLaboratorio As String
    Dim superfamilia As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim oField
    
    FieldExistsInRS = False
    
    On Error GoTo errorEncargos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorEncargos
        
    sql = "SELECT * from encargos LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "proveedor" Then
            FieldExistsInRS = True
        End If
    Next
    
    rs.Close
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE encargos ADD proveedor VARCHAR(255) DEFAULT NULL AFTER laboratorio;"
        
        connMySql.Execute (sql)
    End If
    
    On Error GoTo errorEncargos

    sql = "select * from encargos order by idEncargo Desc Limit 0,1"
       
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

    If rs.EOF Then
        idEncargo = 1
    Else
        idEncargo = rs!idEncargo
    End If

    rs.Close

    On Error GoTo errorEncargos

    sql = "SELECT * From Encargo WHERE year(idFecha) >= 2015 AND IdContador >= " & idEncargo & " Order by IdContador ASC"

    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
     
    Do While Not rs.EOF
        DoEvents
 
        On Error GoTo errorEncargos
    
        idEncargo = rs!idContador
    
        cod_nacional = LTrim(RTrim(rs!XArt_IdArticu))
    
        If IsNull(rs!XCli_IdCliente) Or rs!XCli_IdCliente = "" Then
            cliente = 0
        Else
            cliente = LTrim(RTrim(StripString(rs!XCli_IdCliente)))
        End If
    
        sql = "select * from articu where IdArticu='" & rs!XArt_IdArticu & "'"

        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
        If rs2.EOF Then
            nombre = ""
            familia = ""
            codLaboratorio = ""
            pcoste = 0
            precioMed = 0
            nombreLaboratorio = "<Sin Laboratorio>"
        Else
            nombre = LTrim(RTrim(StripString(rs2!descripcion)))
            
            pcoste = Replace(rs2!puc, ",", ".")
            
            precioMed = Replace(rs2!pvp, ",", ".")
            
            On Error GoTo errorEncargos
            
            sql = "select * from familia where IdFamilia='" & rs2!XFam_IdFamilia & "'"
    
            rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
             
            If rs3.EOF Then
                familia = ""
            Else
                familia = LTrim(RTrim(StripString(rs3!descripcion)))
            End If
        
            rs3.Close
          
            If IsNull(familia) Then
                familia = "<Sin Clasificar>"
            Else
                If Len(familia) = 0 Then
                    familia = "<Sin Clasificar>"
                End If
            End If
                
            If IsNull(rs2!laboratorio) Then
                codLaboratorio = ""
            Else
                codLaboratorio = rs2!laboratorio
            End If
                
            If familia = "<Sin Clasificar>" Then
                superfamilia = "<Sin Clasificar>"
            Else
                On Error GoTo errorEncargos
                
                sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                        " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(familia) & "'"
                        
                rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs3.EOF Then
                    superfamilia = "<Sin Clasificar>"
                Else
                    superfamilia = rs3!descripcion
                End If
                
                rs3.Close
            End If
            
            On Error GoTo errorEncargos
                
            sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs2!proveedorHabitual & "'"
           
            rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            If rs3.EOF Then
                proveedor = ""
            Else
                proveedor = rs3!fis_nombre
            End If
    
            rs3.Close
                            
            If Len(Trim(codLaboratorio)) > 0 Then
                On Error GoTo errorEncargos
                
                sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(codLaboratorio) & "'"
                
                rs3.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                
                If rs3.EOF Then
                    rs3.Close
                
                    On Error GoTo errorEncargos
                
                    sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(codLaboratorio) & "'"
                    
                    rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs3.EOF Then
                        nombreLaboratorio = "<Sin Laboratorio>"
                    Else
                        nombreLaboratorio = rs3!nombre
                    End If
                Else
                    nombreLaboratorio = rs3!nombre
                End If
                
                rs3.Close
            Else
                nombreLaboratorio = "<Sin Laboratorio>"
            End If
        End If
       
        On Error GoTo errorEncargos
           
        id_vendedor = rs!vendedor
                                    
        sql = "select * from vendedor where IdVendedor='" & id_vendedor & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If rs3.EOF Then
            trabajador = ""
        Else
            trabajador = LTrim(RTrim(StripString(rs3!nombre)))
        End If
                    
        rs3.Close
                    
        On Error GoTo errorEncargos
           
        sql = "select * from encargos where IdEncargo=" & idEncargo
        
        rs3.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
                    
        If rs3.EOF Then
            
            On Error GoTo errorEncargos
            
            sql = "INSERT INTO encargos " & "(idEncargo,cod_nacional,nombre,superFamilia,familia,cod_laboratorio,laboratorio,proveedor,pvp,puc,dni,fecha,trabajador,unidades,fechaEntrega,observaciones) VALUES('" & _
                                     idEncargo & "','" & _
                                     cod_nacional & "','" & _
                                     nombre & "','" & _
                                     LTrim(RTrim(StripString(superfamilia))) & "','" & _
                                     familia & "','" & _
                                     LTrim(RTrim(StripString(codLaboratorio))) & "','" & _
                                     LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                                     LTrim(RTrim(StripString(proveedor))) & "','" & _
                                     precioMed & "','" & _
                                     pcoste & "','" & _
                                     cliente & "','" & _
                                     Format(rs!idFecha, "yyyy-MM-dd HH:mm:ss") & "','" & _
                                     trabajador & "','" & _
                                     rs!unidades & "','" & _
                                     Format(rs!FechaEntrega, "yyyy-MM-dd HH:mm:ss") & "','" & _
                                     LTrim(RTrim(StripString(rs!Observaciones))) & "')"
    
            connMySql.Execute (sql)
        End If
            
        rs.MoveNext
        
        rs2.Close
        rs3.Close
    Loop
     
    rs.Close

final:
    GoTo fin
 
errorEncargos:
    Sleep 1500
 
fin:

End Sub

Private Sub Timer_Familias_Timer() '##############3 migrado ####################333333333333
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    On Error GoTo errorFamilias
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorFamilias
    
    sql = "select * from familia"

    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorFamilias
        
        sql = "select * from familia where familia='" & LTrim(RTrim(StripString(rs!descripcion))) & "'"

        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

        If rs2.EOF Then
            On Error GoTo errorFamilias
            
            sql = "INSERT INTO familia " & "(familia) VALUES('" & LTrim(RTrim(StripString(rs!descripcion))) & "')"

            connMySql.Execute (sql)
        End If

        rs2.Close

        rs.MoveNext
    Loop

    rs.Close
    
final:
    GoTo fin
 
errorFamilias:
    Sleep 1500
 
fin:
End Sub

Private Sub Timer_Control_Stock_Inicial_Timer()    '############ migrado #########################3
    Dim precio As String
    Dim pcoste As String
    Dim pvpsiva As String
    Dim stock As String
    Dim stockMinimo As String
    Dim stockMaximo As String
    Dim desc As String
    Dim laboratorio As String
    Dim codArticu As String
    Dim nombreLaboratorio As String
    Dim descripcion As String
    Dim present As String
    Dim activo As String
    
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim FieldExistsInRS1 As Boolean
    Dim FieldExistsInRS2 As Boolean
    Dim FieldExistsInRS3 As Boolean
    Dim FieldExistsInRS4 As Boolean
    Dim FieldExistsInRS5 As Boolean
    Dim FieldExistsInRS6 As Boolean
    Dim FieldExistsInRS7 As Boolean
    Dim FieldExistsInRS8 As Boolean
    Dim FieldExistsInRS9 As Boolean
    Dim oField
    
    FieldExistsInRS = False
    FieldExistsInRS1 = False
    FieldExistsInRS2 = False
    FieldExistsInRS3 = False
    FieldExistsInRS4 = False
    FieldExistsInRS5 = False
    FieldExistsInRS6 = False
    FieldExistsInRS7 = False
    FieldExistsInRS8 = False
    FieldExistsInRS9 = False
    
    On Error GoTo errorControlStockInicial
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorControlStockInicial
        
    sql = "SELECT * from medicamentos LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "laboratorio" Then
            FieldExistsInRS = True
        End If
        If oField.Name = "stock" Then
            FieldExistsInRS1 = True
        End If
        If oField.Name = "puc" Then
            FieldExistsInRS2 = True
        End If
        If oField.Name = "stockMinimo" Then
            FieldExistsInRS3 = True
        End If
        If oField.Name = "presentacion" Then
            FieldExistsInRS4 = True
        End If
        If oField.Name = "fechaCaducidad" Then
            FieldExistsInRS5 = True
        End If
        If oField.Name = "porDondeVoySinStock" Then
            FieldExistsInRS6 = True
        End If
        If oField.Name = "fechaUltimaCompra" Then
            FieldExistsInRS7 = True
        End If
        If oField.Name = "proveedor" Then
            FieldExistsInRS8 = True
        End If
        If oField.Name = "superFamilia" Then
            FieldExistsInRS9 = True
        End If
    Next
    
    rs.Close
    
    On Error GoTo errorControlStockInicial
    
    sql = "SELECT * FROM configuracion WHERE campo = 'fechaActualizacionStockEntrada'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStockInicial
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('fechaActualizacionStockEntrada', NULL)"
        
        connMySql.Execute (sql)
    End If
    
    rs.Close
    
    sql = "SELECT * FROM configuracion WHERE campo = 'fechaActualizacionStockSalida'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStockInicial
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('fechaActualizacionStockSalida', NULL)"
        
        connMySql.Execute (sql)
    End If
    
    rs.Close
    
    sql = "SELECT * FROM configuracion WHERE campo = 'porDondeVoyConStock'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStockInicial
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('porDondeVoyConStock', '0')"
        
        connMySql.Execute (sql)
        
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('porDondeVoySinStock', '0')"
        
        connMySql.Execute (sql)
        
        codArticu = 0
    Else
        codArticu = rs!valor
    End If
    
    rs.Close
    
    On Error GoTo errorControlStockInicial
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE medicamentos ADD laboratorio VARCHAR(255);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS1 = False Then
        sql = "ALTER TABLE medicamentos ADD (pvpSinIva float, iva int (11), stock int (11));"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS2 = False Then
        sql = "ALTER TABLE medicamentos ADD (puc float);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS3 = False Then
        sql = "ALTER TABLE medicamentos ADD (stockMinimo int (11), stockMaximo int (11));"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS4 = False Then
        sql = "ALTER TABLE medicamentos ADD `nombre_laboratorio` varchar(255) DEFAULT NULL AFTER laboratorio;"
        
        connMySql.Execute (sql)
        
        sql = "ALTER TABLE medicamentos ADD (`presentacion` varchar(50) DEFAULT NULL, " & _
                              "`descripcionTienda` text, " & _
                              "`prestashopIdPS` int(10) DEFAULT NULL, " & _
                              "`cargadoPS` tinyint(1) DEFAULT '0', " & _
                              "`fechaCargadoPS` datetime DEFAULT NULL, " & _
                              "`activoPrestashop` tinyint(1) DEFAULT '1', " & _
                              "`actualizadoPS` tinyint(1) DEFAULT '0', " & _
                              "`eliminado` tinyint(1) DEFAULT '0', " & _
                              "`fechaEliminado` datetime DEFAULT NULL);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS5 = False Then
        sql = "ALTER TABLE medicamentos ADD (fechaCaducidad datetime, porDondeVoy TINYINT(1) DEFAULT 0);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS6 = False Then
        sql = "ALTER TABLE medicamentos ADD (porDondeVoySinStock TINYINT(1) DEFAULT 0);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS7 = False Then
        sql = "ALTER TABLE medicamentos ADD (fechaUltimaCompra DATETIME DEFAULT NULL, fechaUltimaVenta DATETIME DEFAULT NULL);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS8 = False Then
        sql = "ALTER TABLE medicamentos ADD proveedor VARCHAR(255) DEFAULT NULL AFTER nombre_laboratorio;"
        
        connMySql.Execute (sql)
        
        sql = "ALTER TABLE medicamentos ADD (baja TINYINT(1) DEFAULT 0);"
        
        connMySql.Execute (sql)
    End If
    
    If FieldExistsInRS9 = False Then
        sql = "ALTER TABLE medicamentos ADD superFamilia VARCHAR(255) DEFAULT NULL AFTER nombre;"
        
        connMySql.Execute (sql)
    End If
    
    'On Error GoTo errorControlStockInicial
    
    'sql = "SELECT cod_nacional FROM medicamentos WHERE porDondeVoy = 1 LIMIT 0,1"
    
    'rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    'If rs.EOF Then
    '    codArticu = -1
    'Else
    '    codArticu = rs!cod_nacional
    'End If
    
    'rs.Close
    
    On Error GoTo errorControlStockInicial
    
    sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
          " WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
          " AND a.IdArticu >= " & codArticu & " AND a.StockActual > 0 ORDER BY a.IdArticu ASC"
          
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoyConStock'"
        
        connMySql.Execute (sql)
        
        sql = "UPDATE medicamentos SET porDondeVoy = 0"
        
        connMySql.Execute (sql)
    End If
    
    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorControlStockInicial
        
        sql = "select * from medicamentos where cod_nacional ='" & rs!IdArticu & "'"
    
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorControlStockInicial
        
        sql = "select * from familia where IdFamilia='" & rs!XFam_IdFamilia & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not IsNull(rs!FechaUltimaEntrada) And rs!FechaUltimaEntrada <> "" Then
            fechaUltimaCompra = Format(rs!FechaUltimaEntrada, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaCompra = "NULL"
        End If
        
        If Not IsNull(rs!FechaUltimaSalida) And rs!FechaUltimaSalida <> "" Then
            fechaUltimaVenta = Format(rs!FechaUltimaSalida, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaVenta = "NULL"
        End If
    
        precio = Replace(rs!pvp, ",", ".")
        
        pcoste = Replace(rs!puc, ",", ".")
        
        pvpsiva = (rs!pvp * 100) / (rs!iva + 100)
        
        pvpsiva = Round(pvpsiva, 2)
        
        pvpsiva = Replace(pvpsiva, ",", ".")
        
        stock = rs!StockActual
        
        stockMinimo = rs!stockMinimo
        
        stockMaximo = rs!stockMaximo
        
        desc = LTrim(RTrim(StripString(rs!descripcion)))
        
        descripcion = ""
        
        If IsNull(rs3!descripcion) Or Len(rs3!descripcion) = 0 Then
            superfamilia = ""
        Else
            On Error GoTo errorControlStockInicial
            
            sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                    " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(rs3!descripcion) & "'"
                    
            rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                superfamilia = ""
            Else
                superfamilia = rs4!descripcion
            End If
            
            rs4.Close
        End If
        
        On Error GoTo errorControlStockInicial
        
        sql = "SELECT * FROM sinonimo WHERE IdArticu = '" & rs!IdArticu & "'"

        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            cod_barras = "847000" + CerosIzq(rs!IdArticu, 6)
        Else
            cod_barras = rs4!Sinonimo
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStockInicial
                
        sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs!proveedorHabitual & "'"
       
        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            proveedor = ""
        Else
            proveedor = rs4!fis_nombre
        End If

        rs4.Close

        If Len(Trim(rs!laboratorio)) > 0 Then
            On Error GoTo errorControlStockInicial
            
            sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(rs!laboratorio) & "'"
            
            rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                rs4.Close
                
                On Error GoTo errorControlStockInicial
            
                sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(rs!laboratorio) & "'"
                
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs4.EOF Then
                    nombreLaboratorio = "<Sin Laboratorio>"
                Else
                    nombreLaboratorio = rs4!nombre
                End If
            Else
                nombreLaboratorio = rs4!nombre
            End If
            
            rs4.Close
        Else
            nombreLaboratorio = "<Sin Laboratorio>"
        End If
        
        On Error GoTo errorControlStockInicial
        
        sql = "SELECT * FROM ESPEPARA WHERE CODIGO = '" & rs!IdArticu & "'"
        
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
        
        If rs4.EOF Then
            present = ""
        Else
            If IsNull(rs4!presentacion) Then
                present = ""
            Else
                present = rs4!presentacion
            End If
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStockInicial
        
        sql = "SELECT t.TEXTO FROM TEXTOS t INNER JOIN TEXTOSESPE te ON te.CODIGOTEXTO = t.CODIGOTEXTO " & _
               "WHERE te.CODIGOESPEPARA = '" & rs!IdArticu & "' ORDER BY te.CODIGOEPIGRAFE"
               
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
        If rs4.EOF Then
            descripcion = ""
        Else
            Do While Not rs4.EOF
                If descripcion = "" Then
                    descripcion = rs4!TEXTO
                Else
                    descripcion = descripcion & " <br> " & rs4!TEXTO
                End If
            
                rs4.MoveNext
            Loop
            
            If Len(descripcion) < 30000 Then
                descripcion = StripString(quitarCaracterCadena(Replace(descripcion, vbCrLf, "<br>"), Chr(0)))
            Else
                descripcion = ""
            End If
        End If
        
        rs4.Close
        
        If CBool(rs!baja) Then
            activo = 0
            baja = 1
        Else
            activo = 1
            baja = 0
        End If
        
        If Not IsNull(rs!fechaCaducidad) And rs!fechaCaducidad <> "" And rs!fechaCaducidad <> 0 Then
            fechaCaducidad = Format(rs!fechaCaducidad, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaCaducidad = ""
        End If
                       
        If rs2.EOF Then
            On Error GoTo errorControlStockInicial
            
            sql = "UPDATE configuracion SET valor = '" & StripString(rs!IdArticu) & "' WHERE campo = 'porDondeVoyConStock'"
               
            connMySql.Execute (sql)
        
            sql = "INSERT INTO medicamentos " & "(cod_barras,cod_nacional,nombre,superFamilia,familia,precio,descripcion,laboratorio,nombre_laboratorio,proveedor,pvpSinIva,iva,stock,puc,stockMinimo,stockMaximo,presentacion,descripcionTienda,activoPrestashop,actualizadoPS,fechaCaducidad,fechaUltimaCompra,fechaUltimaVenta,baja) " & _
                    "VALUES('" & StripString(cod_barras) & "','" & _
                    StripString(rs!IdArticu) & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(superfamilia))) & "','" & _
                    LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                    precio & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(rs!laboratorio))) & "','" & _
                    LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                    LTrim(RTrim(StripString(proveedor))) & "','" & _
                    pvpsiva & "','" & _
                    rs!iva & "','" & _
                    stock & "','" & _
                    pcoste & "','" & _
                    stockMinimo & "','" & _
                    stockMaximo & "','" & _
                    LTrim(RTrim(StripString(present))) & "','" & _
                    LTrim(RTrim(descripcion)) & "', " & _
                    activo & ", 1, '" & fechaCaducidad & "', '" & fechaUltimaCompra & "', '" & fechaUltimaVenta & "', " & baja & ")"
                    
            connMySql.Execute (sql)
        Else
            On Error GoTo errorControlStockInicial
            
            sql = "UPDATE configuracion SET valor = '" & rs!IdArticu & "' WHERE campo = 'porDondeVoyConStock'"
               
            connMySql.Execute (sql)
            
            sqlExtra = ""
            If (LTrim(RTrim(StripString(desc))) <> rs2!nombre _
                        Or precio <> rs2!precio Or LTrim(RTrim(StripString(rs!laboratorio))) <> rs2!laboratorio _
                        Or rs!iva <> rs2!iva Or stock <> rs2!stock _
                        Or LTrim(RTrim(StripString(present))) <> rs2!presentacion _
                        Or LTrim(RTrim(descripcion)) <> rs2!descripcion) Then
                sqlExtra = " cargadoPS = 0, actualizadoPS = 1, "
            End If
            
            sql = "UPDATE medicamentos SET cod_barras = '" & StripString(cod_barras) & "', nombre = '" & LTrim(RTrim(StripString(desc))) & "', superFamilia = '" & LTrim(RTrim(StripString(superfamilia))) & "', familia = '" & LTrim(RTrim(StripString(rs3!descripcion))) & "', " & _
                   "precio = '" & precio & "', descripcion = '" & LTrim(RTrim(StripString(desc))) & "', laboratorio = '" & LTrim(RTrim(StripString(rs!laboratorio))) & "', " & _
                   "nombre_laboratorio = '" & LTrim(RTrim(StripString(nombreLaboratorio))) & "', proveedor = '" & LTrim(RTrim(StripString(proveedor))) & "', " & _
                   "iva = '" & rs!iva & "', pvpSinIva = '" & pvpsiva & "', stock = " & stock & ", puc = '" & pcoste & "', stockMinimo = " & stockMinimo & ", " & _
                   "stockMaximo = " & stockMaximo & ", " & _
                   "presentacion = '" & LTrim(RTrim(StripString(present))) & "', descripcionTienda = '" & LTrim(RTrim(descripcion)) & "', " & sqlExtra & " activoPrestashop = " & activo & ", fechaCaducidad = '" & fechaCaducidad & "', " & _
                   "fechaUltimaCompra = '" & fechaUltimaCompra & "', fechaUltimaVenta = '" & fechaUltimaVenta & "', baja = " & baja & " " & _
                   " WHERE cod_nacional = '" & rs!IdArticu & "'"
                   
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        rs3.Close
     
        rs.MoveNext
    
        If rs.EOF Then
            sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoyConStock'"
            
            connMySql.Execute (sql)
            
            sql = "UPDATE medicamentos SET porDondeVoy = 0"
            
            connMySql.Execute (sql)
        End If
    Loop
    
    rs.Close
    
final:
    GoTo fin
 
errorControlStockInicial:
    Sleep 1500
 
fin:
    
End Sub

Private Sub Timer_Control_Sin_Stock_Inicial_Timer() '############# migrado #####################
    Dim precio As String
    Dim pcoste As String
    Dim pvpsiva As String
    Dim stock As String
    Dim stockMinimo As String
    Dim stockMaximo As String
    Dim desc As String
    Dim laboratorio As String
    Dim codArticu As String
    Dim nombreLaboratorio As String
    Dim descripcion As String
    Dim present As String
    Dim activo As String
    
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    On Error GoTo errorControlStockInicial
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorControlStockInicial
    
    sql = "SELECT * FROM configuracion WHERE campo = 'porDondeVoySinStock'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorControlStockInicial
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('porDondeVoySinStock', '0')"
        
        connMySql.Execute (sql)
        
        codArticu = 0
    Else
        codArticu = rs!valor
    End If
    
    rs.Close
    
    On Error GoTo errorControlStockInicial
    
    sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
          " WHERE a.Descripcion <> 'PENDIENTE DE ASIGNACIÓN' AND a.Descripcion <> 'VENTAS VARIAS' AND a.Descripcion <> '   BASE DE DATOS  3/03/2014' " & _
          " AND a.IdArticu >= " & codArticu & " AND a.StockActual <= 0 ORDER BY a.IdArticu ASC"
          
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoySinStock'"
        
        connMySql.Execute (sql)
        
        sql = "UPDATE medicamentos SET porDondeVoySinStock = 0"
        
        connMySql.Execute (sql)
    End If
    
    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorControlStockInicial
        
        sql = "select * from medicamentos where cod_nacional ='" & rs!IdArticu & "'"
    
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorControlStockInicial
        
        sql = "select * from familia where IdFamilia='" & rs!XFam_IdFamilia & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not IsNull(rs!FechaUltimaEntrada) And rs!FechaUltimaEntrada <> "" Then
            fechaUltimaCompra = Format(rs!FechaUltimaEntrada, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaCompra = "NULL"
        End If
        
        If Not IsNull(rs!FechaUltimaSalida) And rs!FechaUltimaSalida <> "" Then
            fechaUltimaVenta = Format(rs!FechaUltimaSalida, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaUltimaVenta = "NULL"
        End If
    
        precio = Replace(rs!pvp, ",", ".")
        
        pcoste = Replace(rs!puc, ",", ".")
        
        pvpsiva = (rs!pvp * 100) / (rs!iva + 100)
        
        pvpsiva = Round(pvpsiva, 2)
        
        pvpsiva = Replace(pvpsiva, ",", ".")
        
        stock = rs!StockActual
        
        stockMinimo = rs!stockMinimo
        
        stockMaximo = rs!stockMaximo
        
        desc = LTrim(RTrim(StripString(rs!descripcion)))
        
        descripcion = ""

        If IsNull(rs3!descripcion) Or Len(rs3!descripcion) = 0 Then
            superfamilia = ""
        Else
            On Error GoTo errorControlStockInicial
            
            sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                    " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(rs3!descripcion) & "'"
                    
            rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                superfamilia = ""
            Else
                superfamilia = rs4!descripcion
            End If
            
            rs4.Close
        End If
        
        On Error GoTo errorControlStockInicial
            
        sql = "SELECT * FROM sinonimo WHERE IdArticu = '" & rs!IdArticu & "'"

        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            cod_barras = "847000" + CerosIzq(rs!IdArticu, 6)
        Else
            cod_barras = rs4!Sinonimo
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStockInicial
                
        sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs!proveedorHabitual & "'"
       
        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs4.EOF Then
            proveedor = ""
        Else
            proveedor = rs4!fis_nombre
        End If

        rs4.Close

        If Len(Trim(rs!laboratorio)) > 0 Then
            On Error GoTo errorControlStockInicial
            
            sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(rs!laboratorio) & "'"
            
            rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
            If rs4.EOF Then
                rs4.Close
                
                On Error GoTo errorControlStockInicial
            
                sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(rs!laboratorio) & "'"
                
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                If rs4.EOF Then
                    nombreLaboratorio = "<Sin Laboratorio>"
                Else
                    nombreLaboratorio = rs4!nombre
                End If
            Else
                nombreLaboratorio = rs4!nombre
            End If
            
            rs4.Close
        Else
            nombreLaboratorio = "<Sin Laboratorio>"
        End If
        
        On Error GoTo errorControlStockInicial
        
        sql = "SELECT * FROM ESPEPARA WHERE CODIGO = '" & rs!IdArticu & "'"
        
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
        
        If rs4.EOF Then
            present = ""
        Else
            If IsNull(rs4!presentacion) Then
                present = ""
            Else
                present = rs4!presentacion
            End If
        End If
        
        rs4.Close
        
        On Error GoTo errorControlStockInicial
        
        sql = "SELECT t.TEXTO FROM TEXTOS t INNER JOIN TEXTOSESPE te ON te.CODIGOTEXTO = t.CODIGOTEXTO " & _
               "WHERE te.CODIGOESPEPARA = '" & rs!IdArticu & "' ORDER BY te.CODIGOEPIGRAFE"
               
        rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
            
        If rs4.EOF Then
            descripcion = ""
        Else
            Do While Not rs4.EOF
                If descripcion = "" Then
                    descripcion = rs4!TEXTO
                Else
                    descripcion = descripcion & " <br> " & rs4!TEXTO
                End If
            
                rs4.MoveNext
            Loop
            
            If Len(descripcion) < 30000 Then
                descripcion = StripString(quitarCaracterCadena(Replace(descripcion, vbCrLf, "<br>"), Chr(0)))
            Else
                descripcion = ""
            End If
        End If
        
        rs4.Close
        
        If CBool(rs!baja) Then
            activo = 0
            baja = 1
        Else
            activo = 1
            baja = 0
        End If
        
        If Not IsNull(rs!fechaCaducidad) And rs!fechaCaducidad <> "" And rs!fechaCaducidad <> 0 Then
            fechaCaducidad = Format(rs!fechaCaducidad, "yyyy-MM-dd HH:mm:ss")
        Else
            fechaCaducidad = ""
        End If
                       
        If rs2.EOF Then
            On Error GoTo errorControlStockInicial
            
            sql = "UPDATE configuracion SET valor = '" & StripString(rs!IdArticu) & "' WHERE campo = 'porDondeVoySinStock'"
               
            connMySql.Execute (sql)
        
            sql = "INSERT INTO medicamentos " & "(cod_barras,cod_nacional,nombre,superFamilia,familia,precio,descripcion,laboratorio,nombre_laboratorio,proveedor,pvpSinIva,iva,stock,puc,stockMinimo,stockMaximo,presentacion,descripcionTienda,activoPrestashop,actualizadoPS,fechaCaducidad,fechaUltimaCompra,fechaUltimaVenta,baja) " & _
                    "VALUES('" & StripString(cod_barras) & "','" & _
                    StripString(rs!IdArticu) & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(superfamilia))) & "','" & _
                    LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                    precio & "','" & _
                    LTrim(RTrim(StripString(desc))) & "','" & _
                    LTrim(RTrim(StripString(rs!laboratorio))) & "','" & _
                    LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                    LTrim(RTrim(StripString(proveedor))) & "','" & _
                    pvpsiva & "','" & _
                    rs!iva & "','" & _
                    stock & "','" & _
                    pcoste & "','" & _
                    stockMinimo & "','" & _
                    stockMaximo & "','" & _
                    LTrim(RTrim(StripString(present))) & "','" & _
                    LTrim(RTrim(descripcion)) & "', " & _
                    activo & ", 1, '" & fechaCaducidad & "', '" & fechaUltimaCompra & "', '" & fechaUltimaVenta & "'," & baja & ")"
                    
            connMySql.Execute (sql)
        Else
            On Error GoTo errorControlStockInicial
            
            sql = "UPDATE configuracion SET valor = '" & rs!IdArticu & "' WHERE campo = 'porDondeVoySinStock'"
               
            connMySql.Execute (sql)
            
            sqlExtra = ""
            If (LTrim(RTrim(StripString(desc))) <> rs2!nombre _
                        Or precio <> rs2!precio Or LTrim(RTrim(StripString(rs!laboratorio))) <> rs2!laboratorio _
                        Or rs!iva <> rs2!iva Or stock <> rs2!stock _
                        Or LTrim(RTrim(StripString(present))) <> rs2!presentacion _
                        Or LTrim(RTrim(descripcion)) <> rs2!descripcion) Then
                sqlExtra = " cargadoPS = 0, actualizadoPS = 1, "
            End If
            
            sql = "UPDATE medicamentos SET cod_barras = '" & StripString(cod_barras) & "', nombre = '" & LTrim(RTrim(StripString(desc))) & "', superFamilia = '" & LTrim(RTrim(StripString(superfamilia))) & "', familia = '" & LTrim(RTrim(StripString(rs3!descripcion))) & "', " & _
                   "precio = '" & precio & "', descripcion = '" & LTrim(RTrim(StripString(desc))) & "', laboratorio = '" & LTrim(RTrim(StripString(rs!laboratorio))) & "', " & _
                   "nombre_laboratorio = '" & LTrim(RTrim(StripString(nombreLaboratorio))) & "', proveedor = '" & LTrim(RTrim(StripString(proveedor))) & "', " & _
                   "iva = '" & rs!iva & "', pvpSinIva = '" & pvpsiva & "', stock = " & stock & ", puc = '" & pcoste & "', stockMinimo = " & stockMinimo & ", " & _
                   "stockMaximo = " & stockMaximo & ", " & _
                   "presentacion = '" & LTrim(RTrim(StripString(present))) & "', descripcionTienda = '" & LTrim(RTrim(descripcion)) & "', " & sqlExtra & " activoPrestashop = " & activo & ", fechaCaducidad = '" & fechaCaducidad & "', " & _
                   "fechaUltimaCompra = '" & fechaUltimaCompra & "', fechaUltimaVenta = '" & fechaUltimaVenta & "', baja = " & baja & " " & _
                   " WHERE cod_nacional = '" & rs!IdArticu & "'"
                   
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        rs3.Close
     
        rs.MoveNext
        
        If rs.EOF Then
            sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoySinStock'"
            
            connMySql.Execute (sql)
            
            sql = "UPDATE medicamentos SET porDondeVoySinStock = 0"
            
            connMySql.Execute (sql)
        End If
    
    Loop
    
    rs.Close
    
final:
    GoTo fin
 
errorControlStockInicial:
    Sleep 1500
 
fin:
    
End Sub


Private Sub Timer_Listas_Timer() '####### migrado #######################################3
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim oField
    
    FieldExistsInRS = False
    
    On Error GoTo errorListas
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorListas
    
    sql = "SELECT * from listas LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "porDondeVoy" Then
            FieldExistsInRS = True
        End If
    Next
    
    rs.Close
    
    On Error GoTo errorListas
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE listas ADD porDondeVoy TINYINT(1) DEFAULT 0;"
        
        connMySql.Execute (sql)
    End If
    
    On Error GoTo errorListas
    
    sql = "SELECT cod FROM listas WHERE porDondeVoy = 1 LIMIT 0,1"
    
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        codLista = -1
    Else
        codLista = rs!cod
    End If
    
    rs.Close
    
    On Error GoTo errorListas
    
    sql = "SELECT * FROM ListaArticu WHERE idLista > " & codLista
    
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        sql = "UPDATE listas SET porDondeVoy = 0"
        
        connMySql.Execute (sql)
    End If
    
    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorListas
        
        sql = "SELECT * FROM listas WHERE cod = " & rs!idLista
        
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        If rs2.EOF Then
            On Error GoTo errorListas
            
            sql = "UPDATE listas SET porDondeVoy = 0"
               
            connMySql.Execute (sql)
            
            sql = "INSERT INTO listas " & "(cod,lista,porDondeVoy) VALUES('" & rs!idLista & "','" & StripString(rs!descripcion) & "', 1)"
                                     
            connMySql.Execute (sql)
        Else
            On Error GoTo errorListas
            
            sql = "UPDATE listas SET porDondeVoy = 0"
               
            connMySql.Execute (sql)
            
            sql = "UPDATE listas SET lista = '" & StripString(rs!descripcion) & "', porDondeVoy = 1 WHERE cod = " & rs!idLista
                                     
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        On Error GoTo errorListas
    
        sql = "SELECT * FROM ItemListaArticu WHERE XItem_IdLista = " & rs!idLista & " GROUP BY XItem_IdLista, XItem_IdArticu"
        
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        sqlInsert = "INSERT INTO listas_articulos " & "(cod_lista,cod_articulo) VALUES "
        If Not rs3.EOF Then
            On Error GoTo errorListas
            
            sql = "DELETE FROM listas_articulos WHERE cod_lista = " & rs!idLista
        
            connMySql.Execute (sql)
            
            numRegistros = 0
            Do While Not rs3.EOF
                DoEvents
                
                On Error GoTo errorListas
                
                sqlInsert = sqlInsert & "('" & rs3!XItem_IdLista & "','" & rs3!XItem_IdArticu & "')"
                                         
                rs3.MoveNext
                
                If rs3.EOF Or numRegistros = 1000 Then
                    sqlInsert = sqlInsert & ";"
                    
                    connMySql.Execute (sqlInsert)
                    
                    sqlInsert = "INSERT INTO listas_articulos " & "(cod_lista,cod_articulo) VALUES "
                
                    numRegistros = 0
                Else
                    sqlInsert = sqlInsert & ","
                End If
                
                numRegistros = numRegistros + 1
            Loop
            
            'connMySql.Execute (sqlInsert)
        End If
        
        rs3.Close
        
        rs.MoveNext
    Loop
    
    rs.Close
    
final:
    GoTo fin
 
errorListas:
    Sleep 1500
 
fin:
    'procesarTimerListasFechas
    
End Sub

Private Sub Timer_Listas_Fechas_Timer() '################ migrado #############################
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    On Error GoTo errorListas
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorListas
    
    sql = "SELECT * FROM ListaArticu WHERE fecha >= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) AND idLista <> " & codListaTienda
    
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorListas
        
        sql = "SELECT * FROM listas WHERE cod = " & rs!idLista
        
        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        If rs2.EOF Then
            On Error GoTo errorListas
            
            sql = "INSERT INTO listas " & "(cod,lista) VALUES('" & rs!idLista & "','" & StripString(rs!descripcion) & "')"
                                     
            connMySql.Execute (sql)
        Else
            On Error GoTo errorListas
            
            sql = "UPDATE listas SET lista = '" & StripString(rs!descripcion) & "' WHERE cod = " & rs!idLista
                                     
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        On Error GoTo errorListas
        
        sql = "DELETE FROM listas_articulos WHERE cod_lista = " & rs!idLista
        
        connMySql.Execute (sql)
        
        On Error GoTo errorListas
    
        sql = "SELECT * FROM ItemListaArticu WHERE XItem_IdLista = " & rs!idLista & " GROUP BY XItem_IdLista, XItem_IdArticu"
        
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        Do While Not rs3.EOF
            DoEvents
            
            On Error GoTo errorListas
            
            sql = "INSERT INTO listas_articulos " & "(cod_lista,cod_articulo) VALUES('" & _
                                     rs3!XItem_IdLista & "','" & _
                                     rs3!XItem_IdArticu & "')"
                                     
            connMySql.Execute (sql)
            
            rs3.MoveNext
        Loop
        
        rs3.Close
        
        rs.MoveNext
    Loop
    
    rs.Close
    
final:
    GoTo fin
 
errorListas:
    Sleep 1500
 
fin:
    
End Sub

Private Sub Timer_Categorias_PS_Timer() '###################33 migrado ##############################
    Dim sql As String
    Dim padre As String
    Dim padreId As String
        
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    On Error GoTo errorCategoriasPs
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorCategoriasPs
    
    sql = "select * from familia WHERE descripcion NOT IN ('ESPECIALIDAD', 'EFP', 'SIN FAMILIA') AND descripcion NOT LIKE '%ESPECIALIDADES%' AND descripcion NOT LIKE '%Medicamento%'"

    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorCategoriasPs
        
        sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON sf.IdSuperFamilia = fa.IdSuperFamilia " & _
               "WHERE fa.IdFamilia = " & rs!IdFamilia
               
        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If rs2.EOF Then
            padre = "<SIN PADRE>"
        Else
            padre = rs2!descripcion
        End If
        
        rs2.Close
        
        On Error GoTo errorCategoriasPs
        
        sql = "select * from ps_categorias where categoria='" & _
                LTrim(RTrim(StripString(rs!descripcion))) & "' AND padre = '" & LTrim(RTrim(StripString(padre))) & "'"

        rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

        If rs2.EOF Then
            On Error GoTo errorCategoriasPs
    
            sql = "select * from ps_categorias where padre = '" & LTrim(RTrim(StripString(padre))) & "'"
        
            rs3.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
            
            padreId = "NULL"
            If Not rs3.EOF Then
                If Not IsNull(rs3!prestashopPadreId) Then
                    padreId = rs3!prestashopPadreId
                End If
            End If
                                         
            rs3.Close
            
            On Error GoTo errorCategoriasPs
            
            sql = "INSERT INTO ps_categorias " & "(categoria, padre, prestashopPadreId) VALUES('" & _
                                        LTrim(RTrim(StripString(rs!descripcion))) & "', '" & LTrim(RTrim(StripString(padre))) & "', " & padreId & ")"
                                        
            connMySql.Execute (sql)
        End If

        rs2.Close

        rs.MoveNext
    Loop

    rs.Close
    
final:
    GoTo fin
 
errorCategoriasPs:
    Sleep 1500
 
fin:
End Sub

Private Sub Timer_Lista_Tienda_Timer() ''''''''''''''''''''''''' Migrado +++++++++++++++++++++++++++++++
    Dim precio As String
    Dim pcoste As String
    Dim pvpsiva As String
    Dim stock As String
    Dim stockMinimo As String
    Dim stockMaximo As String
    Dim desc As String
    Dim laboratorio As String
    Dim nombreLaboratorio As String
    Dim descripcion As String
    Dim present As String
    Dim activo As String
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Set rs6 = New ADODB.Recordset
    
    If codListaTienda > 0 Then
    
        On Error GoTo errorListas
        
        establecerConexionesBasesDatos
        
        On Error GoTo errorListas
        
        sql = "SELECT * FROM ListaArticu WHERE fecha >= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) AND idLista = " & codListaTienda
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs.EOF Then
            On Error GoTo errorListas
            
            sql = "SELECT * FROM listas WHERE cod = " & rs!idLista
            
            rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
            
            If rs2.EOF Then
                On Error GoTo errorListas
                
                sql = "INSERT INTO listas " & "(cod,lista) VALUES('" & rs!idLista & "','" & StripString(rs!descripcion) & "')"
                                         
                connMySql.Execute (sql)
            Else
                On Error GoTo errorListas
                
                sql = "UPDATE listas SET lista = '" & StripString(rs!descripcion) & "' WHERE cod = " & rs!idLista
                                         
                connMySql.Execute (sql)
            End If
            
            rs2.Close
            
            On Error GoTo errorListas
            
            sql = "DELETE FROM listas_articulos WHERE cod_lista = " & rs!idLista
            
            connMySql.Execute (sql)
            
            On Error GoTo errorListas
        
            sql = "SELECT * FROM ItemListaArticu WHERE XItem_IdLista = " & rs!idLista & " GROUP BY XItem_IdLista, XItem_IdArticu"
            
            rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
            Do While Not rs3.EOF
                DoEvents
                
                On Error GoTo errorListas
                
                sql = "INSERT INTO listas_articulos " & "(cod_lista,cod_articulo) VALUES('" & _
                                         rs3!XItem_IdLista & "','" & _
                                         rs3!XItem_IdArticu & "')"
                                         
                connMySql.Execute (sql)
                
                On Error GoTo errorListas
                
                sql = "select a.*, t.Piva AS iva from articu a INNER JOIN Tablaiva t ON t.IdTipoArt = a.XGrup_IdGrupoIva AND t.IdTipoPro = '05' " & _
                      "INNER JOIN ItemListaArticu li ON li.XItem_IdArticu = a.idArticu AND li.XItem_IdLista = " & codListaTienda & " " & _
                      "WHERE a.idArticu = '" & rs3!XItem_IdArticu & "'"
                      
                rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                On Error GoTo errorListas
                
                sql = "select * from medicamentos where cod_nacional ='" & rs5!IdArticu & "'"
            
                rs2.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
                
                On Error GoTo errorListas
                
                sql = "select * from familia where IdFamilia='" & rs5!XFam_IdFamilia & "'"
            
                rs6.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
                precio = Replace(rs5!pvp, ",", ".")
                
                pcoste = Replace(rs5!puc, ",", ".")
                
                pvpsiva = (rs5!pvp * 100) / (rs5!iva + 100)
                
                pvpsiva = Round(pvpsiva, 2)
                
                pvpsiva = Replace(pvpsiva, ",", ".")
                
                stock = rs5!StockActual
                
                stockMinimo = rs5!stockMinimo
                
                stockMaximo = rs5!stockMaximo
                
                desc = LTrim(RTrim(StripString(rs5!descripcion)))
                
                descripcion = ""
        
                If IsNull(rs6!descripcion) Or Len(rs6!descripcion) = 0 Then
                    superfamilia = ""
                Else
                    On Error GoTo errorListas
                    
                    sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                            " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(rs6!descripcion) & "'"
                            
                    rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs4.EOF Then
                        superfamilia = ""
                    Else
                        superfamilia = rs4!descripcion
                    End If
                    
                    rs4.Close
                End If
                
                On Error GoTo errorListas
                    
                sql = "SELECT * FROM sinonimo WHERE IdArticu = '" & rs5!IdArticu & "'"
        
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                If rs4.EOF Then
                    cod_barras = "847000" + CerosIzq(rs5!IdArticu, 6)
                Else
                    cod_barras = rs4!Sinonimo
                End If
                
                rs4.Close
                
                On Error GoTo errorListas
                        
                sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs5!proveedorHabitual & "'"
               
                rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                If rs4.EOF Then
                    proveedor = ""
                Else
                    proveedor = rs4!fis_nombre
                End If
        
                rs4.Close
        
                If Len(Trim(rs5!laboratorio)) > 0 Then
                    On Error GoTo errorListas
                    
                    sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(rs5!laboratorio) & "'"
                    
                    rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs4.EOF Then
                        rs4.Close
                
                        On Error GoTo errorListas
                    
                        sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(rs5!laboratorio) & "'"
                        
                        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                        
                        If rs4.EOF Then
                            nombreLaboratorio = "<Sin Laboratorio>"
                        Else
                            nombreLaboratorio = rs4!nombre
                        End If
                    Else
                        nombreLaboratorio = rs4!nombre
                    End If
                    
                    rs4.Close
                Else
                    nombreLaboratorio = "<Sin Laboratorio>"
                End If
                
                On Error GoTo errorListas
                
                sql = "SELECT * FROM ESPEPARA WHERE CODIGO = '" & rs5!IdArticu & "'"
                
                rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                
                If rs4.EOF Then
                    present = ""
                Else
                    present = rs4!presentacion
                End If
                
                rs4.Close
                
                On Error GoTo errorListas
                
                sql = "SELECT t.TEXTO FROM TEXTOS t INNER JOIN TEXTOSESPE te ON te.CODIGOTEXTO = t.CODIGOTEXTO " & _
                       "WHERE te.CODIGOESPEPARA = '" & rs5!IdArticu & "' ORDER BY te.CODIGOEPIGRAFE"
                       
                rs4.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                    
                If rs4.EOF Then
                    descripcion = ""
                Else
                    Do While Not rs4.EOF
                        If descripcion = "" Then
                            descripcion = rs4!TEXTO
                        Else
                            descripcion = descripcion & " <br> " & rs4!TEXTO
                        End If
                    
                        rs4.MoveNext
                    Loop
                    
                    If Len(descripcion) < 30000 Then
                        descripcion = StripString(quitarCaracterCadena(Replace(descripcion, vbCrLf, "<br>"), Chr(0)))
                    Else
                        descripcion = ""
                    End If
                End If
                
                rs4.Close
                
                If CBool(rs5!baja) Then
                    activo = 0
                    baja = 1
                Else
                    activo = 1
                    baja = 0
                End If
                               
                If rs2.EOF Then
                    On Error GoTo errorListas
                
                    sql = "INSERT INTO medicamentos " & "(cod_barras,cod_nacional,nombre,superFamilia,familia,precio,descripcion,laboratorio,nombre_laboratorio,proveedor,pvpSinIva,iva,stock,puc,stockMinimo,stockMaximo,presentacion,descripcionTienda,activoPrestashop,actualizadoPS,baja) " & _
                            "VALUES('" & StripString(cod_barras) & "','" & _
                            StripString(rs5!IdArticu) & "','" & _
                            LTrim(RTrim(StripString(desc))) & "','" & _
                            LTrim(RTrim(StripString(superfamilia))) & "','" & _
                            LTrim(RTrim(StripString(rs6!descripcion))) & "','" & _
                            precio & "','" & _
                            LTrim(RTrim(StripString(desc))) & "','" & _
                            LTrim(RTrim(StripString(rs5!laboratorio))) & "','" & _
                            LTrim(RTrim(StripString(nombreLaboratorio))) & "','" & _
                            LTrim(RTrim(StripString(proveedor))) & "','" & _
                            pvpsiva & "','" & _
                            rs5!iva & "','" & _
                            stock & "','" & _
                            pcoste & "','" & _
                            stockMinimo & "','" & _
                            stockMaximo & "','" & _
                            LTrim(RTrim(StripString(present))) & "','" & _
                            LTrim(RTrim(descripcion)) & "', " & _
                            activo & ", 1, " & baja & ")"
                 
                    connMySql.Execute (sql)
                Else
                    If (LTrim(RTrim(StripString(desc))) <> rs2!nombre Or LTrim(RTrim(StripString(rs6!descripcion))) <> rs2!familia _
                        Or precio <> rs2!precio Or LTrim(RTrim(StripString(rs5!laboratorio))) <> rs2!laboratorio _
                        Or rs5!iva <> rs2!iva Or stock <> rs2!stock Or pcoste <> rs2!puc _
                        Or LTrim(RTrim(StripString(present))) <> rs2!presentacion _
                        Or LTrim(RTrim(descripcion)) <> rs2!descripcion) Then
                        
                        On Error GoTo errorListas
                        
                        sql = "UPDATE medicamentos SET cod_barras = '" & StripString(cod_barras) & "', nombre = '" & LTrim(RTrim(StripString(desc))) & "', superFamilia = '" & LTrim(RTrim(StripString(superfamilia))) & "', familia = '" & LTrim(RTrim(StripString(rs6!descripcion))) & "', " & _
                               "precio = '" & precio & "', descripcion = '" & LTrim(RTrim(StripString(desc))) & "', laboratorio = '" & LTrim(RTrim(StripString(rs5!laboratorio))) & "', " & _
                               "nombre_laboratorio = '" & LTrim(RTrim(StripString(nombreLaboratorio))) & "', proveedor = '" & LTrim(RTrim(StripString(proveedor))) & "', " & _
                               "iva = '" & rs5!iva & "', pvpSinIva = '" & pvpsiva & "', stock = " & stock & ", puc = '" & pcoste & "', stockMinimo = " & stockMinimo & ", " & _
                               "stockMaximo = " & stockMaximo & ", " & _
                               "presentacion = '" & LTrim(RTrim(StripString(present))) & "', descripcionTienda = '" & LTrim(RTrim(descripcion)) & "', cargadoPS = 0, actualizadoPS = 1, activoPrestashop = " & activo & ", baja = " & baja & " " & _
                               " WHERE cod_nacional = '" & rs5!IdArticu & "'"
                               
                        connMySql.Execute (sql)
                    End If
                End If
                
                rs2.Close
                
                rs6.Close
             
                rs5.Close
                
                rs3.MoveNext
            Loop
            
            rs3.Close
            
            On Error GoTo errorListas
        
            sql = "UPDATE ListaArticu SET fecha = DATEADD(dd, -1, DATEDIFF(dd, 0, GETDATE())) WHERE idLista = " & codListaTienda
            
            connSqlServer.Execute (sql)
        End If
        
    End If
    
final:
    GoTo fin
 
errorListas:
    Sleep 1500
 
fin:

End Sub

Private Sub Timer_Pedidos_Timer() '''''''''''''''''''' Migrado ++++++++++++++++++++++++++++++++
    Dim familia As String
    Dim pvp As String
    Dim puc As String
    Dim codLaboratorio As String
    Dim nombreLaboratorio As String
    Dim superfamilia As String
    Dim Fecha As String
    Dim FechaPedido As String
    Dim importePvp As String
    Dim importePuc As String

    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Set rs6 = New ADODB.Recordset
    
    On Error GoTo errorPedidos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorPedidos
    
    sql = "SELECT TABLE_NAME AS tipo From information_schema.TABLES WHERE TABLE_SCHEMA = '" & baseRemoto & "' AND TABLE_NAME = 'pedidos'"
    
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If rs.EOF Then
        On Error GoTo errorPedidos
        
        sql = "CREATE TABLE IF NOT EXISTS `lineas_pedidos` (" & _
                "`id` bigint(255) unsigned NOT NULL AUTO_INCREMENT," & _
                "`fechaPedido` datetime DEFAULT NULL," & _
                "`idPedido` bigint(255) DEFAULT NULL," & _
                "`idLinea` bigint(255) DEFAULT NULL," & _
                "`cod_nacional` bigint(255) DEFAULT NULL," & _
                "`descripcion` varchar(255) DEFAULT NULL," & _
                "`familia` varchar(255) DEFAULT NULL," & _
                "`superFamilia` varchar(255) DEFAULT NULL," & _
                "`cantidad` int(11) DEFAULT NULL," & _
                "`pvp` float DEFAULT NULL," & _
                "`puc` float DEFAULT NULL," & _
                "`cod_laboratorio` varchar(50) DEFAULT NULL," & _
                "`laboratorio` varchar(255) DEFAULT NULL," & _
                "PRIMARY KEY (`id`)" & _
                ") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
                
        connMySql.Execute (sql)
        
        On Error GoTo errorPedidos
        
        sql = "CREATE TABLE IF NOT EXISTS `pedidos` (" & _
              "`id` bigint(255) unsigned NOT NULL AUTO_INCREMENT," & _
              "`idPedido` bigint(255) DEFAULT NULL," & _
              "`fechaPedido` datetime DEFAULT NULL," & _
              "`hora` datetime DEFAULT NULL," & _
              "`numLineas` int(11) DEFAULT NULL," & _
              "`importePvp` float DEFAULT NULL," & _
              "`importePuc` float DEFAULT NULL," & _
              "`idProveedor` varchar(50) DEFAULT NULL," & _
              "`proveedor` varchar(255) DEFAULT NULL," & _
              "`trabajador` varchar(255) DEFAULT NULL," & _
              "`sistema` varchar(50) DEFAULT NULL," & _
              "PRIMARY KEY (`id`)" & _
              ") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
                
        connMySql.Execute (sql)
    End If
    
    rs.Close
    
    Dim FieldExistsInRS As Boolean
    Dim oField
    
    FieldExistsInRS = False
    
    On Error GoTo errorPedidos
        
    sql = "SELECT * from lineas_pedidos LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "fechaPedido" Then
            FieldExistsInRS = True
        End If
    Next
    
    rs.Close
    
    If FieldExistsInRS = False Then
        sql = "ALTER TABLE lineas_pedidos ADD fechaPedido DATETIME AFTER id;"
        
        connMySql.Execute (sql)
    End If
    
    On Error GoTo errorPedidos
    
    sql = "select * from pedidos order by idPedido Desc Limit 0,1"
   
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic

    If rs.EOF Then
        rs.Close
        
        On Error GoTo errorPedidos
        
        sql = "SELECT * From Recep WHERE YEAR(Fecha) >= 2015 Order by IdRecepcion ASC"

        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    Else
        pedido = rs!IdPedido
        
        rs.Close
        
        On Error GoTo errorPedidos
        
        sql = "SELECT * From Recep WHERE IdRecepcion >= " & pedido & " AND YEAR(Fecha) >= 2015 Order by IdRecepcion ASC"
    
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    End If

    Do While Not rs.EOF
        DoEvents
        
        FechaPedido = Format(rs!Hora, "yyyy-MM-dd HH:mm:ss")
     
        Fecha = Format(Now, "yyyy-MM-dd HH:mm:ss")
        
        idProveedor = rs!XProv_IdProveedor
        
        On Error GoTo errorPedidos
        
        sql = "SELECT * From proveedor WHERE IDPROVEEDOR = '" & idProveedor & "'"
    
        rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs3.EOF Then
            nombreProveedor = LTrim(RTrim(StripString(rs3!fis_nombre)))
        Else
            nombreProveedor = ""
        End If
        
        rs3.Close
        
        On Error GoTo errorPedidos
            
        sql = "SELECT * FROM vendedor WHERE IdVendedor='" & rs!XVend_IdVendedor & "'"

        rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs5.EOF Then
            trabajador = ""
        Else
            trabajador = rs5!nombre
        End If
        
        rs5.Close
        
        On Error GoTo errorPedidos
        
        sql = "select * from pedidos where idPedido = " & rs!IdRecepcion
        
        rs6.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorPedidos
        
        sql = "SELECT ISNULL(COUNT(IdNLinea),0) AS numLineas, ISNULL(SUM(recibidas*ImportePvp),0) AS importePvp, ISNULL(SUM(importe),0) AS importePuc " & _
              "FROM LINEARECEP WHERE IdRecepcion = " & rs!IdRecepcion & " AND Recibidas <> 0"
        
        rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        numLineas = rs5!numLineas
        
        importePvp = Replace(rs5!importePvp, ",", ".")
        
        importePuc = Replace(rs5!importePuc, ",", ".")
        
        rs5.Close
        
        If rs6.EOF And numLineas > 0 Then
            On Error GoTo errorPedidos
             
            sql = "INSERT INTO pedidos " & "(idPedido,fechaPedido,hora,numLineas,importePvp,importePuc,idProveedor,proveedor,trabajador) VALUES('" & _
                              rs!IdRecepcion & "','" & _
                              FechaPedido & "','" & _
                              Fecha & "','" & _
                              numLineas & "','" & _
                              importePvp & "','" & _
                              importePuc & "','" & _
                              idProveedor & "','" & _
                              nombreProveedor & "','" & _
                              trabajador & "')"
                             
             connMySql.Execute (sql)
        End If
        
        rs6.Close
        
        If numLineas > 0 Then
            On Error GoTo errorPedidos
         
            sql = "select * from LINEARECEP where IdRecepcion = " & rs!IdRecepcion & " AND Recibidas <> 0"
    
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
            Do While Not rs2.EOF
                If (Trim(rs2!XArt_IdArticu) <> "" And Not IsNull(rs2!XArt_IdArticu)) Then
                    On Error GoTo errorPedidos
                    
                    sql = "select * from articu where IdArticu='" & rs2!XArt_IdArticu & "'"
            
                    rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                    
                    If rs3.EOF Then
                        familia = ""
                        superfamilia = ""
                        pvp = 0
                        puc = 0
                        codLaboratorio = ""
                        nombreLaboratorio = "<Sin Laboratorio>"
                    Else
                        On Error GoTo errorPedidos
                        
                        sql = "select * from familia where IdFamilia='" & rs3!XFam_IdFamilia & "'"
                
                        rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                         
                        If rs4.EOF Then
                            familia = ""
                        Else
                            familia = rs4!descripcion
                        End If
                
                        rs4.Close
                        
                        If IsNull(familia) Then
                            familia = "<Sin Clasificar>"
                        Else
                            If Len(familia) = 0 Then
                                familia = "<Sin Clasificar>"
                            End If
                        End If
                
                        If IsNull(rs3!laboratorio) Then
                            codLaboratorio = ""
                        Else
                            codLaboratorio = rs3!laboratorio
                        End If
                
                        If familia = "<Sin Clasificar>" Then
                            superfamilia = "<Sin Clasificar>"
                        Else
                            On Error GoTo errorPedidos
                            
                            sql = "SELECT sf.Descripcion FROM SuperFamilia sf INNER JOIN FamiliaAux fa ON fa.IdSuperFamilia = sf.IdSuperFamilia " & _
                                    " INNER JOIN Familia f ON f.IdFamilia = fa.IdFamilia WHERE f.Descripcion = '" & SqlSafe(familia) & "'"
                                    
                            rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                            
                            If rs5.EOF Then
                                superfamilia = "<Sin Clasificar>"
                            Else
                                superfamilia = rs5!descripcion
                            End If
                            
                            rs5.Close
                        End If
                            
                        If Len(Trim(codLaboratorio)) > 0 Then
                            On Error GoTo errorPedidos
                            
                            sql = "SELECT * FROM LABOR WHERE CODIGO = '" & SqlSafe(codLaboratorio) & "'"
                            
                            rs5.Open sql, connSqlServerBP, adOpenDynamic, adLockBatchOptimistic
                            
                            If rs5.EOF Then
                                rs5.Close
                
                                On Error GoTo errorPedidos
                            
                                sql = "SELECT * FROM laboratorio WHERE codigo = '" & SqlSafe(codLaboratorio) & "'"
                                
                                rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                                
                                If rs5.EOF Then
                                    nombreLaboratorio = "<Sin Laboratorio>"
                                Else
                                    nombreLaboratorio = rs5!nombre
                                End If
                            Else
                                nombreLaboratorio = rs5!nombre
                            End If
                            
                            rs5.Close
                        Else
                            nombreLaboratorio = "<Sin Laboratorio>"
                        End If
                    End If
                    
                    pvp = Replace(rs2!importePvp, ",", ".")
                    
                    puc = Replace(rs2!importePuc, ",", ".")
              
                    On Error GoTo errorPedidos
            
                    sql = "select * from lineas_pedidos where idPedido = " & rs2!IdRecepcion & " AND idLinea = " & rs2!idnlinea
            
                    rs6.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
               
                    If rs6.EOF And Not rs3.EOF Then
                        On Error GoTo errorPedidos
                 
                        sql = "INSERT INTO lineas_pedidos " & "(fechaPedido,idPedido,idLinea,cod_nacional,descripcion,familia,superFamilia,cantidad,pvp,puc,cod_laboratorio,laboratorio) VALUES('" & _
                                          FechaPedido & "','" & _
                                          rs2!IdRecepcion & "','" & _
                                          rs2!idnlinea & "','" & _
                                          StripString(rs3!IdArticu) & "','" & _
                                          LTrim(RTrim(StripString(rs3!descripcion))) & "','" & _
                                          LTrim(RTrim(StripString(familia))) & "','" & _
                                          LTrim(RTrim(StripString(superfamilia))) & "','" & _
                                          rs2!Recibidas & "','" & _
                                          pvp & "','" & _
                                          puc & "','" & _
                                          LTrim(RTrim(StripString(codLaboratorio))) & "','" & _
                                          LTrim(RTrim(StripString(nombreLaboratorio))) & "')"
                                         
                         connMySql.Execute (sql)
                    End If
             
                    rs3.Close
                    rs6.Close
                End If
        
                rs2.MoveNext
            Loop
            
            rs2.Close
        
        End If
   
        rs.MoveNext
    Loop

    rs.Close

final:
    GoTo fin
 
errorPedidos:
    Sleep 1500
 
fin:
End Sub

Private Sub Timer_Sinonimos_Timer()	'################### migrado ##############################
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    On Error GoTo errorSinonimos
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorSinonimos
    
    sql = "SELECT data_type AS tipo " & _
          "From information_schema.Columns " & _
          "WHERE TABLE_SCHEMA = '" & baseRemoto & "' AND TABLE_NAME = 'sinonimos' AND COLUMN_NAME = 'cod_barras'"
          
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If Not rs.EOF Then
        If LCase(rs!tipo) <> "varchar" Then
            sql = "ALTER TABLE sinonimos MODIFY COLUMN cod_nacional VARCHAR(255);"
            
            connMySql.Execute (sql)
            
            sql = "ALTER TABLE sinonimos MODIFY COLUMN cod_barras VARCHAR(255);"
            
            connMySql.Execute (sql)
        End If
    End If
    
    rs.Close
    
    On Error GoTo errorSinonimos
    
    sql = "SELECT * FROM sinonimos LIMIT 0,1"
    
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    insertar = False
    
    If rs.EOF Then
        insertar = True
    Else
        If Format(Now, "HHmm") = "1000" Or Format(Now, "HHmm") = "1230" Or Format(Now, "HHmm") = "1730" Or Format(Now, "HHmm") = "1930" Then
            insertar = True
        End If
    End If
    
    rs.Close
    
    If insertar Then
        On Error GoTo errorSinonimos
            
        sql = "SELECT * FROM Sinonimo"
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        sqlInsert = "INSERT INTO sinonimos " & "(cod_barras,cod_nacional) VALUES "
            
        sql = "TRUNCATE sinonimos"
    
        connMySql.Execute (sql)
        
        numRegistros = 0
        
        Do While Not rs.EOF
            DoEvents
            
            On Error GoTo errorSinonimos
                    
            sqlInsert = sqlInsert & "('" & LTrim(RTrim(StripString(rs!Sinonimo))) & "','" & LTrim(RTrim(StripString(rs!IdArticu))) & "')"
                                 
            rs.MoveNext
                    
            If rs.EOF Or numRegistros = 1000 Then
                sqlInsert = sqlInsert & ";"
                
                connMySql.Execute (sqlInsert)
                
                sqlInsert = "INSERT INTO sinonimos " & "(cod_barras,cod_nacional) VALUES "
                
                numRegistros = 0
            Else
                sqlInsert = sqlInsert & ","
            End If
            
            numRegistros = numRegistros + 1
        Loop
        
        On Error GoTo errorSinonimos
        
        'Set objfso = CreateObject("Scripting.FileSystemObject")

        'Set objOutput = objfso.CreateTextFile("c:\sinonimos.sql")
        
        'objOutput.Write sqlInsert
        
        'objOutput.Close
        
        'connMySql.Execute (sqlInsert)
        
        rs.Close
    End If
    
final:
    GoTo fin
 
errorSinonimos:
    Sleep 1500
 
fin:
    'procesarTimerSinonimos
End Sub

Private Sub Timer_Actualizar_Pendiente_Puntos_Timer() '#################333 current ##########################
    Dim sql As String
    
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    Set rs1 = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    On Error GoTo errorActualizarPP
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorActualizarPP
    
    sql = "SELECT * FROM pendiente_puntos WHERE redencion IS NULL AND YEAR(fechaVenta) >= 2015 GROUP BY idventa ORDER BY idventa ASC LIMIT 0,1000"
    
    rs1.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    Do While Not rs1.EOF
        IdVenta = rs1!IdVenta
        
        On Error GoTo errorActualizarPP
        
        sql = "SELECT * FROM venta WHERE IdVenta = " & IdVenta & " ORDER BY IdVenta ASC"
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
        If Not rs.EOF Then
            tipoPago = rs!TipoVenta
            
            On Error GoTo errorActualizarPP
     
            sql = "SELECT * FROM lineaventa WHERE IdVenta = " & rs!IdVenta
            
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
            descuentoVenta = False
        
            On Error GoTo errorActualizarPP
            
            Do While Not rs2.EOF
                DoEvents
                
                dtoVenta = 0
                If Not descuentoVenta Then
                    dtoVenta = Replace(rs!DescuentoOpera, ",", ".")
                    descuentoVenta = True
                End If
                
                dtoLinea = Replace(rs2!DescuentoLinea, ",", ".")
                
                On Error GoTo errorActualizarPP
                
                sql = "SELECT * FROM LineaVentaReden WHERE IdVenta = " & rs2!IdVenta & " AND IdNLinea = " & rs2!idnlinea
    
                rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
                
                redencion = 0
                If Not rs3.EOF Then
                    redencion = Replace(rs3!redencion, ",", ".")
                End If
                
                rs3.Close
                
                On Error GoTo errorActualizarPP
                
                sql = "SELECT * FROM articu WHERE IdArticu='" & rs2!codigo & "'"
    
                rs3.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
                If rs3.EOF Then
                    proveedor = ""
                Else
                    On Error GoTo errorActualizarPP
                
                    sql = "SELECT * FROM Proveedor WHERE IDProveedor = '" & rs3!proveedorHabitual & "'"
                   
                    rs4.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
            
                    If rs4.EOF Then
                        proveedor = ""
                    Else
                        proveedor = rs4!fis_nombre
                    End If
            
                    rs4.Close
                End If
                
                On Error GoTo errorActualizarPP
                
                sql = "UPDATE pendiente_puntos SET tipoPago = '" & tipoPago & "', proveedor = '" & proveedor & "', " & _
                      "dtoLinea = '" & dtoLinea & "', dtoVenta = '" & dtoVenta & "', redencion = '" & redencion & "' " & _
                      "WHERE IdVenta='" & rs2!IdVenta & "' AND Idnlinea= '" & rs2!idnlinea & "'"
                    
                connMySql.Execute (sql)
                
                rs2.MoveNext
                
                rs3.Close
            Loop
            
            rs2.Close
        Else
            sql = "UPDATE pendiente_puntos SET tipoPago = 'C', redencion = 0 " & _
                  "WHERE IdVenta='" & IdVenta & "'"
            
            connMySql.Execute (sql)
        End If
    
        rs1.MoveNext
        
        rs.Close
    Loop
        
    rs1.Close
        
final:
    GoTo fin
 
errorActualizarPP:
    Sleep 1500
 
fin:
    'procesarTimerActualizarPP
End Sub

Private Sub Timer_Actualizar_Productos_Borrados_Timer()
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Dim FieldExistsInRS As Boolean
    Dim oField
    
    FieldExistsInRS = False
    
    On Error GoTo errorActualizarPB
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorActualizarPB
    
    sql = "SELECT * from medicamentos LIMIT 0,1;"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    For Each oField In rs.Fields
        If oField.Name = "web" Then
            FieldExistsInRS = True
        End If
    Next
    
    rs.Close
    
    sql = "SELECT * FROM configuracion WHERE campo = 'porDondeVoyBorrar'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorActualizarPB
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('porDondeVoyBorrar', '0')"
        
        connMySql.Execute (sql)
        
        codArticu = 0
    Else
        codArticu = rs!valor
    End If
    
    rs.Close
    
    On Error GoTo errorActualizarPB
    
    If FieldExistsInRS Then
        sql = "SELECT cod_nacional FROM medicamentos WHERE web = 0 AND cod_nacional >= " & codArticu & " ORDER BY cod_nacional ASC LIMIT 0,1000"
    Else
        sql = "SELECT cod_nacional FROM medicamentos WHERE cod_nacional >= " & codArticu & " ORDER BY cod_nacional ASC LIMIT 0,1000"
    End If
    
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    If Not rs.EOF Then
        rs.MoveNext
        
        If rs.EOF Then
            rs.MovePrevious
            
            On Error GoTo errorActualizarPB
        
            sql = "SELECT * FROM articu WHERE IdArticu = '" & CerosIzq(rs!cod_nacional, 6) & "'"
        
            rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
            If rs2.EOF Then
                On Error GoTo errorActualizarPB
                
                sql = "DELETE FROM medicamentos WHERE cod_nacional = '" & rs!cod_nacional & "'"
                
                connMySql.Execute (sql)
            End If
            
            rs2.Close
            
            On Error GoTo errorActualizarPB
            
            sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoyBorrar'"
            
            connMySql.Execute (sql)
            
            rs.MoveNext
        Else
            rs.MovePrevious
        End If
    Else
        On Error GoTo errorActualizarPB
        
        sql = "UPDATE configuracion SET valor = '0' WHERE campo = 'porDondeVoyBorrar'"
            
        connMySql.Execute (sql)
    End If
    
    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorActualizarPB
        
        sql = "SELECT * FROM articu WHERE IdArticu = '" & CerosIzq(rs!cod_nacional, 6) & "'"
    
        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
        If rs2.EOF Then
            On Error GoTo errorActualizarPB
            
            sql = "DELETE FROM medicamentos WHERE cod_nacional = '" & rs!cod_nacional & "'"
            
            connMySql.Execute (sql)
        End If
        
        rs2.Close
        
        On Error GoTo errorActualizarPB
        
        sql = "UPDATE configuracion SET valor = '" & rs!cod_nacional & "' WHERE campo = 'porDondeVoyBorrar'"
               
        connMySql.Execute (sql)
        
        rs.MoveNext
    Loop
    
    rs.Close
    
final:
    GoTo fin
 
errorActualizarPB:
    Sleep 1500
 
fin:
    'procesarTimerActualizarProductosBorrados
End Sub

Private Sub Timer_Actualizar_Entregas_Clientes_Timer()
    Dim puesto As String
    Dim numero As String
    Dim dni As String
    Dim fechaVenta As String
    
    Dim sql As String
    
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    Dim rs6 As ADODB.Recordset
    Dim rs7 As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    Set rs5 = New ADODB.Recordset
    
    Set rs6 = New ADODB.Recordset
    
    Set rs7 = New ADODB.Recordset
    
    On Error GoTo errorActualizarEC
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorActualizarEC
    
    sql = "SELECT * FROM configuracion WHERE campo = 'porDondeEntregasClientes'"
    rs.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    On Error GoTo errorActualizarEC
    
    If rs.EOF Then
        sql = "INSERT INTO configuracion (campo, valor) VALUES ('porDondeEntregasClientes', '0')"
        
        connMySql.Execute (sql)
        
        On Error GoTo errorActualizarEC
    
        sql = "SELECT * FROM entregas_clientes GROUP BY idventa ORDER BY idventa DESC LIMIT 0,1"
        
        rs7.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
        If rs7.EOF Then
            rs7.Close
            
            On Error GoTo errorActualizarEC
        
            sql = "SELECT * FROM pendiente_puntos ORDER BY idventa DESC LIMIT 0,1"
               
            rs7.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
            If rs7.EOF Then
                venta = 0
            Else
                venta = rs7!IdVenta
            End If
        
            rs7.Close
        Else
            venta = rs7!IdVenta
        End If
        
        On Error GoTo errorActualizarEC
        
        sql = "UPDATE configuracion SET valor = '" & venta & "' WHERE campo = 'porDondeEntregasClientes'"
        
        connMySql.Execute (sql)
    Else
        venta = rs!valor
    End If
    
    rs.Close
    
    On Error GoTo errorActualizarEC
    
    sql = "SELECT v.* FROM venta v INNER JOIN lineaventavirtual lvv ON lvv.idventa = v.idventa AND (lvv.codigo = 'Pago' OR lvv.codigo = 'A Cuenta') " & _
          "WHERE v.ejercicio >= 2015 AND v.IdVenta <= " & venta & " ORDER BY v.IdVenta DESC"
    
    rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

    Do While Not rs.EOF
        DoEvents
        
        On Error GoTo errorActualizarEC
            
        sql = "UPDATE configuracion SET valor = '" & rs!IdVenta & "' WHERE campo = 'porDondeEntregasClientes'"
        
        connMySql.Execute (sql)
      
        puesto = rs!Maquina
        
        tipoPago = rs!TipoVenta
    
        fechaVenta = Format(rs!FechaHora, "yyyy-MM-dd HH:mm:ss")
        
        dni = LTrim(RTrim(StripString(rs!XClie_IdCliente)))
                
        If IsNull(dni) Or Len(dni) = 0 Then
            dni = 0
        End If
        
        id_vendedor = rs!XVend_IdVendedor
            
        On Error GoTo errorActualizarEC
        
        sql = "SELECT * FROM vendedor WHERE IdVendedor='" & id_vendedor & "'"

        rs5.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic

        If rs5.EOF Then
            trabajador = "NO"
        Else
            trabajador = rs5!nombre
        End If
            
        rs5.Close
        
        ArrayFecha = Split(rs!FechaHora, "/")

        numero_dia = ArrayFecha(0)
        numero_mes = ArrayFecha(1)
        numero_anio = ArrayFecha(2)
        
        numero_Fecha = Split(numero_anio, " ")
        numero_anio = numero_Fecha(0)

        Fecha = Format(rs!FechaHora, "yyyyMMdd")
        
        On Error GoTo errorActualizarEC
        
        sql = "SELECT * FROM lineaventavirtual WHERE IdVenta='" & rs!IdVenta & "' AND (codigo = 'Pago' OR codigo = 'A Cuenta')"

        rs2.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
        
        On Error GoTo errorActualizarEC
        
        Do While Not rs2.EOF
            On Error GoTo errorActualizarEC
        
            sql = "SELECT * FROM entregas_clientes WHERE IdVenta='" & rs2!IdVenta & "' AND Idnlinea= '" & rs2!idnlinea & "'"
       
            rs4.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
        
            numero = Replace(rs2!importeneto, ",", ".")
            
            pvp = Replace(rs2!pvp, ",", ".")
                
            If rs4.EOF Then
                On Error GoTo errorActualizarEC
                
                sql = "INSERT INTO entregas_clientes " & "(idventa,idnlinea,codigo,descripcion,cantidad,precio,tipo,fecha,dni,puesto,trabajador,fechaEntrega,pvp) VALUES('" & _
                                 rs2!IdVenta & "','" & _
                                 rs2!idnlinea & "','" & _
                                 StripString(rs2!codigo) & "','" & _
                                 StripString(rs2!descripcion) & "','" & _
                                 rs2!cantidad & "','" & _
                                 numero & "','" & _
                                 rs2!tipoLinea & "','" & _
                                 Fecha & "','" & _
                                 dni & "','" & _
                                 puesto & "','" & _
                                 trabajador & "','" & _
                                 fechaVenta & "','" & _
                                 pvp & "')"
                                 
                connMySql.Execute (sql)
            End If
        
            rs4.Close
        
            rs2.MoveNext
        Loop
        
        rs.MoveNext
        
        rs2.Close
    Loop
 
    rs.Close
 
final:
    GoTo fin
 
errorActualizarEC:
    Sleep 1500
 
fin:

End Sub

Private Sub Timer_Actualizar_Recetas_Pendientes_Timer()
    Dim sql As String
    
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    Set rs1 = New ADODB.Recordset
    
    Set rs2 = New ADODB.Recordset
    
    Set rs3 = New ADODB.Recordset
    
    Set rs4 = New ADODB.Recordset
    
    On Error GoTo errorActualizarRP
    
    establecerConexionesBasesDatos
    
    On Error GoTo errorActualizarRP
    
    sql = "SELECT * FROM pendiente_puntos WHERE (recetaPendiente IS NULL OR recetaPendiente = 'D') " & _
          "AND YEAR(fechaVenta) >= 2016 GROUP BY idventa, idnlinea ORDER BY idventa ASC LIMIT 0,1000"
    
    rs1.Open sql, connMySql, adOpenDynamic, adLockBatchOptimistic
    
    Do While Not rs1.EOF
        DoEvents
        
        IdVenta = rs1!IdVenta
        
        idnlinea = rs1!idnlinea
        
        On Error GoTo errorActualizarRP
        
        sql = "SELECT * FROM lineaventa WHERE IdVenta = " & IdVenta & " AND IdNLinea = " & idnlinea & " "
        
        rs.Open sql, connSqlServer, adOpenDynamic, adLockBatchOptimistic
    
        If Not rs.EOF Then
            recetaPendiente = rs!recetaPendiente
                
            If IsNull(rs1!recetaPendiente) Then
                sql = "UPDATE pendiente_puntos SET recetaPendiente = '" & recetaPendiente & "' " & _
                      "WHERE IdVenta='" & IdVenta & "' AND Idnlinea= '" & idnlinea & "'"

                connMySql.Execute (sql)
            Else
                If recetaPendiente <> "D" Then
                    On Error GoTo errorActualizarRP
                    
                    sql = "UPDATE pendiente_puntos SET recetaPendiente = '" & recetaPendiente & "' " & _
                          "WHERE IdVenta='" & IdVenta & "' AND Idnlinea= '" & idnlinea & "'"
                        
                    connMySql.Execute (sql)
                End If
            End If
        Else
            If IsNull(rs1!recetaPendiente) Then
                On Error GoTo errorActualizarRP
                
                sql = "UPDATE pendiente_puntos SET recetaPendiente = 'C' " & _
                      "WHERE IdVenta='" & IdVenta & "' AND Idnlinea= '" & idnlinea & "'"
                
                connMySql.Execute (sql)
            End If
        End If
    
        rs1.MoveNext
        
        rs.Close
    Loop
        
    rs1.Close
        
final:
    GoTo fin
 
errorActualizarRP:
    Sleep 1500
 
fin:
    'procesarTimerActualizarPP
End Sub



