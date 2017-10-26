VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ejecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Procesar(ByVal tipo As String)
    Dim iProceso As VBMTEjecutor.Proceso
    Set iProceso = New VBMTEjecutor.Proceso
    Load iProceso
    
    Select Case UCase(tipo)
            Case "PENDIENTE_PUNTOS"
                 iProceso.procesarTimerPendientePuntos
                 
            Case "PRODUCTOS_CRITICOS"
                iProceso.procesarTimerProductosCriticos
                    
            Case "ENCARGOS"
                iProceso.procesarTimerEncargos
                
            Case "FAMILIAS"
                iProceso.procesarTimerFamilias
                
            Case "CLIENTES"
                 'iProceso.actualizarClientes
                 iProceso.setCeroClientes
                 iProceso.procesarTimerClientes
                 
            Case "CONTROL_STOCK"
                iProceso.procesarTimerControlStockInicial
                'iProceso.procesarTimerControlStockFechas
            
            Case "CONTROL_SIN_STOCK"
                iProceso.procesarTimerControlSinStockInicial
                
            Case "CONTROL_STOCK_FECHAS_ENTRADA"
                iProceso.procesarTimerControlStockFechasEntrada
                
            Case "CONTROL_STOCK_FECHAS_SALIDA"
                iProceso.procesarTimerControlStockFechasSalida
                
            Case "LISTAS"
                'iProceso.deleteListas
                iProceso.procesarTimerListas
                
            Case "LISTA_TIENDA"
                'iProceso.deleteListas
                iProceso.procesarTimerListaTienda
                                
            Case "CATEGORIASPS"
                iProceso.procesarTimerCategorias
                
            Case "PEDIDOS"
                iProceso.procesarTimerPedidos
                
            Case "SINONIMOS"
                iProceso.procesarTimerSinonimos
                
            Case "ACTUALIZAR_PP"
                iProceso.procesarTimerActualizarPP
                
            Case "ACTUALIZAR_PRODUCTOS_BORRADOS"
                iProceso.procesarTimerActualizarPB
                
            Case "ACTUALIZAR_ENTREGAS_CLIENTES"
                iProceso.procesarTimerActualizarEntregasClientes
                
            Case "ACTUALIZAR_RECETAS_PENDIENTES"
                iProceso.procesarTimerActualizarRecetasPendientes
                
    End Select
    
    Set iProceso = Nothing
End Sub
