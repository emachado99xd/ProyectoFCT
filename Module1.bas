Attribute VB_Name = "Module1"
Global Base As New ADODB.Connection
Global RsCliente As New Recordset
Global RsDetalleFactura As New Recordset
Global RsFactura As New Recordset
Global RsProductos As New Recordset
Global RsTipoProducto As New Recordset
Global RsTemporal As New Recordset
Sub main()
    With Base
        .CursorLocation = adUseClient
        .Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
        Inicio.Show
    End With
End Sub
Sub Clientes()
    With RsCliente
        If .State = 1 Then .Close
            .Open "select * from Cliente", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub DetalleFactura()
    With RsDetalleFactura
        If .State = 1 Then .Close
            .Open "select * from Detalle_Factura", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub Factura()
    With RsFactura
         If .State = 1 Then .Close
            .Open "select * from Factura", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub Producto()
    With RsProductos
         If .State = 1 Then .Close
            .Open "select * from Producto", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub TipoProducto()
With RsTipoProducto
         If .State = 1 Then .Close
            .Open "select * from Tipo_Producto", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub Temporal()
    With RsTemporal
        If .State = 1 Then Close
            .Open "select * from Temporal", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

