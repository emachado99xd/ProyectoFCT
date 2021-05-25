Attribute VB_Name = "Module1"
Global Base As New ADODB.Connection
Global RsCliente As New Recordset
Global RsDetalleFactura As New Recordset
Global RsFactura As New Recordset
Global RsProductos As New Recordset
Global RsTipoProducto As New Recordset

Sub main()
    With Base
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Base\Base.accdb;Persist Security Info=False"
        Form1.Show
    End With
End Sub
Sub Cliente()
    With RsCliente
        If .State = 1 Then .Close
            .Open "select * from Cliente", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub DetalleFactura()
    With RsDetalleFactura
        If .State = 1 Then .Close
            .Open "select * from Detalle Factura", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub Factura()
    With RsFactura
         If .State = 1 Then .Close
            .Open "select * from Factura", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub Producto()
    With RsProducto
         If .State = 1 Then .Close
            .Open "select * from Producto", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub TipoProducto()
With RsTipoProducto
         If .State = 1 Then .Close
            .Open "select * from Tipo Producto", Base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub



