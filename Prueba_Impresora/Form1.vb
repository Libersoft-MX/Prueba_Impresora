#Region "Importaciones"
Imports System.Drawing                  'Usado para el objeto Image
Imports System.Windows.Forms            'Usado para el DataGridView
Imports System.Drawing.Printing         'Usado para imprimir con PrintDocument

#End Region
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        eCenter.Alignment = StringAlignment.Center
        eCenter.LineAlignment = StringAlignment.Center
        eLeft.Alignment = StringAlignment.Near
        eLeft.LineAlignment = StringAlignment.Center
        eRight.Alignment = StringAlignment.Far
        eRight.LineAlignment = StringAlignment.Center
    End Sub

#Region "Declaraciones de Datos del ticket"
    '***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET *****************

    Private _Logotipo As Image = Nothing                            'Logotipo de la empresa         ID ->1
    Private _Empresa As String = "Aceros Inoxidables Refacciones y Equipos" 'Nombre de la empresa
    Private _Calle As String = "Calle Lino Merino #226"             'Nombre de la calle donde esta ubicada
    Private _Colonia As String = "Colonia Centro"                   'Nombre de la colonia
    Private _Ciudad As String = "Villahermosa Tab. Mex."            'Nombre del ciudad
    Private _Telefono As String = "314-99-06"                       'Telefono
    Private _CP As String = "86000"                                 'Código Postal
    Private _BarCode_Text As String = ""                            'Code 39
    Private _Barcode_Ima As Image = Nothing                         'Imagen del código de barra     ID ->0
    Private _Tabla As DataGridView = Nothing                        'Número del codigo de barra
    Private _Mensaje As String = "¡Gracias por su preferencia!"     'Mensaje de fin de ticket : "Gracias por su preferencia"
    Private _Total As String = "445.00"                             'Total dela venta
    Private _Correo As String = "plasticos_y_derivados@hotmail.com" 'Correo de la empresa
    Private _Cambio As String = "5.50"                              'Cambio de la venta
    Private _Efectivo As String = "500.00"                          'Efectivo con el que se pagó
    Private _Transaccion As String = "Operación"                    'Tipo de transacción que se está realizando
    Private _Tipo_pago As String = "Efectivo"                       'Forma de pago: Efectivo, Cheque
    Private _Fecha As String                                        'Fecha en que se registra la transacción
    Private _Hora As String                                         'Hora en que se registra la transacción

#End Region

#Region "Declaraciones de Funcionamiento de Impresión"
    '***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO *****
    Private WithEvents PD As New PrintDocument                      'Documento a imprimir
    Private PDBody As PrintPageEventArgs = Nothing                  'Cuerpo del documento
    Private _Art As String = "nombre_corto"                         'Indice de la columna articulo en el DataGridView
    Private _Cant As String = "cantidad"                            'Indice de la columna cantidad en el DataGridView
    Private _Sub As String = "subtotal"                             'Indice de la columna subtotal en el DataGridView
    Private _Precio As String = "precio"
    Private _Impresora As String = "POS-58"                         'Nombre de la impresora
    Private _Descuento As String = ""
    Private _ImagenPrint As Boolean = True                          'True imprime logotipo; false imprime código de barra
    Private _AnchoHoja As String = "195"                             'Ancho de la hoja de impresión
    Private _Espacio As Decimal = 5                                 'Espacio entre lineas
    Private _X As Integer = 0                                       'Posición X en la impresión
    Private _Y As Integer = 0                                       'Posición Y en la impresión
    Private AreaImpresion As Rectangle                              'Area de impresión
    Private Titulo_F As New Font("Arial", 12, FontStyle.Bold)       'Fuente de Titulo
    Private Encabezado_F As New Font("Arial", 9, FontStyle.Regular) 'Fuente de encabezado
    Private Cuerpo_F As New Font("Arial", 8, FontStyle.Regular)     'Fuente de cuerpo
    Private Columna_F As New Font("Arial", 8, FontStyle.Bold)       'Fuente de columna
    Private eCenter As New StringFormat()                           'Centra el texto
    Private eLeft As New StringFormat()                             'Alineación a la izquierda
    Private eRight As New StringFormat()                            'Alineación a la derecha
    Private _Aling As StringFormat                                  'Auxiliar para alineación
    Private _PrintView As New PrintPreviewDialog
#End Region

#Region "Operaciones basicas con la impresora"
    Private Function Imprimir() As Boolean
        Try
            If PD.PrinterSettings.IsValid Then
                PD.DocumentName = "Ticket"
                PD.PrinterSettings.PrinterName = _Impresora
                PD.PrintController = New StandardPrintController

                AddHandler PD.PrintPage, AddressOf PrintDocu_PrintPage

                '_PrintView.Document = PD
                '_PrintView.Show()
                PD.Print()
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox("¡Error al intentar imprimir!: " + ex.ToString, vbCritical)
            Return False
        End Try
    End Function
    Private Sub PrintDocu_PrintPage(sender As Object, e As PrintPageEventArgs)
        StartPrint(e)
        If Not IsNothing(_Logotipo) Then
            PrintImage(_Logotipo)
        End If
        If Not _Calle = "" Then
            PrintText(_Calle, Encabezado_F)
        End If
        If Not _Colonia = "" Then
            PrintText(_Colonia + " C.P. " + _CP, Encabezado_F)
        End If
        If Not _Ciudad = "" Then
            PrintText(_Ciudad, Encabezado_F)
        End If
        If Not _Telefono = "" Then
            PrintText("Tel. " + _Telefono, Encabezado_F)
        End If
        If Not _Correo = "" Then
            PrintText("E-mail: " + _Correo, Cuerpo_F)
        End If


        eSpace(2)
        If Not IsNothing(_Barcode_Ima) And Not _Transaccion = "Cotización" Then
            PrintText("Folio: " & _BarCode_Text, Columna_F)
        End If
        eSpace(2)
        PrintText("Transacción: " + _Transaccion, Columna_F)
        PrintBody()

        eSpace(2)
        If Not _Transaccion = "Cotización" Then
            PrintText("Pago en " + _Tipo_pago + " una sola exibición", Columna_F)
            eSpace(2)
            If _Tipo_pago = "Cheque" Then
                PrintText("Cheque: $" + Format((_Efectivo * 1), "##,##0.00"), Columna_F)
            Else
                PrintText("Efectivo: $" + Format((_Efectivo * 1), "##,##0.00"), Columna_F)
                PrintText("Cambio: $" + Format((_Efectivo - _Total), "##,##0.00"), Columna_F)
            End If

            eSpace(2)
        End If

        'If Not _Descuento = "" And Not -Descuento = "0" Then
        '  PrintText("Con el descuento aplicado has ahorrado $" + _Descuento, Columna_F)
        '   eSpace(2)
        'End If

        If Not IsNothing(_Barcode_Ima) And Not _Transaccion = "Cotización" Then
            PrintImage(_Barcode_Ima)
            eSpace(2)
        End If
        If Not _Mensaje = "" Then
            PrintText(_Mensaje, Columna_F)
        End If
        eSpace(6)
        PrintText(".", Cuerpo_F)
        e = EndPrint()
        'PD.Dispose()
    End Sub
    Private Sub PrintBody()
        Dim X1 As Integer
        Dim X2 As Integer
        Dim X3 As Integer
        Dim X4 As Integer
        Dim W1 As Integer
        Dim W2 As Integer
        Dim W3 As Integer
        Dim W4 As Integer
        Dim TF As Decimal
        Dim Lineas As Integer
        Dim Total_ As Decimal = 0
        If Not IsNothing(_Tabla) Then
            Lineas = _Tabla.RowCount
            TF = Cuerpo_F.GetHeight(PDBody.Graphics)
            W4 = PDBody.Graphics.MeasureString("88000", Cuerpo_F).Width
            W3 = W4
            W2 = PDBody.Graphics.MeasureString("500", Cuerpo_F).Width
            W1 = _AnchoHoja - (W4 + W2 + W3 + 10)
            X1 = 0
            X2 = W1 + 5
            X3 = X2 + W2
            X4 = X3 + W3

            eLine()
            AreaImpresion = New Rectangle(X1, _Y, W1, TF)
            PDBody.Graphics.DrawString("Articulo:", Columna_F, Brushes.Black, AreaImpresion, eLeft)
            AreaImpresion = New Rectangle(X2, _Y, W2, TF)
            PDBody.Graphics.DrawString("Cant.", Columna_F, Brushes.Black, AreaImpresion, eCenter)
            AreaImpresion = New Rectangle(X3, _Y, W3, TF)
            PDBody.Graphics.DrawString("P.U.", Columna_F, Brushes.Black, AreaImpresion, eCenter)
            AreaImpresion = New Rectangle(X4, _Y, W3, TF)
            PDBody.Graphics.DrawString("SubT", Columna_F, Brushes.Black, AreaImpresion, eLeft)
            _Y += 10
            eLine()
            Total_ = 0
            For i As Integer = 0 To Lineas - 1
                _Y += TF
                AreaImpresion = New Rectangle(X1, _Y, W1, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Art, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eLeft)
                AreaImpresion = New Rectangle(X2, _Y, W2, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Cant, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eCenter)
                AreaImpresion = New Rectangle(X3, _Y, W3, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Precio, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eCenter)
                AreaImpresion = New Rectangle(X4, _Y, W4, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Sub, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eLeft)
                Total_ += _Tabla.Item(_Sub, i).Value
            Next
            _Total = Format((Total_ * 1), "##,##0.00")
            _Y += TF
            AreaImpresion = New Rectangle(X2 - 20, _Y, _AnchoHoja - (X2 - 20), Columna_F.GetHeight(PDBody.Graphics))
            PDBody.Graphics.DrawString("Total: $" + _Total, Columna_F, Brushes.Black, AreaImpresion, eLeft)

            _Y += Columna_F.GetHeight(PDBody.Graphics) + 10
            eLine()

        End If


    End Sub
    ''' <summary>
    ''' Agrega una linea al documento. Alineación: 0-> Izquierda; 1->Centro; 2-> Derecha
    ''' </summary>
    ''' <param name="Texto">Texto a imprimir</param>
    ''' <param name="Fuente_F">Titulo; Encabezado; Cuerpo; Columna</param>
    ''' <param name="Alineacion">Alineación: 0-> Izquierda; 1->Centro; 2-> Derecha</param>
    ''' <remarks></remarks>
    Private Sub PrintText(ByVal Texto As String, ByVal Fuente_F As Font, Optional ByVal Alineacion As Integer = 1)
        Dim TFuente As Decimal = 12
        If Not IsNothing(PDBody) Then
            Select Case Alineacion
                Case 0
                    _Aling = eLeft
                Case 1
                    _Aling = eCenter
                Case 2
                    _Aling = eRight
                Case Else
                    _Aling = eCenter
            End Select
            TFuente = PDBody.Graphics.MeasureString(Texto, Fuente_F).Height
            If PDBody.Graphics.MeasureString(Texto, Fuente_F).Width > _AnchoHoja Then
                If (PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja) Mod 1 Then
                    TFuente = TFuente * (Fix((PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja) + 1))
                Else
                    TFuente = TFuente * (PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja)
                End If
            Else
                Fuente_F.GetHeight(PDBody.Graphics)
            End If
            AreaImpresion = New Rectangle(_X, _Y, _AnchoHoja, TFuente)
            PDBody.Graphics.DrawString(Texto, Fuente_F, Brushes.Black, AreaImpresion, _Aling)
            _Y = _Y + TFuente
        Else
            MsgBox("¡No se ha indicado el inicio de documento al crear el Ticket!", vbOKOnly + vbExclamation, "Ticket")
        End If
    End Sub

    Private Sub PrintImage(ByVal Imagen As Image)
        Dim ImagenW As Integer
        Dim ImagenH As Integer
        Dim XPos As Integer

        ImagenH = Imagen.Height
        ImagenW = Imagen.Width
        If _AnchoHoja > ImagenW Then
            XPos = (_AnchoHoja - ImagenW) \ 2
        Else
            XPos = 0
        End If
        AreaImpresion = New Rectangle(XPos, _Y, ImagenW, ImagenH)
        PDBody.Graphics.DrawImage(Imagen, AreaImpresion)
        _Y = _Y + ImagenH

    End Sub
    ''' <summary>
    ''' Indica el termino de un impresión
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EndPrint() As PrintPageEventArgs
        PDBody.HasMorePages = False
        Return PDBody
    End Function
    Private Sub eSpace(Optional ByVal num As Integer = 1)
        Dim Fuente_F As New Font("Arial", 10, FontStyle.Regular)
        Dim TFuente As Decimal
        'Dim Texto As String = "- - - - - - - - - - - - - - - - - - - - - - - - - - - -"

        TFuente = Fuente_F.GetHeight(PDBody.Graphics)
        AreaImpresion = New Rectangle(_X, _Y, _AnchoHoja, _Espacio * num)
        PDBody.Graphics.DrawRectangle(Pens.White, AreaImpresion)
        _Y = _Y + (num * _Espacio)
    End Sub
    Private Sub eLine()
        Dim Fuente_F As New Font("Arial", 10, FontStyle.Regular)
        Dim TFuente As Decimal
        Dim Texto As String = "- - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        TFuente = Fuente_F.GetHeight(PDBody.Graphics)
        AreaImpresion = New Rectangle(_X, _Y, _AnchoHoja, TFuente)
        PDBody.Graphics.DrawString(Texto, Fuente_F, Brushes.Black, AreaImpresion, _Aling)
        _Y = _Y + TFuente
    End Sub
    ''' <summary>
    ''' Indica el inicio de la creación de un documento
    ''' </summary>
    ''' <param name="e">PrintPageEventArgs</param>
    ''' <remarks></remarks>
    Private Sub StartPrint(ByVal e As PrintPageEventArgs)
        'PDBody = Nothing
        PDBody = e
    End Sub
#End Region

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Imprimir()
    End Sub
End Class
