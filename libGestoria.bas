Attribute VB_Name = "libGestoria"
Option Explicit


Public mConfig As CFGControl
Public Conn As Connection


Public MiXL As Object  ' Variable que contiene la referencia
    ' de Microsoft Excel.
Public ExcelNoSeEjecutaba As Boolean   ' Indicador para liberación final .
Public ExcelSheet As Object
Public wrk As Excel.Workbook

Public BaseDatos As String
Public Const ValorNulo = "Null"
Public Const FormatoFecha = "yyyy-mm-dd"
Public Const FormatoHora = "hh:mm:ss"

Public TipoListado As Integer
' 1 = listado de comprobacion de venta fruta

Public EsImportaci As Byte
Public NombreHoja As String
Dim Rc As Byte


Public Usuario As Long
Public Fichero As String


Public Sub Main()
Dim I As Integer
'Vemos si ya se esta ejecutando
If App.PrevInstance Then
    MsgBox "Ya se está ejecutando el programa de traspaso a Excel (Tenga paciencia).", vbCritical
    Screen.MousePointer = vbDefault
    Exit Sub
End If


Set mConfig = New CFGControl
If mConfig.Leer = 1 Then
    MsgBox "No configurado"
    End
End If

'Si es importacion o creacion
NombreHoja = Command
'NombreHoja = "/I|arigasol|32000|C:\Users\Monica\Documents\documentacion Arigasol\TRASPASO DE REGAIXO 2015\151215.xlsx|"
'NombreHoja = "/I|arigasol|32000|C:\Users\Monica\Documents\documentacion Arigasol\TRASPASO DE REGAIXO 2015\12.xls|"

I = InStr(1, NombreHoja, "/")
If I = 0 Then
    MsgBox "Mal lanzado el programa", vbExclamation
    End
End If

NombreHoja = Mid(NombreHoja, I + 1)
Select Case Mid(NombreHoja, 1, 1)
    Case "I"
        EsImportaci = 1
End Select

'BaseDatos = Mid(NombreHoja, 3, Len(NombreHoja))
BaseDatos = RecuperaValor(NombreHoja, 2)
If BaseDatos = "" Then
    MsgBox "Falta la base de datos", vbCritical
    End
End If

Usuario = RecuperaValor(NombreHoja, 3)
Fichero = RecuperaValor(NombreHoja, 4)

NombreHoja = ""


'frmTrasArigasol.Text5 = Fichero
'frmTrasArigasol.Show

    NombreHoja = Fichero

    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        If EsImportaci = 1 Then
            If AbrirConexion(BaseDatos) Then
            
                
                'Vamos linea a linea, buscamos su trabajador
                RecorremosLineasFicheroTraspaso
                
            End If
        
            'Cerramos el excel
            CerrarExcel
        
        
        End If
        
        Dim NF As Integer
        NF = FreeFile
        Open App.Path & "\trasarigasol.z" For Output As #NF
        Print #NF, "0"
        Close #NF
    
        
        End
        
    End If


End Sub

Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, Cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(Cadena, J, I - J)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function

Public Function AbrirConexion(BaseDatos As String) As Boolean
Dim cad As String

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Trim(BaseDatos) & ";SERVER=" & mConfig.SERVER & ";"
    cad = cad & ";UID=" & mConfig.User
    cad = cad & ";PWD=" & mConfig.password
'++monica: tema del vista
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    If Err.Number <> 0 Then
        MsgBox "Error en la cadena de conexion" & vbCrLf & BaseDatos, vbCritical
        End
    Else
        AbrirConexion = True
    End If
End Function


Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Sub NombreSQL(ByRef Cadena As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub
'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(Cadena As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, Cadena, ".")
    If I > 0 Then Cadena = Mid(Cadena, 1, I - 1) & Mid(Cadena, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(Cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, Cadena, ",")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "." & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = Cadena
End Function

Public Function TransformaPuntosComas(Cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "," & Mid(Cadena, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = Cadena
End Function

Private Function RecorremosLineasFicheroTraspaso()
Dim FIN As Boolean
Dim I As Long
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
Dim vSql As String
Dim Cadena As String
Dim Posicion As Integer

Dim turno As Integer
Dim Albaran As Long
Dim factura As String
Dim Fecha As String
Dim cliente As String
Dim nomclien As String
Dim tarjeta As Double
Dim matricula As String
Dim km As Long
Dim producto As Integer
Dim nomprodu As String
Dim surtidor As Long
Dim manguera As Long
Dim NSuministro As Long
Dim precio As Currency
Dim descuento As Currency
Dim descuentoporc As Currency
Dim iva As Currency
Dim cantidad As Currency
Dim Importe As Currency
Dim idtipopago As Long
Dim desctipopago As String
Dim nif As String


    vSql = "delete from tmptraspaso where codusu = " & Usuario
    Conn.Execute vSql

    FIN = False
    I = 3
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            LineasEnBlanco = 0
            
            turno = ExcelSheet.Cells(I, 4).Value
            
            Albaran = 0
            If ExcelSheet.Cells(I, 5).Value <> "" Then
                Albaran = Mid(ExcelSheet.Cells(I, 5).Value, 5, Len(ExcelSheet.Cells(I, 5).Value) - 4) ' le quito el TIC1
            End If
            
            If Mid(ExcelSheet.Cells(I, 6).Value, 1, 1) = "'" Then
                factura = Mid(ExcelSheet.Cells(I, 6).Value, 2, Len(ExcelSheet.Cells(I, 6).Value) - 1)
            Else
                factura = ExcelSheet.Cells(I, 6).Value
            End If
            
            Fecha = ""
            If Mid(ExcelSheet.Cells(I, 8).Value, 1, 1) = "'" Then
                Fecha = Mid(ExcelSheet.Cells(I, 8).Value, 2, Len(ExcelSheet.Cells(I, 8).Value) - 1) ' quito la comilla
            Else
                Fecha = ExcelSheet.Cells(I, 8).Value
            End If
            
            '[Monica]04/01/2016: a veces viene el codigo de cliente con un apostrof delante, hay que quitarlo
            If Mid(ExcelSheet.Cells(I, 9).Value, 1, 1) = "'" Then
                cliente = Mid(ExcelSheet.Cells(I, 9).Value, 2, Len(ExcelSheet.Cells(I, 9).Value) - 1)
            Else
                cliente = ExcelSheet.Cells(I, 9).Value
            End If
            
            If Mid(ExcelSheet.Cells(I, 10).Value, 1, 1) = "'" Then
                nomclien = Mid(ExcelSheet.Cells(I, 10).Value, 2, Len(ExcelSheet.Cells(I, 10).Value) - 1)
            Else
                nomclien = ExcelSheet.Cells(I, 10).Value
            End If
            
            tarjeta = 0
            If Mid(ExcelSheet.Cells(I, 11).Value, 1, 1) = "'" Then
                If ExcelSheet.Cells(I, 11).Value <> "'" Then tarjeta = Mid(ExcelSheet.Cells(I, 11).Value, 2, Len(ExcelSheet.Cells(I, 11).Value) - 1)
            Else
                tarjeta = ExcelSheet.Cells(I, 11).Value
            End If
            
            If Mid(ExcelSheet.Cells(I, 13).Value, 1, 1) = "'" Then
                matricula = Mid(ExcelSheet.Cells(I, 13).Value, 2, Len(ExcelSheet.Cells(I, 13).Value) - 1)
            Else
                matricula = ExcelSheet.Cells(I, 13).Value
            End If
            
            km = 0
            If Mid(ExcelSheet.Cells(I, 14).Value, 1, 1) = "'" Then
                If ExcelSheet.Cells(I, 14).Value <> "'" Then km = Mid(ExcelSheet.Cells(I, 14).Value, 2, Len(ExcelSheet.Cells(I, 14).Value) - 1)
            Else
                km = ExcelSheet.Cells(I, 14).Value
            End If
            
            producto = 0
            If Mid(ExcelSheet.Cells(I, 15).Value, 1, 1) = "'" Then
                If ExcelSheet.Cells(I, 15).Value <> "'" Then producto = Mid(ExcelSheet.Cells(I, 15).Value, 2, Len(ExcelSheet.Cells(I, 15).Value) - 1)
            Else
                producto = ExcelSheet.Cells(I, 15).Value
            End If
            
            If Mid(ExcelSheet.Cells(I, 16).Value, 1, 1) = "'" Then
                nomprodu = Mid(ExcelSheet.Cells(I, 16).Value, 2, Len(ExcelSheet.Cells(I, 16).Value) - 1)
            Else
                nomprodu = ExcelSheet.Cells(I, 16).Value
            End If
            
            
            
            surtidor = 0
            If Mid(ExcelSheet.Cells(I, 17).Value, 1, 1) = "'" Then
                If ExcelSheet.Cells(I, 17).Value <> "'" Then surtidor = Mid(ExcelSheet.Cells(I, 17).Value, 2, Len(ExcelSheet.Cells(I, 17).Value) - 1)
            Else
                surtidor = ExcelSheet.Cells(I, 17).Value
            End If
            
            
            
            manguera = 0
            If Mid(ExcelSheet.Cells(I, 18).Value, 1, 1) = "'" Then
                If ExcelSheet.Cells(I, 18).Value <> "'" Then manguera = Mid(ExcelSheet.Cells(I, 18).Value, 2, Len(ExcelSheet.Cells(I, 18).Value) - 1)
            Else
                manguera = ExcelSheet.Cells(I, 18).Value
            End If
            
            
            
            NSuministro = 0
            If ExcelSheet.Cells(I, 19).Value <> "" Then
                NSuministro = Mid(ExcelSheet.Cells(I, 19).Value, 2, Len(ExcelSheet.Cells(I, 19).Value) - 1)
            End If
            precio = ExcelSheet.Cells(I, 20).Value
            descuento = ExcelSheet.Cells(I, 21).Value
            descuentoporc = ExcelSheet.Cells(I, 22).Value
            iva = ExcelSheet.Cells(I, 23).Value
            cantidad = ExcelSheet.Cells(I, 24).Value
            Importe = ExcelSheet.Cells(I, 25).Value
            idtipopago = Mid(ExcelSheet.Cells(I, 26).Value, 2, Len(ExcelSheet.Cells(I, 26).Value) - 1)
            desctipopago = ExcelSheet.Cells(I, 27).Value
            nif = ExcelSheet.Cells(I, 28).Value
            
            vSql = "insert into tmptraspaso (codusu,turno,albaran,factura,fecha,cliente,nomclien,tarjeta,matricula,km,producto,nomprodu, "
            vSql = vSql & "surtidor,manguera,nsuministro,precio,descuento,descuentoporc,iva,cantidad,importe,idtipopago,desctipopago,nif) values ("
            vSql = vSql & Usuario & "," & DBSet(turno, "N") & "," & DBSet(Albaran, "N") & "," & DBSet(factura, "T") & "," & DBSet(Fecha, "T") & ","
            vSql = vSql & DBSet(cliente, "T") & "," & DBSet(nomclien, "T") & "," & DBSet(tarjeta, "N", "S") & "," & DBSet(matricula, "T") & ","
            vSql = vSql & DBSet(km, "N") & "," & DBSet(producto, "N") & "," & DBSet(nomprodu, "T") & "," & DBSet(surtidor, "N") & ","
            vSql = vSql & DBSet(manguera, "N") & "," & DBSet(NSuministro, "T") & "," & DBSet(precio, "N") & "," & DBSet(descuento, "N") & ","
            vSql = vSql & DBSet(descuentoporc, "N") & "," & DBSet(iva, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ","
            vSql = vSql & DBSet(idtipopago, "N") & "," & DBSet(desctipopago, "T") & "," & DBSet(nif, "T") & ")"
            
            Conn.Execute vSql
            
        Else
            
            FIN = True

        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function

