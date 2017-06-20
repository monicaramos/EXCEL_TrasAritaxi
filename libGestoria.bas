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

' esto es una prueba

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
'NombreHoja = "/I|aritaxi2|32000|C:\Users\Monica\Documents\documentacion Aritaxi\RadioTaxi\Servicios.xlsx|"

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


    frmTrasAritaxi.Text5 = Fichero
    frmTrasAritaxi.Show vbModal
    

'    NombreHoja = Fichero
'
'    Rc = AbrirEXCEL
'
'    If Rc = 0 Then
'
'        If EsImportaci = 1 Then
'            If AbrirConexion(BaseDatos) Then
'
'
'                'Vamos linea a linea, buscamos su trabajador
'                RecorremosLineasFicheroTraspaso
'
'            End If
'
'            'Cerramos el excel
'            CerrarExcel
'
'
'        End If
'
'        Dim NF As Integer
'        NF = FreeFile
'        Open App.Path & "\trasaritaxi.z" For Output As #NF
'        Print #NF, "0"
'        Close #NF
'
'
'        End
'
'    End If
'

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

