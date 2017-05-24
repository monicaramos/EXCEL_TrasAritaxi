VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasAritaxi 
   Caption         =   "Traspaso de Llamadas "
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmTrasAritaxi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7260
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameImportar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   4500
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   690
         Width           =   6735
      End
      Begin VB.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1230
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   450
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   900
         Picture         =   "frmTrasAritaxi.frx":1782
         Top             =   420
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTrasAritaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim SQL As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FecAlbaran As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Cajones As String
Dim Calidad(20) As String
Dim Contador As Long
Dim Values As String

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1




Private Sub Command2_Click()
Dim Rc As Byte
Dim I As Integer
Dim Rs1 As ADODB.Recordset
Dim KilosI As Long
Dim b As Boolean
Dim Notas As String

    'IMPORTAR
    If Text5.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
    If Dir(Text5.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    
    NombreHoja = Text5.Text
    'Abrimos excel
    
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
        
'        MsgBox "FIN", vbInformation
        
        Dim NF As Integer
        NF = FreeFile
        Open App.Path & "\trasaritaxi.z" For Output As #NF
        Print #NF, "0"
        Close #NF
        
        
    End If
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    
'    Combo1.ListIndex = Month(Now) - 1
'    Text3.Text = Year(Now)
    FrameImportar.visible = False
    
    Limpiar
    Select Case EsImportaci
    Case 1
        Caption = "Cargar Traspaso de Poste desde fichero excel"
        FrameImportar.visible = True

    End Select
    
    
 
End Sub

Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub

Private Function TransformaComasPuntos(Cadena) As String
Dim cad As String
Dim J As Integer
    
    J = InStr(1, Cadena, ",")
    If J > 0 Then
        cad = Mid(Cadena, 1, J - 1) & "." & Mid(Cadena, J + 1)
    Else
        cad = Cadena
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
'    Text4.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click()
    AbrirDialogo 1
End Sub


Private Sub AbrirDialogo(Opcion As Integer)

    On Error GoTo EA
    
    With Me.CommonDialog1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        .Filter = "EXCEL (*.xls)|*.xls"
        .CancelError = True
        If Opcion <> 1 Then
            .ShowOpen
            If Opcion = 0 Then
            Else
                Text5.Text = .FileName
            End If
        Else
            .ShowSave
        End If
        
        
        
    End With
EA:
End Sub

Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub Image3_Click()
 AbrirDialogo 2
End Sub





Private Function RecorremosLineasFicheroTraspaso()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
Dim vSql As String
Dim Cadena As String
Dim Posicion As Integer

Dim Linea As String
Dim s As String
Dim menErrProceso As String
Dim CadenaAux As String
Dim LlevoFichero As Currency


    On Error GoTo eRecorremosLineasFicheroTraspaso

    RecorremosLineasFicheroTraspaso = False

    vSql = "delete from tmptaxi"
    Conn.Execute vSql


    Linea = "(id,telefono,codclien,codautor,codusuar,nomclien,tipservi,observa1,numeruve,licencia,matricul,"
    Linea = Linea & "dirllama,ciudadre,numllama,puerllama,fecha,hora,idservic,opereser,opedespa,estado,"
    Linea = Linea & "observa2,fecreser,horreser,fecaviso,horaviso,fecllega,horllega,fecocupa,horocupa,fecfinal,"
    Linea = Linea & "horfinal,importtx,impcompr,extcompr,impventa,extventa,distanci,suplemen,imppeaje,imppropi,facturad,"
    '[Monica]03/10/2014: añadimos el taxi del destino
    Linea = Linea & "abonados,validado,destino,error1,error)"



    FIN = False
    I = 2
    LineasEnBlanco = 0
    
    Values = ""
    
    Me.Label1(0).visible = True
    
    
    While Not FIN
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            
            s = ExcelSheet.Cells(I, 1).Value
            If IsNumeric(s) Then
                ArmarCadena menErrProceso, I
                menErrProceso = CadenaAux
            Else
                menErrProceso = menErrProceso & " " & CadenaAux
            End If
            
            
            LlevoFichero = LlevoFichero + Len(menErrProceso)
            Me.Label1(0).Caption = "Linea " & I
            Me.Label1(0).Refresh
            DoEvents
            
            If Len(Values) > 100000 Then
                'quitamos la ultima coma
                Values = Mid(Values, 1, Len(Values) - 1)
                SQL = "INSERT INTO tmptaxi " & Linea & " VALUES " & Values
                Conn.Execute SQL
                Values = ""
            End If
        Else
            FIN = True
        End If
        
        I = I + 1
    Wend
    RecorremosLineasFicheroTraspaso = True
    Exit Function
            
eRecorremosLineasFicheroTraspaso:
    MsgBox "Error en Recorremos Lineas Fichero Traspaso " & Err.Description
End Function


'*******************************************
' RADIOTAXI
'*******************************************

Private Sub ArmarCadena(Cadena As String, ByRef I As Long)
Dim Telefono As String
Dim Values1 As String
Dim Error As String
Dim Error1 As Byte
Dim FechaHora As String

Dim Valor As Double
Dim Fecha As String
Dim Hora As String
Dim Vehiculo As String



    
    
    FechaHora = ExcelSheet.Cells(I, 94).Value
    
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)
    Vehiculo = ExcelSheet.Cells(I, 19).Value 'Mid(menErrProceso, 293, 4)


    Error1 = 0
    Error = ""
    'armamos los registros segun la cadena
    Telefono = ExcelSheet.Cells(I, 93).Value
    'telefono
    
    Values1 = I
    
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
                    


    Telefono = ExcelSheet.Cells(I, 9).Value
    'codclien
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
            Values1 = Values1 & ",NULL"
            Error1 = 1
            Error = "codclien con formato incorrecto"
    Else
            Values1 = Values1 & "," & CInt(Telefono)
    End If
    
    Telefono = ExcelSheet.Cells(I, 13).Value
    'codautor"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    Telefono = ExcelSheet.Cells(I, 11).Value
    'codusuar"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    
    Telefono = ExcelSheet.Cells(I, 10).Value
    'nomclien"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    Telefono = ExcelSheet.Cells(I, 36).Value
    'tipservi"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        If Telefono = "0" Then
            Values1 = Values1 & ",0"
        ElseIf Telefono = "1" Then
                Values1 = Values1 & ",1"
        Else
            Values1 = Values1 & ",NULL"
            Error1 = "1"
            Error = "tipservi con formato incorrecto"
        End If
    End If
    
    Telefono = ExcelSheet.Cells(I, 34).Value
    'observa1"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    'numeruve"
    If Not IsNumeric(Vehiculo) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "Vehiculo con formato incorrecto"
    Else
        Values1 = Values1 & "," & CInt(Vehiculo) + 10000
    End If

    Telefono = ExcelSheet.Cells(I, 20).Value
    'licencia"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    Telefono = ExcelSheet.Cells(I, 21).Value
    'matricul"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    Telefono = ExcelSheet.Cells(I, 28).Value
    'dirllama"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    
    Telefono = ExcelSheet.Cells(I, 29).Value
    'ciudadre"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    
    Telefono = ""
    'numllama"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    
    Telefono = ""
    'puerllama"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    'fecha"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "Falta fecha"
    ElseIf Not IsDate(Fecha) Then
            Values1 = Values1 & ",NULL"
            Error1 = 1
            Error = "Fecha con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If
    
    'hora"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "Falta hora"
    ElseIf Not IsDate(Hora) Then
            Values1 = Values1 & ",NULL"
            Error1 = 1
            Error = "Hora con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If

    Telefono = ExcelSheet.Cells(I, 1).Value
    'idservic"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If


    Telefono = ExcelSheet.Cells(I, 58).Value
    'opereser"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    
    Telefono = ExcelSheet.Cells(I, 59).Value
    'opedespa"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If


    '****** NO HE ENCONTRADO EL ESTADO
    '******
    Telefono = "" 'Trim(Mid(Cadena, 481, 4))
    'estado"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    '***** HASTA AQUI

    Telefono = ExcelSheet.Cells(I, 35).Value
    'observa2"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If

    
    FechaHora = ExcelSheet.Cells(I, 94).Value
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)
    
    'fecreser"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Fecha) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "fecha reserva con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If

    'horreser"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Hora) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "hora reserva con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If

'???????????
    
    
    FechaHora = ExcelSheet.Cells(I, 26).Value
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)
    'fecaviso"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Fecha) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "fecha aviso con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If

    'horaviso"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Hora) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "hora aviso con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If


    FechaHora = ExcelSheet.Cells(I, 24).Value
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)

    'fecllega"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Fecha) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "fecha llegada con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If
    
    'horllega"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Hora) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "hora llegada con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If


    FechaHora = ExcelSheet.Cells(I, 25).Value
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)

    'fecocupa"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Fecha) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "fecha ocupa con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If
    
    'horocupa"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Hora) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "hora ocupa con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If


    FechaHora = ExcelSheet.Cells(I, 27).Value
    Fecha = Mid(FechaHora, 1, 10)
    Hora = Mid(FechaHora, 12, 8)

    'fecfinal"
    If Fecha = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Fecha) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "fecha final con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Fecha), FormatoFecha), "T")
    End If
    
    'horfinal"
    If Hora = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsDate(Hora) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "hora final con formato incorrecto"
    Else
        Values1 = Values1 & "," & DBSet(Format(CDate(Hora), FormatoHora), "T")
    End If


    Telefono = ExcelSheet.Cells(I, 47).Value
    'importtx"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "importe tx con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
    
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If


'????????????????????????????


    Telefono = ExcelSheet.Cells(I, 41).Value
    'impcompr"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "importe compra con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
    
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If



    Telefono = ExcelSheet.Cells(I, 42).Value
    'extcompr"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "extcompr con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
        
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ExcelSheet.Cells(I, 37).Value
    'impventa"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "importe venta con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
    
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ExcelSheet.Cells(I, 38).Value
    'extventa"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "extventa con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
        
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ExcelSheet.Cells(I, 48).Value
    'distanci"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "distancia con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ExcelSheet.Cells(I, 95).Value
    'suplemen"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "suplemento con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
        
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ""
    'imppeaje"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "importe peaje con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
    
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If
    
    Telefono = ""
    'imppropi"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "importe propina con formato incorrecto"
    Else
        If InStr(1, Telefono, ",") > 0 Then
            Valor = ImporteFormateado(Telefono)
        Else
            Valor = CDbl(TransformaPuntosComas(Telefono))
        End If
    
        Values1 = Values1 & "," & DBSet(Valor, "N")
    End If

    Telefono = ExcelSheet.Cells(I, 39).Value
    'facturad"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "facturado con formato incorrecto"
    Else
        Values1 = Values1 & "," & CInt(Telefono)
    End If

    Telefono = ExcelSheet.Cells(I, 43).Value
    'abonados"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "abonado con formato incorrecto"
    Else
        Values1 = Values1 & "," & CInt(Telefono)
    End If

    Telefono = ExcelSheet.Cells(I, 46).Value
    'validado"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    ElseIf Not IsNumeric(Telefono) Then
        Values1 = Values1 & ",NULL"
        Error1 = 1
        Error = "validado con formato incorrecto"
    Else
        Values1 = Values1 & "," & CInt(Telefono)
    End If

    '[Monica]03/10/2014: añadimos el destino del servicio
    Telefono = Trim(ExcelSheet.Cells(I, 32).Value & " " & ExcelSheet.Cells(I, 33).Value)
    'destino"
    If Telefono = "" Then
        Values1 = Values1 & ",NULL"
    Else
        Values1 = Values1 & "," & DBSet(Telefono, "T")
    End If
    '++
    
    'error1,error
    Values1 = "(" & Values1 & "," & Error1 & "," & DBSet(Error, "T") & "),"
    
    Values = Values & Values1

EInsert:
    If Err.Number <> 0 Then
        MsgBox "Error al insertar en error1. " & Err.Description
    End If
End Sub






