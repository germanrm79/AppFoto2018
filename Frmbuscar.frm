VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmbuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "Frmbuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4830
      Top             =   4170
   End
   Begin VB.OptionButton OptNombrE 
      Caption         =   "Nombre Completo"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   210
      TabIndex        =   12
      Top             =   780
      Width           =   5325
      Begin VB.TextBox TxtNombreComP 
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Top             =   330
         Width           =   4995
      End
      Begin VB.TextBox TxtNombreS 
         Height          =   285
         Left            =   3450
         TabIndex        =   3
         Top             =   330
         Width           =   1725
      End
      Begin VB.TextBox TxtMaternO 
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox TxtPaternO 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2445
         Left            =   150
         TabIndex        =   13
         Top             =   690
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   4313
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   2295
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Estatus"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Numtra"
            Object.Width           =   776
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   270
      Width           =   1155
   End
   Begin VB.CommandButton btncancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2430
      TabIndex        =   5
      Top             =   4140
      Width           =   1335
   End
   Begin VB.CommandButton btnsel 
      Caption         =   "Seleccionar"
      Height          =   315
      Left            =   930
      TabIndex        =   4
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   3255
      Left            =   180
      TabIndex        =   8
      Top             =   690
      Width           =   5295
      Begin VB.TextBox Txtbusca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   480
         Width           =   4995
      End
      Begin VB.ListBox Lstelementos 
         Enabled         =   0   'False
         Height          =   2205
         Left            =   150
         TabIndex        =   9
         Top             =   870
         Width           =   4965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numtra:"
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   60
      Width           =   555
   End
End
Attribute VB_Name = "Frmbuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim busca As Boolean, BanBusquedA As Byte
Dim PatErnO As String

Private Sub btncancel_Click()
Me.Hide
End Sub

Private Sub btnsel_Click()
Dim strTemp As String
Dim con As String

'aqui es despues de haber presionado el click en el boton de seleccion!!!
'verificamos que este algo seleccionado
 If ListView1.SelectedItem.Index <> -1 Then
     ' ponemos el relojito
       Screen.MousePointer = 11
     
      ' en al forma de consulta ponemos el nombre correspodiente
      'Frmconsulta.txtnombre.Text = Trim(Lstelementos.List(Lstelementos.ListIndex))
      'Frmconsulta.txtnumtra.Text = Lstelementos.ItemData(Lstelementos.ListIndex)
      'Frmconsulta.txtnumtra_KeyPress (13)
      frmDataEnv.txtNumtra.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
      'Debug.Print "Paso Numero de catalogo:" & Time
      'TxtNombreComP.Text = Empty
      'TxtPaternO.Text = Empty
      'TxtMaternO.Text = Empty
      'TxtNombreS.Text = Empty
      'Debug.Print "Se presiona enter:" & Time
      frmDataEnv.txtnumtra_KeyPress (13)
      'Debug.Print "salio del enter:" & Time
      Screen.MousePointer = 0
  Me.Hide
 End If

End Sub

Private Sub Command1_Click()
   ListView1.ListItems(ListView1.ListItems.Count).Ghosted = True
End Sub

Private Sub Form_Load()
   CARGA
   BanBusquedA = 3
   TxtNombreComP.Visible = True
   'ListView1.ColumnHeaders(1).Width = 0
   'ListView1.ColumnHeaders(2).Width = 0
   'ListView1.ColumnHeaders(3).Width = 0
   ListView1.ColumnHeaders(1).Width = 4300
   ListView1.ColumnHeaders(2).Width = 0
   'ListView1.ColumnHeaders(6).Width = 400
   TxtPaternO.Enabled = False
   TxtMaternO.Enabled = False
   TxtNombreS.Enabled = False
   TxtNombreComP.Enabled = True
   
End Sub

Public Sub CARGA()
Dim Item As ListItem
'se pone el boton de seleccion disabled
 btnsel.Enabled = False
   'cuandos se carga la forma cargamos todos los nombres
   ' de los trabajadores en orden alfabetico
   'Lstelementos.Clear
   ListView1.ListItems.Clear
   With Dte
     .personal
      Do While Not .rsPersonal.EOF
        If Trim(.rsPersonal!NOMBRE) <> Empty Then
           Set Item = ListView1.ListItems.Add(, ("#" & .rsPersonal!numtra), Trim(.rsPersonal!NOMBRE))
            'Item.SubItems(1) = Trim(.rsPersonal!APELL_MAT)
           'Item.SubItems(2) = Trim(.rsPersonal!NOMBRE_1)
           Item.SubItems(1) = Trim(.rsPersonal!numtra)
           Item.SubItems(2) = Trim(.rsPersonal!NOMBRE)
           'Item.SubItems(3) = Trim(.rsPersonal!numtra)
       End If
         .rsPersonal.MoveNext
       Loop
       
     If .rsPersonal.State <> 0 Then .rsPersonal.Close
'''     .Connections(1).CommandTimeout = 0
'''      If .Connections(1).State <> 0 Then .Connections(1).Close
     'If .Connections(2).State <> 0 Then .Connections(2).Close
     
  End With
    
 Screen.MousePointer = 0
End Sub



Private Sub ListView1_Click()
    'TxtPaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index)
    'TxtMaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
    'TxtNombreS.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
    Text1.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView1.SortKey = ColumnHeader.Index - 1
  ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
'Debug.Print "inicio" & Time
  btnsel.Value = True
'Debug.Print "Fin " & Time
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'  Case vbKeyDown, vbKeyUp
'      TxtPaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index)
'      TxtMaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
'      TxtNombreS.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
'End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyDown, vbKeyUp
      'TxtPaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index)
      'TxtMaternO.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
      'TxtNombreS.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
      Text1.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
  Case vbKeyReturn
       btnsel.Value = True
  Case vbKeySpace
      TxtPaternO.SetFocus
End Select

End Sub

Private Sub Lstelementos_Click()
If Lstelementos.ListIndex <> -1 Then
 Text1.Text = Lstelementos.ItemData(Lstelementos.ListIndex)
End If
End Sub

Private Sub Lstelementos_DblClick()
  
  If Lstelementos.ListIndex <> -1 Then
    Txtbusca.Text = Lstelementos.List(Lstelementos.ListIndex)
    btnsel.Value = True
  End If
  
End Sub

Private Sub Lstelementos_KeyPress(KeyAscii As Integer)
 busca = True
 ' aqui es para cuando presionemos enter ya se seleccione automaticamente
 If KeyAscii = 13 Then
  If Lstelementos.ListIndex <> -1 Then
    Txtbusca.Text = Lstelementos.List(Lstelementos.ListIndex)
    btnsel.Value = True
  End If
 End If
End Sub

Private Sub OptNombrE_Click(Index As Integer)

TxtNombreComP.Text = ""
TxtPaternO.Text = ""
TxtMaternO.Text = ""
TxtNombreS.Text = ""

ListView1.ColumnHeaders(5).Width = 0

Select Case Index
Case 0
   TxtNombreComP.Visible = False
   BanBusquedA = 0
   ListView1.ColumnHeaders(1).Width = 1300
   ListView1.ColumnHeaders(2).Width = 1300
   ListView1.ColumnHeaders(3).Width = 1600
   ListView1.ColumnHeaders(4).Width = 0
   ListView1.ColumnHeaders(5).Width = 0
   ListView1.ColumnHeaders(6).Width = 400
   TxtPaternO.Enabled = True
   TxtMaternO.Enabled = True
   TxtNombreS.Enabled = True
   TxtNombreComP.Enabled = False
   
Case 1
   BanBusquedA = 1
   TxtNombreComP.Visible = False
   ListView1.ColumnHeaders(1).Width = 1300
   ListView1.ColumnHeaders(2).Width = 1300
   ListView1.ColumnHeaders(3).Width = 1600
   ListView1.ColumnHeaders(4).Width = 0
   ListView1.ColumnHeaders(5).Width = 0
   ListView1.ColumnHeaders(6).Width = 400
   
   TxtPaternO.Enabled = True
   TxtMaternO.Enabled = True
   TxtNombreS.Enabled = True
   TxtNombreComP.Enabled = False
   
Case 2
   BanBusquedA = 3
   TxtNombreComP.Visible = True
   ListView1.ColumnHeaders(1).Width = 0
   ListView1.ColumnHeaders(2).Width = 0
   ListView1.ColumnHeaders(3).Width = 0
   ListView1.ColumnHeaders(4).Width = 4300
   ListView1.ColumnHeaders(5).Width = 0
   ListView1.ColumnHeaders(6).Width = 400
   TxtPaternO.Enabled = False
   TxtMaternO.Enabled = False
   TxtNombreS.Enabled = False
   TxtNombreComP.Enabled = True
   
End Select

    ListView1.SortKey = BanBusquedA
    ListView1.Sorted = True
End Sub



Private Sub TxtApellPaT_KeyDown(KeyCode As Integer, Shift As Integer)
 busca = True
End Sub

Private Sub Timer1_Timer()
'  If TxtPaternO.Text = PatErnO Then
'   ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
'   'Timer1.Enabled = False
'  End If
End Sub

Private Sub Txtbusca_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim aun As Integer
  busca = True
  
 'estos procesos son para validar si se presiono la felcha arriba o abajo
 'para poder movernos entre la lista desde el textbox
  If KeyCode = 40 Then
    aun = Lstelementos.ListIndex + 1
    
      If aun > Lstelementos.ListCount - 1 Then
        aun = Lstelementos.ListCount - 1
      End If
        Lstelementos.ListIndex = aun
        Txtbusca.Text = Lstelementos.List(Lstelementos.ListIndex)
        busca = False
  End If
  
  If KeyCode = 38 Then
    aun = Lstelementos.ListIndex - 1
    
      If aun < 0 Then
        aun = 0
      End If
        Lstelementos.ListIndex = aun
        Txtbusca.Text = Lstelementos.List(Lstelementos.ListIndex)
        busca = False
  End If
  
  If KeyCode = 13 Then
    btnsel.Value = True
  End If
  
  
  
End Sub

Private Sub Txtbusca_KeyUp(KeyCode As Integer, Shift As Integer)
    
  If Txtbusca.Text = "" Then
   btnsel.Enabled = False
  End If
  If busca = True Then
    Lstelementos.ListIndex = BuscaNombrELstBoX(Lstelementos, Trim(Txtbusca.Text), 1)
  End If
  
End Sub


Function BuscaNombrE(Lstbox As ListView, PatErnO As String, Materno As String, Nombres As String, NombreCompletO As String, Tipobusqueda As Byte) As Long
 Dim i As Long
   BuscaNombrE = ListView1.SelectedItem.Index
   
'   If StrBusca = "" Then
'      Exit Function
'   End If
   
   For i = 1 To Lstbox.ListItems.Count
     'DoEvents
      Select Case Tipobusqueda
      
        Case 0, 1, 2
           If UCase(Trim(Lstbox.ListItems(i))) Like UCase(Trim(PatErnO)) & "*" And UCase(Trim(Lstbox.ListItems.Item(i).SubItems(1))) Like UCase(Trim(Materno)) & "*" _
           And UCase(Trim(Lstbox.ListItems.Item(i).SubItems(2))) Like UCase(Trim(Nombres)) & "*" Then
             BuscaNombrE = i
              Lstbox.ListItems(i).Selected = True
              Text1.Text = Lstbox.ListItems(i).SubItems(4)
              Lstbox.ListItems(i).EnsureVisible
              Exit Function
           End If
           
      Case 3
       
         If UCase(Trim(Lstbox.ListItems.Item(i).SubItems(2))) Like UCase(Trim(NombreCompletO)) & "*" Then
              BuscaNombrE = i
              Lstbox.ListItems(i).Selected = True
              Text1.Text = Lstbox.ListItems(i).SubItems(1)
              Lstbox.ListItems(i).EnsureVisible
              Exit Function
           End If
           
      End Select

   Next i
   
End Function

Private Sub TxtMaternO_Change()
  If TxtPaternO.Text = "" And TxtMaternO.Text = "" And TxtNombreS.Text = "" Then
     btnsel.Enabled = False
  Else
    btnsel.Enabled = True
  End If
  
'  If busca = True Then
'    ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
'  End If
End Sub

Private Sub TxtMaternO_KeyDown(KeyCode As Integer, Shift As Integer)
 busca = True
 Select Case KeyCode
    Case vbKeySpace
      TxtNombreS.SetFocus
    Case vbKeyReturn
       btnsel.Value = True
    Case vbKeyDown
      ListView1.SetFocus
   End Select
End Sub

Private Sub TxtNombreComP_Change()
If TxtNombreComP.Text = "" Then
   btnsel.Enabled = False
  End If
  If busca = True Then
    ListView1.ListItems.Item(BuscaNombrE(ListView1, "", "", "", Trim(TxtNombreComP.Text), 3)).Selected = True
  End If
End Sub

Private Sub TxtNombreComP_KeyDown(KeyCode As Integer, Shift As Integer)
 busca = True
 
  Select Case KeyCode
    Case vbKeyReturn
      If TxtNombreComP.Text <> "" Then
       btnsel.Value = True
      End If
    Case vbKeyEscape
       Me.Hide
  End Select
  
End Sub

Private Sub TxtNombreS_Change()

  If TxtPaternO.Text = "" And TxtMaternO.Text = "" And TxtNombreS.Text = "" Then
     btnsel.Enabled = False
  Else
    btnsel.Enabled = True
  End If
  
'  If busca = True Then
'    ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
'  End If
  
End Sub


Function BuscaNombrELstBoX(Lstbox As ListBox, StrBusca As String, Tipobusqueda As Byte) As Long
 Dim i As Long
 
   BuscaNombrELstBoX = 1
   
   If StrBusca = "" Then
      Exit Function
   End If
   
   For i = 1 To Lstbox.ListCount
          
          Select Case Tipobusqueda
          Case 0
            If UCase(Trim(Lstbox.List(i))) Like UCase(StrBusca) & "*" Then
               BuscaNombrELstBoX = i
               Exit Function
            End If
           Case 1
           
              If UCase(Trim(Lstbox.List(i))) Like "*_" & UCase(StrBusca) & "*" Then
                 BuscaNombrELstBoX = i
                 Exit Function
              End If
           End Select
           
   Next i
   
   
End Function

Private Sub TxtNombreS_KeyDown(KeyCode As Integer, Shift As Integer)
 busca = True
 
  Select Case KeyCode
    Case vbKeyReturn
       btnsel.Value = True
    Case vbKeyDown
       ListView1.SetFocus
  End Select
  
End Sub

Private Sub TxtPaternO_Change()

  If TxtPaternO.Text = "" And TxtMaternO.Text = "" And TxtNombreS.Text = "" Then
     btnsel.Enabled = False
  Else
    btnsel.Enabled = True
  End If
'  If busca = True Then
'    'ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
'  End If
'  PatErnO = TxtPaternO.Text
  
End Sub



Private Sub TxtPaternO_KeyUp(KeyCode As Integer, Shift As Integer)

  busca = True
  Timer1.Enabled = True
  Select Case KeyCode
    Case vbKeySpace
         TxtMaternO.SetFocus
         DoEvents
         ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
    Case vbKeyReturn
         If TxtPaternO.Text <> "" Then
          btnsel.Value = True
         End If
    Case vbKeyDown
         ListView1.SetFocus
    Case vbKeyEscape
         Me.Hide
    Case Else
     ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
  End Select
  
End Sub

Private Sub TxtPaternO_Validate(Cancel As Boolean)
 If TxtPaternO.Text <> "" Then
  ListView1.ListItems.Item(BuscaNombrE(ListView1, Trim(TxtPaternO.Text), Trim(TxtMaternO.Text), Trim(TxtNombreS.Text), "", 0)).Selected = True
 End If
End Sub
