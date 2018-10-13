VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDataEnv 
   Caption         =   "Modulo Toma de Fotografia"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdFoto 
         Height          =   735
         Left            =   120
         Picture         =   "frmDataEnv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdFirma 
         Height          =   735
         Left            =   1800
         Picture         =   "frmDataEnv.frx":0AF6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtnumtra 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   3
         ToolTipText     =   "Teclee el numero de trabajador y presione [enter]"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnbuscatrab 
         Caption         =   "-->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4560
         Picture         =   "frmDataEnv.frx":10DB
         TabIndex        =   2
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox Txtnombre 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   525
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Nombre"
         Top             =   1440
         Width           =   5895
      End
      Begin MSComDlg.CommonDialog Dialog2 
         Left            =   2640
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog Dialog1 
         Left            =   3360
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "jpg"
         DialogTitle     =   "Selecciona nueva fotografia.."
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2745
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Numtra."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   2880
         TabIndex        =   4
         Top             =   1080
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rutaFotos As String
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Function FileExists(Strruta As String) As Boolean
On Error GoTo error
  Dim fso, mensaje
  Set fso = CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(Strruta)) Then
    FileExists = True
  Else
    FileExists = False
  End If
  Exit Function
error:   MsgBox "Falla:" & Err.Number & " " & Err.Description & " " & Err.Source
  FileExists = False
  
End Function
Private Function ResizeImage(ByVal Original As WIA.ImageFile, ByVal WidthPixels As Long, ByVal HeightPixels As Long) As WIA.ImageFile
    'Scale the photo to fit supplied dimensions w/o distortion.
    With New WIA.ImageProcess
        .Filters.Add .FilterInfos!Crop.FilterID
        .Filters.Add .FilterInfos!Scale.FilterID
        .Filters.Add .FilterInfos!Convert.FilterID
        With .Filters(1).Properties
            '!PreserveAspectRatio = True by default, so just:
            '!MaximumWidth = WidthPixels
            '!MaximumHeight = HeightPixels
            '!Left = 750
            '!Top = 150
            '!Right = 750
            '!Bottom = 150
            !Left = (Original.Width - Original.Height) / 2
            '!Top = 150
            !Right = (Original.Width - Original.Height) / 2
            
            '!Bottom = 150
            
        End With
        With .Filters(2).Properties
            !PreserveAspectRatio = True
            !MaximumWidth = 250
            !MaximumHeight = 250
        End With
        
        With .Filters(3).Properties
            !FormatID = wiaFormatJPEG
            !Quality = 100
        End With
        Set ResizeImage = .Apply(Original)
    End With
End Function



Private Sub btnbuscatrab_Click()
Frmbuscar.Show 1
End Sub

Private Sub cmdFirma_Click()
If txtNumtra <> "" Then
    frmFirma.Show 1
Else
        MsgBox "¡Favor de consultar un numero de trabajador antes de actualizar!", vbExclamation
End If
End Sub

Private Sub CmdFoto_Click()
        Dim imgPhoto As WIA.ImageFile
        Dim FileName  As String
        Dim Confirmacion, A As Integer
        If txtNumtra <> "" Then
            Dialog1.DialogTitle = "Seleccionar fotografia"
            Dialog1.Filter = "Import Files|*.jpg"
            'Dialog1.InitDir = "D:\prograsRechum\Sicoda\fotos" '//Needs to be "My Computer"
            Dialog1.InitDir = rutaFotos & "fotos" '//Needs to be "My Computer"
            
            'Dialog1.CancelError = False
            Dialog1.ShowOpen
            
            
            If Err.Number = &H7FF3 Then
                MsgBox "CANCELADO"
            Else
            If Dialog1.FileName <> "" Then
                FileName = Dialog1.FileName
                Set imgPhoto = New WIA.ImageFile
                imgPhoto.LoadFile FileName
                With Picture1
                     'MsgBox "scale Width:" & .ScaleWidth & " --- " & .ScaleMode & "   scale  height:" & .ScaleHeight & " --- " & .ScaleMode
                    Set imgPhoto = ResizeImage(imgPhoto, ScaleX(.ScaleWidth, .ScaleMode, vbPixels), ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
                    Set .Picture = imgPhoto.FileData.Picture
                    Confirmacion = MsgBox("Confirme actualizacion de fotografia", vbYesNo)
                    If Confirmacion = 6 Then
                        If Dir(rutaFotos & "\fotos reducidas\" & Val(txtNumtra) & ".jpg") <> "" Then
                            Confirmacion = MsgBox("La foto ya existe en el Servidor." & vbCrLf & "¿Desea actualizarla?", vbYesNo)
                            If Confirmacion = 6 Then
                                A = MoveFile(Trim$(rutaFotos & "\fotos reducidas\" & Val(txtNumtra) & ".jpg"), Trim(rutaFotos & "\fotos reducidas\R\" & Val(txtNumtra) & ".jpg"))
                                
                                'Dialog2.Filter = "JPG|*.jpg|BMP|*.bmp|GIF|*.gif|PNG|*.png|Todos los archivos|*.*"
                                '.ShowSave
                                'Dialog2.FileName = "D:\Credencial 2002\CANON T3I\TMP2017\nuevas2018\tmp2\" & Val(txtNumtra) & ".jpg"
                                SavePicture Picture1.Picture, rutaFotos & "\fotos\" & Val(txtNumtra) & ".jpg"
                                SavePicture Picture1.Picture, rutaFotos & "\fotos reducidas\" & Val(txtNumtra) & ".jpg"
                                'SavePicture Picture1.Picture, rutaFotos & "r\" & Val(txtNumtra) & ".bmp"
                                'SavePicture Picture1.Picture, "D:\Credencial 2002\CANON T3I\TMP2017\" & Val(txtNumtra) & ".bmp" 'DUPLICAR CON NUMTRA ACT EN TMP2017
                                'SavePicture Picture1.Picture, Dialog2.FileName 'DUPLICAR CON NUMTRA ACT EN TMP2017
                                'SavePicture Picture1.Picture, "\\PIFIVOSTRO6\Credencial 2002\CANON T3I\Foto2016\" & Val(txtNUMTRA) & ".jpg"
                                imgFoto.Picture = LoadPicture(rutaFotos & "\fotos\" & Val(txtNumtra) & ".jpg")
                                'MsgBox "Foto Actualizada Exitosamente!", vbExclamation
                            Else
                                MsgBox "¡Actualizacion Cancelada!", vbExclamation
                            End If
                            
                        Else
                            SavePicture Picture1.Picture, rutaFotos & "\fotos\" & Val(txtNumtra) & ".jpg"
                            SavePicture Picture1.Picture, rutaFotos & "\fotos reducidas\" & Val(txtNumtra) & ".jpg"
                            'SavePicture Picture1.Picture, rutaFotos & "r\" & Val(txtNumtra) & ".bmp"
                            'SavePicture Picture1.Picture, "D:\Credencial 2002\CANON T3I\TMP2017\" & Val(txtNumtra) & ".bmp" 'DUPLICAR EN TMP2017
                            'SavePicture Picture1.Picture, "\\PIFIVOSTRO6\Credencial 2002\CANON T3I\Foto2016\" & Val(txtNUMTRA) & ".jpg"
                            imgFoto.Picture = LoadPicture(rutaFotos & "\fotos\" & Val(txtNumtra) & ".jpg")
                            'MsgBox "Foto Actualizada Exitosamente!", vbExclamation
                        End If
                        
                    Else
                        MsgBox "¡Actualizacion Cancelada!", vbExclamation
                    End If
                End With
            End If
            End If
            Else
            MsgBox "¡Favor de consultar un numero de trabajador antes de actualizar!", vbExclamation
        End If

End Sub

Private Sub Form_Load()
rutaFotos = "D:\prograsrechum\bd\"
End Sub

Private Sub Picture2_Click()

End Sub

Public Sub txtnumtra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

      If txtNumtra.Text <> "" Then
        ' se pone el relojito
        Screen.MousePointer = 11
          
          With Dte
               'se bare la tabla de personal
               '.Personal
               'se busca por el numero de trabajador
               
               '.rspersonal.Find "numtra= " & txtnumtra.Text
                 ' si no lo encunetra despliega un mensaje de error
                 
                'Debug.Print "Antes de buscarnumtra" & Time
                '.Connections(1).Open
                .buscarNumtra CLng(Trim(txtNumtra.Text))
                '.buscarNumtra CLng(Trim(txtnumtra.Text))
               ' Debug.Print "DEspues de buscarnumtra" & Time
                 
                 If .rsbuscarNumtra.EOF = True Then
                   
                   TxtNombre.Text = ""
                   imgFoto.Picture = LoadPicture(rutaFotos & "\fotos reducidas\" & "no.jpg")
                   MsgBox "Empleado no registrado"
                   
                   
                   
                 Else
                 V_Numtra = txtNumtra.Text
                 StrNumtra = txtNumtra.Text
                 
                     'Despues se despleigan todos los datos encontrados segun el permiso que tenga de acceso
                         TxtNombre.Text = Trim(.rsbuscarNumtra!NOMBRE)
                         If FileExists(rutaFotos & "\fotos reducidas\" & Trim(.rsbuscarNumtra!numtra) & ".jpg") Then
                            imgFoto.Picture = LoadPicture(rutaFotos & "\fotos reducidas\" & Trim(.rsbuscarNumtra!numtra) & ".jpg")
                          Else
                            imgFoto.Picture = LoadPicture(rutaFotos & "\fotos reducidas\" & "no.jpg")
                          End If
                          
                          
                          
                
                End If
             
            .rsbuscarNumtra.Close
            'se pone el cursor normal
            Screen.MousePointer = 0
        End With
      End If 'if del txtnnumtra no este vacio
End If ' end if del band

'''''''''''''''''''''''''''
' con esto se restringe al textbox del numtra que solo escriba numeros
If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
   KeyAscii = 0
End If

End Sub


