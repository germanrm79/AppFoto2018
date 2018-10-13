VERSION 5.00
Object = "{6E19FD9B-8238-4532-806B-DF15C03FC17B}#1.1#0"; "esW25COM.ocx"
Begin VB.Form frmFirma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capturar Firma"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Activar pad"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   615
      Left            =   7080
      TabIndex        =   2
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   7575
      End
      Begin VB.TextBox txtNumtra 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin eSign3Ctl.esCapture IntegriSign1 
         Height          =   5295
         Left            =   240
         OleObjectBlob   =   "frmFirma.frx":0000
         TabIndex        =   4
         Top             =   720
         Width           =   9615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   8880
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
End
Attribute VB_Name = "frmFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SignData As String
Dim RetVal As Integer
Const iSTR$ = "0123456789"
Public rutaFirmas As String

Private Sub Command1_Click()
Unload Me
End Sub

Sub inicializarFirma()
    SignData = vbNullString
    'set the sign dialog caption
    IntegriSign1.SignDlgCaption = "IntegriSign Signature Capture"
    'set the data to be associated with the signature
    'IntegriSign1.HashData = txtHashData.Text
    'set whether cross mark needs to be displayed if content is modified
    'If chkAppOptions(5).Value = 0 Then
    '    IntegriSign1.Cross = Cross_OFF
    'Else
        IntegriSign1.Cross = Cross_ON
   ' End If
    
    'If chkEnableWhiteSpaceRemoval.Value = 1 Then
        IntegriSign1.EnableWhiteSpaceRemoval = True
        'If chkMaxEnlargementFactor.Value = 1 Then
            IntegriSign1.EnableMaxEnlargementFeature = True
            
            IntegriSign1.MaxEnlargementFactor = "2.0"
        'Else
        '    IntegriSign1.EnableMaxEnlargementFeature = False
        'End If
    'Else
        'IntegriSign1.EnableWhiteSpaceRemoval = False
    'End If
    
    'initiate the act of signing
    IntegriSign1.EnableAntiAliasing = False
    IntegriSign1.SignThickness = 2
    IntegriSign1.StartSign 1, False
    'If chkAppOptions(6).Value = 1 Then
    '    lblft.Caption = vbNullString
    'End If

End Sub
Sub saveSign(numtra As String)
    If IntegriSign1.IsSigned = 0 Then
        MsgBox "No existe firma capturada por favor capture primero antes de guardar!", vbInformation, "IntegriSign"
        Exit Sub
    End If
    'save the image of the signature to a image file
    'the second parameter specifies the image type 0-BMP and 1-JPEG
    Dim ImageFileName As String
    Dim ImageFileType As Integer
    'for bmp
    'ImageFileName = "c:\sign.bmp"
    'ImageFileType = 0
    
    'for jpeg
    'ImageFileName = App.Path & "\Firmas\" & numtra & ".gif"
    ImageFileName = rutaFirmas & numtra & ".gif"
    ImageFileType = 1
    'IntegriSign1.SaveToFile ImageFileName, 180, 110, ImageFileType
    IntegriSign1.SaveToFile ImageFileName, 235, 124, 2, , 1
    MsgBox "Signature image saved successfully." & vbCrLf & "The image file name is " & ImageFileName, vbInformation, "IntegriSign"
End Sub

Private Sub Command2_Click()
Call saveSign(txtNumtra)
Unload Me
End Sub

Private Sub Command3_Click()
 IntegriSign1.ClearSign
End Sub

Private Sub Command4_Click()
    Call inicializarFirma
End Sub

Private Sub Form_Load()
    Call inicializarFirma
    txtNumtra = frmDataEnv.txtNumtra
    TxtNombre = frmDataEnv.TxtNombre
    rutaFirmas = "D:\prograsrechum\bd\firmas\"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
IntegriSign1.CloseConnection
End Sub
