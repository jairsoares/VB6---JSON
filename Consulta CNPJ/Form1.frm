VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA CNPJ"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtCidade 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   16
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtBairro 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   14
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox txtComplemento 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLogradouro 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   6855
   End
   Begin VB.TextBox txtCEP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFantasia 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   9855
   End
   Begin VB.TextBox txtRazao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   9855
   End
   Begin VB.CommandButton botConsultar 
      Caption         =   "CONSULTAR"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtCNPJ 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Text            =   "12042457000148"
      ToolTipText     =   "CNPJ"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "UF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   9000
      TabIndex        =   19
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "CIDADE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   17
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "BAIRRO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "COMPLEMENTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "NUMERO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9000
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "LOGRADOURO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "CEP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "NOME FANTASIA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "RAZÃO SOCIAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botConsultar_Click()

Dim JSON As String

JSON = FazConsultaCNPJ(txtCNPJ.Text)

txtRazao.Text = ObtenDados(JSON, "nome")
txtFantasia.Text = ObtenDados(JSON, "fantasia")
txtCEP.Text = ObtenDados(JSON, "cep")
txtLogradouro.Text = ObtenDados(JSON, "logradouro")
txtNumero.Text = ObtenDados(JSON, "numero")
txtComplemento.Text = ObtenDados(JSON, "complemento")
txtBairro.Text = ObtenDados(JSON, "bairro")
txtCidade.Text = ObtenDados(JSON, "municipio")
txtUF.Text = ObtenDados(JSON, "uf")


End Sub

