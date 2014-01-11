VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00FFFF80&
   Caption         =   "Home"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7530
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDrop 
      Height          =   315
      ItemData        =   "frmHome.frx":187B1
      Left            =   3120
      List            =   "frmHome.frx":18914
      TabIndex        =   5
      Text            =   "Pick A Champion!"
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      Height          =   1860
      Left            =   3000
      ScaleHeight     =   1800
      ScaleWidth      =   1800
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "                                      Exit!"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   6000
      Picture         =   "frmHome.frx":18CCC
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Label lblNorm 
      BackStyle       =   0  'Transparent
      Caption         =   "                              Back To Norm"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   6000
      Picture         =   "frmHome.frx":1C888
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lblOnTop 
      BackStyle       =   0  'Transparent
      Caption         =   "                               Stay On Top"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   6000
      Picture         =   "frmHome.frx":20444
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "     Show                  Cooldowns!"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblQCD 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   360
      Picture         =   "frmHome.frx":24000
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Image Image6 
      Height          =   630
      Left            =   2880
      Picture         =   "frmHome.frx":27BBC
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   660
      Left            =   1920
      Picture         =   "frmHome.frx":2B9CD
      Top             =   5400
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   1920
      Picture         =   "frmHome.frx":2F263
      Top             =   4560
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   660
      Left            =   1920
      Picture         =   "frmHome.frx":329EA
      Top             =   3720
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   1920
      Picture         =   "frmHome.frx":363BC
      Top             =   2880
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblRCD 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblECD 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblWCD 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image8 
      Height          =   630
      Left            =   2880
      Picture         =   "frmHome.frx":39E1E
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image7 
      Height          =   630
      Left            =   2880
      Picture         =   "frmHome.frx":3DC2F
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   630
      Left            =   2880
      Picture         =   "frmHome.frx":41A40
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Aatrox()
    lblQCD.Caption = " 16 / 15 / 14 / 13 / 12"
    lblWCD.Caption = " 0.5s "
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 100 / 85 / 70"
    picIcon = LoadPicture("Aatrox.gif")
End Sub
Sub Ahri()
    lblQCD.Caption = " 7s "
    lblWCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblECD.Caption = " 12s "
    lblRCD.Caption = " 110 / 95 / 80"
    picIcon = LoadPicture("Ahri.gif")
End Sub
Sub Akali()
    lblQCD.Caption = " 6 / 5.5 / 5 / 4.5 / 4"
    lblWCD.Caption = " 20s"
    lblECD.Caption = " 7 / 6 / 5 / 4 / 3"
    lblRCD.Caption = " 2 / 1.5 / 1"
    picIcon = LoadPicture("Akali.gif")
End Sub
Sub Alistar()
    lblQCD.Caption = " 17 / 16 / 15 / 14 / 13"
    lblWCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 120 / 100 / 80"
    picIcon = LoadPicture("Alistar.gif")
End Sub
Sub Amumu()
    lblQCD.Caption = " 16 / 14 / 12 / 10 / 8"
    lblWCD.Caption = " 1s "
    lblECD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblRCD.Caption = " 150 / 130 / 110"
    picIcon = LoadPicture("Amumu.gif")
End Sub
Sub Anivia()
    lblQCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblWCD.Caption = " 25s "
    lblECD.Caption = " 5s"
    lblRCD.Caption = " 6s"
    picIcon = LoadPicture("Anivia.gif")
End Sub
Sub Annie()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 8s "
    lblECD.Caption = " 10s "
    lblRCD.Caption = " 120s "
    picIcon = LoadPicture("Annie.gif")
End Sub
Sub Ashe()
    lblQCD.Caption = " 0s (Toggle)"
    lblWCD.Caption = " 16 / 13 / 10 / 7 / 4"
    lblECD.Caption = " 60s "
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Ashe.gif")
End Sub
Sub Blitzcrank()
    lblQCD.Caption = " 20 / 19 / 18 / 17 / 16"
    lblWCD.Caption = " 15s "
    lblECD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblRCD.Caption = " 30s"
    picIcon = LoadPicture("Blitzcrank.gif")
End Sub
Sub Brand()
    lblQCD.Caption = " 8 / 7.5 / 7 / 6.5 / 6"
    lblWCD.Caption = " 10s "
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 105 / 90 / 75"
    picIcon = LoadPicture("Brand.gif")
End Sub
Sub Caitlyn()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 20 / 17 / 14 / 11 / 8"
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 90 / 75 / 60"
    picIcon = LoadPicture("Caitlyn.gif")
End Sub
Sub Cassiopeia()
    lblQCD.Caption = " 3s"
    lblWCD.Caption = " 9s "
    lblECD.Caption = " 5s (0.5s If hit on poisoned target)"
    lblRCD.Caption = " 130 / 120 / 110"
    picIcon = LoadPicture("Cassiopeia.gif")
End Sub
Sub ChoGath()
    lblQCD.Caption = " 9s"
    lblWCD.Caption = " 13s "
    lblECD.Caption = " 0s (Toggle)"
    lblRCD.Caption = " 60s"
    picIcon = LoadPicture("ChoGath.gif")
End Sub
Sub Corki()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 26 / 23 / 20 / 17 / 14"
    lblECD.Caption = " 16s"
    lblRCD.Caption = " 12 / 10 / 8 (2s Between each rocket shot)"
    picIcon = LoadPicture("Corki.gif")
End Sub
Sub Darius()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 8s "
    lblECD.Caption = " 24 / 21 / 18 / 15 / 12"
    lblRCD.Caption = " 120 / 100 / 80 (Available for 12s after kill)"
    picIcon = LoadPicture("Darius.gif")
End Sub
Sub Diana()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 10s "
    lblECD.Caption = " 26 / 24 / 22 / 20 / 18"
    lblRCD.Caption = " 25 / 20 / 15 (No CD on Moonlight Target)"
    picIcon = LoadPicture("Diana.gif")
End Sub
Sub DrMundo()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 4s "
    lblECD.Caption = " 7s"
    lblRCD.Caption = " 75s"
    picIcon = LoadPicture("DrMundo.gif")
End Sub
Sub Draven()
    lblQCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblWCD.Caption = " 12s (CD Refreshed when Axe Caught)"
    lblECD.Caption = " 18 / 17 / 16 / 15 / 14"
    lblRCD.Caption = " 110 / 100 / 90"
    picIcon = LoadPicture("Draven.gif")
End Sub
Sub Elise()
    lblQCD.Caption = " 6s (6s)"
    lblWCD.Caption = " 12s (12s) "
    lblECD.Caption = " 14 / 13 / 12 / 11 / 10 (26/24/22/20/18)"
    lblRCD.Caption = " 4s"
    picIcon = LoadPicture("Elise.gif")
End Sub
Sub Evelynn()
    lblQCD.Caption = " 1.5s"
    lblWCD.Caption = " 15s "
    lblECD.Caption = " 9s"
    lblRCD.Caption = " 150 / 120 / 90"
    picIcon = LoadPicture("Evelynn.gif")
End Sub
Sub Ezreal()
    lblQCD.Caption = " 6 / 5.5 / 5 / 4.5 / 4"
    lblWCD.Caption = " 9s "
    lblECD.Caption = " 19 / 17 / 15 / 13 / 11"
    lblRCD.Caption = " 80s"
    picIcon = LoadPicture("Ezreal.gif")
End Sub
Sub FiddleSticks()
    lblQCD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblWCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblECD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblRCD.Caption = " 150 / 140 / 130"
    picIcon = LoadPicture("FiddleSticks.gif")
End Sub
Sub Fiora()
    lblQCD.Caption = " 16 / 14 / 12 / 10 / 8"
    lblWCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblECD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblRCD.Caption = " 130 / 120 / 110"
    picIcon = LoadPicture("Fiora.gif")
End Sub
Sub Fizz()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 10s "
    lblECD.Caption = " 16 / 14 / 12 / 10 / 8"
    lblRCD.Caption = " 100 / 85 / 70"
    picIcon = LoadPicture("Fizz.gif")
End Sub
Sub Galio()
    lblQCD.Caption = " 7s"
    lblWCD.Caption = " 13s "
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 150 / 135 / 120"
    picIcon = LoadPicture("Galio.gif")
End Sub
Sub Gangplank()
    lblQCD.Caption = " 5s"
    lblWCD.Caption = " 22 / 21 / 20 / 19 / 18"
    lblECD.Caption = " 20s"
    lblRCD.Caption = " 120 / 115 / 110"
    picIcon = LoadPicture("Gangplank.gif")
End Sub
Sub Garen()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 24 / 23 / 22 / 21 / 20"
    lblECD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblRCD.Caption = " 160 / 120 / 80"
    picIcon = LoadPicture("Garen.gif")
End Sub
Sub Gragas()
    lblQCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblWCD.Caption = " 25s "
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Gragas.gif")
End Sub
Sub Graves()
    lblQCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblWCD.Caption = " 20 / 19 / 18 / 17 / 16"
    lblECD.Caption = " 22 / 20 / 18 / 16 / 14"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Graves.gif")
End Sub
Sub Hecarim()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 22 / 21 / 20 / 19 / 18"
    lblECD.Caption = " 24 / 22 / 20 / 18 / 16"
    lblRCD.Caption = " 140 / 120 / 100"
    picIcon = LoadPicture("Hecarim.gif")
End Sub
Sub Heimerdinger()
    lblQCD.Caption = " 24 / 23 / 22 / 21 / 20 (1s CD Between Cast)"
    lblWCD.Caption = " 11s "
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 100 / 80 / 60"
    picIcon = LoadPicture("Heimerdinger.gif")
End Sub
Sub Irelia()
    lblQCD.Caption = " 14 / 12 / 10 / 8 / 6"
    lblWCD.Caption = " 15s "
    lblECD.Caption = " 8s"
    lblRCD.Caption = " 70 / 60 / 50"
    picIcon = LoadPicture("Irelia.gif")
End Sub
Sub Janna()
    lblQCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblWCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 150 / 135 / 120"
    picIcon = LoadPicture("Janna.gif")
End Sub
Sub Jarvan()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 20 / 18 / 16 / 14 / 12"
    lblECD.Caption = " 13s"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Jarvan.gif")
End Sub
Sub Jax()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 7 / 6 / 5 / 4 / 3"
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 80s"
    picIcon = LoadPicture("Jax.gif")
End Sub
Sub Jayce()
    lblQCD.Caption = " 16 / 14 / 12 / 10 / 8 (8s)"
    lblWCD.Caption = " 10s (14/12/10/8/6) "
    lblECD.Caption = " 14 / 13 / 12 / 11 / 10 (16s)"
    lblRCD.Caption = " 6s"
    picIcon = LoadPicture("Jayce.gif")
End Sub
Sub Jinx()
    lblQCD.Caption = " 0.9s"
    lblWCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblECD.Caption = " 24 / 22 / 20 / 18 / 16"
    lblRCD.Caption = " 90 / 75 / 60"
    picIcon = LoadPicture("Jinx.gif")
End Sub
Sub Karma()
    lblQCD.Caption = " 7 / 6.5 / 6 / 5.5 / 5"
    lblWCD.Caption = " 16 / 15 / 14 / 13 / 12"
    lblECD.Caption = " 10"
    lblRCD.Caption = " 45 / 42 / 39 / 36 (CD Lowered By Auto(1s) & Spells(2s)"
    picIcon = LoadPicture("Karma.gif")
End Sub
Sub Karthus()
    lblQCD.Caption = " 1s"
    lblWCD.Caption = " 18s "
    lblECD.Caption = " 0s (Toggle)"
    lblRCD.Caption = " 200 / 180 / 160"
    picIcon = LoadPicture("Karthus.gif")
End Sub
Sub Kassadin()
    lblQCD.Caption = " 9s"
    lblWCD.Caption = " 12s"
    lblECD.Caption = " 6s"
    lblRCD.Caption = " 7 / 6 / 5"
    picIcon = LoadPicture("Kassadin.gif")
End Sub
Sub Katarina()
    lblQCD.Caption = " 10 / 9.5 / 9 / 8.5 / 8"
    lblWCD.Caption = " 4s "
    lblECD.Caption = " 12 / 10.5 / 9 / 7.5 / 6"
    lblRCD.Caption = " 60 / 52.5 / 45"
    picIcon = LoadPicture("Katarina.gif")
End Sub
Sub Kayle()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 15s "
    lblECD.Caption = " 16s"
    lblRCD.Caption = " 90 / 75 / 60"
    picIcon = LoadPicture("Kayle.gif")
End Sub
Sub Kennen()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 14 / 12 / 10 / 8 / 6"
    lblECD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblRCD.Caption = " 120s"
    picIcon = LoadPicture("Kennen.gif")
End Sub
Sub Khazix()
    lblQCD.Caption = " 3.5s"
    lblWCD.Caption = " 8s "
    lblECD.Caption = " 22 / 20 / 18 / 16 / 14 (Reset on K/A when evolved)"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Khazix.gif")
End Sub
Sub KogMaw()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 17s "
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 2 / 1.5 / 1"
    picIcon = LoadPicture("KogMaw.gif")
End Sub
Sub LeBlanc()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblECD.Caption = " 14 / 12.5 / 10 / 9.5 / 8"
    lblRCD.Caption = " 40 / 32 / 24"
    picIcon = LoadPicture("LeBlanc.gif")
End Sub
Sub LeeSin()
    lblQCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblWCD.Caption = " 9s "
    lblECD.Caption = " 10s "
    lblRCD.Caption = " 90 / 75 / 60"
    picIcon = LoadPicture("LeeSin.gif")
End Sub
Sub Leona()
    lblQCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblWCD.Caption = " 14s "
    lblECD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblRCD.Caption = " 90 / 75 / 60"
    picIcon = LoadPicture("Leona.gif")
End Sub
Sub Lissandra()
    lblQCD.Caption = " 6 / 5.5 / 5 / 4.5 / 4"
    lblWCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblECD.Caption = " 24 / 21 / 18 / 15 / 12"
    lblRCD.Caption = " 130 / 105 / 80"
    picIcon = LoadPicture("Lissandra.gif")
End Sub
Sub Lucian()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 100 / 75 / 50"
    picIcon = LoadPicture("Lucian.gif")
End Sub
Sub Lulu()
    lblQCD.Caption = " 7s"
    lblWCD.Caption = " 18 / 16.5 / 15 / 13.5 / 12"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 110 / 95 / 80"
    picIcon = LoadPicture("Lulu.gif")
End Sub
Sub Lux()
    lblQCD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblWCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 80 / 65 / 50"
    picIcon = LoadPicture("Lux.gif")
End Sub
Sub Malphite()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 14s"
    lblECD.Caption = " 7s"
    lblRCD.Caption = " 130 / 115 / 100"
    picIcon = LoadPicture("Malphite.gif")
End Sub
Sub Malzahar()
    lblQCD.Caption = " 9s"
    lblWCD.Caption = " 14s"
    lblECD.Caption = " 15 / 13 / 11 / 9 / 7"
    lblRCD.Caption = " 120 / 100 / 80"
    picIcon = LoadPicture("Malzahar.gif")
End Sub
Sub Maokai()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 13s"
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 40 / 30 / 20"
    picIcon = LoadPicture("Maokai.gif")
End Sub
Sub MasterYi()
    lblQCD.Caption = " 18 / 17 / 16 / 15 / 14 (CD Lowered By Auto Attack)"
    lblWCD.Caption = " 35s"
    lblECD.Caption = " 18 / 17 / 16 / 15 / 14"
    lblRCD.Caption = " 75s"
    picIcon = LoadPicture("MasterYi.gif")
End Sub
Sub MissFortune()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 16s"
    lblECD.Caption = " 15s"
    lblRCD.Caption = " 120 / 110 / 100"
    picIcon = LoadPicture("MissFortune.gif")
End Sub
Sub MordeKaiser()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 20 / 18 / 16 / 14 / 12"
    lblECD.Caption = " 6s"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("MordeKaiser.gif")
End Sub
Sub Morgana()
    lblQCD.Caption = " 11s"
    lblWCD.Caption = " 10s"
    lblECD.Caption = " 23 / 21 / 19 / 17 / 15"
    lblRCD.Caption = " 120 / 110 / 100"
    picIcon = LoadPicture("Morgana.gif")
End Sub
Sub Nami()
    lblQCD.Caption = " 14 / 13 / 12 / 11 / 10"
    lblWCD.Caption = " 9s"
    lblECD.Caption = " 11s"
    lblRCD.Caption = " 120 / 110 / 100"
    picIcon = LoadPicture("Nami.gif")
End Sub
Sub Nasus()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 120s"
    picIcon = LoadPicture("Nasus.gif")
End Sub
Sub Nautilus()
    lblQCD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblWCD.Caption = " 22 / 21 / 20 / 19 / 18"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 140 / 110 / 80"
    picIcon = LoadPicture("Nautilus.gif")
End Sub
Sub Nidalee()
    lblQCD.Caption = " 6s (5s)"
    lblWCD.Caption = " 18s (3.5s)"
    lblECD.Caption = " 10s (6s)"
    lblRCD.Caption = " 4s"
    picIcon = LoadPicture("Nidalee.gif")
End Sub
Sub Nocturne()
    lblQCD.Caption = " 10s"
    lblWCD.Caption = " 20 / 18 / 16 / 14 / 12"
    lblECD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblRCD.Caption = " 180 / 140 / 100"
    picIcon = LoadPicture("Nocturne.gif")
End Sub
Sub Nunu()
    lblQCD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblWCD.Caption = " 15s"
    lblECD.Caption = " 6s"
    lblRCD.Caption = " 110 / 100 / 90"
    picIcon = LoadPicture("Nunu.gif")
End Sub
Sub Olaf()
    lblQCD.Caption = " 8s (Picking up lowers CD by 4.5s)"
    lblWCD.Caption = " 16s"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8 (Autos Lower CD by 1s)"
    lblRCD.Caption = " 120 / 100 / 80"
    picIcon = LoadPicture("Olaf.gif")
End Sub
Sub Orianna()
    lblQCD.Caption = " 6 / 5.25 / 4.5 / 3.75 / 3"
    lblWCD.Caption = " 9s"
    lblECD.Caption = " 9s"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Orianna.gif")
End Sub
Sub Pantheon()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblECD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblRCD.Caption = " 150 / 135 / 120"
    picIcon = LoadPicture("Pantheon.gif")
End Sub
Sub Poppy()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 12s"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 140 / 120 / 100"
    picIcon = LoadPicture("Poppy.gif")
End Sub
Sub Quinn()
    lblQCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblWCD.Caption = " 50 / 45 / 40 / 35 / 30"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 140 / 110 / 80"
    picIcon = LoadPicture("Quinn.gif")
End Sub
Sub Rammus()
    lblQCD.Caption = " 10s"
    lblWCD.Caption = " 14s"
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 60s"
    picIcon = LoadPicture("Rammus.gif")
End Sub
Sub Renekton()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblECD.Caption = " 18 / 17 / 16 / 15 / 14"
    lblRCD.Caption = " 120"
    picIcon = LoadPicture("Renekton.gif")
End Sub
Sub Rengar()
    lblQCD.Caption = " 8 / 7.5 / 7 / 6.5 / 6"
    lblWCD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 140 / 105 / 70"
    picIcon = LoadPicture("Rengar.gif")
End Sub
Sub Riven()
    lblQCD.Caption = " 13s"
    lblWCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblECD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblRCD.Caption = " 110 / 80 / 50"
    picIcon = LoadPicture("Riven.gif")
End Sub
Sub Rumble()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 6s"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 105 / 90 / 75"
    picIcon = LoadPicture("Rumble.gif")
End Sub
Sub Ryze()
    lblQCD.Caption = " 3.5s (Reduced By Other Spell Casts by 1s)"
    lblWCD.Caption = " 14s (Reduced By Other Spell Casts by 1s)"
    lblECD.Caption = " 14s (Reduced By Other Spell Casts by 1s)"
    lblRCD.Caption = " 70 / 60 / 50 (Reduced By Other Spell Casts by 1s)"
    picIcon = LoadPicture("Ryze.gif")
End Sub
Sub Sejuani()
    lblQCD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblWCD.Caption = " 11 / 10 / 9 / 8 / 7"
    lblECD.Caption = " 11s"
    lblRCD.Caption = " 130 / 115 / 100"
    picIcon = LoadPicture("Sejuani.gif")
End Sub
Sub Shaco()
    lblQCD.Caption = " 11s (After Leaving Stealth)"
    lblWCD.Caption = " 16s"
    lblECD.Caption = " 8s"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Shaco.gif")
End Sub
Sub Shen()
    lblQCD.Caption = " 6 / 5.5 / 5 / 4.5 / 4"
    lblWCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblECD.Caption = " 16 / 14 / 12 / 10 / 8"
    lblRCD.Caption = " 200 / 180 / 160"
    picIcon = LoadPicture("Shen.gif")
End Sub
Sub Shyvana()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 12s (Shyvanas Auto's increase duration)"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 1/2/3 (Fury gained per second per level, Autos gain 2 fury)"
    picIcon = LoadPicture("Shyvana.gif")
End Sub
Sub Singed()
    lblQCD.Caption = " 0s (Toggle)"
    lblWCD.Caption = " 14s"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 100s"
    picIcon = LoadPicture("Singed.gif")
End Sub
Sub Sion()
    lblQCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblWCD.Caption = " 8s (CD Starts after shield expires)"
    lblECD.Caption = " 0s (Toggle)"
    lblRCD.Caption = " 90s"
    picIcon = LoadPicture("Sion.gif")
End Sub
Sub Sivir()
    lblQCD.Caption = " 9s"
    lblWCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblECD.Caption = " 22 / 19 / 16 / 13 / 10"
    lblRCD.Caption = " 120 / 100 / 80"
    picIcon = LoadPicture("Sivir.gif")
End Sub
Sub Skarner()
    lblQCD.Caption = " 3.5s (Reduced 0.5s By Autos)"
    lblWCD.Caption = " 18s (Reduced 0.5s By Autos)"
    lblECD.Caption = " 10s (Reduced 0.5s By Autos)"
    lblRCD.Caption = " 130 / 120 / 110 (Reduced 0.5s By Autos)"
    picIcon = LoadPicture("Skarner.gif")
End Sub
Sub Sona()
    lblQCD.Caption = " 7s"
    lblWCD.Caption = " 7s"
    lblECD.Caption = " 7s"
    lblRCD.Caption = " 140 / 120 / 100"
    picIcon = LoadPicture("Sona.gif")
End Sub
Sub Soraka()
    lblQCD.Caption = " 2.5s"
    lblWCD.Caption = " 20s"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 160 / 145 / 130"
    picIcon = LoadPicture("Soraka.gif")
End Sub
Sub Swain()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblECD.Caption = " 10s"
    lblRCD.Caption = " 8s"
    picIcon = LoadPicture("Swain.gif")
End Sub
Sub Syndra()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblECD.Caption = " 18 / 16.50 / 15 / 13.5 / 12"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Syndra.gif")
End Sub
Sub Talon()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 10s"
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 75 / 65 / 55"
    picIcon = LoadPicture("Talon.gif")
End Sub
Sub Taric()
    lblQCD.Caption = " 18 / 17 / 16 / 15 / 14 (Reduced 2 Seconds on Auto After Cast)"
    lblWCD.Caption = " 10s (Reduced 2 Seconds on Auto After Cast)"
    lblECD.Caption = " 18 / 17 / 16 / 15 / 14 (Reduced 2 Seconds on Auto After Cast)"
    lblRCD.Caption = " 100 / 75 / 50 (Reduced 2 Seconds on Auto After Cast)"
    picIcon = LoadPicture("Taric.gif")
End Sub
Sub Teemo()
    lblQCD.Caption = " 8s"
    lblWCD.Caption = " 17s"
    lblECD.Caption = " N/A "
    lblRCD.Caption = " 35 / 31 / 27 (1s CD between Casts)"
    picIcon = LoadPicture("Teemo.gif")
End Sub
Sub Thresh()
    lblQCD.Caption = " 20 / 18 / 16 / 14 / 12"
    lblWCD.Caption = " 22 / 20.5 / 19 / 17.5 / 16"
    lblECD.Caption = " 9s"
    lblRCD.Caption = " 150 / 140 / 130"
    picIcon = LoadPicture("Thresh.gif")
End Sub
Sub Tristana()
    lblQCD.Caption = " 20s"
    lblWCD.Caption = " 22 / 20 / 18 / 16 / 14 (CD Reset After Kill or Assist)"
    lblECD.Caption = " 16s"
    lblRCD.Caption = " 60s"
    picIcon = LoadPicture("Tristana.gif")
End Sub
Sub Trundle()
    lblQCD.Caption = " 4s"
    lblWCD.Caption = " 15s"
    lblECD.Caption = " 23 / 20 / 17 / 14 / 11"
    lblRCD.Caption = " 80 / 70 / 60"
    picIcon = LoadPicture("Trundle.gif")
End Sub
Sub Tryndamere()
    lblQCD.Caption = " 12s"
    lblWCD.Caption = " 14s"
    lblECD.Caption = " 13/12/11/10/9 (Reduced By Criticals(1s for Minions, 2s for Champs)"
    lblRCD.Caption = " 110 / 100 / 90"
    picIcon = LoadPicture("Tryndamere.gif")
End Sub
Sub TwistedFate()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 6s"
    lblECD.Caption = " N/A"
    lblRCD.Caption = " 180 / 150 / 120"
    picIcon = LoadPicture("TwistedFate.gif")
End Sub
Sub Twitch()
    lblQCD.Caption = " 16s"
    lblWCD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 120 / 110 / 100"
    picIcon = LoadPicture("Twitch.gif")
End Sub
Sub Udyr()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 6s"
    lblECD.Caption = " 6s"
    lblRCD.Caption = " 6s"
    picIcon = LoadPicture("Udyr.gif")
End Sub
Sub Urgot()
    lblQCD.Caption = " 2s"
    lblWCD.Caption = " 16 / 15 / 14 / 13 / 12"
    lblECD.Caption = " 15 / 14 / 13 / 12 / 11"
    lblRCD.Caption = " 120s"
    picIcon = LoadPicture("Urgot.gif")
End Sub
Sub Varus()
    lblQCD.Caption = " 16 / 14 / 12 / 10 / 8"
    lblWCD.Caption = " N/A "
    lblECD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Varus.gif")
End Sub
Sub Vayne()
    lblQCD.Caption = " 6 / 5 / 4 / 3 / 2"
    lblWCD.Caption = " N/A "
    lblECD.Caption = " 20 / 18 / 16 / 14 / 12"
    lblRCD.Caption = " 100 / 85 / 70"
    picIcon = LoadPicture("Vayne.gif")
End Sub
Sub Veigar()
    lblQCD.Caption = " 8 / 7 / 6 / 5 / 4"
    lblWCD.Caption = " 10s"
    lblECD.Caption = " 20 / 19 / 18 / 17 / 16"
    lblRCD.Caption = " 130 / 110 / 90"
    picIcon = LoadPicture("Veigar.gif")
End Sub
Sub Vi()
    lblQCD.Caption = " 18 / 15.5 / 13 / 10.5 / 8"
    lblWCD.Caption = " N/A "
    lblECD.Caption = " 14 / 12.5 / 11 / 9.5 / 8 (1s Cooldown between casts)"
    lblRCD.Caption = " 130 / 105 / 80"
    picIcon = LoadPicture("Vi.gif")
End Sub
Sub Viktor()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 17 / 16 / 15 / 14 / 13"
    lblECD.Caption = " 13 / 12 / 11 / 10 / 9"
    lblRCD.Caption = " 120s"
    picIcon = LoadPicture("Viktor.gif")
End Sub
Sub Vladimir()
    lblQCD.Caption = " 10 / 8.5 / 7 / 5.5 / 4"
    lblWCD.Caption = " 26 / 23 / 20 / 17 / 14"
    lblECD.Caption = " 4.5s"
    lblRCD.Caption = " 150 / 135 / 120"
    picIcon = LoadPicture("Vladimir.gif")
End Sub
Sub Volibear()
    lblQCD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblWCD.Caption = " 18s"
    lblECD.Caption = " 11s"
    lblRCD.Caption = " 100 / 90 / 80"
    picIcon = LoadPicture("Volibear.gif")
End Sub
Sub WarWick()
    lblQCD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblWCD.Caption = " 24 / 22 / 20 / 18 / 16"
    lblECD.Caption = " 4s (Toggle)"
    lblRCD.Caption = " 90 / 80 / 70"
    picIcon = LoadPicture("Warwick.gif")
End Sub
Sub Wukong()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 18 / 16 / 14 / 12 / 10"
    lblECD.Caption = " 8s"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Wukong.gif")
End Sub
Sub Xerath()
    lblQCD.Caption = " 7 / 6.5 / 6 / 5.5 / 5"
    lblWCD.Caption = " 20 / 16 / 12 / 8 / 4"
    lblECD.Caption = " 12 / 11 / 10 / 9 / 8"
    lblRCD.Caption = " 80 / 70 / 60"
    picIcon = LoadPicture("Xerath.gif")
End Sub
Sub XinZhao()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5 "
    lblWCD.Caption = " 16 / 15 / 14 / 13 / 12 (Reduced by 1s every Auto Q)"
    lblECD.Caption = " 14 / 13 / 12 / 11 / 10 (Reduced by 1s every Auto Q)"
    lblRCD.Caption = " 120 / 110 / 100 (Reduced by 1s every Auto Q)"
    picIcon = LoadPicture("XinZhao.gif")
End Sub
Sub Yasuo()
    lblQCD.Caption = " 5 / 4.75 / 4.5 / 4.25 / 4"
    lblWCD.Caption = " 26 / 24 / 22 / 20 / 18"
    lblECD.Caption = " 0.5/0.4/0.3/0.2/0.1 (Duration Mark Lasts; 10/9/8/7/6)"
    lblRCD.Caption = " 80 / 55 / 30"
    picIcon = LoadPicture("Yasuo.gif")
End Sub
Sub Yorick()
    lblQCD.Caption = " 9 / 8 / 7 / 6 / 5"
    lblWCD.Caption = " 12s"
    lblECD.Caption = " 10 / 9 / 8 / 7 / 6"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Yorick.gif")
End Sub
Sub Zac()
    lblQCD.Caption = " 9 / 8.5 / 8 / 7.5 / 7"
    lblWCD.Caption = " 4s"
    lblECD.Caption = " 24 / 21 / 18 / 15 / 12"
    lblRCD.Caption = " 130 / 115 / 100"
    picIcon = LoadPicture("Zac.gif")
End Sub
Sub Zed()
    lblQCD.Caption = " 6s"
    lblWCD.Caption = " 18/17/16/15/14 (Reduced if struck by Shadow Slash by 2s"
    lblECD.Caption = " 4s"
    lblRCD.Caption = " 120 / 100 / 80"
    picIcon = LoadPicture("Zed.gif")
End Sub
Sub Ziggs()
    lblQCD.Caption = " 6 / 5.5 / 5 / 4.5 / 4"
    lblWCD.Caption = " 26 / 24 / 22 / 20 / 18"
    lblECD.Caption = " 16s"
    lblRCD.Caption = " 120 / 105 / 90"
    picIcon = LoadPicture("Ziggs.gif")
End Sub
Sub Zilean()
    lblQCD.Caption = " 10s (Reduced By Rewind By 10s)"
    lblWCD.Caption = " 18 / 15 / 12 / 9 / 6"
    lblECD.Caption = " 20s (Reduced By Rewind By 10s)"
    lblRCD.Caption = " 180s (Reduced By Rewind By 10s)"
    picIcon = LoadPicture("Zilean.gif")
End Sub
Sub Zyra()
    lblQCD.Caption = " 7 / 6.5 / 6 / 5.5 / 5"
    lblWCD.Caption = " 16.7 / 15.4 / 14.1 / 12.9 / 11.7"
    lblECD.Caption = " 12s"
    lblRCD.Caption = " 130 / 120 / 110"
    picIcon = LoadPicture("Zyra.gif")
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdShow_Click()
    If cboDrop = "Aatrox" Then
        Aatrox
    End If
    If cboDrop = "Ahri" Then
        Ahri
    End If
        If cboDrop = "Akali" Then
        Akali
    End If
    If cboDrop = "Alistar" Then
        Alistar
    End If
        If cboDrop = "Amumu" Then
        Amumu
    End If
    If cboDrop = "Anivia" Then
        Anivia
    End If
        If cboDrop = "Annie" Then
        Annie
    End If
    If cboDrop = "Ashe" Then
        Ashe
    End If
        If cboDrop = "Blitzcrank" Then
        Blitzcrank
    End If
    If cboDrop = "Brand" Then
        Brand
    End If
        If cboDrop = "Caitlyn" Then
        Caitlyn
    End If
    If cboDrop = "Cassiopeia" Then
        Cassiopeia
    End If
        If cboDrop = "ChoGath" Then
        ChoGath
    End If
    If cboDrop = "Corki" Then
        Corki
    End If
        If cboDrop = "Darius" Then
        Darius
    End If
    If cboDrop = "Diana" Then
        Diana
    End If
        If cboDrop = "DrMundo" Then
        DrMundo
    End If
    If cboDrop = "Draven" Then
        Draven
    End If
        If cboDrop = "Elise" Then
        Elise
    End If
    If cboDrop = "Evelynn" Then
        Evelynn
    End If
        If cboDrop = "Ezreal" Then
        Ezreal
    End If
    If cboDrop = "FiddleSticks" Then
        FiddleSticks
    End If
        If cboDrop = "Fiora" Then
        Fiora
    End If
    If cboDrop = "Fizz" Then
        Fizz
    End If
        If cboDrop = "Galio" Then
        Galio
    End If
        If cboDrop = "Gangplank" Then
        Gangplank
    End If
    If cboDrop = "Garen" Then
        Garen
    End If
        If cboDrop = "Gragas" Then
        Gragas
    End If
    If cboDrop = "Graves" Then
        Graves
    End If
        If cboDrop = "Hecarim" Then
        Hecarim
    End If
    If cboDrop = "Heimerdinger" Then
        Heimerdinger
    End If
        If cboDrop = "Irelia" Then
        Irelia
    End If
    If cboDrop = "Janna" Then
        Janna
    End If
        If cboDrop = "Jarvan" Then
        Jarvan
    End If
    If cboDrop = "Jax" Then
        Jax
    End If
        If cboDrop = "Jayce" Then
        Jayce
    End If
    If cboDrop = "Jinx" Then
        Jinx
    End If
        If cboDrop = "Karma" Then
        Karma
    End If
    If cboDrop = "Karthus" Then
        Karthus
    End If
        If cboDrop = "Kassadin" Then
        Kassadin
    End If
    If cboDrop = "Katarina" Then
        Katarina
    End If
        If cboDrop = "Kayle" Then
        Kayle
    End If
        If cboDrop = "Kennen" Then
        Kennen
    End If
    If cboDrop = "KhaZix" Then
        Khazix
    End If
        If cboDrop = "KogMaw" Then
        KogMaw
    End If
    If cboDrop = "LeBlanc" Then
        LeBlanc
    End If
        If cboDrop = "LeeSin" Then
        LeeSin
    End If
    If cboDrop = "Leona" Then
        Leona
    End If
        If cboDrop = "Lissandra" Then
        Lissandra
    End If
    If cboDrop = "Lucian" Then
        Lucian
    End If
        If cboDrop = "Lulu" Then
        Lulu
    End If
    If cboDrop = "Lux" Then
        Lux
    End If
        If cboDrop = "Malphite" Then
        Malphite
    End If
    If cboDrop = "Malzahar" Then
        Malzahar
    End If
        If cboDrop = "Maokai" Then
        Maokai
    End If
    If cboDrop = "MasterYi" Then
        MasterYi
    End If
        If cboDrop = "MissFortune" Then
        MissFortune
    End If
    If cboDrop = "Mordekaiser" Then
        MordeKaiser
    End If
        If cboDrop = "Morgana" Then
        Morgana
    End If
        If cboDrop = "Nami" Then
        Nami
    End If
    If cboDrop = "Nasus" Then
        Nautilus
    End If
        If cboDrop = "Nidalee" Then
        Nidalee
    End If
    If cboDrop = "Nocturne" Then
        Nocturne
    End If
        If cboDrop = "Nunu" Then
        Nunu
    End If
    If cboDrop = "Olaf" Then
        Olaf
    End If
        If cboDrop = "Orianna" Then
        Orianna
    End If
    If cboDrop = "Pantheon" Then
        Pantheon
    End If
        If cboDrop = "Poppy" Then
        Poppy
    End If
    If cboDrop = "Quinn" Then
        Quinn
    End If
        If cboDrop = "Rammus" Then
        Rammus
    End If
    If cboDrop = "Renekton" Then
        Renekton
    End If
        If cboDrop = "Rengar" Then
        Rengar
    End If
    If cboDrop = "Riven" Then
        Riven
    End If
        If cboDrop = "Rumble" Then
        Rumble
    End If
    If cboDrop = "Ryze" Then
        Ryze
    End If
        If cboDrop = "Sejuani" Then
        Sejuani
    End If
        If cboDrop = "Shaco" Then
        Shaco
    End If
    If cboDrop = "Shen" Then
        Shen
    End If
        If cboDrop = "Shyvana" Then
        Shyvana
    End If
    If cboDrop = "Singed" Then
        Singed
    End If
        If cboDrop = "Sion" Then
        Sion
    End If
    If cboDrop = "Sivir" Then
        Sivir
    End If
        If cboDrop = "Skarner" Then
        Skarner
    End If
    If cboDrop = "Sona" Then
        Sona
    End If
        If cboDrop = "Soraka" Then
        Soraka
    End If
    If cboDrop = "Swain" Then
        Swain
    End If
        If cboDrop = "Syndra" Then
        Syndra
    End If
    If cboDrop = "Talon" Then
        Talon
    End If
        If cboDrop = "Taric" Then
        Taric
    End If
    If cboDrop = "Teemo" Then
        Teemo
    End If
        If cboDrop = "Thresh" Then
        Thresh
    End If
    If cboDrop = "Tristana" Then
        Tristana
    End If
        If cboDrop = "Trundle" Then
        Trundle
    End If
        If cboDrop = "Tryndamere" Then
        Tryndamere
    End If
    If cboDrop = "TwistedFate" Then
        TwistedFate
    End If
        If cboDrop = "Twitch" Then
        Twitch
    End If
    If cboDrop = "Udyr" Then
        Udyr
    End If
        If cboDrop = "Urgot" Then
        Urgot
    End If
    If cboDrop = "Varus" Then
        Varus
    End If
        If cboDrop = "Vayne" Then
        Vayne
    End If
    If cboDrop = "Veigar" Then
        Veigar
    End If
        If cboDrop = "Vi" Then
        Vi
    End If
    If cboDrop = "Viktor" Then
        Viktor
    End If
        If cboDrop = "Vladimir" Then
        Vladimir
    End If
    If cboDrop = "Volibear" Then
        Volibear
    End If
        If cboDrop = "Warwick" Then
        WarWick
    End If
    If cboDrop = "Wukong" Then
        Wukong
    End If
        If cboDrop = "Xerath" Then
        Xerath
    End If
    If cboDrop = "XinZhao" Then
        XinZhao
    End If
        If cboDrop = "Yasuo" Then
        Yasuo
    End If
        If cboDrop = "Yorick" Then
        Yorick
    End If
    If cboDrop = "Zac" Then
        Zac
    End If
        If cboDrop = "Zed" Then
        Zed
    End If
    If cboDrop = "Ziggs" Then
        Ziggs
    End If
        If cboDrop = "Zilean" Then
        Zilean
    End If
    If cboDrop = "Zyra" Then
        Zyra
    End If
    picIcon.Visible = True
    Image1.Visible = True
    Image2.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    lblQCD.Visible = True
    lblWCD.Visible = True
    lblECD.Visible = True
    lblRCD.Visible = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cboDrop = "Aatrox" Then
        Aatrox
    End If
    If cboDrop = "Ahri" Then
        Ahri
    End If
        If cboDrop = "Akali" Then
        Akali
    End If
    If cboDrop = "Alistar" Then
        Alistar
    End If
        If cboDrop = "Amumu" Then
        Amumu
    End If
    If cboDrop = "Anivia" Then
        Anivia
    End If
        If cboDrop = "Annie" Then
        Annie
    End If
    If cboDrop = "Ashe" Then
        Ashe
    End If
        If cboDrop = "Blitzcrank" Then
        Blitzcrank
    End If
    If cboDrop = "Brand" Then
        Brand
    End If
        If cboDrop = "Caitlyn" Then
        Caitlyn
    End If
    If cboDrop = "Cassiopeia" Then
        Cassiopeia
    End If
        If cboDrop = "ChoGath" Then
        ChoGath
    End If
    If cboDrop = "Corki" Then
        Corki
    End If
        If cboDrop = "Darius" Then
        Darius
    End If
    If cboDrop = "Diana" Then
        Diana
    End If
        If cboDrop = "DrMundo" Then
        DrMundo
    End If
    If cboDrop = "Draven" Then
        Draven
    End If
        If cboDrop = "Elise" Then
        Elise
    End If
    If cboDrop = "Evelynn" Then
        Evelynn
    End If
        If cboDrop = "Ezreal" Then
        Ezreal
    End If
    If cboDrop = "FiddleSticks" Then
        FiddleSticks
    End If
        If cboDrop = "Fiora" Then
        Fiora
    End If
    If cboDrop = "Fizz" Then
        Fizz
    End If
        If cboDrop = "Galio" Then
        Galio
    End If
        If cboDrop = "Gangplank" Then
        Gangplank
    End If
    If cboDrop = "Garen" Then
        Garen
    End If
        If cboDrop = "Gragas" Then
        Gragas
    End If
    If cboDrop = "Graves" Then
        Graves
    End If
        If cboDrop = "Hecarim" Then
        Hecarim
    End If
    If cboDrop = "Heimerdinger" Then
        Heimerdinger
    End If
        If cboDrop = "Irelia" Then
        Irelia
    End If
    If cboDrop = "Janna" Then
        Janna
    End If
        If cboDrop = "Jarvan" Then
        Jarvan
    End If
    If cboDrop = "Jax" Then
        Jax
    End If
        If cboDrop = "Jayce" Then
        Jayce
    End If
    If cboDrop = "Jinx" Then
        Jinx
    End If
        If cboDrop = "Karma" Then
        Karma
    End If
    If cboDrop = "Karthus" Then
        Karthus
    End If
        If cboDrop = "Kassadin" Then
        Kassadin
    End If
    If cboDrop = "Katarina" Then
        Katarina
    End If
        If cboDrop = "Kayle" Then
        Kayle
    End If
        If cboDrop = "Kennen" Then
        Kennen
    End If
    If cboDrop = "KhaZix" Then
        Khazix
    End If
        If cboDrop = "KogMaw" Then
        KogMaw
    End If
    If cboDrop = "LeBlanc" Then
        LeBlanc
    End If
        If cboDrop = "LeeSin" Then
        LeeSin
    End If
    If cboDrop = "Leona" Then
        Leona
    End If
        If cboDrop = "Lissandra" Then
        Lissandra
    End If
    If cboDrop = "Lucian" Then
        Lucian
    End If
        If cboDrop = "Lulu" Then
        Lulu
    End If
    If cboDrop = "Lux" Then
        Lux
    End If
        If cboDrop = "Malphite" Then
        Malphite
    End If
    If cboDrop = "Malzahar" Then
        Malzahar
    End If
        If cboDrop = "Maokai" Then
        Maokai
    End If
    If cboDrop = "MasterYi" Then
        MasterYi
    End If
        If cboDrop = "MissFortune" Then
        MissFortune
    End If
    If cboDrop = "Mordekaiser" Then
        MordeKaiser
    End If
        If cboDrop = "Morgana" Then
        Morgana
    End If
        If cboDrop = "Nami" Then
        Nami
    End If
    If cboDrop = "Nasus" Then
        Nautilus
    End If
        If cboDrop = "Nidalee" Then
        Nidalee
    End If
    If cboDrop = "Nocturne" Then
        Nocturne
    End If
        If cboDrop = "Nunu" Then
        Nunu
    End If
    If cboDrop = "Olaf" Then
        Olaf
    End If
        If cboDrop = "Orianna" Then
        Orianna
    End If
    If cboDrop = "Pantheon" Then
        Pantheon
    End If
        If cboDrop = "Poppy" Then
        Poppy
    End If
    If cboDrop = "Quinn" Then
        Quinn
    End If
        If cboDrop = "Rammus" Then
        Rammus
    End If
    If cboDrop = "Renekton" Then
        Renekton
    End If
        If cboDrop = "Rengar" Then
        Rengar
    End If
    If cboDrop = "Riven" Then
        Riven
    End If
        If cboDrop = "Rumble" Then
        Rumble
    End If
    If cboDrop = "Ryze" Then
        Ryze
    End If
        If cboDrop = "Sejuani" Then
        Sejuani
    End If
        If cboDrop = "Shaco" Then
        Shaco
    End If
    If cboDrop = "Shen" Then
        Shen
    End If
        If cboDrop = "Shyvana" Then
        Shyvana
    End If
    If cboDrop = "Singed" Then
        Singed
    End If
        If cboDrop = "Sion" Then
        Sion
    End If
    If cboDrop = "Sivir" Then
        Sivir
    End If
        If cboDrop = "Skarner" Then
        Skarner
    End If
    If cboDrop = "Sona" Then
        Sona
    End If
        If cboDrop = "Soraka" Then
        Soraka
    End If
    If cboDrop = "Swain" Then
        Swain
    End If
        If cboDrop = "Syndra" Then
        Syndra
    End If
    If cboDrop = "Talon" Then
        Talon
    End If
        If cboDrop = "Taric" Then
        Taric
    End If
    If cboDrop = "Teemo" Then
        Teemo
    End If
        If cboDrop = "Thresh" Then
        Thresh
    End If
    If cboDrop = "Tristana" Then
        Tristana
    End If
        If cboDrop = "Trundle" Then
        Trundle
    End If
        If cboDrop = "Tryndamere" Then
        Tryndamere
    End If
    If cboDrop = "TwistedFate" Then
        TwistedFate
    End If
        If cboDrop = "Twitch" Then
        Twitch
    End If
    If cboDrop = "Udyr" Then
        Udyr
    End If
        If cboDrop = "Urgot" Then
        Urgot
    End If
    If cboDrop = "Varus" Then
        Varus
    End If
        If cboDrop = "Vayne" Then
        Vayne
    End If
    If cboDrop = "Veigar" Then
        Veigar
    End If
        If cboDrop = "Vi" Then
        Vi
    End If
    If cboDrop = "Viktor" Then
        Viktor
    End If
        If cboDrop = "Vladimir" Then
        Vladimir
    End If
    If cboDrop = "Volibear" Then
        Volibear
    End If
        If cboDrop = "Warwick" Then
        WarWick
    End If
    If cboDrop = "Wukong" Then
        Wukong
    End If
        If cboDrop = "Xerath" Then
        Xerath
    End If
    If cboDrop = "XinZhao" Then
        XinZhao
    End If
        If cboDrop = "Yasuo" Then
        Yasuo
    End If
        If cboDrop = "Yorick" Then
        Yorick
    End If
    If cboDrop = "Zac" Then
        Zac
    End If
        If cboDrop = "Zed" Then
        Zed
    End If
    If cboDrop = "Ziggs" Then
        Ziggs
    End If
        If cboDrop = "Zilean" Then
        Zilean
    End If
    If cboDrop = "Zyra" Then
        Zyra
    End If
    picIcon.Visible = True
    Image1.Visible = True
    Image2.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    lblQCD.Visible = True
    lblWCD.Visible = True
    lblECD.Visible = True
    lblRCD.Visible = True
    Image5.Visible = True
    Image6.Visible = True
    Image7.Visible = True
    Image8.Visible = True
End Sub

Private Sub lblExit_Click()
    End
End Sub

Private Sub lblNorm_Click()
         Dim lR As Long
         lR = SetTopMostWindow(frmHome.hwnd, False)
End Sub

Private Sub lblOnTop_Click()
         Dim lR As Long
         lR = SetTopMostWindow(frmHome.hwnd, True)
End Sub
