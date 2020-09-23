VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paletted Bitmap Example"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   795
      Left            =   3900
      TabIndex        =   2
      Top             =   60
      Width           =   2475
      Begin VB.Label lblINFO 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Palette"
      Height          =   2655
      Left            =   3900
      TabIndex        =   0
      Top             =   900
      Width           =   2475
      Begin VB.PictureBox picPALETTE 
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   2235
         TabIndex        =   1
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.Image imgSIZE 
      Height          =   495
      Left            =   2160
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgDISPLAY 
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   60
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      ToolTipText     =   "Drag a bitmap into this window to load it and display its information"
      Top             =   60
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' This code is copyright 2002 Jeff 'Kuja' Katz

' You can use this code in any freeware product without charge
' if you are going to use any of the below in a commercial product,
' please email me so we can work things out. My email is
' jeff@katzonline.net

Private lng16_color_palette(15) As Long      ' variable to hold our 16  color palettes
Private lng256_color_palette(255) As Long    ' variable to hold our 256 color palettes
Private boolDepth As Boolean                 ' which type of bitmap is it again? False = 16 true =256


''' The load picture sub (takes vFN as the argument for which filename to load) '''
Private Sub LoadPic(vFN As String)
    
On Error Resume Next ' again, prevent unseen errors
        
    If Not LCase(Trim(Mid(vFN, Len(vFN) - 2))) = "bmp" Then ' make sure the file is a bitmap
        MsgBox "Only bitmaps, please. No " & LCase(Trim(Mid(vFN, Len(vFN) - 3))) & "'s allowed.", vbCritical '' Tell the user so
        Exit Sub ' drop out of the sub
    End If
    
    imgDISPLAY.Picture = LoadPicture(vFN) ' display the picture in the box
    imgSIZE.Picture = LoadPicture(vFN)    ' image size is invisible, we do this so we can extrapolate the sizes.
    Open vFN For Binary As #1 ' opens the file for binary input

        bmpfileheader = Input(14, 1) ' input the first 14 bits of the file header
                                     ' I could concievably check this to make sure
                                     ' the file is a bitmap.
        bmpinfohdr = Input(14, 1)    ' Some information about the bitmap
        bitcount = Input(2, 1)       ' The bit-depth (color type)
        bmpinfohdr = Input(24, 1)    ' More information before we get to the palette data

        If bitcount = (Chr(4) + Chr(0)) Then 'It was a 4bit (16 color) image
            
            For i = 0 To 15          ' Loop 16 times
                b = Asc(Input(1, 1)) ' Palette entry i blue value
                g = Asc(Input(1, 1)) ' Palette entry i green value
                r = Asc(Input(1, 1)) ' Palette entry i red value
                res = Input(1, 1)    ' unused information
                lng16_color_palette(i) = r + g * 256 + b * 65536 ' store the rgb long value for this palette entry
            Next i
        
        boolDepth = False            ' Signifies we just loaded a 16 color bitmap

        ElseIf bitcount = (Chr(8) + Chr(0)) Then 'It was a 8bit (256 color) image
            
            For i = 0 To 255         ' Loop 256 times
                b = Asc(Input(1, 1)) ' Palette entry i blue value
                g = Asc(Input(1, 1)) ' Palette entry i green value
                r = Asc(Input(1, 1)) ' Palette entry i red value
                res = Input(1, 1)    ' unused information
                lng256_color_palette(i) = r + g * 256 + b * 65536 ' store the rgb long value for this palette entry
            Next i
        boolDepth = True             ' Signifies we just loaded a 256 color bitmap
    
        Else ' Must have been monochrome or 24bit
    
            MsgBox "Bitmap was not 16 or 256 colors.", vbCritical ' Tell the user they messed up
            imgDISPLAY.Picture = Me.Picture ' Clear the image box

        End If

Close #1 ' Its imperitive we close the file handle.
End Sub

''' Generally called after loading an image, this sub outputs the palette to a box on the form
Sub DisplayPalette()
  
    If boolDepth Then z = 255 Else z = 15 ' Make sure to set up an appropriate environment for drawing
    
    For i = 0 To z                        ' Set up a for loop to draw each color
        
        ' if its 16 color, then make 16 vertical bars of the colors of the palette
        If Not boolDepth Then picPALETTE.Line (i * (picPALETTE.ScaleWidth / 16), 0)-((i + 1) * (picPALETTE.ScaleWidth / 16), picPALETTE.ScaleHeight), lng16_color_palette(i), BF
        ' if its 256 color, make 16 rows of 16 columns containing the colors
        If boolDepth Then picPALETTE.Line ((i Mod 16) * (picPALETTE.ScaleWidth / 16), (Int(i / 16)) * (picPALETTE.ScaleHeight / 16))-(((i Mod 16) * (picPALETTE.ScaleWidth / 16)) + (picPALETTE.ScaleWidth / 16), ((Int(i / 16)) * (picPALETTE.ScaleHeight / 16)) + (picPALETTE.ScaleHeight / 16)), lng256_color_palette(i), BF
    Next

End Sub

''' Handles what happens when the user drops a file onto the picture '''
Private Sub imgDISPLAY_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next ' Stops any pesky errors,
                     ' in case the user does something to bork up the code below

If (Data.GetFormat(vbCFFiles) = True) Then ' Is the thing they dropped a file (or a set of files)?
 Dim vFN                                   ' Sets up a variable to hold each file name
 For Each vFN In Data.Files                ' Cycle through the list of file names
  LoadPic (CStr(vFN))                            ' Preform the loadpic function on each file name
 Next vFN                                  ' Keep track of which vFN we are on
End If

    DisplayPalette                         ' Now show the palette
    
    If boolDepth Then                      ' This if statement updates the label
        lblINFO.Caption = Me.ScaleX(imgSIZE.Width, 1, 3) & "x" & Me.ScaleY(imgSIZE.Height, 1, 3) & ", 256 Colors"
    Else                                   ' Makes sure the output is in pixels
        lblINFO.Caption = Me.ScaleX(imgSIZE.Width, 1, 3) & "x" & Me.ScaleY(imgSIZE.Height, 1, 3) & ", 16 Colors"
    End If

End Sub

''' Shows the user the RGB values for the palette entry their mouse is over '''

Private Sub picPALETTE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DisplayPalette ' Just in case something erased our pretty picture
    
    lngColor = picPALETTE.Point(X, Y) ' Get the color
    
    r = lngColor Mod 256              ' Get the r value
    g = (lngColor \ 256) Mod 256      ' get the g value
    b = (lngColor \ 65536) Mod 256    ' get the b value
    
    Frame1.Caption = "Palette - (" & r & "," & b & "," & g & ")" ' update the caption
    
End Sub
