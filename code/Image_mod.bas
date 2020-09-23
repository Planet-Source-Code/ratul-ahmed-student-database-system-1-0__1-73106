Attribute VB_Name = "Image_mod"
'API Calls
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Const STRETCHMODE = vbPaletteModeNone

Dim JPEGp As clsJPEGparser  'Set Jpeg Class name
Dim picX As Long    'Picture X
Dim picY As Long    'Picture Y

'High Quality Photo Straching
Public Function Streach_pic(picfile As String, originalPicbox As PictureBox, ResizedPicbox As PictureBox, pWidth As Integer, pHeight As Integer, oHight As Integer, oWidth As Integer)

Dim nSizeX As Long
Dim nSizeY As Long
Dim pcX As Integer
Dim pcY As Integer

            '/Get Picture Properties/
    '-----------------------------
    Set JPEGp = New clsJPEGparser
    With JPEGp
        .ParseJpegFile (picfile)
        picX = .XsizePicture    ' Get X Size (Width)
        picY = .YsizePicture    ' Get Y Size (Height)
    End With
    '------------------------------
    
    ResizedPicbox.Cls
    
    originalPicbox.Height = picY    ' Resize Picture box
    originalPicbox.Width = picX     ' Resize Picture box
    
    pcY = picY - (picX * 0.1)
    
    nSizeX = picX
    nSizeY = picY

    If oHight = 0 Then oHight = picY
    If oWidth = 0 Then oWidth = picX
    
    If picfile = "" Then
        On Error Resume Next
        Call SetStretchBltMode(ResizedPicbox.hdc, STRETCHMODE)
        Call StretchBlt(ResizedPicbox.hdc, 0, 0, pWidth, pHeight, originalPicbox.hdc, 0, 0, oWidth, oHight, vbSrcCopy)
        ResizedPicbox.Refresh
    Else
        On Error Resume Next
        'Load the picture
        originalPicbox.Picture = LoadPicture(picfile)
        'Reduce it with SetStretchBltMode
        Call SetStretchBltMode(ResizedPicbox.hdc, STRETCHMODE)
        Call StretchBlt(ResizedPicbox.hdc, 0, 0, pWidth, pHeight, originalPicbox.hdc, 0, 0, oWidth, oHight, vbSrcCopy)
        ResizedPicbox.Refresh
    End If
    
End Function


Public Function Resize_image(imgFile As String, originalPicbox As PictureBox, tmpPicbox As PictureBox)

Dim npicX As Long
Dim npicY As Long
Dim picbX As Long
Dim picbY As Long

originalPicbox.Cls
tmpPicbox.Cls
            '/Get Picture Properties/
    '-----------------------------
    Set JPEGp = New clsJPEGparser
    With JPEGp
        .ParseJpegFile (imgFile)
        picX = .XsizePicture    ' Get X Size (Width)
        picY = .YsizePicture    ' Get Y Size (Height)
    End With
    '------------------------------
    
    frmimgedit.orx.Caption = picX
    frmimgedit.ory.Caption = picY
    
    npicX = picX '- (picX * 0.1)
    npicY = picY '- (picY * 0.1)
    
    If (picX > 184) Or (picY > 224) Then
    
        Do Until npicY <= originalPicbox.ScaleWidth
            npicX = npicX - (npicX * 0.1)
            npicY = npicY - (npicY * 0.1)
        Loop
    
        tmpPicbox.Height = picY
        tmpPicbox.Width = picX
    
        picbY = (originalPicbox.ScaleHeight - npicY) / 2
        picbX = (originalPicbox.ScaleWidth - npicX) / 2
    
        'MsgBox npicX & "x" & npicY
        'Load the picture
        tmpPicbox.Picture = LoadPicture(imgFile)
        'Reduce it with SetStretchBltMode
        Call SetStretchBltMode(originalPicbox.hdc, STRETCHMODE)
        Call StretchBlt(originalPicbox.hdc, picbX, picbY, npicX, npicY, tmpPicbox.hdc, 0, 0, picX, picY, vbSrcCopy)
        originalPicbox.Refresh
    Else
        MsgBox "The Picture Size Should be Larger then 250x250 pixel", vbExclamation, "Error"
    End If
    
    
    
End Function



