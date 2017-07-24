Attribute VB_Name = "Procedures"
Option Explicit
Public AspectRatio As Double
Public hGLRC&
Public obj&
Public startx&, starty&
Public anglez&, anglex&, border&, titlebar&, height&, width&, koeff&
Public Speed!
Public Sub InitOpenGL()
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim R&, pos!(0 To 3)
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    '��������� ��������� �������� ����������� ������� � ����, �� ����� ��������� ���� PFD_DRAW_TO_WINDOW
    '�������, ��� ����� ����� ������������ ����� ����� OpenGL - PFD_SUPPORT_OPENGL
    '�, ��� ��� ������������ ��� ������ ������ (Back � Front), ������������� ���� PFD_DOUBLEBUFFER
    'PFD_TYPE_RGBA - ��� ������������� �������� ����� RGBA ��� ��������
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    'pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24 '����� ������� ���������� � ������ ������ �����
    pfd.cDepthBits = 16 '������ ������ ������� (��� z)
    pfd.iLayerType = PFD_MAIN_PLANE '��� ���������
    R = ChoosePixelFormat(Trackball.hDC, pfd) '������ �� ����� ������ ����������� ������� ��� �����-��������. ��� �������
    ' ���������� ����������� ��� ���������� ��������� ���������� ������ �������
    If R = 0 Then '���� ��������� �� �������
        MsgBox "ChoosePixelFormat failed"
        Exit Sub
    End If
    R = SetPixelFormat(Trackball.hDC, R, pfd) '����������� ����������� �������
    hGLRC = wglCreateContext(Trackball.hDC) '�������� ������������ ��������� ��� �����-��������
    wglMakeCurrent Trackball.hDC, hGLRC '��������� �������� ������������ �������
    glClearColor 0, 0, 0, 1 '������������� ������ ������������ ���� ����

    glClearDepth 1 '������� ������ �������
    glEnable GL_DEPTH_TEST '��������� ����� ������� ������

    glEnable glcColorMaterial
    glColorMaterial faceFront, GL_AMBIENT_AND_DIFFUSE '��������� ������� ���������
    
    glEnable GL_LIGHTING '"���������" ���������
    glEnable glcLight0 '"���������" ��������� ���������
    pos(0) = 10: pos(1) = 10: pos(2) = 10: pos(3) = 1   '��������� �����
    glLightfv ltLight0, lpmPosition, pos(0) '����������� �����

    AspectRatio = 1 '����������� ������. ����� ��� ��������� ����� ������

    obj = gluNewQuadric '�������� ������ ��� ���������
End Sub
Public Sub TrackballMouseMove(X As Single, Y As Single)
    anglez = anglez + Speed / 8 * (X - startx) '������� ����������� ��������� ������������ ���������� ��������� ��� �������
    anglex = anglex + Speed / 8 * (Y - starty)
    If anglez < 0 Then anglez = 0 '���� ���� ������� �� ���������� �������, �� ���� "��������������"=�������� �������� ������������
    If anglez > 360 Then anglez = 360
    If anglex < 0 Then anglex = 0
    If anglex > 360 Then anglex = 360
    If (anglez > 0) And (anglez < 360) And (anglex > 0) And (anglex < 360) Then Draw '����������� ���� ��� ���������� ����� ��������
    PicField.Image1.Left = anglez / 360 * PicField.ScaleWidth - PicField.Image1.width / 2 '������� ����� � �������� ����������
    PicField.Image1.Top = anglex / 360 * PicField.ScaleHeight - PicField.Image1.height / 2
    Trackball.Label1.Caption = PicField.Image1.Left + PicField.Image1.width / 2
    Trackball.Label2.Caption = PicField.Image1.Top + PicField.Image1.height / 2
    startx = X
    starty = Y
End Sub
Public Sub DrawTable()
    glBegin (GL_POLYGON) '������
        glNormal3f 0, 1, 0 '������� ����� ��� ���������
        glColor3f 1, 1, 1 '���� - �����
        glVertex3f 0, -1, 1 '�������� �����
        glVertex3f 2, 0, 0 '������
        glVertex3f 0, 1, -2 '������
        glVertex3f -2, 0, 0 '�����
    glEnd
    glBegin (GL_POLYGON) '����� ������
        glNormal3f -1, 0, 1
        glColor3f 1, 1, 1
        glVertex3f 0, -1, 1 '�������� �����
        glVertex3f -2, 0, 0 '�����
        glVertex3f -2, -0.25, 0 '
        glVertex3f 0, -1.25, 1
    glEnd
    glBegin (GL_POLYGON)
        glNormal3f 0.5, 0, 0.5
        glColor3f 1, 1, 1
        glVertex3f 0, -1, 1 '�������� �����
        glVertex3f 0, -1.25, 1
        glVertex3f 2, -0.25, 0
        glVertex3f 2, 0, 0 '������
    glEnd
End Sub
Public Sub PicFieldSetCursor()
    PicField.MousePointer = 99
    PicField.MouseIcon = LoadPicture(App.Path & "\hand.ico")
End Sub
Public Sub PicFieldEscape(Keycode)
    Select Case (Keycode)
    Case 27: Unload PicField
    End Select
End Sub
Public Sub TrackballEscape(Keycode)
    Select Case (Keycode)
    Case 27: Unload Trackball: Unload PicField
    End Select
End Sub
Public Sub Sizing(ByVal W&, ByVal H&)
    If H = 0 Then H = 1
    AspectRatio = W / H
    glViewport 0, 0, W, H
    glMatrixMode mmProjection
    glLoadIdentity
    gluPerspective 45, AspectRatio, 0.5, 200
    gluLookAt 0, 0, 6, 0, 0, 0, 0, 1, 0
    glMatrixMode mmModelView
    glLoadIdentity
End Sub
Public Sub TrackballSetCursor()
    Trackball.MousePointer = 99
    Trackball.MouseIcon = LoadPicture(App.Path & "\hand.ico")
    PicField.Image1.Picture = LoadPicture(App.Path & "\cross.cur")
    Trackball.SpeedEight.Checked = True
End Sub
Public Sub UnloadGL()
    If hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC
    End If
End Sub
Public Sub Draw()
    wglMakeCurrent Trackball.hDC, hGLRC
    glClear clrColorBufferBit Or clrDepthBufferBit
    glMatrixMode (mmModelView)
    glLoadIdentity
    glPushMatrix
        glRotatef anglex, 1, 0, 0
        glRotatef anglez, 0, 1, 0
        glColor3f 1, 0, 0
        gluSphere obj, 1, 20, 20
    glPopMatrix
    DrawTable
    glFlush
    SwapBuffers Trackball.hDC
End Sub
Public Sub TrackballMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    startx = X
    starty = Y
    If Trackball.SpeedOne.Checked Then Speed = 1
    If Trackball.SpeedTwo.Checked Then Speed = 2
    If Trackball.SpeedFour.Checked Then Speed = 4
    If Trackball.SpeedEight.Checked Then Speed = 8
    If Trackball.Speed16.Checked Then Speed = 16
End If
If Button = 2 Then
    On Error GoTo ErrorHandler
    Trackball.CommonDialog1.CancelError = True
    Trackball.CommonDialog1.Filter = "�������|*.jpg;*.bmp|"
    Trackball.CommonDialog1.ShowOpen

    border = (PicField.width - PicField.ScaleWidth * Screen.TwipsPerPixelX) / 2
    titlebar = PicField.height - PicField.ScaleHeight * Screen.TwipsPerPixelY

    PicField.Picture = LoadPicture(Trackball.CommonDialog1.FileName)
    height = PicField.ScaleY(PicField.Picture.height, vbHimetric, vbPixels)
    width = PicField.ScaleX(PicField.Picture.width, vbHimetric, vbPixels)
    koeff = height / width

    If width > (Screen.width - Trackball.width - Trackball.Left) / Screen.TwipsPerPixelX Or height > Screen.height / Screen.TwipsPerPixelY Then
        MsgBox "�� ��������� ��������� �����������, ������� ������ ��������� ������� ������! ���������� ����������� �� �������������!", vbOKOnly, "��������������"
        If height > width Then
        PicField.height = Screen.height
        PicField.width = PicField.height / koeff

        Else:
            PicField.width = Screen.width - Trackball.width - Trackball.Left
            PicField.height = PicField.width * koeff
        End If
        Else:   PicField.width = width * Screen.TwipsPerPixelX + 2 * border
                    PicField.height = height * Screen.TwipsPerPixelY + titlebar
    
    End If
    PicField.Top = 0
    PicField.Left = Trackball.Left + Trackball.width
    PicField.Image1.ZOrder vbBringToFront

    anglez = 0: anglex = 0
    PicField.Image1.Left = anglez - PicField.Image1.width / 2
    PicField.Image1.Top = anglex - PicField.Image1.height / 2
    PicField.Show

ErrorHandler:
    If Err.Number = 32755 Then
        Exit Sub
    End If
End If
End Sub
Public Sub UnCheckMenu()
    Trackball.SpeedOne.Checked = False
    Trackball.SpeedTwo.Checked = False
    Trackball.SpeedFour.Checked = False
    Trackball.SpeedEight.Checked = False
    Trackball.Speed16.Checked = False
End Sub
