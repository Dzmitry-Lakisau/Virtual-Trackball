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
    'Поскольку собираюсь выводить графические объекты в окно, то нужно выставить флаг PFD_DRAW_TO_WINDOW
    'сообщаю, что буфер кадра поддерживает вывод через OpenGL - PFD_SUPPORT_OPENGL
    'и, так как используется два буфера кадров (Back и Front), устанавливаем флаг PFD_DOUBLEBUFFER
    'PFD_TYPE_RGBA - для использования цветовой схемы RGBA для пикселей
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    'pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24 'число битовых плоскостей в каждом буфере цвета
    pfd.cDepthBits = 16 'размер буфера глубины (ось z)
    pfd.iLayerType = PFD_MAIN_PLANE 'тип плоскости
    R = ChoosePixelFormat(Trackball.hDC, pfd) 'запрос на выбор самого подходящего формата для формы-трекбола. Эта функция
    ' возвращает подобранный для указанного контекста устройства индекс формата
    If R = 0 Then 'если подобрать не удалось
        MsgBox "ChoosePixelFormat failed"
        Exit Sub
    End If
    R = SetPixelFormat(Trackball.hDC, R, pfd) 'выставление пиксельного формата
    hGLRC = wglCreateContext(Trackball.hDC) 'создание графического контекста для формы-трекбола
    wglMakeCurrent Trackball.hDC, hGLRC 'созданный контекст выставляется текущим
    glClearColor 0, 0, 0, 1 'устанавливает черный непрозрачный цвет фона

    glClearDepth 1 'очистка буфера глубины
    glEnable GL_DEPTH_TEST 'включение теста глубины кадров

    glEnable glcColorMaterial
    glColorMaterial faceFront, GL_AMBIENT_AND_DIFFUSE 'установка свойств материала
    
    glEnable GL_LIGHTING '"включение" освещения
    glEnable glcLight0 '"включение" источника освещения
    pos(0) = 10: pos(1) = 10: pos(2) = 10: pos(3) = 1   'положение лампы
    glLightfv ltLight0, lpmPosition, pos(0) 'направление света

    AspectRatio = 1 'соотношение сторон. нужно для установки точки обзора

    obj = gluNewQuadric 'объявляю обьект как каркасный
End Sub
Public Sub TrackballMouseMove(X As Single, Y As Single)
    anglez = anglez + Speed / 8 * (X - startx) 'перевод перемещения указателя относительно начального положения при нажатии
    anglex = anglex + Speed / 8 * (Y - starty)
    If anglez < 0 Then anglez = 0 'если углы выходят за допустимые границы, то углы "замораживаются"=вращение трекбола прекращается
    If anglez > 360 Then anglez = 360
    If anglex < 0 Then anglex = 0
    If anglex > 360 Then anglex = 360
    If (anglez > 0) And (anglez < 360) And (anglex > 0) And (anglex < 360) Then Draw 'перерисовка шара при допустимых углах поворота
    PicField.Image1.Left = anglez / 360 * PicField.ScaleWidth - PicField.Image1.width / 2 'перевод углов в экранные координаты
    PicField.Image1.Top = anglex / 360 * PicField.ScaleHeight - PicField.Image1.height / 2
    Trackball.Label1.Caption = PicField.Image1.Left + PicField.Image1.width / 2
    Trackball.Label2.Caption = PicField.Image1.Top + PicField.Image1.height / 2
    startx = X
    starty = Y
End Sub
Public Sub DrawTable()
    glBegin (GL_POLYGON) 'крышка
        glNormal3f 0, 1, 0 'нормаль нужна для освещения
        glColor3f 1, 1, 1 'цвет - белый
        glVertex3f 0, -1, 1 'передняя точка
        glVertex3f 2, 0, 0 'правая
        glVertex3f 0, 1, -2 'задняя
        glVertex3f -2, 0, 0 'левая
    glEnd
    glBegin (GL_POLYGON) 'левая стенка
        glNormal3f -1, 0, 1
        glColor3f 1, 1, 1
        glVertex3f 0, -1, 1 'передняя точка
        glVertex3f -2, 0, 0 'левая
        glVertex3f -2, -0.25, 0 '
        glVertex3f 0, -1.25, 1
    glEnd
    glBegin (GL_POLYGON)
        glNormal3f 0.5, 0, 0.5
        glColor3f 1, 1, 1
        glVertex3f 0, -1, 1 'передняя точка
        glVertex3f 0, -1.25, 1
        glVertex3f 2, -0.25, 0
        glVertex3f 2, 0, 0 'правая
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
    Trackball.CommonDialog1.Filter = "Графика|*.jpg;*.bmp|"
    Trackball.CommonDialog1.ShowOpen

    border = (PicField.width - PicField.ScaleWidth * Screen.TwipsPerPixelX) / 2
    titlebar = PicField.height - PicField.ScaleHeight * Screen.TwipsPerPixelY

    PicField.Picture = LoadPicture(Trackball.CommonDialog1.FileName)
    height = PicField.ScaleY(PicField.Picture.height, vbHimetric, vbPixels)
    width = PicField.ScaleX(PicField.Picture.width, vbHimetric, vbPixels)
    koeff = height / width

    If width > (Screen.width - Trackball.width - Trackball.Left) / Screen.TwipsPerPixelX Or height > Screen.height / Screen.TwipsPerPixelY Then
        MsgBox "Вы пытаетесь загрузить изображение, которое больше свободной области экрана! Правильное отображение не гарантируется!", vbOKOnly, "Предупреждение"
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
