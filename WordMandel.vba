Function calcMandel(x, y, max)
    zx = x
    zy = y
    i = max
    zx2 = zx ^ 2
    zy2 = zy ^ 2
    While (i > 0 And zx2 + zy2 < 4)
        zy = 2 * zx * zy + y
        zx = zx2 - zy2 + x
        zx2 = zx ^ 2
        zy2 = zy ^ 2
        i = i - 1
    Wend
    chars$ = " .,-:;!=o¤*%Ox?$X#@SHNWM"
    nchars = Len(chars$)
    idx = 1 + Int((nchars - 1) * (max - i) / max)
    calcMandel = Mid$(chars$, idx, 1)
End Function


Sub Mandel()
    cx = -0.613
    cy = 0
    wx = 2.85
    nc = 124
    nl = 102
    
    wy = wx / 1.45
    sx = cx - (wx / 2)
    sy = cy - (wy / 2)
    dx = wx / nl
    dy = wy / nc
    rx = sx
    ActiveDocument.Content.Delete
    With ActiveDocument.Content
        .Font.Name = "Courier New"
        .Font.Size = 6
        For char_x = 0 To nl - 1
            ry = sy
            For char_y = 0 To nc - 1
                Rem .Collapse Direction:=wdCollapseEnd
                .InsertAfter Text:=calcMandel(rx, ry, 127)
                ry = ry + dy
            Next char_y
            .InsertAfter Text:=vbCrLf
            Application.ScreenRefresh
            rx = rx + dx
        Next char_x
    End With
End Sub
