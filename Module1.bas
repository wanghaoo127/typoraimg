Attribute VB_Name = "Module1"

Sub MySql_连接() '引用需勾选Microsoft Activex Data Objects 6. 1 Library和Microsoft Activex Data Objects Recordset 2.8 Library
    Dim PI As Double
    PI = 3.14159265258979
    
    z = 0
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    '配置连接串
    conn.ConnectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;DB=hebei_wells;UID=root;PWD=1023;OPTION=3;"
    conn.Open
    rs.Open "SELECT  * FROM  `well_structure` WHERE  well_structure.`区域` ='康定' AND   well_structure.`井名` ='定向井'  ORDER BY well_structure.`区域` ASC, well_structure.`井名` ASC ", conn ' and well_structure.`井型`='直井'
    arr = rs.GetRows 'well_structure.`井名` ='绿源2井' AND   well_structure.`井型` ='定向井' AND
    '关闭连接
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing
    ReDim brr(1 To UBound(arr, 2) + 1, 1 To UBound(arr, 1) + 1)
    For i = 1 To UBound(arr, 2) + 1
        For j = 1 To UBound(arr, 1) + 1
            brr(i, j) = arr(j - 1, i - 1)
            If IsNull(brr(i, j)) = True Then brr(i, j) = 0
        Next j
    Next i
    
    
    'Dim p00(2) As Double, p0(2) As Double, p1(2) As Double, p2(2) As Double, p3(2) As Double, p4(2) As Double, p5(2) As Double, p6(2) As Double '井身结构中的点
    'Dim p7(2) As Double, p8(2) As Double, p9(2) As Double, p10(2) As Double, p11(2) As Double, p12(2) As Double
    
    Dim p11(2) As Double, p12(2) As Double, p13(2) As Double, p14(2) As Double, p15(2) As Double '井身结构中的点
    Dim p21(2) As Double, p22(2) As Double, p23(2) As Double, p24(2) As Double, p25(2) As Double
    Dim p31(2) As Double, p32(2) As Double, p33(2) As Double, p34(2) As Double, p35(2) As Double
    Dim p41(2) As Double, p42(2) As Double, p43(2) As Double, p44(2) As Double, p45(2) As Double
    
    Dim p11l(2) As Double, p12l(2) As Double, p13l(2) As Double, p14l(2) As Double, p15l(2) As Double '井身结构中的点
    Dim p21l(2) As Double, p22l(2) As Double, p23l(2) As Double, p24l(2) As Double, p25l(2) As Double
    Dim p31l(2) As Double, p32l(2) As Double, p33l(2) As Double, p34l(2) As Double, p35l(2) As Double
    Dim p41l(2) As Double, p42l(2) As Double, p43l(2) As Double, p44l(2) As Double, p45l(2) As Double
    
    Dim p11p(2) As Double, p13p(2) As Double, p21p(2) As Double, p23p(2) As Double, p31p(2)  As Double, p33p(2) As Double, p41p(2)  As Double, p43p(2) As Double  '标注的端点
    Dim bzwz(2) As Double '标注的位置
    
    Dim pointsarray(0 To 5) As Double '水泥固井标注位置
    Dim dimpoint(2) As Double
    
    Dim yinxian1(2) As Double, yinxian2(2) As Double
    Dim L_chicun(2) As Double, R_chicun(2) As Double
    
    Dim L_dayin(2) As Double, R_dayin(2) As Double '打印图框的对角点
    
    
    Dim m0(2) As Double, m1(2) As Double, m2(2) As Double, m3(2) As Double, m4(2) As Double '轴线的端点
    
    Dim b1(2) As Double, b2(2) As Double, b3(2) As Double, b4(2) As Double  '表格中的点
    Dim b5(2) As Double, b6(2) As Double, b7(2) As Double, b8(2) As Double
    Dim b9(2) As Double, b10(2) As Double, b11(2) As Double, b12(2) As Double
    Dim b13(2) As Double, b14(2) As Double, b15(2) As Double, b16(2) As Double
    
    Dim b1move(2) As Double, b2move(2) As Double, b3move(2) As Double, b4move(2) As Double  'b向下移动的点
    Dim b5move(2) As Double, b6move(2) As Double, b7move(2) As Double, b8move(2) As Double
    Dim b9move(2) As Double, b10move(2) As Double, b11move(2) As Double, b12move(2) As Double
    
    
    Dim p_tx1(2) As Double, p_tx2(2) As Double, p_tx3(2) As Double, p_tx4(2) As Double, p_tx5(2) As Double '表头的文字位置
    Dim p_tx6(2) As Double, p_tx7(2) As Double, p_tx8(2) As Double, p_tx9(2) As Double, p_tx10(2) As Double '10为井名的位置
    
    'Dim p(0 To 6, 0 To 2) As Double '点
    Dim l1, l2, l3, l4, l5, l6 As AcadLine
    Dim pl1, pl2 As AcadPolyline
    
    Dim box(0) As AcadLWPolyline
    
    Dim lk(1 To 7) As Integer '列宽
    Dim bk As Integer '表宽，lk的sum
    Dim V_bt As Integer '表头的高
    
    Dim yanxing(1 To 4) As String
    
    V_bt = 100 '表头的高
    lk(1) = 100  '列宽 界
    lk(2) = lk(1) '系
    lk(3) = lk(1) '组
    lk(4) = 150  '垂深
    lk(5) = lk(4) '垂厚
    lk(6) = 700  '成井结构
    lk(7) = 400  '岩性描述
    
    
    For i = LBound(lk, 1) To UBound(lk, 1)
        bk = bk + lk(i) '计算表宽
    Next i
    
    '--------------------------------------------------------设置字体样式
    'Dim NewTextStyleObj   As AcadTextStyle
    'Dim TextStyleName   As String
    
    'TextStyleName = "biaotou"
    'Set NewTextStyleObj = ThisDrawing.TextStyles.Add(TextStyleName)
    'NewTextStyleObj.Height = 20
    'NewTextStyleObj.TextGenerationFlag = 0
    
    'Dim biaotou As AcadTextStyle
    ' Set biaotou = ThisDrawing.TextStyles.Add("表头样式")
    'biaotou.Height = 20
    ' biaotou.TextGenerationFlag = 0
    
    '--------------------------------------------------------删除绘图区域内的元素
    Dim zoomLowLeft(0 To 2) As Double
    Dim zoomUpRight(0 To 2) As Double
    zoomLowLeft(0) = -100000
    zoomLowLeft(1) = -300000
    zoomUpRight(0) = 200000
    zoomUpRight(1) = 2000
    
    Dim sel1 As AcadSelectionSet '定义选择集对象
    'Set sel1 = ThisDrawing.SelectionSets.Add("s" & Timer) '新建一个选择集
    'Call sel1.Select(5, zoomLowLeft, zoomUpRight) '全部选中 acSelectionSetAll
    'sel1.Highlight (True) '显示选择的对象
    'sel1.Erase
    '--------------------------------------------------------设置每幅图的原点
    Dim yuandian(1 To 150, 0 To 2) As Double
    
    For i = 1 To 15
        For j = 1 To 10
            yuandian((i - 1) * 10 + j, 0) = j * 3000
            yuandian((i - 1) * 10 + j, 1) = (i - 1) * 3000 * (-1)
        Next j
    Next i
    
    '--------------------------------------------------------设置每幅图的原点
    Set lay_biaoge = ThisDrawing.Layers.Add("表格")
    Set lay_jingshen = ThisDrawing.Layers.Add("井身结构")
    Set lay_biaogewaikuang = ThisDrawing.Layers.Add("表格外框")
    Set lay_biaowenzi = ThisDrawing.Layers.Add("表格文字")
    Set lay_biaozhu = ThisDrawing.Layers.Add("标注")
    Set lay_tianchong = ThisDrawing.Layers.Add("填充")
    Set lay_dayin = ThisDrawing.Layers.Add("图框_打印")
    
    
    lay_biaoge.Lineweight = 100
    lay_jingshen.Lineweight = 80
    lay_biaogewaikuang.Lineweight = 200
    
    
    '--------------------------------------------------------循环开始
    
    For i = 1 To UBound(brr, 1)

        
        
        '--------------------------------------------------------将库里的数赋值给变量
        quyu = brr(i, 1)
        wellname = brr(i, 2)
        jingbie = brr(i, 3)
        jingxing = brr(i, 4)
        jingshen = brr(i, 5)
        jingyan1 = brr(i, 6)
        jingyan2 = brr(i, 7)
        jingyan3 = brr(i, 8)
        jingyan4 = brr(i, 9)
        jingyan1_d = brr(i, 10)
        jingyan2_d = brr(i, 11)
        jingyan3_d = brr(i, 12)
        jingyan4_d = brr(i, 13)
        taoguan1 = brr(i, 14)
        taoguan2 = brr(i, 15)
        taoguan3 = brr(i, 16)
        taoguan4 = brr(i, 17)
        taoguan1_start = brr(i, 18)
        taoguan1_end = brr(i, 19)
        taoguan2_start = brr(i, 20)
        taoguan2_end = brr(i, 21)
        taoguan3_start = brr(i, 22)
        taoguan3_end = brr(i, 23)
        taoguan4_start = brr(i, 24)
        taoguan4_end = brr(i, 25)
        diceng1_d = brr(i, 26)
        diceng2_d = brr(i, 27)
        diceng3_d = brr(i, 28)
        diceng4_d = brr(i, 29)
        diceng1_V = brr(i, 30)
        diceng2_V = brr(i, 31)
        diceng3_V = brr(i, 32)
        diceng4_V = brr(i, 33)
        
        chuihou1 = diceng1_V               '计算垂厚
        chuihou2 = diceng2_V - diceng1_V
        chuihou3 = diceng3_V - diceng2_V
        
        If diceng4_d <> 0 Then
            chuihou4 = diceng4_V - diceng3_V
        End If
        
        If jingyan4_d - jingyan3_d >= 100 Then
            bg = jingyan4_d
        ElseIf jingyan4_d > 0 And jingyan4_d - jingyan3_d < 100 Then
            bg = jingyan4_d + 100
        Else
            bg = jingyan3_d
        End If
        
        If jingxing = "定向井" Then
            lk(6) = 900
            bk = 1900
        Else
            lk(6) = 700
            bk = 1700
        End If
        
        For k = 1 To 4
            yanxing(k) = brr(i, 33 + k)     '岩性描述
        Next k
        
        
        '--------------------------------------------------------表格框架的点
        
        b1(0) = yuandian(i, 0)
        b1(1) = yuandian(i, 1)
        
        b2(0) = b1(0) + bk
        b2(1) = b1(1)
        
        b3(0) = b2(0)
        b3(1) = b2(1) - bg ' (chuihou1 + chuihou2 + chuihou3 + chuihou4)
        
        b4(0) = b1(0)
        b4(1) = b3(1)
        
        b5(0) = b1(0)
        b5(1) = b1(1) + V_bt
        
        b6(0) = b2(0)
        b6(1) = b2(1) + V_bt
        
        b7(0) = b5(0) + lk(1) + lk(2) + lk(3)
        b7(1) = b5(1)
        
        b8(0) = b5(0) + lk(1) + lk(2) + lk(3)
        b8(1) = b4(1)
        
        b9(0) = b1(0) + lk(1)
        b9(1) = b1(1) - diceng1_d
        
        b10(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5)
        b10(1) = b9(1)
        
        b11(0) = b1(0)
        b11(1) = b1(1) - diceng2_d
        
        b12(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5)
        b12(1) = b11(1)
        
        b13(0) = b2(0) - lk(7)
        b13(1) = b2(1) - diceng1_d
        
        b14(0) = b2(0)
        b14(1) = b13(1)
        
    '------------------------------打印框
    L_dayin(0) = b1(0) - 50
    L_dayin(1) = b1(1) + 300
    R_dayin(0) = b3(0) + 50
    R_dayin(1) = b3(1) - 50
    ThisDrawing.ActiveLayer = lay_dayin
            Set box(0) = drawbox(L_dayin, R_dayin)
            'Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            'hobj.AppendInnerLoop box
            'hobj.color = 253
            'box(0).Delete
        
        
        ThisDrawing.ActiveLayer = lay_biaoge
        
        Set bx1 = ThisDrawing.ModelSpace.AddLine(b1(), b2())
        'Set bx2 = ThisDrawing.ModelSpace.AddLine(b2(), b3())
        'Set bx3 = ThisDrawing.ModelSpace.AddLine(b3(), b4())
        Set bx4 = ThisDrawing.ModelSpace.AddLine(b1(), b4())
        'Set bx5 = ThisDrawing.ModelSpace.AddLine(b1(), b5())
        'Set bx6 = ThisDrawing.ModelSpace.AddLine(b2(), b6())
        Set bx7 = ThisDrawing.ModelSpace.AddLine(b7(), b8())
        Set bx8 = ThisDrawing.ModelSpace.AddLine(b10(), b9())
        Set bx9 = ThisDrawing.ModelSpace.AddLine(b11(), b12())
        Set bx10 = ThisDrawing.ModelSpace.AddLine(b14(), b13())
        
        
        offsetpline = bx1.Offset(V_bt)
        offsetpline = bx4.Offset(Abs(lk(1)))
        offsetpline = bx4.Offset(lk(1) * 2)
        offsetpline = bx7.Offset(lk(4))
        offsetpline = bx7.Offset(lk(4) + lk(5))
        offsetpline = bx7.Offset(lk(4) + lk(5) + lk(6))
        offsetpline = bx7.Offset(lk(4) + lk(5) + lk(6) + lk(7))
        offsetpline = bx8.Offset(diceng3_d - diceng1_d)
        offsetpline = bx10.Offset(diceng2_d - diceng1_d)
        offsetpline = bx10.Offset(diceng3_d - diceng1_d)
        
        ThisDrawing.ActiveLayer = lay_biaogewaikuang
        Set wk1 = ThisDrawing.ModelSpace.AddLine(b4(), b5())
        Set wk2 = ThisDrawing.ModelSpace.AddLine(b6(), b5())
        Set wk3 = ThisDrawing.ModelSpace.AddLine(b3(), b6())
        Set wk4 = ThisDrawing.ModelSpace.AddLine(b3(), b4())
        
        
        '--------------------------------------------------------表格中的文字
        
        ThisDrawing.ActiveLayer = lay_biaowenzi
        H = 70 '文字的高度
        
        p_tx1(0) = b1(0)  'tx=text
        p_tx1(1) = b1(1) + H
        
        p_tx2(0) = b1(0) + lk(1) + lk(2) + lk(3)
        p_tx2(1) = b1(1) + H
        
        p_tx3(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4)
        p_tx3(1) = b1(1) + H
        
        p_tx4(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5)
        p_tx4(1) = b1(1) + H
        
        p_tx5(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6)
        p_tx5(1) = b1(1) + H
        
        p_tx6(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6) + lk(7)
        p_tx6(1) = b1(1) + H
        
        p_tx10(0) = p_tx1(0)
        p_tx10(1) = p_tx1(1) + 200
        
        weizigaodu = 30
        'ThisDrawing.ActiveTextStyle = 表头样式
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx10, lk(1) + lk(2) + lk(3) + lk(4) + lk(5), wellname) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu * 2
        MTextObj.AttachmentPoint = 2
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx1, lk(1) + lk(2) + lk(3), "地  层") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.AttachmentPoint = 2
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx2, lk(4), "垂深/m") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.AttachmentPoint = 2
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx3, lk(5), "垂厚/m")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.AttachmentPoint = 2
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx4, lk(6), "井身结构")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.AttachmentPoint = 2
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx5, lk(7), "岩性描述")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.AttachmentPoint = 2
        'Set MTextObj = ThisDrawing.ModelSpace.AddMText(p_tx6, lk(8), "成井结构") '(位置，文字宽度，字符串)
        'MTextObj.Height = weizigaodu
        'MTextObj.AttachmentPoint = 2
        
        weizigaodu = 25
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "新" & vbCrLf & "生" & vbCrLf & "界")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0)
        b1move(1) = b1(1) - 100
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "第" & vbCrLf & "四" & vbCrLf & "系")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 50
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1)
        b1move(1) = b1(1) - 100
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "新" & vbCrLf & "近" & vbCrLf & "系")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 50
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1)
        b1move(1) = b1(1) - 50 - diceng1_d
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "平" & vbCrLf & "原" & vbCrLf & "组")  '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 50
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2)
        b1move(1) = b1(1) - 100
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "明" & vbCrLf & "化" & vbCrLf & "镇" & vbCrLf & "组") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 50
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2)
        b1move(1) = b1(1) - 50 - diceng1_d
        MTextObj.Move b1, b1move
        
        If diceng3_d - diceng2_d > 400 Then
            hangjianjv = 100
        ElseIf diceng3_d - diceng2_d > 300 Then
            hangjianjv = 50
        Else
            hangjianjv = 40
        End If
        
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "中" & vbCrLf & "元" & vbCrLf & "古" & vbCrLf & "界") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        'MTextObj.defined_height = diceng2_d
        MTextObj.LineSpacingDistance = hangjianjv '(bg - jingyan2_d - 100 * 2 - 4 * weizigaodu) / 6
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0)
        b1move(1) = b1(1) - diceng2_d - hangjianjv
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "蓟" & vbCrLf & "县" & vbCrLf & "系") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = hangjianjv
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1)
        b1move(1) = b1(1) - diceng2_d - hangjianjv
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "雾" & vbCrLf & "迷" & vbCrLf & "山" & vbCrLf & "组") '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = hangjianjv
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2)
        b1move(1) = b1(1) - diceng2_d - hangjianjv
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(4), diceng1_V) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3)
        b1move(1) = b1(1) - diceng1_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(4), diceng2_V) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3)
        b1move(1) = b1(1) - diceng2_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(4), diceng3_V) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3)
        b1move(1) = b1(1) - diceng3_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(5), chuihou1) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4)
        b1move(1) = b1(1) - diceng1_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(5), chuihou2) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4)
        b1move(1) = b1(1) - diceng2_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(5), chuihou3) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 5
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4)
        b1move(1) = b1(1) - diceng3_d + 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(7) - 20, yanxing(1)) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        'MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 4  '3-右对齐，4-左对齐，5-居中对齐
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6) + 20
        b1move(1) = b1(1) - 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(7) - 20, yanxing(2)) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        'MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 4  '3-右对齐，4-左对齐，5-居中对齐
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6) + 20
        b1move(1) = b1(1) - diceng1_d - 50
        MTextObj.Move b1, b1move
        
        Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(7) - 20, yanxing(3)) '(位置，文字宽度，字符串)
        MTextObj.Height = weizigaodu
        'MTextObj.LineSpacingDistance = 150
        MTextObj.AttachmentPoint = 4  '3-右对齐，4-左对齐，5-居中对齐
        b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6) + 20
        b1move(1) = b1(1) - diceng2_d - 50
        MTextObj.Move b1, b1move
        
        If diceng4_d <> 0 Then
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(1), "地层4")  '(位置，文字宽度，字符串)
            MTextObj.Height = weizigaodu
            'MTextObj.LineSpacingDistance = 150
            MTextObj.AttachmentPoint = 5
            b1move(0) = b1(0) + lk(1) + lk(2)
            b1move(1) = b1(1) - bg + 50
            MTextObj.Move b1, b1move
            
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(4), diceng4_V) '(位置，文字宽度，字符串)
            MTextObj.Height = weizigaodu
            MTextObj.LineSpacingDistance = 150
            MTextObj.AttachmentPoint = 5
            b1move(0) = b1(0) + lk(1) + lk(2) + lk(3)
            b1move(1) = b1(1) - bg + 50
            MTextObj.Move b1, b1move
            
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(5), chuihou4) '(位置，文字宽度，字符串)
            MTextObj.Height = weizigaodu
            MTextObj.LineSpacingDistance = 150
            MTextObj.AttachmentPoint = 5
            b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4)
            b1move(1) = b1(1) - bg + 50
            MTextObj.Move b1, b1move
            
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(b1, lk(7) - 20, yanxing(4)) '(位置，文字宽度，字符串)
            MTextObj.Height = weizigaodu
            'MTextObj.LineSpacingDistance = 150
            MTextObj.AttachmentPoint = 4  '3-右对齐，4-左对齐，5-居中对齐
            b1move(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + lk(6) + 20
            b1move(1) = b1(1) - bg + 50
            MTextObj.Move b1, b1move
            
        End If
        
        
        
        Debug.Print MTextObj.LineSpacingDistance
        
        
        
        If jingxing = "直井" Then
            '--------------------------------------------------------井身结构中的点++++直井
            
            ThisDrawing.ActiveLayer = lay_jingshen
            
            m0(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + 300
            m0(1) = b1(1)
            
            m1(0) = m0(0)
            m1(1) = m0(1) - diceng1_d + 10
            
            '------------------------------------一开
            p11(0) = m0(0) + jingyan1 / 2
            p11(1) = m0(1)
            
            p12(0) = p11(0)
            p12(1) = p11(1) - jingyan1_d
            
            p13(0) = p11(0) - (jingyan1 - taoguan1) / 2
            p13(1) = p11(1)
            
            p14(0) = p13(0)
            p14(1) = p13(1) - jingyan1_d
            
            '------------------------------------二开
            p21(0) = m0(0) + jingyan2 / 2
            p21(1) = p12(1)
            
            p22(0) = p21(0)
            p22(1) = p21(1) - jingyan2_d + jingyan1_d
            
            p23(0) = m0(0) + taoguan2 / 2
            p23(1) = m0(1) - taoguan2_start
            
            p24(0) = p23(0)
            p24(1) = p23(1) - taoguan2_end + taoguan2_start
            
            p25(0) = p13(0)
            p25(1) = p23(1)
            
            '------------------------------------三开
            p31(0) = m0(0) + jingyan3 / 2
            p31(1) = p22(1)
            
            p32(0) = p31(0)
            p32(1) = p31(1) - jingyan3_d + jingyan2_d
            
            p33(0) = m0(0) + taoguan3 / 2
            p33(1) = m0(1) - taoguan3_start
            
            p34(0) = p33(0)
            p34(1) = p33(1) - taoguan3_end + taoguan3_start
            
            p35(0) = p24(0)
            p35(1) = p33(1)
            
            '--------------------------------------------------------井身结构中的填充
            
            ThisDrawing.ActiveLayer = lay_tianchong
            
            Set box(0) = drawbox(p21, p25)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            hobj.Mirror m0, m1
            
            Set box(0) = drawbox(p13, p12)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            hobj.Mirror m0, m1
            
            Set box(0) = drawbox(p23, p22)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            hobj.Mirror m0, m1
            
            ThisDrawing.ActiveLayer = lay_jingshen
            
            Set l11 = ThisDrawing.ModelSpace.AddLine(p11(), p12())
            Set l12 = ThisDrawing.ModelSpace.AddLine(p13(), p14())
            Set l13 = ThisDrawing.ModelSpace.AddLine(p12(), p21())
            Set l14 = ThisDrawing.ModelSpace.AddLine(p11(), p13())
            
            Set l21 = ThisDrawing.ModelSpace.AddLine(p21(), p22())
            Set l22 = ThisDrawing.ModelSpace.AddLine(p23(), p24())
            Set l23 = ThisDrawing.ModelSpace.AddLine(p22(), p31())
            Set l24 = ThisDrawing.ModelSpace.AddLine(p25(), p23())
            
            Set l31 = ThisDrawing.ModelSpace.AddLine(p31(), p32())
            Set l32 = ThisDrawing.ModelSpace.AddLine(p33(), p34())
            
            Set l34 = ThisDrawing.ModelSpace.AddLine(p35(), p33())
            
            l11.Mirror m0, m1
            l12.Mirror m0, m1
            l13.Mirror m0, m1
            l14.Mirror m0, m1
            l21.Mirror m0, m1
            l22.Mirror m0, m1
            l23.Mirror m0, m1
            l24.Mirror m0, m1
            l31.Mirror m0, m1
            l32.Mirror m0, m1
            'l33.Mirror m0, m1
            l34.Mirror m0, m1
            
            '------------------------------------四开     ++++++++++++
            If jingyan4 <> 0 Or taoguan4 <> 0 Then
                
                p41(0) = p32(0) - (jingyan3 - jingyan4) / 2 '0
                p41(1) = p32(1)
                
                p42(0) = p41(0)
                p42(1) = p41(1) - jingyan4_d + jingyan3_d
                
                p43(0) = m0(0) + taoguan4 / 2
                p43(1) = m0(1) - taoguan4_start
                
                p44(0) = p43(0)
                p44(1) = p43(1) - taoguan4_end + taoguan4_start
                
                p45(0) = p34(0)
                p45(1) = p43(1)
                
                Set l33 = ThisDrawing.ModelSpace.AddLine(p32(), p41())
                Set l41 = ThisDrawing.ModelSpace.AddLine(p41(), p42())
                Set l42 = ThisDrawing.ModelSpace.AddLine(p43(), p44()) '
                Set l43 = ThisDrawing.ModelSpace.AddLine(p42(), p41())
                Set l44 = ThisDrawing.ModelSpace.AddLine(p45(), p43())
                
                l33.Mirror m0, m1
                l41.Mirror m0, m1
                l42.Mirror m0, m1
                l43.Mirror m0, m1
                l44.Mirror m0, m1
                
            End If
            
            'If jingxing = "定向井" Then
            '    l4.Rotate p3, 0.2     '旋转直线
            'End If
            
            
            
            '--------------------------------------------------------井身结构中的标注
            
            ThisDrawing.ActiveLayer = lay_biaozhu
            di = 150
            ding = 100
            
            p13p(0) = 2 * m0(0) - p13(0)
            p13p(1) = p13(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - ding
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p13, p13p, bzwz)
            biaozhu.TextOverride = taoguan1 & "mm"
            
            p11p(0) = 2 * m0(0) - p11(0)
            p11p(1) = p11(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - di
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p11, p11p, bzwz)
            biaozhu.TextOverride = jingyan1 & "mm"
            
            p21p(0) = 2 * m0(0) - p21(0)
            p21p(1) = p21(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - di - jingyan1_d
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p21, p21p, bzwz)
            biaozhu.TextOverride = jingyan2 & "mm"
            
            p23p(0) = 2 * m0(0) - p23(0)
            p23p(1) = p23(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - ding - jingyan1_d
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p23, p23p, bzwz)
            biaozhu.TextOverride = taoguan2 & "mm"
            
            p31p(0) = 2 * m0(0) - p31(0)
            p31p(1) = p31(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - di - jingyan2_d
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p31, p31p, bzwz)
            biaozhu.TextOverride = jingyan3 & "mm"
            
            If taoguan3_start <> 0 Then
                p33p(0) = 2 * m0(0) - p33(0)
                p33p(1) = p33(1)
                bzwz(0) = m0(0)
                bzwz(1) = m0(1) - ding - jingyan2_d
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p33, p33p, bzwz)
                biaozhu.TextOverride = taoguan3 & "mm"
            End If
            
            If jingyan4_d <> 0 Then
                p41p(0) = 2 * m0(0) - p41(0)
                p41p(1) = p41(1)
                bzwz(0) = m0(0)
                bzwz(1) = m0(1) - di - jingyan3_d
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p41, p41p, bzwz)
                biaozhu.TextOverride = jingyan4 & "mm"
            End If
            
            If taoguan4_start <> 0 Then
                p43p(0) = 2 * m0(0) - p43(0)
                p43p(1) = p43(1)
                bzwz(0) = m0(0)
                bzwz(1) = m0(1) - ding - jingyan3_d
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p43, p43p, bzwz)
                biaozhu.TextOverride = taoguan4 & "mm"
            End If
            
            '--------------------------------------------------------井身结构中的标注深度数据
            yinxian2(0) = m0(0) + 350
            b1move(0) = yinxian2(0) - 20
            
            '------------------------------井深
            yinxian1(0) = p12(0)
            yinxian1(1) = m0(1) - bg
            yinxian2(1) = yinxian1(1)
            'Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, jingshen & "m")  '(位置，文字宽度，字符串)
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
            b1move(1) = yinxian2(1) + 20
            MTextObj.Move yinxian2, b1move
            
            '------------------------------一开
            yinxian1(0) = p12(0)
            yinxian1(1) = p12(1)
            yinxian2(1) = p12(1)
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, jingyan1_d & "m") '(位置，文字宽度，字符串)
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
            b1move(1) = yinxian2(1)
            MTextObj.Move yinxian2, b1move
            
            yinxian1(0) = p12(0)
            yinxian1(1) = p25(1)
            yinxian2(1) = p25(1)
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, taoguan2_start & "m") '(位置，文字宽度，字符串)
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
            b1move(1) = yinxian2(1) + 20
            MTextObj.Move yinxian2, b1move
            
            '------------------------------二开
            yinxian1(0) = p22(0)
            yinxian1(1) = p22(1)
            yinxian2(1) = p22(1)
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, jingyan2_d & "m") '(位置，文字宽度，字符串)
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
            b1move(1) = yinxian2(1)
            MTextObj.Move yinxian2, b1move
            
            '------------------------------三开
            If jingyan4_d > 0 Then
                yinxian1(0) = p32(0)
                yinxian1(1) = p32(1)
                yinxian2(1) = p32(1)
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, jingyan3_d & "m") '(位置，文字宽度，字符串)
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
                b1move(1) = yinxian2(1)
                MTextObj.Move yinxian2, b1move
            End If
            
            
            If taoguan3_start > 0 Then
                yinxian1(0) = p22(0)
                yinxian1(1) = p33(1)
                yinxian2(1) = p33(1)
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, taoguan3_start & "m") '(位置，文字宽度，字符串)
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
                b1move(1) = yinxian2(1) + 20
                MTextObj.Move yinxian2, b1move
            End If
            
            If taoguan4_start > 0 Then
                yinxian1(0) = p32(0)
                yinxian1(1) = p43(1)
                yinxian2(1) = p43(1)
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 20, taoguan4_start & "m") '(位置，文字宽度，字符串)
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3  '3-右对齐，4-左对齐，5-居中对齐
                b1move(1) = yinxian2(1) + 20
                MTextObj.Move yinxian2, b1move
            End If
            
            '-------------------------水泥固井
            
            pointsarray(0) = p23(0) + 15
            pointsarray(1) = p23(1) - (jingyan2_d - jingyan1_d) / 2
            pointsarray(3) = pointsarray(0) + 100
            pointsarray(4) = pointsarray(1)
            dimpoint(0) = pointsarray(3)
            dimpoint(1) = pointsarray(4)
            Set MLEADERR = ThisDrawing.ModelSpace.AddMLeader(pointsarray, 0)
            MLEADERR.TextString = "水泥固井"
            MLEADERR.TextHeight = 20
            MLEADERR.TextWidth = 20
            MLEADERR.ArrowheadSize = 0
            
            
            
        ElseIf jingxing = "定向井" Then
            '--------------------------------------------------------------------------------------------------------------------------------井身结构中的点++++定向井
            ThisDrawing.ActiveLayer = lay_jingshen
            theta = 8 / 180 * PI
            
            'sin4 = Sin(theta)
            
            m0(0) = b1(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + 300
            m0(1) = b1(1)
            
            m1(0) = m0(0)
            m1(1) = m0(1) - jingyan1_d
            
            m2(0) = m0(0) + Tan(theta) * (jingyan2_d - jingyan1_d)
            m2(1) = m0(1) - jingyan2_d
            
            m3(0) = m0(0) + Tan(theta) * (jingyan3_d - jingyan1_d)
            m3(1) = m0(1) - jingyan3_d
            
            
            If jingyan4 > 0 Then
                m4(0) = m0(0) + Tan(theta) * (jingyan4_d - jingyan1_d)
                m4(1) = m0(1) - jingyan4_d
            End If
            
            '------------------------------------一开
            p11(0) = m0(0) + jingyan1 / 2
            p11(1) = m0(1)
            
            p12(0) = p11(0)
            p12(1) = p11(1) - jingyan1_d
            
            p13(0) = p11(0) - (jingyan1 - taoguan1) / 2
            p13(1) = p11(1)
            
            p14(0) = p13(0)
            p14(1) = p13(1) - jingyan1_d
            
            '-----------------------------
            p11l(0) = m0(0) - jingyan1 / 2
            p11l(1) = p11(1)
            
            p12l(0) = p11l(0)
            p12l(1) = p12(1)
            
            p13l(0) = p11l(0) + (jingyan1 - taoguan1) / 2
            p13l(1) = p13(1)
            
            p14l(0) = p13l(0)
            p14l(1) = p14(1)
            
            '------------------------------------二开
            p21(0) = m1(0) + jingyan2 / 2
            p21(1) = m1(1)
            
            p22(0) = m2(0) + jingyan2 / 2
            p22(1) = m2(1)
            
            p23(0) = m0(0) + taoguan2 / 2 - Tan(theta) * (jingyan1_d - taoguan2_start)
            p23(1) = m0(1) - taoguan2_start
            
            p24(0) = p23(0) + Tan(theta) * (taoguan2_end - taoguan2_start)
            p24(1) = p23(1) - taoguan2_end + taoguan2_start
            
            p25(0) = p13(0)
            p25(1) = p23(1)
            '-----------------------------
            p21l(0) = m1(0) - jingyan2 / 2
            p21l(1) = p21(1)
            
            p22l(0) = m2(0) - jingyan2 / 2
            p22l(1) = p22(1)
            
            p23l(0) = p23(0) - taoguan2
            p23l(1) = p23(1)
            
            p24l(0) = p24(0) - taoguan2
            p24l(1) = p24(1)
            
            p25l(0) = p13l(0)
            p25l(1) = p25(1)
            '------------------------------------三开
            p31(0) = m2(0) + jingyan3 / 2
            p31(1) = m2(1)
            
            p32(0) = m3(0) + jingyan3 / 2
            p32(1) = m3(1)
            
            p33(0) = m2(0) + taoguan3 / 2 - Tan(theta) * (jingyan2_d - taoguan3_start)
            p33(1) = m0(1) - taoguan3_start
            
            p34(0) = p33(0) + Tan(theta) * (taoguan3_end - taoguan3_start)
            p34(1) = p33(1) - taoguan3_end + taoguan3_start
            
            p35(0) = p23(0) + Tan(theta) * (taoguan3_start - taoguan2_start)
            p35(1) = p33(1)
            '---------------------------------
            p31l(0) = p31(0) - jingyan3
            p31l(1) = p31(1)
            
            p32l(0) = p32(0) - jingyan3
            p32l(1) = p32(1)
            
            p33l(0) = p33(0) - taoguan3
            p33l(1) = p33(1)
            
            p34l(0) = p34(0) - taoguan3
            p34l(1) = p34(1)
            
            p35l(0) = p35(0) - taoguan2
            p35l(1) = p35(1)
            
            
            If jingyan4 > 0 Then
                
                p41(0) = m3(0) + jingyan4 / 2
                p41(1) = m3(1)
                
                p42(0) = m4(0) + jingyan4 / 2
                p42(1) = m4(1)
                
                p41l(0) = p41(0) - jingyan4
                p41l(1) = p41(1)
                
                p42l(0) = p42(0) - jingyan4
                p42l(1) = p42(1)
                
                If taoguan4_start > 0 Then
                    p43(0) = m3(0) + taoguan4 / 2 - Tan(theta) * (jingyan3_d - taoguan4_start)
                    p43(1) = m0(1) - taoguan4_start
                    
                    p44(0) = p43(0) + Tan(theta) * (taoguan4_end - taoguan4_start)
                    p44(1) = p43(1) - taoguan4_end + taoguan4_start
                    
                    p43l(0) = p43(0) - taoguan4
                    p43l(1) = p43(1)
                    
                    p44l(0) = p44(0) - taoguan4
                    p44l(1) = p44(1)
                End If
                
            End If
            
            ThisDrawing.ActiveLayer = lay_tianchong
            Set box(0) = drawbox(p13, p12)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            hobj.Mirror m0, m1
            
            Set box(0) = duobianxing(p23, p24, p22, p21, p14, p25)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            
            Set box(0) = duobianxing(p23l, p24l, p22l, p21l, p14l, p25l)
            Set hobj = ThisDrawing.ModelSpace.AddHatch(0, "solid", True)
            hobj.AppendInnerLoop box
            hobj.color = 253
            box(0).Delete
            
            
            ThisDrawing.ActiveLayer = lay_jingshen
            Set l11 = ThisDrawing.ModelSpace.AddLine(p11(), p12())
            Set l12 = ThisDrawing.ModelSpace.AddLine(p13(), p14())
            Set l13 = ThisDrawing.ModelSpace.AddLine(p12(), p21())
            Set l14 = ThisDrawing.ModelSpace.AddLine(p11(), p13())
            
            Set l11l = ThisDrawing.ModelSpace.AddLine(p11l(), p12l())
            Set l12l = ThisDrawing.ModelSpace.AddLine(p13l(), p14l())
            Set l13l = ThisDrawing.ModelSpace.AddLine(p12l(), p21l())
            Set l14l = ThisDrawing.ModelSpace.AddLine(p11l(), p13l())
            
            Set l21 = ThisDrawing.ModelSpace.AddLine(p21(), p22())
            Set l22 = ThisDrawing.ModelSpace.AddLine(p23(), p24())
            Set l23 = ThisDrawing.ModelSpace.AddLine(p22(), p31())
            Set l24 = ThisDrawing.ModelSpace.AddLine(p25(), p23())
            
            Set l21l = ThisDrawing.ModelSpace.AddLine(p21l(), p22l())
            Set l22l = ThisDrawing.ModelSpace.AddLine(p23l(), p24l())
            Set l23l = ThisDrawing.ModelSpace.AddLine(p22l(), p31l())
            Set l24l = ThisDrawing.ModelSpace.AddLine(p25l(), p23l())
            
            
            Set l31 = ThisDrawing.ModelSpace.AddLine(p31(), p32())
            Set l31l = ThisDrawing.ModelSpace.AddLine(p31l(), p32l())
            
            If taoguan3_start > 0 Then
                Set l32 = ThisDrawing.ModelSpace.AddLine(p33(), p34())
                Set l32l = ThisDrawing.ModelSpace.AddLine(p33l(), p34l())
                
            End If
            
            If jingyan4 = 0 Then
                Set lding3 = ThisDrawing.ModelSpace.AddLine(p33(), p33l())
            ElseIf jingyan4 > 0 Then
                Set l34 = ThisDrawing.ModelSpace.AddLine(p35(), p33())
                Set l34l = ThisDrawing.ModelSpace.AddLine(p35l(), p33l())
                Set l41 = ThisDrawing.ModelSpace.AddLine(p41(), p42())
                Set l41l = ThisDrawing.ModelSpace.AddLine(p41l(), p42l())
                Set l33 = ThisDrawing.ModelSpace.AddLine(p32(), p41())
                Set l33l = ThisDrawing.ModelSpace.AddLine(p32l(), p41l())
                If taoguan4_start > 0 Then
                    Set l42 = ThisDrawing.ModelSpace.AddLine(p43(), p44())
                    Set l42l = ThisDrawing.ModelSpace.AddLine(p43l(), p44l())
                    Set lding4 = ThisDrawing.ModelSpace.AddLine(p43l(), p43())
                End If
            End If
            
            '--------------------------------------------------------井身结构中的标注
            
            ThisDrawing.ActiveLayer = lay_biaozhu
            di = 150
            ding = 100
            
            p13p(0) = 2 * m0(0) - p13(0)
            p13p(1) = p13(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - ding
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p13, p13p, bzwz)
            biaozhu.TextOverride = taoguan1 & "mm"
            
            p11p(0) = 2 * m0(0) - p11(0)
            p11p(1) = p11(1)
            bzwz(0) = m0(0)
            bzwz(1) = m0(1) - di
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(p11, p11p, bzwz)
            biaozhu.TextOverride = jingyan1 & "mm"
            
            '----------------------定向井二开
            
            R_chicun(0) = p23(0) + (taoguan2_end - taoguan2_start - 150) / 2 * Tan(theta)
            R_chicun(1) = p23(1) - (taoguan2_end - taoguan2_start - 150) / 2
            L_chicun(0) = R_chicun(0) - taoguan2
            L_chicun(1) = R_chicun(1)
            bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
            bzwz(1) = R_chicun(1)
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
            biaozhu.TextOverride = taoguan2 & "mm"
            
            R_chicun(0) = p21(0) + (jingyan2_d - jingyan1_d) / 2 * Tan(theta)
            R_chicun(1) = p21(1) - (jingyan2_d - jingyan1_d) / 2
            L_chicun(0) = R_chicun(0) - jingyan2
            L_chicun(1) = R_chicun(1)
            bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
            bzwz(1) = R_chicun(1)
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
            biaozhu.TextOverride = jingyan2 & "mm"
            
            '---------------------------
            
            If taoguan3_start > 0 Then
                R_chicun(0) = p33(0) + (taoguan3_end - taoguan3_start - 150) / 2 * Tan(theta)
                R_chicun(1) = p33(1) - (taoguan3_end - taoguan3_start - 150) / 2
                L_chicun(0) = R_chicun(0) - taoguan3
                L_chicun(1) = R_chicun(1)
                bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
                bzwz(1) = R_chicun(1)
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
                biaozhu.TextOverride = taoguan3 & "mm"
            End If
            
            R_chicun(0) = p31(0) + (jingyan3_d - jingyan2_d) / 2 * Tan(theta)
            R_chicun(1) = p31(1) - (jingyan3_d - jingyan2_d) / 2
            L_chicun(0) = R_chicun(0) - jingyan3
            L_chicun(1) = R_chicun(1)
            bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
            bzwz(1) = R_chicun(1)
            Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
            biaozhu.TextOverride = jingyan3 & "mm"
            
            If taoguan4_start > 0 Then
                R_chicun(0) = p43(0) + (taoguan4_end - taoguan4_start - 150) / 2 * Tan(theta)
                R_chicun(1) = p43(1) - (taoguan4_end - taoguan4_start - 150) / 2
                L_chicun(0) = R_chicun(0) - taoguan4
                L_chicun(1) = R_chicun(1)
                bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
                bzwz(1) = R_chicun(1)
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
                biaozhu.TextOverride = taoguan4 & "mm"
            End If
            
            If jingyan4 > 0 Then
                R_chicun(0) = p41(0) + (jingyan4_d - jingyan3_d) / 2 * Tan(theta)
                R_chicun(1) = p41(1) - (jingyan4_d - jingyan3_d) / 2
                L_chicun(0) = R_chicun(0) - jingyan4
                L_chicun(1) = R_chicun(1)
                bzwz(0) = (R_chicun(0) + L_chicun(0)) / 2
                bzwz(1) = R_chicun(1)
                Set biaozhu = ThisDrawing.ModelSpace.AddDimAligned(R_chicun, L_chicun, bzwz)
                biaozhu.TextOverride = jingyan4 & "mm"
            End If
            
            
            
            pointsarray(0) = p23(0) + 15 + (taoguan2_end - taoguan2_start) / 2 * Tan(theta)
            pointsarray(1) = p23(1) - (jingyan2_d - jingyan1_d) / 2
            pointsarray(3) = pointsarray(0) + 100
            pointsarray(4) = pointsarray(1)
            dimpoint(0) = pointsarray(3)
            dimpoint(1) = pointsarray(4)
            Set MLEADERR = ThisDrawing.ModelSpace.AddMLeader(pointsarray, 0)
            MLEADERR.TextString = "水泥固井"
            MLEADERR.TextHeight = 20
            MLEADERR.TextWidth = 20
            MLEADERR.ArrowheadSize = 0
            '--------------------------------------------------------井身结构中的标注深度数据
            
            '------------------------------井深
            yinxian1(0) = b4(0) + lk(1) + lk(2) + lk(3) + lk(4) + lk(5) + 500
            yinxian1(1) = b4(1)
            yinxian2(0) = yinxian1(0) + 150
            yinxian2(1) = yinxian1(1)
            b1move(0) = yinxian2(0) - 20
            b1move(1) = yinxian2(1) + 30
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, "斜深:" & jingshen & "m")
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3
            MTextObj.Move yinxian2, b1move
            
            
            '------------------------------一开
            yinxian1(0) = p25(0) + (jingyan1 - taoguan1) / 2 'gai
            yinxian1(1) = p25(1) 'gai
            yinxian2(0) = yinxian1(0) + 200
            yinxian2(1) = yinxian1(1)
            b1move(0) = yinxian1(0)
            b1move(1) = yinxian1(1) + 30
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, " " & taoguan2_start & "m")
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3
            MTextObj.Move yinxian2, b1move
            
            yinxian1(0) = p12(0)
            yinxian1(1) = p12(1) 'gai
            yinxian2(0) = yinxian1(0) + 200
            yinxian2(1) = yinxian1(1)
            b1move(0) = yinxian1(0)
            b1move(1) = yinxian1(1)
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, " " & jingyan1_d & "m")
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3
            MTextObj.Move yinxian2, b1move
            
            '------------------------------二开
            If taoguan3_start > 0 Then
                yinxian1(0) = p35(0) + (jingyan2 - taoguan2) / 2 'gai
                yinxian1(1) = p35(1) 'gai
                yinxian2(0) = yinxian1(0) + 200
                yinxian2(1) = yinxian1(1)
                b1move(0) = yinxian1(0)
                b1move(1) = yinxian1(1) + 30
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, "斜深:" & taoguan3_start & "m")
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3
                MTextObj.Move yinxian2, b1move
            End If
            
            yinxian1(0) = p22(0)
            yinxian1(1) = p22(1) 'gai
            yinxian2(0) = yinxian1(0) + 200
            yinxian2(1) = yinxian1(1)
            b1move(0) = yinxian1(0)
            b1move(1) = yinxian1(1)
            Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
            Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, "斜深:" & jingyan2_d & "m")
            MTextObj.Height = 20
            MTextObj.AttachmentPoint = 3
            MTextObj.Move yinxian2, b1move
            
            
            '------------------------------三开
            If taoguan4_start > 0 Then
                yinxian1(0) = p43(0) + (jingyan3 - taoguan4) / 2 'gai
                yinxian1(1) = p43(1) 'gai
                yinxian2(0) = yinxian1(0) + 200
                yinxian2(1) = yinxian1(1)
                b1move(0) = yinxian1(0)
                b1move(1) = yinxian1(1) + 30
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, "斜深:" & taoguan4_start & "m")
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3
                MTextObj.Move yinxian2, b1move
            End If
            
            If jingyan4 > 0 Then
                yinxian1(0) = p32(0)
                yinxian1(1) = p32(1) 'gai
                yinxian2(0) = yinxian1(0) + 200
                yinxian2(1) = yinxian1(1)
                b1move(0) = yinxian1(0)
                b1move(1) = yinxian1(1)
                Set yx = ThisDrawing.ModelSpace.AddLine(yinxian1(), yinxian2())
                Set MTextObj = ThisDrawing.ModelSpace.AddMText(yinxian2, 200, "斜深:" & jingyan3_d & "m")
                MTextObj.Height = 20
                MTextObj.AttachmentPoint = 3
                MTextObj.Move yinxian2, b1move
            End If
            
        End If
        
        
        ThisDrawing.Application.ZoomAll
        
    Next i
    
    
    
    
    '--------------------------------------------------------修改标注的样式
    Dim newdimstyle As AcadDimStyle
    dimsize = 10
    Set newdimstyle = ThisDrawing.DimStyles.Add("新新标注")
    With ThisDrawing
        .SetVariable "DIMLFAC", 1 '全局比例因子
        .SetVariable "DIMSCALE", 1 '线性比例因子
        .SetVariable "DIMTXSTY", "STANDARD"  '文字样式
        .SetVariable "DIMDEC", 3 '文字小数位数
        .SetVariable "DIMDSEP", "." '文字分隔符用小数点
        .SetVariable "DIMTIH", 0  '文字与尺寸线对齐
        .SetVariable "DIMTXT", 2.5 * dimsize '文字高度
        .SetVariable "DIMASZ", 2.5 * dimsize '箭头大小
        .SetVariable "DIMDLI", 3.75 * dimsize '基线尺寸线的间距
        .SetVariable "DIMEXE", 1.25 * dimsize * 0 '尺寸界线超出尺寸线的距离
        .SetVariable "DIMEXO", 0.625 * dimsize * 0 '尺寸界线偏移量
        .SetVariable "DIMGAP", 0.625 * dimsize '文字偏移量
    End With
    newdimstyle.CopyFrom ThisDrawing '复制系统变量到标注样式
    ThisDrawing.ActiveDimStyle = newdimstyle '设为当前标注样式
    For Each ent In ThisDrawing.ModelSpace
        If InStr(ent.ObjectName, "Dim") > 0 Or InStr(ent.ObjectName, "MLEADER") > 0 Then  '对象类型名称中含有Dim，可视其为标注对象
            ent.StyleName = "新新标注"
        End If
    Next
    
    
    
End Sub

Function drawbox(p1, p2) As AcadLWPolyline '用对角线画矩形
    Dim boxp(0 To 7) As Double
    boxp(0) = p1(0): boxp(1) = p1(1)
    boxp(2) = p1(0): boxp(3) = p2(1)
    boxp(4) = p2(0): boxp(5) = p2(1)
    boxp(6) = p2(0): boxp(7) = p1(1)
    Set drawbox = ThisDrawing.ModelSpace.AddLightWeightPolyline(boxp)
    drawbox.Closed = True
End Function

Function duobianxing(p1, p2, p3, p4, p5, p6) As AcadLWPolyline '6点多边形
    Dim boxp(0 To 11) As Double
    boxp(0) = p1(0): boxp(1) = p1(1)
    boxp(2) = p2(0): boxp(3) = p2(1)
    boxp(4) = p3(0): boxp(5) = p3(1)
    boxp(6) = p4(0): boxp(7) = p4(1)
    boxp(8) = p5(0): boxp(9) = p5(1)
    boxp(10) = p6(0): boxp(11) = p6(1)
    Set duobianxing = ThisDrawing.ModelSpace.AddLightWeightPolyline(boxp)
    duobianxing.Closed = True
End Function


Sub Example_AddPolyLine()
    Dim xyz() As Double
    n = 6 '顶点个数
    ReDim xyz(0 To 3 * n - 1)
    xyz(0) = 4: xyz(1) = 7: xyz(2) = 0
    xyz(3) = 5: xyz(4) = 7: xyz(5) = 0
    xyz(6) = 6: xyz(7) = 7: xyz(8) = 0
    xyz(9) = 4: xyz(10) = 6: xyz(11) = 0
    xyz(12) = 5: xyz(13) = 6: xyz(14) = 0
    xyz(15) = 6: xyz(16) = 6: xyz(17) = 0
    Set mLineObj = ThisDrawing.ModelSpace.AddPolyline(xyz)
    ThisDrawing.Application.ZoomAll
    MsgBox "A new PolyLine has been added to the drawing"
End Sub

Sub Myaddline()
    Dim ln As AcadLine
    Dim startPt(2) As Double, EndPt(2) As Double
    startPt(0) = 0
    startPt(1) = 0
    startPt(0) = 100
    startPt(1) = 50
    Set ln = ThisDrawing.ModelSpace.AddLine(startPt(), EndPt())
    ln.color = acRed
    ZoomAll
End Sub

