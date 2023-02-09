VERSION 5.00
Begin VB.Form R76101120 
   Caption         =   "Hw4"
   ClientHeight    =   6156
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   15708
   LinkTopic       =   "Form2"
   ScaleHeight     =   6156
   ScaleWidth      =   15708
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   11880
      TabIndex        =   18
      Top             =   1440
      Width           =   3500
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   8040
      TabIndex        =   16
      Top             =   1440
      Width           =   3500
   End
   Begin VB.ListBox List9 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   4200
      TabIndex        =   13
      Top             =   1440
      Width           =   3500
   End
   Begin VB.ListBox List7 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   4200
      TabIndex        =   12
      Top             =   7080
      Width           =   3500
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4536
      Left            =   12120
      TabIndex        =   8
      Top             =   7080
      Width           =   3855
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4536
      Left            =   8040
      TabIndex        =   5
      Top             =   7080
      Width           =   3855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   3500
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4416
      Left            =   360
      TabIndex        =   3
      Top             =   7080
      Width           =   3500
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Text            =   "pima.txt"
      Top             =   300
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "NBC"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "KNN"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   12000
      TabIndex        =   17
      Top             =   1080
      Width           =   1692
   End
   Begin VB.Label Label9 
      Caption         =   "Entropy_Based Interval"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   15
      Top             =   1080
      Width           =   2652
   End
   Begin VB.Label Label8 
      Caption         =   "Entropy_Based data"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   14
      Top             =   6720
      Width           =   3372
   End
   Begin VB.Label Label5 
      Caption         =   "Equal_Frequency Interval"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   12240
      TabIndex        =   11
      Top             =   6840
      Width           =   1812
   End
   Begin VB.Label Label4 
      Caption         =   "Equal_Frequency data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8160
      TabIndex        =   10
      Top             =   6720
      Width           =   1692
   End
   Begin VB.Label Label7 
      Caption         =   "k-fold accuracy rate"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   10680
      TabIndex        =   9
      Top             =   480
      Width           =   3012
   End
   Begin VB.Label Label3 
      Caption         =   "Equal_Width Interval"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   2412
   End
   Begin VB.Label Label2 
      Caption         =   "Equal_Width data"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   6
      Top             =   6720
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "R76101120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String, out_file As String, nstr As String
Dim out_rec As String
Dim att(768, 9) As Double
Dim origin_att(768, 9), dummy As Double
Dim equal_width(9, 9) As Integer
Dim entropy_based(9, 9) As Integer
Dim max(9, 1), min(9, 1), bigM, bound, Interval, f_sort(768, 2) As Double
Dim i, j, num, pos, bin, Sorted, ii, jj, cut_len, cut_size, p As Integer
Dim val, key, a, b As Integer
Dim cut_array(768) As Double

Sub Entropy(EA() As Double)
    Dim cut As Integer
    Dim size, l_num, r_num, cut_index, NL, NR As Integer
    Dim T_Ent, Le, Re, cut_point, prob, Ent, Min_Ent, Min_cut_point, delta As Double
    Dim countEA(1), countL(1), countR(1) As Integer
    Min_Ent = bigM
    T_Ent = 0
    cut = 0

    size = UBound(EA) - LBound(EA)
    If size > 1 Then
        For i = 1 To size - 1
            If EA(i, 1) <> EA(i + 1, 1) Then
                cut = cut + 1
            End If
        Next
    End If
    
    
    If cut = 0 Then
        'List8.AddItem "stop"
        'List8.AddItem "-----------------------"
    Else
        For j = 0 To 1
            countEA(j) = 0
        Next
        For i = 1 To size
            countEA(EA(i, 1)) = countEA(EA(i, 1)) + 1
        Next
       
        For i = 0 To 1
            prob = countEA(i) / size
            If prob <> 0 Then
                T_Ent = T_Ent - prob * Math.Log(prob) / Math.Log(2)
            End If
        Next
        
        For i = 1 To size - 1
            If EA(i, 1) <> EA(i + 1, 1) Then
                cut_point = (EA(i, 0) + EA(i + 1, 0)) / 2
                For j = 0 To 1
                    countL(j) = 0
                    countR(j) = 0
                Next
                For j = 1 To size
                    'If j < i + 1 Then
                    If EA(j, 0) <= cut_point Then
                        countL(EA(j, 1)) = countL(EA(j, 1)) + 1
                    Else
                        countR(EA(j, 1)) = countR(EA(j, 1)) + 1
                    End If
                Next
                NL = 0
                NR = 0
                For j = 0 To 1
                    NL = NL + countL(j)
                    NR = NR + countR(j)
                Next
                If NL = 0 Or NR = 0 Then
                    Exit For
                End If

                Le = 0
                Re = 0
                For p = 0 To 1
                    'prob = countL(p) / i
                    prob = countL(p) / NL
                    If prob <> 0 Then
                        Le = Le - prob * Math.Log(prob) / Math.Log(2)
                    End If
                Next
                For p = 0 To 1
                    'prob = countR(p) / (size - i)
                    prob = countR(p) / NR
                    If prob <> 0 Then
                        Re = Re - prob * Math.Log(prob) / Math.Log(2)
                    End If
                Next
                
                Ent = Le * NL / size + Re * NR / size
                'Ent = Le * i / size + Re * (size - i) / size
                Dim k, k1, k2 As Integer
                Dim reject As Double
                'reject condition
                k = 0
                k1 = 0
                k2 = 0
                For p = 0 To 1
                    If countEA(p) <> 0 Then
                        k = k + 1
                    End If
                    If countL(p) <> 0 Then
                        k1 = k1 + 1
                    End If
                    If countR(p) <> 0 Then
                        k2 = k2 + 1
                    End If
                Next

                delta = (Math.Log(3 ^ k - 2) / Math.Log(2)) - k * T_Ent + k1 * Le + k2 * Re
                reject = T_Ent - Ent - ((Math.Log(size - 1) / Math.Log(2)) + delta) / size
                If Min_Ent >= Ent And reject > 0 Then
                    Min_Ent = Ent
                    Min_cut_point = cut_point
                    'cut_index = i + 1
                    cut_index = NL + 1
                End If
            End If
        Next
        If Min_Ent <> bigM Then
            cut_len = cut_len + 1

            cut_array(cut_len) = Min_cut_point

            l_num = cut_index - 1
            r_num = size - l_num
            Dim left() As Double
            Dim right() As Double
            ReDim left(l_num, 2) As Double
            ReDim right(r_num, 2) As Double
            'List8.AddItem l_num & " " & r_num
            Dim l, r As Integer
            l = 1
            r = 1
            For i = 1 To size
                If i < cut_index Then
                    left(l, 0) = EA(i, 0)
                    left(l, 1) = EA(i, 1)
                    l = l + 1
                Else
                    right(r, 0) = EA(i, 0)
                    right(r, 1) = EA(i, 1)
                    r = r + 1
                End If
            Next
            Call Entropy(left)
            Call Entropy(right)
        End If
    End If
End Sub

Sub BAY(b_att() As Double)
    Dim ran(768), size, num_bay, temp, a, b, c, d As Integer
    Dim train_data() As Double
    Dim test_data() As Double
    Dim Cmax, bestC, pC(2), pxc(8, 2), total_pxc, correct, v(8) As Double
    Dim predict() As Double
    Dim test_num, train_num As Integer
    Randomize Timer
    Dim totalc As Integer
    totalc = 0
    
    size = UBound(b_att) - LBound(b_att)
    For i = 1 To size
        For j = 1 To 8
            If v(j) <= b_att(i, j) Then
                v(j) = b_att(i, j)
            End If
        Next
    Next
    'List5.AddItem v(1) & " " & v(2) & " " & v(3) & " " & v(4) & " " & v(5) & " " & v(6) & " " & v(7) & " " & v(8) & " " & v(9)
    For i = 1 To size
        ran(i) = i
    Next
    For i = size To 1 Step -1
        'num_bay = Fix(Rnd() * i) + 1
        num_bay = Int(size * Rnd())
        If i <> num_bay Then
            temp = ran(i)
            ran(i) = ran(num_bay)
            ran(num_bay) = temp
        End If
    Next
    
    Dim fold(4) As Integer
    fold(0) = 154
    fold(1) = 154
    fold(2) = 154
    fold(3) = 153
    fold(4) = 153
    Dim fcount As Integer
    fcount = 0
    
    For i = 0 To 616 Step fold(fcount) '5 fold cross validation
        
        ReDim test_data(0, 0)
        ReDim train_data(0, 0)
        For j = 1 To 768
            If j > i And j <= i + fold(fcount) Then
                ReDim test_data(UBound(test_data) + 1, 9) As Double
            Else
                ReDim train_data(UBound(train_data) + 1, 9) As Double
            End If
        Next
        
        test_num = 1
        train_num = 1
        For j = 1 To 768
            If j > i And j <= i + fold(fcount) Then
                For b = 1 To 9
                    test_data(test_num, b) = b_att(ran(j), b)
                Next
                test_num = test_num + 1
            Else
                For b = 1 To 9
                    train_data(train_num, b) = b_att(ran(j), b)
                Next
                train_num = train_num + 1
            End If
        Next
        ReDim predict(UBound(test_data)) As Double
        For b = 1 To 2
            pC(b) = 0
            For a = 1 To UBound(train_data)
                If train_data(a, 9) = b - 1 Then
                    pC(b) = pC(b) + 1
                End If
            Next
            'List6.AddItem pc(b)
            'pc(b) = pc(b) / UBound(train_data)
        Next
        'List6.AddItem "+++++++++++++++++++++++++++"
        For a = 1 To UBound(test_data) '¿ù¤F
            For c = 1 To 8
                For d = 1 To 2
                    pxc(c, d) = 0
                Next
            Next
            
            For c = 1 To 8
                For b = 1 To UBound(train_data)
                    If test_data(a, c) = train_data(b, c) Then
                        pxc(c, train_data(b, 9)) = pxc(c, train_data(b, 9)) + 1
                    End If
                Next
            Next
            
            
            For c = 1 To 8
                For d = 1 To 2
                    'List6.AddItem c & " " & d & " " & pxc(c, d)
                    pxc(c, d) = (pxc(c, d) + 1) / (pC(d) + v(c) + 1)
                    'List6.AddItem c & " " & d & " " & pxc(c, d)
                Next
                'List6.AddItem "---------------------------------"
            Next
            Cmax = -9999
            bestC = 0
            For d = 1 To 2
                total_pxc = 1
                For c = 1 To 8
                    total_pxc = total_pxc * pxc(c, d)
                Next
                total_pxc = total_pxc * pC(d) / UBound(train_data)
                If Cmax <= total_pxc Then
                    Cmax = total_pxc
                    bestC = d - 1
                End If
            Next
            predict(a) = bestC
            'List6.AddItem predict(a)
        Next
        'List6.AddItem "-------------------------"
        correct = 0
        For a = 1 To UBound(test_data)
            If test_data(a, 9) = predict(a) Then
                correct = correct + 1
                totalc = totalc + 1
            End If
        Next
        correct = correct / UBound(test_data)
        List6.AddItem "fold" & (fcount + 1) & " : " & correct
        
        fcount = fcount + 1
    Next
    'List6.AddItem "totalc" & " : " & totalc
    
End Sub
Sub KNN(k_att() As Double)
    Dim ran(768), size, num_bay, temp, a, b, c, d As Integer
    Dim train_data() As Double
    Dim test_data() As Double
    Dim Cmax, bestC, pxc(8, 2), total_pxc, correct, v(8) As Double
    Dim predict() As Double
    Dim test_num, train_num As Integer
    
    'Randomize 'Timer
    Dim totalc As Integer
    totalc = 0
    
    size = UBound(k_att) - LBound(k_att)
    For i = 1 To size
        For j = 1 To 8
            If v(j) <= k_att(i, j) Then
                v(j) = k_att(i, j)
            End If
        Next
    Next
    'List5.AddItem v(1) & " " & v(2) & " " & v(3) & " " & v(4) & " " & v(5) & " " & v(6) & " " & v(7) & " " & v(8) & " " & v(9)
    For i = 1 To size
        ran(i) = i
    Next
    For i = size To 1 Step -1
        'num_bay = Fix(Rnd() * i) + 1
        num_bay = Int(size * Rnd() + 1)
        If i <> num_bay Then
            temp = ran(i)
            ran(i) = ran(num_bay)
            ran(num_bay) = temp
        End If
    Next
    
    Dim distance() As Double
    Dim disindex() As Integer
    Dim fold(5) As Integer
    fold(0) = 154
    fold(1) = 154
    fold(2) = 154
    fold(3) = 153
    fold(4) = 153
    fold(5) = 153
    Dim fcount As Integer
    fcount = 0
    
    For i = 0 To 4 '616 Step fold(fcount + 1) '5 fold cross validation
      If i < 3 Then
                ReDim test_data(0, 0)
                ReDim train_data(0, 0)
                For j = 1 To 768
                    If j > i * 154 And j <= i * 154 + 154 Then
                        ReDim test_data(UBound(test_data) + 1, 9) As Double
                    Else
                        ReDim train_data(UBound(train_data) + 1, 9) As Double
                    End If
                Next
                
                test_num = 1
                train_num = 1
                For j = 1 To 768
                    If j > i * 154 And j <= i * 154 + 154 Then
                        For b = 1 To 9
                            test_data(test_num, b) = k_att(ran(j), b)
                        Next
                        test_num = test_num + 1
                    Else
                        For b = 1 To 9
                            train_data(train_num, b) = k_att(ran(j), b)
                        Next
                        train_num = train_num + 1
                    End If
                Next
        ElseIf i = 3 Then
                ReDim test_data(0, 0)
                ReDim train_data(0, 0)
                For j = 1 To 768
                    If j > 462 And j <= 615 Then
                        ReDim test_data(UBound(test_data) + 1, 9) As Double
                    Else
                        ReDim train_data(UBound(train_data) + 1, 9) As Double
                    End If
                Next
                
                test_num = 1
                train_num = 1
                For j = 1 To 768
                    If j > 462 And j <= 615 Then
                        For b = 1 To 9
                            test_data(test_num, b) = k_att(ran(j), b)
                        Next
                        test_num = test_num + 1
                    Else
                        For b = 1 To 9
                            train_data(train_num, b) = k_att(ran(j), b)
                        Next
                        train_num = train_num + 1
                    End If
                Next
        ElseIf i = 4 Then
                ReDim test_data(0, 0)
                ReDim train_data(0, 0)
                For j = 1 To 768
                    If j > 615 And j <= 768 Then
                        ReDim test_data(UBound(test_data) + 1, 9) As Double
                    Else
                        ReDim train_data(UBound(train_data) + 1, 9) As Double
                    End If
                Next
                
                test_num = 1
                train_num = 1
                For j = 1 To 768
                    If j > 615 And j <= 768 Then
                        For b = 1 To 9
                            test_data(test_num, b) = k_att(ran(j), b)
                        Next
                        test_num = test_num + 1
                    Else
                        For b = 1 To 9
                            train_data(train_num, b) = k_att(ran(j), b)
                        Next
                        train_num = train_num + 1
                    End If
                Next
        
        End If
    
        Dim dis, powsum As Double
        ReDim distance(UBound(train_data)) As Double
        ReDim disindex(UBound(train_data)) As Integer
        powsum = 0
        For a = 1 To UBound(test_data)
            For b = 1 To UBound(train_data)
                For c = 1 To 8
                    powsum = powsum + (test_data(a, c) - train_data(b, c)) ^ 2
                Next
                dis = Sqr(powsum)
                distance(b) = dis
                disindex(b) = b
            Next
        
            Dim tempindex As Double
            For c = 0 To UBound(distance) - 1
                For j = c + 1 To UBound(distance)
                    If distance(j) < distance(c) Then
                        tempindex = disindex(c)
                        disindex(c) = disindex(j)
                        disindex(j) = tempindex
                    End If
                Next
            Next
            
            Dim C0, C1, pC As Integer
            pC = -1
            For c = 1 To 5
                If train_data(disindex(c), 9) = 0 Then
                    C0 = C0 + 1
                ElseIf train_data(disindex(c), 9) = 1 Then
                    C1 = C1 + 1
                End If
            Next
            If C0 > C1 Then
                pC = 0
            Else
                pC = 1
            End If
            C0 = 0
            C1 = 0
            
            If test_data(a, 9) = pC Then
                correct = correct + 1
                totalc = totalc + 1
                'List5.AddItem "totalc" & " : " & totalc
            End If
        Next
        
        
        correct = correct / UBound(test_data)
        List5.AddItem "fold" & (fcount + 1) & " : " & correct
        fcount = fcount + 1
    Next
   ' List5.AddItem "totalc" & " : " & totalc
    
End Sub

Private Sub Partition_click()
    bigM = 100000000
    For i = 1 To 9
        max(i, 0) = -bigM
        min(i, 0) = bigM
        For j = 0 To 9
            equal_width(i, j) = 0
            entropy_based(i, j) = 0
        Next
    Next
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    List7.Clear
    List9.Clear
    'check whether the file name is empty
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        'check whether the data file exists
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            Open in_file For Input As #1
            num = 1
            Do While Not EOF(1)
                For i = 1 To 9
                    'If i = 1 Then
                        'Input #1, dummy
                    'Else
                        Input #1, origin_att(num, i)
                        att(num, i) = origin_att(num, i)
                        If i < 12 Then
                            If max(i, 0) < att(num, i) Then
                                max(i, 0) = att(num, i)
                                max(i, 1) = num
                            End If
                            If min(i, 0) > att(num, i) Then
                                min(i, 0) = att(num, i)
                                min(i, 1) = num
                            End If
                        End If
                    'End If
                    
                Next
                
                'List1.AddItem att(num, 1) & " " & att(num, 2) & " " & att(num, 3) & " " & att(num, 4) & " " & att(num, 5) & " " & att(num, 6) & " " & att(num, 7) & " " & att(num, 8) & " " & att(num, 9) & " " & att(num, 10)
                num = num + 1
            Loop
            max(9, 0) = -bigM
            min(9, 0) = bigM
            Close #1
        End If
    End If
    bin = 10
    
    'Equal_Width
    For i = 1 To 768
        For j = 1 To 9
            If max(j, 0) <> -bigM Then
                Interval = (max(j, 0) - min(j, 0)) / bin
                pos = (att(i, j) - min(j, 0)) / Interval
                pos = -Fix(-(pos - 0.00001))
                'List2.AddItem att(i, j) & " " & min(j, 0) & " " & Interval & " " & pos
                att(i, j) = pos
                equal_width(j, att(i, j)) = equal_width(j, att(i, j)) + 1
            Else
                equal_width(j, att(i, j)) = equal_width(j, att(i, j)) + 1
            End If
        Next
        List1.AddItem att(i, 1) & " " & att(i, 2) & " " & att(i, 3) & " " & att(i, 4) & " " & att(i, 5) & " " & att(i, 6) & " " & att(i, 7) & " " & att(i, 8) & " " & att(i, 9) '& " " & att(i, 10)
    Next
    For i = 1 To 9
        If max(i, 0) <> -bigM Then
            Interval = (max(i, 0) - min(i, 0)) / bin
            List2.AddItem "A" & i
            List2.AddItem 0 & ":[" & min(i, 0) & " , " & min(i, 0) + Interval & "]"
            min(i, 0) = min(i, 0) + Interval
            For j = 2 To 10
                List2.AddItem j - 1 & ":(" & min(i, 0) & " , " & min(i, 0) + Interval & "]"
                min(i, 0) = min(i, 0) + Interval
            Next
        End If
    Next
    
    List6.AddItem "[Equal Width]"
    
    Call BAY(att)
    
    
    
    
    For i = 1 To 768
        For j = 1 To 9
            att(i, j) = origin_att(i, j)
        Next
    Next
    'Entropy Based
    For ii = 1 To 9
        If (ii <> 9) Then
            Dim ent_array(768, 2) As Double
            
            For jj = 1 To 768
                ent_array(jj, 0) = att(jj, ii)
                ent_array(jj, 1) = att(jj, 9)
            Next
            
            For a = 2 To 768
                val = ent_array(a, 0)
                key = ent_array(a, 1)
                b = a - 1
                While b > 0 And ent_array(b, 0) > val
                    ent_array(b + 1, 0) = ent_array(b, 0)
                    ent_array(b + 1, 1) = ent_array(b, 1)
                    b = b - 1
                Wend
                ent_array(b + 1, 0) = val
                ent_array(b + 1, 1) = key
            Next
            
            cut_len = 0
            cut_size = UBound(cut_array) - LBound(cut_array)
            For a = 1 To cut_size
                cut_array(a) = -1
            Next
            
            Call Entropy(ent_array)
            
            'List8.AddItem "====================================="
            'List9.AddItem UBound(cut_array) & " " & LBound(cut_array)
            
            For a = 2 To cut_size
                val = cut_array(a)
                b = a - 1
                While b > 0 And cut_array(b) > val
                    cut_array(b + 1) = cut_array(b)
                    b = b - 1
                Wend
                cut_array(b + 1) = val
            Next
            
            Dim save As Double
            Dim print_cut_point As New Collection
            print_cut_point.Add ent_array(1, 0)
            
            For a = 1 To cut_size
                If cut_array(a) <> -1 Then
                    'save = (cut_array(a) + cut_array(a + 1)) / 2
                    print_cut_point.Add cut_array(a)
                End If
            Next
            
            print_cut_point.Add ent_array(768, 0)
            List9.AddItem "A" & ii & ":"
            
            For a = 1 To print_cut_point.Count - 1
                If a = 1 Then
                    List9.AddItem a - 1 & ":[" & print_cut_point(a) & " , " & print_cut_point(a + 1) & "]"
                Else
                    List9.AddItem a - 1 & ":(" & print_cut_point(a) & " , " & print_cut_point(a + 1) & "]"
                End If
            Next
            
            For jj = 1 To 768
                If att(jj, ii) <= print_cut_point(2) Then
                    att(jj, ii) = 0
                Else
                    For a = 2 To print_cut_point.Count - 1
                        If att(jj, ii) > print_cut_point(a) And att(jj, ii) <= print_cut_point(a + 1) Then
                            att(jj, ii) = a - 1
                            Exit For
                        End If
                    Next
                End If
            Next
            
            For a = 1 To print_cut_point.Count
                print_cut_point.Remove (1)
            Next
            
        End If
    Next
    
    For i = 1 To 768
        List7.AddItem att(i, 1) & " " & att(i, 2) & " " & att(i, 3) & " " & att(i, 4) & " " & att(i, 5) & " " & att(i, 6) & " " & att(i, 7) & " " & att(i, 8) & " " & att(i, 9) ' & " " & att(i, 10)
    Next
    List6.AddItem "[Entropy Based]"
    Call BAY(att)
    
    Call KNN(att)
End Sub


