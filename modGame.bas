Attribute VB_Name = "modGame"
Option Explicit

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Const Play_W = 32
Public Const Play_H = 32

Public Const P_Right = Play_W * 0
Public Const P_Left = Play_W * 1

Public Const P_Stand = Play_H * 0
Public Const P_Walk = Play_H * 1
Public Const P_Punch = Play_H * 2
Public Const P_Kick = Play_H * 3
Public Const P_Toss = Play_H * 4
Public Const P_Blast = Play_H * 5
Public Const P_Up = Play_H * 6
Public Const P_Fall = Play_H * 7

Public Const WP_Blast = 0
Public Const WP_Bomb = 1
Public Const WP_Poo = 2

Public Type Player
Act As Long
X As Double
Y As Double
XS As Double
YS As Double
Dir As Long
Ani As Long
AniL As Long
OnG As Long
ID As Long
Jp As Long
HP As Double
MP As Double
Reload As Long
T As Long
End Type

Type Shot
X As Double
Y As Double
XS As Double
YS As Double
ID As Long
T As Long
Bounce As Long
Act As Long
End Type

Type Explo
X As Double
Y As Double
Fr As Long
FrL As Long
Act As Long
End Type

Type Bot
X As Double
Y As Double
dX As Double
dY As Double
T As Long
Act As Long
End Type

Public P(1 To 2) As Player
Public S(1 To 50) As Shot
Public Running As Boolean

Function MoveShots()
Dim I As Long, I2 As Long, opID As Long, XDr As Long, YDr As Long, X1, X2, Y1, Y2
For I = 1 To 50
    If S(I).Act = True Then
        S(I).X = S(I).X + S(I).XS
        S(I).Y = S(I).Y + S(I).YS
        If S(I).T = WP_Bomb Then
            S(I).YS = S(I).YS + 0.1
        End If
        opID = IIf(S(I).ID = 1, 2, 1)
        Select Case S(I).T
        Case WP_Blast
            If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y) = vbBlack Then
                S(I).Act = False
                frmGame.bMask.DrawWidth = 5
                frmGame.bSprite.DrawWidth = 5
                frmGame.bMask.PSet (S(I).X, S(I).Y), vbWhite
                frmGame.bSprite.PSet (S(I).X, S(I).Y), vbBlack
            End If
            For I2 = 1 To 2
                X1 = S(I).X
                X2 = P(I2).X + 10
                Y1 = S(I).Y
                Y2 = P(I2).Y
                If X2 > X1 And X2 < X1 + 16 And Y2 > Y1 And Y2 < Y1 + 8 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                    P(I2).HP = P(I2).HP - 10
                    S(I).Act = False
                End If
            Next I2
        Case WP_Bomb
            If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y + S(I).YS) = vbBlack Then
                XDr = 0
                YDr = 0
                If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y) Then XDr = 1
                If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y + S(I).YS) Then YDr = 1
                If XDr = 1 Then S(I).XS = S(I).XS * -0.8
                If YDr = 1 Then S(I).YS = S(I).YS * -0.8
            End If
        Case WP_Poo
            S(I).YS = S(I).YS + 0.35
            frmGame.Board.PSet (S(I).X, S(I).Y), &H4080&
            For I2 = 1 To 2
                X1 = S(I).X - 2.5
                X2 = P(I2).X + 10
                Y1 = S(I).Y - 2.5
                Y2 = P(I2).Y
                If X2 > X1 And X2 < X1 + 5 And Y2 > Y1 And Y2 < Y1 + 5 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                    P(I2).HP = P(I2).HP - 4
                    S(I).Act = False
                End If
            Next I2
            If S(I).Bounce > 0 Then
                If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y + S(I).YS) = vbBlack Then
                    XDr = 0
                    YDr = 0
                    If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y) Then XDr = 1
                    If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y + S(I).YS) Then YDr = 1
                    If XDr = 1 Then S(I).XS = S(I).XS * -0.8
                    If YDr = 1 Then S(I).YS = S(I).YS * -0.8
                    S(I).Bounce = S(I).Bounce - 1
                End If
            Else
                If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y) = vbBlack Then
                    S(I).Act = False
                    frmGame.bMask.DrawWidth = 5
                    frmGame.bSprite.DrawWidth = 5
                    frmGame.bMask.PSet (S(I).X - S(I).XS / 2, S(I).Y - S(I).YS / 2), vbBlack
                    frmGame.bSprite.PSet (S(I).X - S(I).XS / 2, S(I).Y - S(I).YS / 2), &H4080&
                End If
            End If
        End Select
        
    End If
Next I
End Function

Function pBlast(P As Player)
Dim I As Long
If P.Reload > 0 Then Exit Function
If P.MP < 5 Then Exit Function
P.Ani = P_Blast
P.AniL = 5
Select Case P.T
Case 0 'Classic

Case 1 'Poopsmith
    For I = 1 To 50
        If S(I).Act = False Then
            S(I).X = IIf(P.Dir = P_Left, P.X, P.X + 27)
            S(I).Y = P.Y - 3
            S(I).XS = IIf(P.Dir = P_Left, -(Rnd * 2.5 + 1), (Rnd * 2.5 + 1))
            S(I).YS = -(Rnd * 2.5 + 1)
            S(I).T = WP_Poo
            S(I).Act = True
            S(I).Bounce = 2
            P.Reload = 6
            P.MP = P.MP - 5
            Exit For
        End If
    Next I
End Select
End Function

Function pPunch(Pl As Player)
Dim opID As Long
If Pl.Reload > 0 Then Exit Function
If Pl.MP < 2 Then Exit Function
opID = IIf(Pl.ID = 1, 2, 1)
Pl.Ani = P_Punch
Pl.AniL = 10
Pl.Reload = 10
Dim X1, X2, Y1, Y2
X1 = Pl.X
X2 = P(opID).X
Y1 = Pl.Y
Y2 = P(opID).Y
If X2 > X1 And X2 < X1 + Play_W And Y2 > Y1 And Y2 < Y1 + Play_H Or X1 > X2 And X1 < X2 + Play_W And Y1 > Y2 And Y1 < Y2 + Play_H Then
    P(opID).HP = P(opID).HP - 2
End If
End Function

Function pKick(Pl As Player)
Dim opID As Long
If Pl.Reload > 0 Then Exit Function
If Pl.Reload < 5 Then Exit Function
opID = IIf(Pl.ID = 1, 2, 1)
Pl.Ani = P_Kick
Pl.Reload = 15
Pl.AniL = 15
Dim X1, X2, Y1, Y2
X1 = Pl.X
X2 = P(opID).X
Y1 = Pl.Y
Y2 = P(opID).Y
If X2 > X1 And X2 < X1 + Play_W And Y2 > Y1 And Y2 < Y1 + Play_H Or X1 > X2 And X1 < X2 + Play_W And Y1 > Y2 And Y1 < Y2 + Play_H Then
    P(opID).HP = P(opID).HP - 5
End If
End Function

Function GetKeys(P As Player, J, L, R, Pn, K, B, G, C)
If GetAsyncKeyState(J) Then
    pJump P
End If
If GetAsyncKeyState(L) Then
    pWalk P, P_Left
End If
If GetAsyncKeyState(R) Then
    pWalk P, P_Right
End If
If GetAsyncKeyState(Pn) Then
    pPunch P
End If
If GetAsyncKeyState(Pn) Then
    pPunch P
End If
If GetAsyncKeyState(K) Then
    pKick P
End If
If GetAsyncKeyState(B) Then
    pBlast P
End If
End Function

Function pWalk(P As Player, D As Long)
P.OnG = False
P.XS = IIf(D = P_Right, 2, -2)
P.Dir = D
If P.Jp = True Then Exit Function
P.Ani = IIf(P.Ani = P_Stand, P_Walk, P_Stand)
P.AniL = 2
End Function

Function pJump(P As Player)
If P.Act = True And P.OnG = True Then
    P.Jp = True
    P.OnG = False
    P.Ani = P_Up
    P.YS = -3.5
End If
End Function

Function NewPlayer(P As Player, X, Y, ID, T)
P.Reload = 0
P.X = X
P.Y = Y
P.XS = 0
P.YS = 0
P.Ani = P_Stand
P.AniL = 0
P.Dir = P_Right
P.OnG = False
P.ID = ID
P.HP = 50
P.T = T
P.Act = True
End Function

Function MovePlayers()
Dim I As Long, Sle As Long
For I = 1 To 2
    If P(I).Act = True Then
        P(I).Reload = P(I).Reload - 1
        If P(I).Reload < 0 Then P(I).Reload = 0
        P(I).MP = P(I).MP + 0.5: If P(I).MP > 80 Then P(I).MP = 80
        P(I).AniL = P(I).AniL - 1
        If P(I).AniL <= 0 Then
            P(I).Ani = P_Stand
            P(I).AniL = 5
        End If
        If P(I).Jp = True Then P(I).Ani = P_Up
        If P(I).YS < -0.5 Then P(I).Ani = P_Up
        If P(I).YS > 2 Then P(I).Ani = P_Fall
        If P(I).OnG = False And GetPixel(frmGame.bMask.hDC, P(I).X + Play_W / 2, P(I).Y) = vbBlack Then
            P(I).Y = P(I).Y + 1
            P(I).YS = -P(I).YS
        End If
        If P(I).OnG = False And GetPixel(frmGame.bMask.hDC, P(I).X + Play_W / 2, P(I).Y + Play_H) = vbBlack Then
            P(I).Jp = False
            P(I).Y = P(I).Y - 1
            P(I).YS = 0
            P(I).OnG = True
        End If
        If GetPixel(frmGame.bMask.hDC, P(I).X + P(I).XS + Play_W * 0.8, P(I).Y + Play_H / 2) = vbBlack Then
            P(I).XS = 0
        End If
        If GetPixel(frmGame.bMask.hDC, P(I).X + P(I).XS + Play_W * 0.2, P(I).Y + Play_H / 2) = vbBlack Then
            P(I).XS = 0
        End If
        If P(I).X + P(I).XS + (Play_W * 0.75) > 320 Then P(I).XS = 0
        If P(I).X + P(I).XS < -(Play_W / 4) Then P(I).XS = 0
        P(I).X = P(I).X + P(I).XS
        If P(I).OnG = False Then
            P(I).YS = P(I).YS + 0.1
        End If
        P(I).Y = P(I).Y + P(I).YS
        If P(I).OnG = True Then
            P(I).XS = P(I).XS * 0.75
        End If
        BitBlt frmGame.Board.hDC, P(I).X, P(I).Y, Play_W, Play_H, frmGame.SM(P(I).T).hDC, P(I).Dir, P(I).Ani, vbSrcAnd
        BitBlt frmGame.Board.hDC, P(I).X, P(I).Y, Play_W, Play_H, frmGame.SS(P(I).T).hDC, P(I).Dir, P(I).Ani, vbSrcInvert
        frmGame.pHP(I - 1).Cls
        frmGame.pHP(I - 1).Line (0, 0)-(frmGame.pHP(I - 1).ScaleWidth * (P(I).HP / 50), frmGame.pHP(I - 1).Height), &H80FF80, BF
        frmGame.pEN(I - 1).Cls
        frmGame.pEN(I - 1).Line (0, 0)-(frmGame.pEN(I - 1).ScaleWidth * (P(I).MP / 80), frmGame.pEN(I - 1).Height), &HFFC0FF, BF
    End If
Next I
End Function
