Attribute VB_Name = "Engine"
Option Explicit
Const poptop As Integer = 399
Dim pop(400) As Double
Dim scor(400) As Long
Dim tim1, tim2
Private samplecount As Integer ' total count of samples /8 to get groups
Dim allSamples(12, 20) As Integer
Dim excludedoubles(80) As Boolean
Dim psiCount(80) As Long
Public bestscore As Integer
Dim lastsam(400) As sample
Dim precStats(20, 80) As Long
Dim bestout As String
Public oLn(12, 80) As Integer
Public killloop As Boolean

Public Sub makePop()
killloop = False
Dim tzl, qzl
For tzl = 0 To 12
For qzl = 0 To 80
oLn(tzl, qzl) = 3
Next qzl
Next tzl
qzl = tzl
'Double
'(double-precision floating-point) 8 bytes -1.79769313486232E308 to
'-4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values
'
Dim d As Double, popsize
popsize = poptop
Randomize Timer

For tzl = 0 To popsize
d = (Rnd * (1.79769313486232 ^ Int(Rnd * 101)))
If Rnd < 0.5 Then d = -1 * d
If d = 0# Then d = (Rnd * (1.79769313486232 ^ Int(Rnd * 101)))
pop(tzl) = d
Next tzl
bestscore = 10000
End Sub

Public Function runEngine(numgens As Integer, typ As Integer) As Boolean
Dim tzl, stl As Integer, qzl As Integer
Dim tym1, tym2
    bestscore = 10000
    Erase precStats
    Erase psiCount
    tym1 = Timer
For tzl = 1 To numgens
   Select Case typ
   Case 1
   runPop2
   Case 2
   runPop
   End Select
   
    If tzl / 2 = Int(tzl / 2) Then Call breed3 Else Call Breed4
    Next tzl
For stl = 0 To 80
   ' Form2.LS2 stl, psiCount(stl)
Next stl
'Form2.Chartit2

Dim most(20) As Integer
Dim best(20) As Long
Dim topten(20) As Integer, ttBest, ttz
Dim lastbest, nextbest, bestit
'Dim least
'Dim worst(20) As Long
Dim bst As String

For qzl = 0 To 19
best(qzl) = 70
For stl = 1 To 80
    If precStats(qzl, stl) > best(qzl) Then most(qzl) = stl: best(qzl) = precStats(qzl, stl)
  '  Form1.loadGrid3 qzl + 1, stl, precStats(qzl, stl)
    
   ' Form1.colorGrid3 qzl + 1, stl, RGB(precStats(qzl, stl) * 1.3, precStats(qzl, stl) * 1.3, precStats(qzl, stl) * 1.3)
  '  If precStats(qzl, stl) * 1.3 > 255 Then Form1.colorGrid3 qzl + 1, stl, RGB(255, 0, 0)
Next stl
Next qzl

For stl = 1 To 80
    If psiCount(stl) > ttBest Then ttBest = psiCount(stl): bestit = stl
Next stl ' determines best
topten(0) = bestit
lastbest = ttBest

For qzl = 1 To 19
    ttBest = 0
    For stl = 1 To 80
        If psiCount(stl) > ttBest And psiCount(stl) < lastbest Then
        ttBest = psiCount(stl)
        topten(qzl) = stl
        End If
    Next stl
   
    lastbest = ttBest
Next qzl

Dim top10 As String
For qzl = 0 To 19
If qzl = 10 Then top10 = top10 & "##"
top10 = top10 & CStr(topten(qzl)) & ":"
Next qzl
'Form2.loadList2 "top:" & top10
Select Case typ
Case 1
KenoStats.loadList "O:" & top10
Case 2
KenoStats.loadList "C:" & top10
End Select
Erase precStats
'Debug.Print "The mostest bestest:"

'Dim sl1, sl2, tempy
'For sl1 = 0 To 19
'For sl2 = 0 To 19
'If most(sl2) > most(sl1) Then tempy = most(sl2): most(sl2) = most(sl1): most(sl1) = tempy
'Next sl2
'Next sl1

'For qzl = 0 To 19
''Debug.Print "::"; most(qzl);
'bst = bst & ": " & most(qzl)
'Next qzl
'bst = bst & " best score;" & bestscore & " bestout:" & bestout
'Form2.loadLabel1 bst
KenoStats.loadTime CStr(Timer - tym1)

runEngine = True
End Function

Public Sub loadsampl(sampl As Integer, ndex As Integer, samcount As Integer)
Dim tzl, ut As String
'For tzl = 0 To 19
allSamples(samcount, ndex) = sampl
oLn(samcount, sampl) = 1
samplecount = samcount
End Sub
Public Sub testit()
Dim tzl, ut As String
For tzl = 0 To 80
If oLn(samplecount, tzl) = 1 Then ut = ut & "T" Else ut = ut & "F"

Next tzl
KenoStats.loadList ut
End Sub

Sub sort()
'Debug.Print "sort"
Dim gap, doneflag, Index, tempop As Double, temscor As Integer
 gap = Int(poptop / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For Index = 1 To 200 - gap
     If scor(Index) > scor(Index + gap) Then
     tempop = pop(Index): temscor = scor(Index)
     pop(Index) = pop(Index + gap): scor(Index) = scor(Index + gap)
     pop(Index + gap) = tempop: scor(Index + gap) = temscor

      doneflag = 0
     End If
    Next Index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
'Debug.Print pop(2).scor; "CCC"; pop(3).scor
'Debug.Print
End Sub

'Public Sub Breed()
'sort
''Rnd (0)
''Randomize
'Dim mdna As Double
'Dim fdna As Double
'Dim mrna As String
'Dim frna As String
'
'Dim newpop(2000) As double
'Dim tested(2000) As Boolean
'Dim pk1 As Integer, pk2 As Integer, pk3 As Integer, pk4 As Integer
'Dim melsloop, femsloop
'Dim pick1 As double, pick2 As double, pick3 As double, pick4 As double
'Dim mel As double, fem As double, offcount As Integer, tzl
'
'For melsloop = 1 To 20
'    mel = pop(melsloop)
'    For femsloop = 21 To 45
'        fem = pop(femsloop)
'        mdna = mel.codex
'        fdna = fem.codex
'    If mdna = fdna Then fem = pop(Int(Rnd * 100) + 50): fdna = fem.codex
'    Dim melsign, femsign
'    If mdna < 0 Then melsign = -1 Else melsign = 1
'    If fdna < 0 Then femsign = -1 Else femsign = 1 ' resolves misplacement of sign
'        mrna = CStr(Abs(mdna))
'        frna = CStr(Abs(fdna))
'        If mrna = 0 Or frna = 0 Then Stop
'        ' now gotta parse the . E +
'        Dim mgd As String, fgd As String
'        Dim mre, fre, mex
'        Dim fex, ogd, ore, oex
'        ' {male, female, offspring} greatest-digit, remainder and exponent
'        Dim moff As double, foff As double, dexter1, dexter2, dex1a, dex2a
'        Dim xover As Integer, codx As String
'
'
'
'
'        dexter1 = InStr(mrna, "E")
'        dexter2 = InStr(frna, "E")
'        If dexter1 <> 0 Then
'           If Left(mrna, 1) <> "-" Then mgd = Left(mrna, 1) Else mgd = Left(mrna, 2)
'            mre = Mid(mrna, 3, dexter1 - 2)
'            mex = Right(mrna, Abs(dexter1 - Len(mrna)))
'            Else ' male codex has no exponent
'            mgd = Left(mrna, 1)
'            mre = Right(mrna, Len(mrna) - 2)
'            mex = "+00"
'        End If
'
'        If dexter2 <> 0 Then
'           If Left(frna, 1) <> "-" Then fgd = Left(frna, 1) Else fgd = Left(frna, 2)
'            fre = Mid(frna, 3, dexter2 - 2)
'            fex = Right(frna, Abs(dexter2 - Len(frna)))
'            Else ' female codex has no exponent
'           If Left(frna, 1) <> "-" Then fgd = Left(frna, 1) Else fgd = Left(mrna, 2)
'            fre = Right(frna, Len(frna) - 2)
'            fex = "+00"
'        End If
'
'        ' double check start, crossover and endpoint figures in left,mid,right statements
'        ' and at concatenations
'        ' first child
'        ogd = mgd
'        xover = Len(mre) / 2 'Int(Rnd * (Len(mre) - 1)) + 3
'        ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'        oex = mex
'        If Left(ore, 1) <> "." Then codx = ogd & "." & ore & oex Else codx = ogd & ore & oex
'        If codx = 0 Then Stop
'        newpop(offcount).codex = CDbl(codx)
'        offcount = offcount + 1
'        '2nd child
'        ogd = fgd
'        xover = Len(mre) / 2 'Int(Rnd * (Len(mre) - 1)) + 3
'        ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'        oex = mex
'        If Left(ore, 1) <> "." Then codx = ogd & "." & ore & oex Else codx = ogd & ore & oex
'        If codx = 0 Then Stop
'        newpop(offcount).codex = CDbl(codx)
'        offcount = offcount + 1
'        '3rd child
'        ogd = mgd
'        xover = Len(mre) / 2 'Int(Rnd * (Len(mre) - 1)) + 3
'        ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'        oex = fex
'        If Left(ore, 1) <> "." Then codx = ogd & "." & ore & oex Else codx = ogd & ore & oex
'        If codx = 0 Then Stop
'        newpop(offcount).codex = CDbl(codx)
'        offcount = offcount + 1
'        '4th child
'        ogd = fgd
'        xover = Len(mre) / 2 'Int(Rnd * (Len(mre) - 1)) + 3
'        ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'        oex = fex
'        If Left(ore, 1) <> "." Then codx = ogd & "." & ore & oex Else codx = ogd & ore & oex
'        If codx = 0 Then Stop
'        newpop(offcount).codex = CDbl(codx)
'        offcount = offcount + 1
'    Next femsloop
'
'Next melsloop
'
'
'Dim mutrat
'For tzl = 0 To poptop
'    pop(tzl) = newpop(tzl)
'    If pop(tzl).codex = 0 Then Stop
'    mutrat = Rnd
'    If mutrat < 0.005 Then pop(tzl).codex = pop(tzl).codex + (5 - (Rnd * 10)): 'Debug.Print "mut";: 'Beep
'Next tzl
'
'
'End Sub
'
'Public Sub Breed2()
'Dim newpop(2000) As double
'Dim tested(2000) As Boolean
'Dim pk1 As Integer, pk2 As Integer, pk3 As Integer, pk4 As Integer
'Dim breedloop, pick1 As double, pick2 As double, pick3 As double, pick4 As double
'Dim mel As double, fem As double, offcount As Integer, tzl
'For breedloop = 1 To 500 ' 2 offspring for every winner
'repick1:
'pk1 = Int(Rnd * 2000)
'If tested(pk1) Then GoTo repick1
'pick1 = pop(pk1)
'repick2:
'pk2 = Int(Rnd * 2000)
'If tested(pk2) Then GoTo repick2
'pick2 = pop(pk2)
'repick3:
'pk3 = Int(Rnd * 2000)
'If tested(pk2) Then GoTo repick3
'pick3 = pop(pk3)
'repick4:
'pk4 = Int(Rnd * 2000)
'If tested(pk4) Then GoTo repick4
'pick4 = pop(pk4)
'tested(pk1) = True
'tested(pk2) = True
'tested(pk3) = True
'tested(pk4) = True
'If pick1.scor < pick2.scor Then
'mel = pick1
'Else
'mel = pick2
'End If
'If pick3.scor < pick4.scor Then
'fem = pick3
'Else
'fem = pick4
'End If
'Dim mdna As Double
'Dim fdna As Double
'Dim mrna As String
'Dim frna As String
'mdna = mel.codex
'fdna = fem.codex
'If mdna = fdna Then fem = pop(Int(Rnd * 100) + 50): fdna = fem.codex
'mrna = CStr(mdna)
'frna = CStr(fdna)
'' now gotta parse the . E +
'Dim mgd, fgd, mre, fre, mex, fex, ogd, ore, oex
'' {male, female, offspring} greatest-digit, remainder and exponent
'Dim moff As double, foff As double, dexter1, dexter2
'Dim xover As Integer, codx As String
'dexter1 = InStr(frna, "E")
'dexter2 = InStr(mrna, "E")
'If dexter2 <> 0 Then
'    mgd = Left(mrna, 1)
'    mre = Mid(mrna, 3, dexter2 - 2)
'    mex = Right(mrna, Abs(dexter2 - Len(mrna)))
'    Else ' male codex has no exponent
'    mgd = Left(mrna, 1)
'    mre = Right(mrna, Len(mrna) - 2)
'    mex = "+00"
'End If
'
'If dexter1 <> 0 Then
'    fgd = Left(frna, 1)
'    fre = Mid(frna, 3, dexter1 - 2)
'    fex = Right(frna, Abs(dexter1 - Len(frna)))
'    Else ' female codex has no exponent
'    fgd = Left(frna, 1)
'    fre = Right(frna, Len(frna) - 2)
'    fex = "+00"
'End If
'' double check start, crossover and endpoint figures in left,mid,right statements
'' and at concatenations
'
'' first child
'ogd = mgd
'xover = Int(Rnd * (Len(mre) - 1)) + 2
'ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'oex = mex
'codx = ogd & "." & ore & oex
'newpop(offcount).codex = CDbl(codx)
'offcount = offcount + 1
''2nd child
'ogd = fgd
'xover = Int(Rnd * (Len(mre) - 1)) + 2
'ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'oex = mex
'codx = ogd & "." & ore & oex
'newpop(offcount).codex = CDbl(codx)
'offcount = offcount + 1
''3rd child
'ogd = mgd
'xover = Int(Rnd * (Len(mre) - 1)) + 2
'ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'oex = fex
'codx = ogd & "." & ore & oex
'newpop(offcount).codex = CDbl(codx)
'offcount = offcount + 1
''4th child
'ogd = fgd
'xover = Int(Rnd * (Len(mre) - 1)) + 2
'ore = Left(mre, xover) + Right(fre, Len(mre) - xover)
'oex = fex
'codx = ogd & "." & ore & oex
'newpop(offcount).codex = CDbl(codx)
'offcount = offcount + 1
'Next breedloop
'Dim mutrat
'For tzl = poptop To 0 Step -1
'    pop(tzl) = newpop(revc)
'    revc = revc + 1
'    mutrat = Rnd
'    If mutrat < 0.04 Then pop(tzl).codex = pop(tzl).codex + (0.05 - (Rnd * 0.1)): 'Debug.Print "mut";: 'Beep
'Next tzl
'
'End Sub
Public Sub breed3()
sort
Dim mdna As Double
Dim fdna As Double
Dim mrna As String
Dim frna As String
Dim newpop(1000) As Double
Dim tested(1000) As Boolean
Dim pk1 As Integer, pk2 As Integer, pk3 As Integer, pk4 As Integer
Dim breedloop, pick1 As Double, pick2 As Double, pick3 As Double, pick4 As Double
Dim mel As Double, fem As Double, offcount As Integer, tzl
Dim melsloop, femsloop
For melsloop = 1 To 10
mel = pop(melsloop)
For femsloop = 11 To 35

fem = pop(femsloop)

mdna = mel
fdna = fem
If mdna = fdna Then fem = pop(Int(Rnd * 50) + 35): fdna = fem

mrna = CStr(mdna)
frna = CStr(fdna)
' now gotta parse the . E +
Dim mgd, fgd, mre, fre, mex, fex, ogd, ore, oex
' {male, female, offspring} greatest-digit, remainder and exponent
Dim moff As Double, foff As Double, dexter1, dexter2
Dim xover1 As Double, xover2 As Double, codx As Double
' 1st child
xover1 = (Rnd * 10)
xover2 = 10 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 10
If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
' 2nd
xover1 = (Rnd * 10)
xover2 = 10 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 10
If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
'3rd
xover1 = (Rnd * 10)
xover2 = 10 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 10
If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
'4th
xover1 = (Rnd * 10)
xover2 = 10 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 10
If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1

Next femsloop
Next melsloop
Dim mutrat
For tzl = 0 To poptop
    pop(tzl) = newpop(tzl)
    mutrat = Rnd
    If mutrat < 0.02 Then
    If Rnd < 0.5 Then pop(tzl) = pop(tzl) + ((10 * Rnd) ^ (Rnd * 20)) Else pop(tzl) = pop(tzl) - (Rnd ^ (Rnd * 10))
    End If
Next tzl
End Sub
Public Sub Breed4()
Dim newpop(2000) As Double
Dim tested(2000) As Boolean
Dim pk1 As Integer, pk2 As Integer, pk3 As Integer, pk4 As Integer
Dim breedloop, pick1 As Double, pick2 As Double, pick3 As Double, pick4 As Double
Dim mel As Double, fem As Double, offcount As Integer, tzl
Rnd
Randomize
For breedloop = 1 To 100 ' 2 offspring for every winner
Do
pk1 = Int(Rnd * 399) + 1
pick1 = pop(pk1)
Loop While tested(pk1)

Do
pk2 = Int(Rnd * 399) + 1
pick2 = pop(pk2)
Loop While tested(pk2)
Do
pk3 = Int(Rnd * 399) + 1
pick3 = pop(pk3)
Loop While tested(pk3)
Do
pk4 = Int(Rnd * 399) + 1
pick4 = pop(pk4)
Loop While tested(pk4)

tested(pk1) = True
tested(pk2) = True
tested(pk3) = True
tested(pk4) = True
If scor(pk1) < scor(pk2) Then
mel = pick1
Else
mel = pick2
End If
If scor(pk3) < scor(pk4) Then
fem = pick3
Else
fem = pick4
End If
Dim mdna As Double
Dim fdna As Double
Dim mrna As String
Dim frna As String
mdna = mel
fdna = fem
If mdna = fdna Then fem = pop(Int(Rnd * 50) + 50): fdna = fem

mrna = CStr(mdna)
frna = CStr(fdna)
' now gotta parse the . E +
Dim mgd, fgd, mre, fre, mex, fex, ogd, ore, oex
' {male, female, offspring} greatest-digit, remainder and exponent
Dim dexter1, dexter2
Dim xover1 As Double, xover2 As Double, codx As Double
' 1st child
xover1 = (Rnd * 20)
xover2 = 20 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 20
'If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
' 2nd
xover1 = (Rnd * 20)
xover2 = 20 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 20
'If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
'3rd
xover1 = (Rnd * 20)
xover2 = 20 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 20
'If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1
'4th
xover1 = (Rnd * 20)
xover2 = 20 - xover1
codx = ((mdna * xover1) + (fdna * xover2)) / 20
If codx = 0 Then Stop
newpop(offcount) = codx
offcount = offcount + 1

Next breedloop
Dim mutrat
For tzl = 0 To poptop
    pop(tzl) = newpop(tzl)
    mutrat = Rnd
    If mutrat < 0.001 Then
    If Rnd < 0.5 Then pop(tzl) = pop(tzl) + ((10 * Rnd) ^ (Rnd * 20)) Else pop(tzl) = pop(tzl) - (Rnd ^ (Rnd * 10))
    End If
Next tzl
End Sub

Public Sub runPop()
Dim tzl As Integer, samco As Integer, stl As Integer
Dim tsam, xyz
'ReDim psamp(samplecount) As sample ' population output sample

KenoStats.loadTime "Running..."
For tzl = 1 To poptop

'If pop(tzl) = 0 Then Stop
Rnd (-1)
Randomize pop(tzl) ' herein lies the holy of holys wherein the codex becomes transmogrified
'For samco = 0 To samplecount
Erase excludedoubles

For stl = 0 To 19
'
Do
tsam = Int(Rnd * 80) + 1 ' temp sample
Loop While excludedoubles(tsam)
excludedoubles(tsam) = True
xyz = allSamples(samplecount, stl) 'the list of inputs
If oLn(samplecount, tsam) Then scor(tzl) = scor(tzl) - 2

'If xyz = tsam Then scor(tzl) = scor(tzl) - 1

Next stl
   
'Next samco
Erase excludedoubles

For stl = 0 To 19
' and the residue of power is gathered here
Do
    tsam = Int(Rnd * 80) + 1 ' 1-80
   Loop While excludedoubles(tsam)
    lastsam(tzl) = tsam
    precStats(stl, tsam) = precStats(stl, tsam) + 1
    excludedoubles(tsam) = True
    psiCount(tsam) = psiCount(tsam) + 1
Next stl

Next tzl
End Sub
Public Sub runPop2()
Dim tzl As Integer, samco As Integer, stl As Integer
Dim tsam, xyz
'ReDim psamp(samplecount) As sample ' population output sample

KenoStats.loadTime "Running..."
For tzl = 1 To poptop

'If pop(tzl) = 0 Then Stop
Rnd (-1)
Randomize pop(tzl) ' herein lies the holy of holys wherein the codex becomes transmogrified
'For samco = 0 To samplecount
Erase excludedoubles

For stl = 0 To 19
'
Do
tsam = Int(Rnd * 80) + 1 ' temp sample
Loop While excludedoubles(tsam)
excludedoubles(tsam) = True
xyz = allSamples(samplecount, stl) 'the list of inputs
If oLn(samplecount, tsam) Then scor(tzl) = scor(tzl) - 2

If xyz = tsam Then scor(tzl) = scor(tzl) - 1

Next stl
   
'Next samco
Erase excludedoubles

For stl = 0 To 19
' and the residue of power is gathered here
Do
    tsam = Int(Rnd * 80) + 1 ' 1-80
   Loop While excludedoubles(tsam)
    lastsam(tzl) = tsam
    precStats(stl, tsam) = precStats(stl, tsam) + 1
    excludedoubles(tsam) = True
    psiCount(tsam) = psiCount(tsam) + 1
Next stl

'If scor(tzl) < bestscore Then
'    bestout = ""
'    Beep
'    bestscore = scor(tzl)
'    Debug.Print "NBS"; bestscore
'    Dim zzl
'    For zzl = 0 To 19
'    Debug.Print ":"; lastsam(tzl).sampl(zzl);
'    bestout = bestout & "| " & CStr(lastsam(tzl))
'    Next zzl
'    Debug.Print
'End If

'DoEvents
'KenoStats.loadTime tzl
Next tzl

End Sub
'
'Public Sub savepop(fylname As String)
'Dim tzl, whpop As wholePop
'Open fylname For Binary As #1
'
'For tzl = 0 To poptop - 1
'whpop.wpop(tzl - 1) = pop(tzl - 1)
'Next tzl
'Put 1, , whpop
'Close 1
'
'End Sub
'
'
'Public Function loadPop(fylname As String) As Boolean
'On Error GoTo erhandl
'Dim tzl, whpop As wholePop
'Open fylname For Binary As #1
'Get 1, , whpop
'Close 1
'For tzl = 0 To poptop
'pop(tzl) = whpop.wpop(tzl)
'Next tzl
'loadPop = True
'Exit Function
'erhandl:
'MsgBox "error" & Err.Description
'
'End Function

'Public Sub saveData(Fyl As String)
'Dim tzl, evs As evSamp
'For tzl = 0 To samplecount - 1
'evs.samples(tzl) = allSamples(tzl)
'
'Next tzl
'evs.sampcount = samplecount
'Open Fyl For Binary As #1
'
'Put 1, , evs
'Close 1
'
'End Sub
'Public Sub loadData(Fyl As String)
'Dim tzl, evs As evSamp
'Open Fyl For Binary As #1
'Get 1, , evs
'Close 1
'samplecount = evs.sampcount
'ReDim allSamples(samplecount)
'For tzl = 0 To samplecount - 1
'allSamples(tzl) = evs.samples(tzl)
'Next tzl
'Dim qzl
'For tzl = 0 To samplecount - 1
'For qzl = 0 To 19
'Debug.Print allSamples(tzl).sampl(qzl); " ";
'Next qzl
'Debug.Print
'Next tzl
'
'End Sub
Public Sub mutateAll()
For tzl = 0 To poptop
If Rnd < 0.5 Then pop(tzl) = pop(tzl) + ((10 * Rnd) ^ (Rnd * 20)) Else pop(tzl).codex = pop(tzl).codex - (Rnd ^ (Rnd * 10))
Next tzl
End Sub
Public Sub saturate()
Dim d As Double, best As Integer, scor, tym
tym = Timer
Dim ary(2000) As Double, smp, qzl
Randomize Timer
Dim tzl, dzl
Do While tzl < 400
scor = 0
d = (Rnd * (1.79769313486232 ^ Int(Rnd * 101)))
If Rnd < 0.5 Then d = -1 * d

Rnd (-1)
Randomize d
For qzl = 0 To samplecount
For dzl = 0 To 19
smp = Int(Rnd * 80) + 1
If allSamples(qzl).orderless_entry(smp) Then scor = scor - 1
Next dzl
Next qzl
If scor < -3 Then ary(tzl).codex = d: tzl = tzl + 1: ' Form1.tymer tzl 'Beep
'DoEvents
Loop
For tzl = 0 To 399
pop(tzl) = ary(tzl)
Next tzl
'Form1.tymer "r" & Timer - tym
'Debug.Print
End Sub
Public Sub scan(threshold As Integer, limit As Integer)
KenoStats.loadTime "scanning..."
'oLn(0, 0) = 1
'KenoStats.loadList CStr(oLn(0, 0))
Erase psiCount
Dim d As Double, best As Integer, scree, tym, stl
tym = Timer
Dim smp As Integer, qzl, top20(20) As Integer, ttBest
Dim sml, tsam, bestit, lastbest
Dim tzl As Integer, dzl
Rnd
Randomize
Dim escount
Do While tzl < limit
If killloop Then Exit Do
Randomize
scree = 0
escount = escount + 1
If escount > 10000 Then Exit Do
d = (Rnd * 1.79769313486232) ^ (Rnd * 100)
'KenoStats.loadList CStr(d)
If Rnd < 0.5 Then d = -1 * d
Rnd (-1)
Randomize d
For qzl = 0 To samplecount
For dzl = 0 To 19
smp = Int(Rnd * 80) + 1
'KenoStats.Text2.Text = CStr(smp)
'KenoStats.loadTime CStr(oLn(qzl, smp)) & ":" & CStr(qzl) & ":" & CStr(smp)
If oLn(qzl, smp) = 1 Then scree = scree - 1: ' KenoStats.Text4.Text = "XX"
Next dzl
Next qzl
'KenoStats.loadTime scree
If scree < -1 * threshold * (samplecount + 1) Then
Erase excludedoubles
tzl = tzl + 1
For stl = 0 To 19
' and the residue of power is gathered here
Do
    tsam = Int(Rnd * 80) + 1 ' 1-80
Loop While excludedoubles(tsam)
    excludedoubles(tsam) = True
    psiCount(tsam) = psiCount(tsam) + 1
Next stl
End If
Loop
ttBest = 0
For stl = 1 To 80
    If psiCount(stl) > ttBest Then ttBest = psiCount(stl): bestit = stl
Next stl ' determines best
top20(0) = bestit
lastbest = ttBest

For qzl = 1 To 19
    ttBest = 0
    For stl = 1 To 80
        If psiCount(stl) > ttBest And psiCount(stl) < lastbest Then
        ttBest = psiCount(stl)
        top20(qzl) = stl
        End If
    Next stl
    lastbest = ttBest
Next qzl
Dim topit

For qzl = 0 To 19
If qzl = 10 Then topit = topit & "##"
topit = topit & CStr(top20(qzl)) & "."
Next qzl
KenoStats.loadList "S:" & topit
KenoStats.loadTime Timer - tym

End Sub
Public Function givePopcodex(ndex As Integer) As Double
givePopcodex = pop(ndex)
End Function
Public Sub scan2()

End Sub
