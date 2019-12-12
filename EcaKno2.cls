VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EcaKno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim dynarray$(300), samp$(300), trigr$(300), npt$(300), a5$(300) ',mask$(300)
Dim sc(300) As Integer
Dim config As String, chromo As String, sample As String, trigger As String
Public ynput As Integer
Dim compdata As String
Dim compdat(24) As String
Dim cdatk As Integer, proc_config As Integer, lnmsk As Integer
Dim lnaa As Integer, lnsamp As Integer, lntrig As Integer, lnin As Integer
Dim ibit As Boolean, sbit As Boolean, rbit As Boolean
Dim catype As Integer, alreadydone(300) As Integer
' Dim result As String
Dim kount_trigger As Integer, ktr As Integer
Dim trigbit As Boolean, prevtim As Date, hightime As Date
Dim worst As Single, mutrat As Integer
Public led As Integer
Public best As Single
Dim scoreprop(8000) As Integer
Dim badmask As String, goodmask As String
Dim scoretype As Boolean, bytescore As Single, bitscore As Single
Dim scorproplim As Integer, bestmut As Integer
Dim i As Integer, numberofinputarrays
' copyright Guy Giesbrecht april 20, 1998
Dim curiter As Integer
Public expout As Double
Public current1st As Integer, current2nd As Integer, current3rd As Integer
Public current4th As Integer, currentnext As Integer

Public Sub run(updata As String)
Call parse(updata)
ynput = nput
Next t

' fylname$ = InputBox("load pop", , "Yes")
' If fylname$ = "Yes" Then Call lodr Else Call makepop
Call makepop

Debug.Print "main"
Debug.Print Len(compdata)

youknow$ = "copyright Guy Giesbrecht april 20,1998"

' If Len(updata) < 128 Then Call makepop

' Call bioconfig

Call evolv(compdata)
' pass in data, invoke evolver
End Sub

Public Sub continue(expout As Double, ynput As Double)
Debug.Print "continue"
Call evolv(expout)
' called after initialization and first run by ducerone
' on each consecutive generation. returns control to ducerone
' for time check. other than that, basically the same as main

End Sub

Public Sub makepop()
Debug.Print "makepop"

Rem establish random population
Randomize Time
For i = 1 To 300
' main chromo string
For t = 1 To lnaa
r% = Int(Rnd * 256)
e$ = Chr(r%)
dynarray$(i) = dynarray$(i) + e$
Next t
' skip these parts if arrays are fixed.
' sample type
For t = 1 To lnsamp
r = Int((Rnd * 20) / 10)
eb$ = LTrim(Str(r))
samp$(i) = samp$(i) + eb$
Next t
' trigger sequence
For t = 1 To lntrig
r = Int((Rnd * 20) / 10)
ec$ = LTrim(Str(r))
trigr$(i) = trigr$(i) + ec$
Next t

' input type
For t = 1 To lnin
r = Int((Rnd * 20) / 10)
ec$ = LTrim(Str(r))
npt$(i) = npt$(i) + ec$
Next t
'For t = 1 To lnmsk
'ec$ = LTrim(Str(Int(Rnd * 2)))
'mask(i) = mask(i) + ec$
'Next t
a5$(i) = dynarray$(i) + samp$(i) + trigr$(i) + npt$(i) ' + mask(i)
Debug.Print a5$(i)
' config scheme needs writting
Next i


End Sub

Public Sub bioconfig()
Debug.Print "config"

proc_config = 0

' use config data to arrange execution of CA

End Sub
Public Sub parse(d2p As String)

Debug.Print "parse"
cfgdat$ = d2p
' bytes 1-4 length of components
lnaa = Asc(Mid(cfgdat, 1, 1))        ' active array
lnsamp = Asc(Mid(cfgdat, 2, 1))     ' sample
lntrig = Asc(Mid(cfgdat, 3, 1)) * 2   ' trigger
lnin = Asc(Mid(cfgdat, 4, 1))       ' input
'lnmsk = Asc(Mid(cfgdat, 5, 1))
' parse a population of bit strings from a string
' byte array
' config, chromo, sample, trigger and input
' part of config is the ratio or size for next 4

End Sub
Public Sub parse2()
For t = 1 To 300
' dynarray$ samp$, trigr$, npt$
dynarray$(t) = Left(a5$(t), lnaa)
samp$(t) = Mid(a5$(t), lnaa + 1, lnsamp)
trigr$(t) = Mid(a5$(t), lnaa + lnsamp + 2, lntrig)
npt$(t) = Mid(a5$(t), lnaa + lnsamp + lntrig + 3, lnin)
'mask$(t) = Right(a5$(t), lnmsk)
Next t
'ReDim compdat(Len(compdata) / 20)

End Sub
Public Function doneyet(limit As Single) As Boolean
doneyet = (best >= limit)

End Function
Public Sub evolv(expo As Double)
worst = 0#: best = 8000#
Dim amr As Integer
Debug.Print Len(compdata)
Debug.Print "evolv"
' data arrays need specifying
current1st = 0
current2nd = 1
current3rd = 2
current4th = 3
currentnext = 4
begn:
begcal:
best = 8000
current1st = current1st + 1
If current1st > numberofinputarrays Then current1st = 1
current2nd = current2nd + 1
If current2nd > numberofinputarrays Then current2nd = 1
current3rd = current3rd + 1
If current3rd > numberofinputarrays Then current3rd = 1
current4th = current4th + 1
If current4th > numberofinputarrays Then current4th = 1
currentnext = currentnext + 1
If currentnext > numberofinputarrays Then currentnext = 1: geniter = geniter + 1
Randomize best
' led = Int(Rnd * (lnaa + lnsamp + lntrig + lnin)) + 1 ' led is bit to mutate
' led = led + 1
' If led > (lnaa + lnsamp + lntrig + lnin) Then led = 1
mutrat = 0
curiter = curiter + 1
If curiter > 24 Then curiter = 1
For i = 1 To 300
'If alreadydone(i) > 0 Then GoTo skipr
If Len(dynarray$(i)) = 0 Then Stop
' operations loop for each string (see scheme)
sc(i) = process2(i)
If sc(i) = 0 Then Debug.Print "hoozah!!!!!": Call savr: Do While True: Beep: Loop
' also scoring routine for each string
If sc(i) < best Then hightime = Time
If sc(i) <= best Then best = sc(i): goodmask = a5$(i)
If sc(i) > worst Then worst = sc(i): badmask = a5$(i)
scoreprop(sc(i)) = scoreprop(sc(i)) + 1
'If sc(i) = 2440 Then a5$(i) = goodmask 'mutate(a5$(i)): sc(i) = sc(i) - 4
'If scoreprop(sc(i)) > 500 Then a5$(i) = mutate(a5$(i)): sc(i) = sc(i) - 5: alreadydone(i) = alreadydone(i) + 1
  
' If sc(i) = best And i / 2 = Int(i / 2) And scoreprop(best) > 1500 Then a5$(i) = mutate(a5$(i)): sc(i) = sc(i) - 2
If sc(i) <= best + 5 Then mutrat = mutrat + 1
skipr:
alreadydone(i) = alreadydone(i) + 1

'If mutrat > 80 And sc(p) = best Then sc(p) = sc(p) - 10
Next i
' Form1.Label1.Caption = Str(best)
Randomize Time
' competition routine
' Debug.Print "tourney . . ."
Call tourney
' Beep

Debug.Print geniter; " "; currentnext
If current1st / 5 = Int(current1st / 5) Then Call savr

End Sub

Public Sub savr()
' INPUT "file name"; fi$: fi$ = fi$ + ".dat"
a5$(0) = goodmask: sc(0) = best
thepop$ = ""
For tzl = 0 To 300
thepop$ = thepop$ + a5$(tzl)
Next tzl
Open "kno2pop3.dat" For Output As #1
'For z = 0 To 300
Write #1, thepop$ 'a5$(z) ,sc(z)
'Next z
'Write #1, "0", 0#
'Write #1, lnaa, lnsamp, lntrig, lnin
Close
End Sub
Public Function nsavr() As String
a5$(0) = goodmask: sc(0) = best
thepop$ = ""
For tzl = 0 To 300
thepop$ = thepop$ + a5$(tzl)
Next tzl
nsavr = thepop$

End Function
Public Sub lodr()
' INPUT "file name"; fi$: fi$ = fi$ + ".dat"
Open "kno2pop1.dat" For Input As #1
For z = 1 To 300
Input #1, a5$(z), sc(z)
Debug.Print a5$(z)
Next z
Input #1, extra$, extr#
Input #1, sep$, sp#
Input #1, lnaa, lnsamp, lntrig, lnin
Close
Call parse2
End Sub

Public Sub breed(a5$())
Debug.Print "breed"
Dim flag(150), f$(15), m$(5)
' change for variable size strings
' proportional splicing when pop has various sizes
' and multi-parent offspring
k = 152 ' popsize

For q = 1 To 15
f$(q) = a5$(q)
Next
For q = 1 To 5
m$(q) = a5$(q + 15)
Next
For ma = 1 To 5
For fe = 1 To 15

fi$ = f$(fe)
mi$ = m$(ma)

If fi$ = mi$ Then fi$ = a5$(Int(Rnd * 10) + 20)
If Len(fi$) = 0 Or Len(mi$) = 0 Then Stop
l% = Int(Rnd * Len(mi$) + 1): r% = Abs(l% - Len(mi$))
l1% = Int(Rnd * Len(fi$) + 1): r1% = Abs(l1% - Len(fi$))
o$ = Left$(fi$, l%) + Right$(mi$, r%)
p$ = Left$(mi$, l1%) + Right$(fi$, r1%)
a5$(k) = o$
a5$(k - 1) = p$
Debug.Print p$; k - 1, o$; k
k = k + 2
brnxt: Next fe, ma
For t = 135 To 300
md = Rnd * 100
If md > 20 Then
   Debug.Print t;
   o$ = a5$(t)
   a5$(t) = mutate(o$)
   End If
Next t
Call parse2
End Sub

Public Function mutate(o$)
Dim pled As Integer
result$ = ""
Randomize best

Dim rid As Integer

led = led + 1
If led > Len(o$) Then led = 1
 pled = led ' store led for after tml loop
'nb2m = 10 - Int(Rnd * 14)
'If nb2m < 1 Then nb2m = 1
'For tml = 1 To nb2m
'If tml > 1 Then led = Int(Rnd * Len(o$)) + 1
'led = Int(Rnd * Len(o$)) + 1

rid = Len(o$) - led
If led < lnaa Then GoTo chrmut
If Mid$(o$, led, 1) = "1" Then mu$ = "0" Else mu$ = "1"
If led = Len(o$) Then mutat$ = Left(o$, led - 1) + mu$ Else mutat$ = Left$(o$, led) + mu$ + Right$(o$, rid - 1)
'Next tml
GoTo doner
chrmut:
nuchar = (Asc(Mid(o$, led, 1)) + 1)
If nuchar > 255 Then nuchar = 1
mutat$ = Left$(o$, led) + Chr(nuchar) + Right$(o$, rid - 1)
If mutat$ = "" Then Stop
Debug.Print "mut"; led
doner:
mutate = mutat$
End Function

Sub sort()
Debug.Print "sort"
 gap = Int(150 / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For Index = 1 To 150 - gap
     If sc(Index) < sc(Index + gap) Then
     tmpsc = sc(Index)
     sc(Index) = sc(Index + gap)
     sc(Index + gap) = tmpsc
     tmpb$ = a5$(Index)
     a5$(Index) = a5$(Index + gap)
     a5$(Index + gap) = tmpb$
     
     ' SWAP bsc(Index), bsc(Index + gap)
     ' SWAP b$(Index), b$(Index + gap)
      doneflag = 0
     End If
    Next Index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
Beep
Call parse2
End Sub

Sub synth(dynarray$())
' synthesizer program emulates predictor for next set of numbers

End Sub

Public Sub tourney()
Dim mwinr As Integer, fwinr As Integer
'mutrat = mutrat / 3
Randomize Time
Debug.Print "tourn"
Dim flag(300) As Integer
For q = 1 To 300: flag(q) = 0: Next
If mutrat > 90 Then mutrat = 90
If mutrat < 24 Then mutrat = 24
For q = 1 To 75
pick1: r1 = Int(Rnd * 300) + 1: If flag(r1) = 1 Then GoTo pick1
flag(r1) = 1
pick2: r2 = Int(Rnd * 300) + 1: If flag(r2) = 1 Then GoTo pick2
flag(r2) = 1
pick3: r3 = Int(Rnd * 300) + 1: If flag(r3) = 1 Then GoTo pick3
flag(r3) = 1
pick4: r4 = Int(Rnd * 300) + 1: If flag(r4) = 1 Then GoTo pick4

flag(r4) = 1
If sc(r1) < sc(r2) Then
ma$ = a5$(r1): pk1 = r2: mwinr = r1
Else: ma$ = a5$(r2): pk1 = r1: mwinr = r2
End If

If sc(r3) < sc(r4) Then
fe$ = a5$(r3): pk2 = r4: fwinr = r3
Else: fe$ = a5$(r4): pk2 = r3: fwinr = r4
End If

If fe$ = ma$ Then fe$ = mutate(fe$)
If Len(fe$) = 0 Or Len(ma$) = 0 Then Stop
If Rnd < 0.25 Then
Debug.Print "e/o";
o$ = "": p$ = ""
   For ch = 1 To Len(ma$) Step 2
    
     o$ = o$ + Mid(ma$, ch, 1) + Mid(fe$, ch + 1, 1)
     p$ = p$ + Mid(fe$, ch, 1) + Mid(ma$, ch + 1, 1)
    
   Next ch
Else
le = Int(Rnd * Len(ma$) + 1): ri = Abs(le - Len(ma$))
lez = Int(Rnd * Len(fe$) + 1): riz = Abs(lez - Len(fe$))
o$ = Left$(fe$, le) + Right$(ma$, ri)
p$ = Left$(ma$, lez) + Right$(fe$, riz)
End If
If o$ = p$ Then mutate (p$)
md = Int(Rnd * 100)
If md < mutrat Then o$ = mutate(o$): p$ = mutate(p$)

a5$(pk1) = o$
a5$(pk2) = p$
alreadydone(pk1) = 0
alreadydone(pk2) = 0
If sc(mwinr) = sc(fwinr) Then a5$(mwinr) = mutate(a5$(mwinr)): alreadydone(mwinr) = 0
 If alreadydone(mwinr) > 1 Then
    a5$(mwinr) = mutate(a5$(mwinr)): alreadydone(mwinr) = 0
 End If
 If alreadydone(fwinr) > 1 Then
    a5$(fwinr) = mutate(a5$(fwinr)): alreadydone(fwinr) = 0
 End If

Next q
Debug.Print best; " mutrat"; mutrat; "  "; Time; "  prev:"; prevtim; "   last high"; hightime
prevtim = Time
' For z = 500 To 1300
' Debug.Print z; "="; scoreprop(z); "  |";
' Next z

Call parse2
End Sub
Public Function runtimeconfig()

End Function
Public Function process2(p As Integer) As Double
Dim mask(4) As String, indat(2) As String
Dim result(4) As String
Dim skor As Integer
res$ = ""
mask(1) = Left(dynarray$(p), 20)
mask(2) = Mid(dynarray$(p), 20, 20)
mask(3) = Mid(dynarray$(p), 40, 20)
'mask(4) = Mid(dynarray$(p), 60, 20)
chdata = handlinput()
indat(1) = Left(chdata, 20)
indat(2) = Right(chdata, 20)
Dim kbyt As Integer, mbyt As Integer, rbyt As Integer
trigk = 0
For t = 1 To 2
For z = 1 To Len(indat(1))
kbyt = Asc(Mid(indat(1), z, 1))
mbyt = Asc(Mid(mask(1), z, 1))
trigk = trigk + 2
catp = bits2int(Mid(trigr(p), trigk, 2)) 'TriggerCheck(z, t)
Select Case catp
Case 1
rbyt = kbyt Xor mbyt
Case 2
rbyt = kbyt Or mbyt
Case 3
rbyt = kbyt And mbyt
Case 4
rbyt = kbyt Eqv mbyt
End Select
result(t) = result(t) + Chr(rbyt)
Next z
Next t
For z = 1 To Len(result$(1))
kbyt = Asc(Mid(result(1), z, 1))
mbyt = Asc(Mid(result(2), z, 1))
trigk = trigk + 2
catp = bits2int(Mid(trigr(p), trigk, 2)) 'TriggerCheck(z, t)
Select Case catp
Case 1
rbyt = kbyt Xor mbyt
Case 2
rbyt = kbyt Or mbyt
Case 3
rbyt = kbyt And mbyt
Case 4
rbyt = kbyt Eqv mbyt
End Select
result(3) = result(3) + Chr(rbyt)
Next z

For z = 1 To Len(result(3))
kbyt = Asc(Mid(result(3), z, 1))
mbyt = Asc(Mid(mask(3), z, 1))
trigk = trigk + 2
catp = bits2int(Mid(trigr(p), trigk, 2))
Select Case catp
Case 1
rbyt = kbyt Xor mbyt
Case 2
rbyt = kbyt Or mbyt
Case 3
rbyt = kbyt And mbyt
Case 4
rbyt = kbyt Eqv mbyt
End Select
result(4) = result(4) + Chr(rbyt)
Next z

process2 = skor
End Function
Public Function process(p As Integer, compdata As String) As Single
' Beep
Dim skore As Single
cdatk = 1
Dim kiter As Integer

' Debug.Print "process"; p
' this part uses proc_config initialized in config() to construct the
' proper individual to be processed
'proc_config = 0
'Select Case proc_config
'Case 0
'ind$ = dynarray$(p) ' just the active array
'Case 1
'ind$ = dynarray$(p) + samp$(p) ' active array and sampler
'Case 2
'ind$ = dynarray$(p) + samp$(p) + trigr$(p) ' aa, sampler and trigger
'Case 3
'ind$ = a5$(p) ' aa, sampler, trigger and input types
' Case etc. other combos , segments, sliders here
'End Select

dind$ = ind$
' Debug.Print "*"; ind$
' process strings here according to rules and types
' several types of binary processes each one rotated
' thru by trigger: current_rule(kount_trigger)
' input mechanism? Expansive CA rules?
' config regulators thru this sub
' shiftdbit set here or in trigger?
kiter = 0
trig$ = trigr(p)

' kount_trigger = 2
Do While cdatk < Len(compdata)
kount_trigger = TriggerCheck(ind$, trig$)
kiter = kiter + 1
forclogin$ = ""
For tzl = 1 To 20
forclogin$ = forclogin$ + Right(bytes2bits(Mid(compdata, cdatk + tzl, 1)), 7) 'Mid(npt(p), 4, Len(npt(p)) - 4)
Next tzl
cdatk = cdatk + 20
'Dim bt24s As Integer
'bt24s = bits2int(Left(npt(p), 3)) + 1
ind$ = handlinput(ind$, forclogin$, kiter)


Select Case kount_trigger
Case 0
' Debug.Print " 0and"; ind$
' rule type rshift & xor
inres$ = and0(ind$)  ' pass indiv string to function
' and get back intermediate result

Case 1
' Debug.Print " 1or"; ind$
' rule type lshift & xor
inres$ = or1(ind$)
Case 2
' Debug.Print "2nab"; ind$

inres$ = nabors2(ind$)
Case 3
' Debug.Print "3not"; ind$
' rshift and equality
inres$ = not3(ind$)
Case 4
' Debug.Print "4xor"; ind$
' xor neighbors either each side of current bit
inres$ = xor4(ind$)

Case 5
' Debug.Print " 5eq"; ind$
inres$ = eq5(ind$)
Case 6
' Debug.Print "6nh"; ind$
'inres$ = naborhood(ind$)
inres$ = rshift(ind$)
Case 7
' Debug.Print "7exp"; ind$
'inres$ = expansive(ind$)
inres$ = lshift(ind$)
End Select

'If proc_config = 2 Or proc_config = 3 Then trig$ = Mid(inres$, lnaa + lnsamp + 2, lntrig)
'If proc_config > 0 Then insamp$ = Mid(inres$, lnaa + 1, lnsamp) Else insamp$ = samp(p)
 insamp$ = samp(p)
'If kiter / 2 <> Int(kiter / 2) Then GoTo sckipper
'res$ = Left(inres$, lnaa)
sampl$ = sampler(inres$, p, insamp$)
skore = skore + score(sampl$, Mid(compdata, cdatk, 20))
ind$ = inres$
'sckipper:
Loop
Debug.Print p; " "; Right(dind$, 70); " "; skore

' Debug.Print "               ########### "; skore; " ###############"

process = skore
End Function

Public Function handlinput() As String
' make ynput a string of bits?
handlinput = int2bits(ynput)

End Function

Public Function TriggerCheck(iter As String, bytnum As String) As Integer
numz = bits2int(Left(iter, bits2int(bytnum)))

TriggerCheck = Int(numz)
End Function

Public Function sampler(indi$, p As Integer, insamp$) As String

' config regulators here too
' samp$() defines sample config or mask
' skip using samp(p) if fixed type sampler. determined in config
ksamp = 0
result$ = ""
If indi$ = String(Len(indi$), "0") Then indi$ = String(Len(indi$) - 1, "0") + "1"
Do While Len(result$) < 140
For t = 1 To Len(indi$)
sampspot = sampspot + 1
If sampspot > Len(indi$) Then sampspot = ksamp
ksamp = ksamp + 1
If ksamp > Len(insamp$) Then ksamp = 1
If Mid(insamp$, ksamp, 1) = "1" Then result$ = result$ + Mid(indi$, t, 1)
Next t
Loop
trunclen = 140
result$ = Left(result$, trunclen)
sampler = result$
Exit Function
End Function
Public Function score(sampl$, expout As String) As Single
Dim skor As Integer, pick As Integer, xout As Integer, xbyt As Integer
xbyt = 1
For ts = 1 To 140 Step 7
pick = bits2int(Mid(sampl$, ts, 7))
If pick > 80 Then pick = pick - 80
xout = Asc(Mid(expout, xbyt, 1))
xbyt = xbyt + 1
skor = skor + Abs(pick - xout)
dif$ = LTrim(Str(xout)) + "#" + LTrim(Str(pick)) + "."
'Debug.Print dif$;
Next ts
'Debug.Print
score = skor
' Debug.Print cdatk
End Function
Public Function getsomebytes(numbytes As Integer, compdata As String) As String

thebytes = Mid(compdata, cdatk, numbytes)
cdatk = cdatk + numbytes
' Debug.Print cdatk; "#"
getsomebytes = thebytes
End Function
Public Function bytes2bits(byts2con$) As String

' convert some bytes to a string of 1's and 0's
Dim thisbyte As Integer, cbytes As Integer, tbits As Integer
Dim thebits As String
bits$ = ""
For cbytes = 1 To Len(byts2con$)
thisbyte = Asc(Mid(byts2con, cbytes, 1))
thebits = ""
ysum = 0
For tbits = 7 To 0 Step -1
ysum = ysum + 2 ^ tbits
If ysum > thisbyte Then thebits = "0" + thebits Else thebits = "1" + thebits
Next tbits
bits$ = thebits + bits$
Next cbytes
bytes2bits = bits$

End Function
Public Sub int2bits(int2con As Integer)
If int2con > 256 Then Stop
thebits = ""
ysum = 0
For tbits = 7 To 0 Step -1
ysum = ysum + 2 ^ tbits
If ysum > int2con Then thebits = "0" + thebits Else thebits = "1" + thebits
Next tbits
bits$ = thebits + bits$

End Sub
Public Function bits2int(bits2con As String) As Integer
resint = 0
bitnum = Len(bits2con)
' Debug.Print "b2i"; bits2con
' convert a string of bits to int
For t = 1 To Len(bits2con)
bitnum = bitnum - 1
If Mid(bits2con, t, 1) = "1" Then resint = resint + 2 ^ bitnum
Next t
bits2int = resint
End Function
Public Function bits2longer(bits2con As String) As Double
Dim reslong As Double
reslong = 0#
bitnum = Len(bits2con)
' Debug.Print "b2i"; bits2con
' convert a string of bits to int
For t = 1 To Len(bits2con)
bitnum = bitnum - 1
If Mid(bits2con, t, 1) = "1" Then reslong = reslong + 2 ^ bitnum
Next t
bits2longer = reslong
End Function
Public Function bits2double(bits2con As String) As Double
Debug.Print Len(bits2con)
If Len(bits2con) <> 64 Then bits2double = 0#: Exit Function

sign$ = Left(bits2con, 1) ' sign bit
bits2con = Right(bits2con, Len(bits2con) - 1)
expon$ = Left(bits2con, 11) ' exponents
mant$ = Right(bits2con, Len(bits2con) - 11) ' mantissa
Dim resdub As Double
If sign$ = "1" Then thesign = 1 Else thesign = -1
exponent = bits2int(expon$)
mantissa = 1 / bits2longer(mant$)
bits2double = thesign * (mantissa * 10 ^ exponent)



End Function

Public Function rshift(indi$) As String
' Debug.Print " RI#"; indi$
rindiv$ = Right(indi$, 1) + Left(indi$, Len(indi$) - 1)
' rightshifts indiv
rshift = rindiv$
End Function

Public Function lshift(indi$) As String
' Debug.Print " LE#"; indi$
lindiv$ = Right(indi$, Len(indi$) - 1) + Left(indi$, 1)
' leftshifts indiv
lshift = lindiv$
End Function
Public Function shiftdbit(ind2s$) As String
bits2$ = Right(ind2s$, 2)
Select Case bits2$
Case "00"
result$ = rshift(ind2s$)
Case "01"
result$ = rshift(rshift(ind2s$))
Case "10"
result$ = lshift(ind2s$)
Case "11"
result$ = lshift(lshift(ind2s$))
End Select
shiftdbit = result$
End Function

Public Function and0(indiv$)
' Debug.Print "and0*"; indiv$
Dim result As String
result = ""
Dim ibit As Boolean, sbit As Boolean, rbit As Boolean
sind$ = rshift(indiv$)

For t = 1 To Len(indiv$)
ibit = (Mid(indiv$, t, 1) = "1")
sbit = (Mid(sind$, t, 1) = "1")
rbit = (ibit And sbit)

If rbit Then result = result + "1" Else result = result + "0"
Next t
and0 = result
End Function

Public Function or1(indiv$)
result$ = ""
' Debug.Print "or1*"; indiv$
sind$ = shiftdbit(indiv$)
For t = 1 To Len(indiv$)
ibit = (Mid(indiv$, t, 1) = "1")
sbit = (Mid(sind$, t, 1) = "1")
rbit = (ibit Or sbit)
If rbit Then result$ = result$ + "1" Else result$ = result$ + "0"
Next t
or1 = result$
End Function

Public Function nabors2(indiv$) As String
' Debug.Print "nabors2"; indiv$
result$ = ""
Dim bit1 As Boolean, bit2 As Boolean, rbit As Boolean
For t = 1 To Len(indiv$)
ct = t
Select Case ct
Case Is < 3
n2bit$ = Mid(indiv$, t + 1, 2)
If t = 1 Then p2bit$ = Right(indiv$, 2) Else p2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Is > Len(indiv$) - 2
p2bit$ = Mid(indiv$, t - 2, 2)
If t = Len(indiv$) Then n2bit$ = Left(indiv$, 2) Else n2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Else
n2bit$ = Mid(indiv$, t + 1, 2)
p2bit$ = Mid(indiv$, t - 2, 2)
End Select

Dim numkey As Integer
'mask$ = Left(indiv$, 16) + Right(indiv$, 16)
For ml = 64 To 128 Step 2
 mask$ = mask$ + Mid(indiv$, ml, 1)
Next ml
Key$ = p2bit$ + Mid(indiv$, t, 1) + n2bit$
numkey = bits2int(Key$) + 1
result$ = result$ + Mid(mask$, numkey, 1)
Next t
nabors2 = result$
Exit Function


' old nabors function
bit1 = (Left(n2bit$, 1) = 1): bit2 = (Right(n2bit$, 1) = 1)
Select Case p2bit$
Case "00"
rbit = bit1 Xor bit2
Case "10"
rbit = (bit1 = bit2)
Case "01"
rbit = bit1 Or bit2
Case "11"
rbit = bit1 And bit2
End Select
If rbit Then result$ = result$ + "1" Else result$ = result$ + "0"
' Next t
nabors2 = result$


End Function

Public Function not3(indiv$) As String
' Debug.Print "not3"; indiv$
For t = 1 To Len(indiv$)
ibit = (Mid(indiv$, t, 1) = "1")
If (Not ibit) Then result$ = result$ + "1" Else result$ = result$ + "0"
Next t
not3 = result$
End Function

Public Function xor4(indiv$)
' *** not a finalised shiftdbit function !!!!!!
' Debug.Print "xor4"; indiv$
' Debug.Print "ashf"; sind$
sind$ = shiftdbit(indiv$)
For t = 1 To Len(indiv$)
ibit = (Mid(indiv$, t, 1) = "1")
sbit = (Mid(sind$, t, 1) = "1")
 rbit = (ibit Xor sbit)
' Debug.Print "ibit"; ibit; " xor sbit "; sbit; " result$ "; rbit
If rbit Then result$ = result$ + "1" Else result$ = result$ + "0"
                      Next t
' Debug.Print "axor"; result$
xor4 = result$
End Function

Public Function eq5(indiv$) As String
' Debug.Print "..eq5"; indiv$
sind$ = shiftdbit(indiv$)
For t = 1 To Len(indiv$)
ibit = (Mid(indiv$, t, 1) = "1")
sbit = (Mid(sind$, t, 1) = "1")
 rbit = (ibit = sbit)
If rbit Then result$ = result$ + "1" Else result$ = result$ + "0"
Next t
eq5 = result$
End Function

Public Function naborhood(indiv$) As String

result$ = ""
' noffset established in parse from primary ca string?
For t = 1 To Len(indiv$)
ct = t
Select Case ct
Case Is < 3
n2bit$ = Mid(indiv$, t + 1, 2)
If t = 1 Then p2bit$ = Right(indiv$, 2) Else p2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Is > Len(indiv$) - 2
p2bit$ = Mid(indiv$, t - 2, 2)
If t = Len(indiv$) Then n2bit$ = Left(indiv$, 2) Else n2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Else
n2bit$ = Mid(indiv$, t + 1, 2)
p2bit$ = Mid(indiv$, t - 2, 2)
End Select
bit1$ = Right(p2bit$, 1): bit2$ = Left(n2bit$, 1)
c2bit$ = bit1$ + bit2$
Select Case c2bit$
Case "00"
rebit$ = Left(p2bit$, 1)
Case "01"
rebit$ = bit1$
Case "10"
rebit$ = bit2$
Case "11"
rebit$ = Right(n2bit$, 1)
End Select
result$ = result$ + rebit$

Next t
naborhood = result$

End Function

Public Function expansive(indiv$) As String
' Debug.Print "expando"; indiv$
result$ = ""
For t = 1 To Len(indiv$) Step 2
ct = t
Select Case ct
Case Is < 3
n2bit$ = Mid(indiv$, t + 1, 2)
If t = 1 Then p2bit$ = Right(indiv$, 2) Else p2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Is > Len(indiv$) - 2
p2bit$ = Mid(indiv$, t - 2, 2)
If t = Len(indiv$) Then n2bit$ = Left(indiv$, 2) Else n2bit$ = Right(indiv$, 1) + Left(indiv$, 1)

Case Else
n2bit$ = Mid(indiv$, t + 1, 2)
p2bit$ = Mid(indiv$, t - 2, 2)
End Select

If Mid(indiv$, t, 1) = "1" Then result$ = result$ + p2bit$ Else result$ = result$ + n2bit$

Next t

expansive = result$
End Function

Public Function getbestpop() As String
For i = 1 To 300
pop$ = pop$ + convert(a5$(i)) ' converts bitstring to bytes
Next i
getbestpop = pop$
End Function

Public Function convert(stocon As String) As String
Dim thisbit As Integer, byt As Integer
For t = 1 To Len(stocon) Step 8
byt = 0
bt$ = Mid(stocon, t, 8)
thisbit = 8
For t2 = 7 To 0 Step -1
If Mid(bt$, thisbit, 1) = 1 Then byt = byt + 2 ^ t2
thisbit = thisbit - 1
Next t2
converted$ = converted$ + Chr(byt)
Next t
convert = converted$
End Function

Private Sub Class_Initialize()
' setup stuff
End Sub

Private Sub Class_Terminate()
' clean up stuff
End Sub
