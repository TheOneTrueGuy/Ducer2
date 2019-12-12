VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Evoca1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim dynarray$(600), samp$(600), trigr$(600), npt$(600), a5$(600)
Dim sc(600) As Single
Dim config As String, chromo As String, sample As String, trigger As String
Dim ynput As String
Dim cdatk As Integer, proc_config As Integer
Dim lnaa As Integer, lnsamp As Integer, lntrig As Integer, lnin As Integer
Dim ibit As Boolean, sbit As Boolean, rbit As Boolean
Dim catype As Integer, alreadydone(600) As Integer
Dim kount_trigger As Integer, ktr As Integer
Dim trigbit As Boolean, prevtim As Date, hightime As Date
Dim worst As Single, mutrat As Integer
Dim scoreprop(1600) As Integer
Dim badmask As String, goodmask As String
Dim scoretype As Boolean, bytescore As Single, bitscore As Single
Dim scorproplim As Integer, bestmut As Integer
Dim i As Integer
Public updata As String, compdata As String
Public Event checkform()
Public led As Integer
Public best As Single, toppop$
' copyright Guy Giesbrecht april 20, 1999
Public Sub Main(compdat As String, popul As String)
compdata = compdat
Debug.Print "main"
Debug.Print Len(compdata)
cdatk = 1
youknow$ = "copyright Guy Giesbrecht april 20,1998"
lnaa = 128: lnsamp = 32
lntrig = 32: lnin = 32
If popul = "" Then makepop
Call evolv(compdata)
' pass in data, invoke evolver
End Sub


Public Sub makepop()
Debug.Print "makepop"
Rem establish random population
Randomize Time
For i = 1 To 600
' main chromo string
For t = 1 To lnaa
r% = Int((Rnd * 20) / 10)
e$ = LTrim(Str(r%))
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
a5$(i) = dynarray$(i) + samp$(i) + trigr$(i) + npt$(i)
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
cfgdat = Left(d2p, 20): d2p3$ = Right(d2p, Len(d2p) - 20)
' bytes 1-4 length of components
lnaa = Asc(Mid(cfgdat, 1, 1))       ' active array
lnsamp = Asc(Mid(cfgdat, 2, 1))     ' sample
lntrig = Asc(Mid(cfgdat, 3, 1))     ' trigger
lnin = Asc(Mid(cfgdat, 4, 1))       ' input

' parse a population of bit strings from a string
' byte array
' config, chromo, sample, trigger and input
' part of config is the ratio or size for next 4

End Sub
Public Sub parse2()
For t = 1 To 600
' dynarray$ samp$, trigr$, npt$
dynarray$(t) = Left(a5$(t), lnaa)
samp$(t) = Mid(a5$(t), lnaa + 1, lnsamp)
trigr$(t) = Mid(a5$(t), lnaa + lnsamp + 2, lntrig)
npt$(t) = Right(a5$(t), lnin)
Next t


End Sub
Public Function doneyet(limit As Single) As Boolean
doneyet = (best >= limit)

End Function
Public Sub evolv(compdata As String)
worst = 800#: best = 0#
scoretype = True
Debug.Print Len(compdata)
Debug.Print "evolv"
' data arrays need specifying
begn:
begcal:
Randomize best
' led = Int(Rnd * (lnaa + lnsamp + lntrig + lnin)) + 1 ' led is bit to mutate
' led = led + 1
' If led > (lnaa + lnsamp + lntrig + lnin) Then led = 1
mutrat = 0

' scoretype = Not scoretype
If scoretype Then
    bytescore = 1: bitscore = 0.0001
Else: bytescore = 0.001: bitscore = 1
End If
bitscore = 1
bytescore = 0

For i = 1 To 600
If alreadydone(i) > 0 Then GoTo skipr
If Len(dynarray$(i)) = 0 Then Stop
' operations loop for each string (see scheme)
sc(i) = process(i, compdata)
If sc(i) >= 1800 Then Debug.Print "hoozah!!!!!": Call savr: Stop
' also scoring routine for each string
If sc(i) > best Then hightime = Time
If sc(i) >= best Then best = sc(i): goodmask = a5$(i): toppop$ = a5$(i)
If sc(i) < worst Then worst = sc(i): badmask = a5$(i)
scoreprop(sc(i)) = scoreprop(sc(i)) + 1
'If sc(i) = 1568 Then a5$(i) = mutate(a5$(i)): sc(i) = sc(i) - 4
 If scoreprop(sc(i)) > 300 And Rnd < 0.95 Then a5$(i) = mutate(a5$(i)): sc(i) = sc(i) - 5 ' alreadydone(i) = alreadydone(i) + 1
  
' If sc(i) = best And i / 2 = Int(i / 2) And scoreprop(best) > 3000 Then a5$(i) = mutate(a5$(i)): sc(i) = sc(i) - 2
    

If sc(i) >= best - 2 Then mutrat = mutrat + 1
skipr:
alreadydone(i) = alreadydone(i) + 1

'If mutrat > 80 And sc(p) = best Then sc(p) = sc(p) - 10
RaiseEvent checkform
Next i
' Form1.Label1.Caption = Str(best)
Randomize Time
' competition routine
' Debug.Print "tourney . . ."
Call tourney
Beep
End Sub

Public Sub savr()
' INPUT "file name"; fi$: fi$ = fi$ + ".dat"
a5$(0) = goodmask: sc(0) = best
Open "ntpop.dat" For Output As #1
For z = 0 To 600
Write #1, a5$(z), sc(z)
Next z
Write #1, "0", 0#
Write #1, lnaa, lnsamp, lntrig, lnin
Close
End Sub

Public Sub lodr()
' INPUT "file name"; fi$: fi$ = fi$ + ".dat"
Open "ntpop.dat" For Input As #1
For z = 1 To 600
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
Dim flag(300), f$(15), m$(5)
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
For t = 135 To 600
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
If Rnd < 0.09 Then
    mutate = rshift(o$): Debug.Print "rmut"
     Exit Function
End If
Dim rid As Integer

led = led + 1
If led > Len(o$) Then led = 1
 pled = led ' store led for after tml loop
nb2m = 10 - Int(Rnd * 14)
If nb2m < 1 Then nb2m = 1
For tml = 1 To nb2m
If tml > 1 Then led = Int(Rnd * Len(o$)) + 1
'led = Int(Rnd * Len(o$)) + 1

rid = Len(o$) - led
If Mid$(o$, led, 1) = "1" Then mu$ = "0" Else mu$ = "1"
If led = Len(o$) Then mutat$ = Left(o$, led - 1) + mu$ Else mutat$ = Left$(o$, led) + mu$ + Right$(o$, rid - 1)
o$ = mutat$
Next tml
led = pled
If mutat$ = "" Then Stop
 
 If Rnd > 0.92 Then
   For trn = 1 To Len(mutat$)
   If Rnd > 0.75 Then GoTo same
   If Mid(goodmask, trn, 1) <> Mid(mutat$, trn, 1) Then GoTo invert Else GoTo same
invert:
   If Mid(mutat$, trn, 1) = "0" Then result$ = result$ + "1" Else result$ = result$ + "0"
   GoTo doner
same:
   result$ = result$ + Mid(mutat$, trn, 1)
doner:
   Next trn
   mutat$ = result$
   
End If
 If Rnd > 0.96 Then mutat$ = not3(mutat$)
Debug.Print "mut"; led
mutate = mutat$
End Function

Sub sort()
Debug.Print "sort"
 gap = Int(300 / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For Index = 1 To 300 - gap
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
mutrat = mutrat / 4
Randomize Time
Debug.Print "tourn"
Dim flag(600) As Integer
For q = 1 To 600

flag(q) = 0: Next
If mutrat > 90 Then mutrat = 90
If mutrat < 4 Then mutrat = 4
For q = 1 To 150
pick1: r1 = Int(Rnd * 600) + 1: If flag(r1) = 1 Then GoTo pick1
flag(r1) = 1
pick2: r2 = Int(Rnd * 600) + 1: If flag(r2) = 1 Then GoTo pick2
flag(r2) = 1
pick3: r3 = Int(Rnd * 600) + 1: If flag(r3) = 1 Then GoTo pick3
flag(r3) = 1
pick4: r4 = Int(Rnd * 600) + 1: If flag(r4) = 1 Then GoTo pick4

flag(r4) = 1
If sc(r1) > sc(r2) Then
ma$ = a5$(r1): pk1 = r2: mwinr = r1
Else: ma$ = a5$(r2): pk1 = r1: mwinr = r2
End If

If sc(r3) > sc(r4) Then
fe$ = a5$(r3): pk2 = r4: fwinr = r3
Else: fe$ = a5$(r4): pk2 = r3: fwinr = r4
End If

If fe$ = ma$ Then fe$ = mutate(fe$)
If Len(fe$) = 0 Or Len(ma$) = 0 Then Stop
If Rnd < 0.05 Then
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
 If alreadydone(fwinr) > 2 Then
    a5$(fwinr) = mutate(a5$(fwinr)): alreadydone(fwinr) = 0
 End If

Next q
Debug.Print best; " mutrat"; mutrat; "  "; Time; "  prev:"; prevtim; "   last high"; hightime
prevtim = Time
' For z = 500 To 1600
' Debug.Print z; "="; scoreprop(z); "  |";
' Next z

Call parse2
End Sub
Public Function runtimeconfig()

End Function
Public Function process(p As Integer, compdata As String) As Single
' Beep
Dim skore As Single
cdatk = 1


Dim kiter As Long
' Debug.Print "process"; p
' this part uses proc_config initialized in config() to construct the
' proper individual to be processed
proc_config = 0
Select Case proc_config
Case 0
ind$ = dynarray$(p) ' just the active array
Case 1
ind$ = dynarray$(p) + samp$(p) ' active array and sampler
Case 2
ind$ = dynarray$(p) + samp$(p) + trigr$(p) ' aa, sampler and trigger
Case 3
ind$ = a5$(p) ' aa, sampler, trigger and input types
' Case etc. other combos , segments, sliders here
End Select
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
Do While cdatk < 200 'Len(compdata) ' loop until samples done
kount_trigger = TriggerCheck(ind$, trig$)
forclogin$ = Mid(npt(p), 4, Len(npt(p)) - 4)
Dim bt24s As Integer
bt24s = bits2int(Left(npt(p), 3)) + 1
ind$ = handlinput(ind$, forclogin$, bt24s, kiter)

' hows about applying this to a segment of the string?
' reiteratively, in blocks etc.
kiter = kiter + 1

 
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
' not
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
inres$ = naborhood(ind$)
Case 7
' Debug.Print "7exp"; ind$
inres$ = expansive(ind$)


End Select

' parse inres for components

res$ = Left(inres$, lnaa)
If proc_config = 2 Or proc_config = 3 Then trig$ = Mid(inres$, lnaa + lnsamp + 2, lntrig)
If proc_config > 0 Then insamp$ = Mid(inres$, lnaa + 1, lnsamp) Else insamp$ = samp(p)
sampl$ = sampler(res$, p, insamp$, compdata)
skore = skore + score(sampl$, compdata)

ind$ = inres$ ' replace previous ind with this iter.
' or not depending on config etc
Loop

Debug.Print p; " "; Left(dind$, 90); " "; skore

' Debug.Print "               ########### "; skore; " ###############"

process = skore
End Function

Public Function handlinput(indiv$, forlog$, bits2forc As Integer, priter As Long) As String
result$ = ""
'If Left(forlog$, 1) = "1" Then GoTo forced_logic Else GoTo inputer
' inputer handles input ambiguities
forced_logic:
If Mid(forlog$, 2, 1) = "1" Then GoTo triggered Else GoTo cycle
triggered:
 If Mid(forlog$, 3, 1) = "1" Then GoTo temporal_trig Else GoTo samp_trig
temporal_trig:
itertrig = bits2int(Mid(forlog$, 3, 3))
invins$ = Mid(forlog$, 7, 1)
phrase2handl$ = Right(forlog$, Len(forlog$) - 7)
If priter / itertrig = Int(priter / itertrig) Then
  GoTo cycle
Else: handlinput = indiv$
Exit Function
End If
 
samp_trig:
samploc = Len(indiv$) * (1 / (bits2int(Mid(forlog$, 4, 5)) + 1))
samplen = bits2int(Mid(forlog$, 10, 3)) + 1
samptrig$ = Mid(indi$, samploc, samplen)
trigphrase$ = Mid(forlog$, 14, samplen)
invins$ = Mid(forlog$, 15, 1)
phrase2handl$ = Right(forlog$, Len(forlog$) - (15 + samplen))
If samptrig$ = trigphrase$ Then
 GoTo cycle
Else: handlinput = indiv$
Exit Function
End If
cycle:
If phrase2handl$ = "" Then phrase2handl = Right(forlog$, Len(forlog$) - 4)
If invins$ = "" Then invins$ = Mid(forlog$, 3, 1)
lin = Len(indiv$)
For t = 1 To lin
phk = phk + 1
If phk > Len(phrase2handl$) Then phk = 1
If t / bits2forc = Int(t / bits2forc) Then GoTo forcbit Else GoTo skipit
forcbit:
   If invins$ = "1" Then GoTo invert Else GoTo insert
invert:
   If Mid(indiv$, t, 1) = "1" Then result$ = result$ + "0" Else result$ = result$ + "1"
    GoTo skipskipit
insert:
   result$ = result$ + Mid(phrase2handl$, phk, 1)
   GoTo skipskipit
   
skipit:
result$ = result$ + Mid(indiv$, t, 1)

skipskipit:
Next t
handlinput = result$
End Function

Public Function TriggerCheck(ind$, trigger$) As Integer
' trigr$(p) is trigger sequence
' also test to toggle shiftdbit
' skip using trig(p) if fixed type trigger
Dim trg As Integer, ktr As Integer

' trigconfig$ = Left(trigr$(p), 4)
' trigfig = bits2int(trigconfig$)
' ktrig$ = Mid(ind$, Int(Len(ind$) / 2), 3)
' TriggerCheck = bits2int(ktrig$)
' Exit Function
' ktrig$ = Mid(ind$, Int(Len(ind$) / 2), 3)
'triglen = bits2int(Left(trigger$, 3)) + 1
trig2$ = trigger$ 'Right(trigger$, Len(trigger$) - 3)
'Dim actrig(8) As String
'For t = 1 To Len(trigger$) Step 3
ktr = ktr + 3
If ktr > Len(trigger$) Then ktr = 1
ktrig$ = Mid(trigger$, ktr, 3)
'If ktr > 8 Then ktr = 1
'actrig(ktr) = Mid(trig2$, t, 3)
'Next t
Select Case ktrig$
Case "000" 'actrig(1)
trg = 0
Case "001" 'actrig(2)
trg = 1
Case "010" 'actrig(3)
trg = 2
Case "011" 'actrig(4)
trg = 3
Case "100" 'actrig(5)
trg = 4
Case "101" 'actrig(6)
trg = 5
Case "110" 'actrig(7)
trg = 6
Case "111" 'actrig(8)
trg = 7
Case Else
trg = kount_trigger
End Select
' Debug.Print trg;
TriggerCheck = trg
End Function

Public Function sampler(indi$, p As Integer, insamp$, compdata As String) As String

' config regulators here too
' samp$() defines sample config or mask
' skip using samp(p) if fixed type sampler. determined in config
ksamp = 0
result$ = ""
For t = 1 To Len(indi$)
ksamp = ksamp + 1
If ksamp > Len(insamp$) Then ksamp = 1
If Mid(insamp$, ksamp, 1) = "1" Then result$ = result$ + Mid(indi$, t, 1)
Next t
trunclen = Int(Len(result$) / 8) * 8
If trunclen = 0 Then
   Do While Len(result$) < 8
   result$ = "0" + result$: Loop
  trunclen = 8
End If
result$ = Left(result$, trunclen)
sampler = result$
Exit Function
' old score method
Dim samplace As Integer, samplen As Integer, sampsize As Integer
samplace = Int(Len(indi$) * (1 / (0.5 / (1 + bits2int(Right(insamp$, 3)))) + (bits2int(Left(insamp$, 3)))))
' ratio sampler?
samplen = bits2int(Mid(insamp$, Len(insamp$) - 11, 3))
sampsize = bits2int(Mid(insamp$, Len(insamp$) - 16, 3)) + 1 ' + (bits2int(Mid(insamp$, Len(insamp$) - 22, 2)) + 1)
' bits from distinct locations
Do Until samplace + (sampsize * 8) < Len(indi$)
 samplace = samplace - sampsize * 8
Loop
If samplace < 1 Then samplace = 1
samptype$ = Mid(indi$, Len(indi$) * 0.25, 2)
samptype$ = "00"
Select Case samptype$
Case "00"
sampbyte$ = Mid(indi$, samplace, sampsize * 8)
'
Case "01"
result$ = ""
For t2 = 1 To Len(indi$) Step 2
result$ = result$ + Mid(indi$, t2, 1)
Next t2
sampbyte$ = Left(result$, sampsize * 8)
'
Case "10"
result$ = ""
For t2 = 2 To Len(indi$) Step 2
result$ = result$ + Mid(indi$, t2, 1)
Next t2
sampbyte$ = Left(result$, sampsize * 8)
'
Case "11"
result$ = ""
For t1 = 1 To Len(indi$) Step 16
result$ = result$ + Mid(indi$, t1, 8)
Next t1
sampbyte$ = Left(result$, sampsize * 8)
End Select
sampler = sampbyte$
Stop
End Function
Public Function score(sampl$, compdata As String) As Single
Dim skor As Single
Dim sampsize As Integer
sampsize = Len(sampl$) / 8

eosamp$ = getsomebytes(sampsize, compdata)
If Len(eosamp$) * 8 < Len(sampl$) Then sampl$ = Left(sampl$, Len(eosamp$) * 8)

somebyte$ = bytes2bits(eosamp$)
' Debug.Print Len(somebyte$), Len(sampbyte$), Len(eosamp$), Len(indi$)

If Len(somebyte$) <> Len(sampl$) Then Stop
' asx = "somebytes" + Str(Len(somebyte$)) + " sampbytes" + Str(Len(sampbyte$))
' asxi = InputBox(asx)

For t = 1 To Len(somebyte$)
If Mid(somebyte$, t, 1) = Mid(sampl$, t, 1) Then skor = skor + bitscore
Next t

For t = 1 To Len(somebyte$) Step 8
If Mid(somebyte$, t, 8) = Mid(sampl$, t, 8) Then skor = skor + bytescore
Next t

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
mask$ = Left(indiv$, 16) + Right(indiv$, 16)
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
For i = 1 To 600
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
