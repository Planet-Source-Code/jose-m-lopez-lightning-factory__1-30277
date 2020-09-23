Attribute VB_Name = "mdlProcessing"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'Sleep 1 = Sleeps for one millisecond
Public booStop As Boolean        'request to stop playing textboxes
Public booStopped As Boolean     'status of Main routine
Public booHighlightOn As Boolean 'status of option buttons, highlight script as playing
Public intbooAngleInc As Integer 'status of option buttons, use angle increment
Public strFileName As String     'holds filename when saving script files
Const Pi = 3.14159265358979      'Pi

Public Sub rtnPlasma1()
Dim intLowerAngle As Integer        'user selected lower angle
Dim intUpperAngle As Integer        'user selected upper angle
Dim intDesiredRadius As Integer     'user selected radius
Dim intQtyFades As Integer          'user selected fade count
Dim intQtyRaysPerDisplay As Integer 'user selected ray count
Dim intQtySegments As Integer       'user selected segment count
Dim intXStart As Integer            'user selected ray x starting point
Dim intYStart As Integer            'user selected ray y starting point
Dim dblTimes As Long                'user selected display count
Dim intXRnd As Integer              'user selected linear x deviation
Dim intYRnd As Integer              'user selected linear y deviation
Dim intXRndHalf As Integer          'user selected linear x deviation/2
Dim intYRndHalf As Integer          'user selected linear y deviation/2
Dim intDrawWidth As Integer         'user selected drawwidth
Dim intSleepSegments                'user selected delay between segments in milliseconds
Dim intSleepRays                    'user selected delay between rays in milliseconds
Dim intColorInnerR As Integer, intColorInnerG As Integer, intColorInnerB As Integer 'user selected inner circle colors
Dim intColorOuterR As Integer, intColorOuterG As Integer, intColorOuterB As Integer 'user selected outer circle colors
Dim lngColorOuter As Long 'holds outer circle color

Dim intLinearXinc As Integer 'holds x axis random increments according to user selected lower/upper angle
Dim intLinearYinc As Integer 'holds y axis random increments according to user selected lower/upper angle
Dim intRndXinc As Integer 'holds x axis random increments according to user selected "x rnd dev"
Dim intRndYinc As Integer 'holds y axis random increments according to user selected "y rnd dev"
Dim intAngle As Integer 'used temporarily when checking user input angle limits
Dim intSegments As Integer 'holds segment count on for/next loop
Dim intRayPerDisplay As Integer 'holds ray count on for/next loop
Dim dblCount As Long 'used to compare to dblTimes for display count
Dim X As Integer 'used miscellaneously
Dim xx As Integer, yy As Integer 'hold X and Y axis coordinates
Dim xCoordinates() As Integer 'array holds X axis coordinates
Dim yCoordinates() As Integer 'array holds Y axis coordinates
Dim intFadeStack As Integer 'used temporarily when pushing the fade stack
Dim dblRadiuses() As Double 'array holds the radius for each ray point, used to determine if we should change color according to user selection
Dim dblTotalXinc As Double 'Holds the linear x rnd increment plus the angle x rnd increment, from a 0,0 center
Dim dblTotalYinc As Double 'Holds the linear y rnd increment plus the angle y rnd increment, from a 0,0 center
Dim dblRadius As Double 'holds the calculated radius by(Asqrd+Bsqrd=Csqrd) using dblTotalXinc and dblTotalYinc
Dim intFadeDecrementR As Integer, intFadeDecrementG As Integer, intFadeDecrementB As Integer 'holds the color fading

'Get user inputs from textboxes
'**************************
           intXStart = frmControls.txtParameters(0).Text 'ray x start
           intYStart = frmControls.txtParameters(1).Text 'ray y start
       intLowerAngle = frmControls.txtParameters(2).Text 'lower pie angle
       intUpperAngle = frmControls.txtParameters(3).Text 'upper pie angle
intQtyRaysPerDisplay = frmControls.txtParameters(4).Text 'rays in each display
      intQtySegments = frmControls.txtParameters(5).Text 'points in ray
         intQtyFades = frmControls.txtParameters(6).Text 'How many fade steps, 0 and negative = no fade and no clear
    intDesiredRadius = frmControls.txtParameters(7).Text 'radius from where we may alter the inside or outside color
             intXRnd = frmControls.txtParameters(8).Text 'randomness for ray points' x deviation
             intYRnd = frmControls.txtParameters(9).Text 'randomness for ray points' y deviation
        intXRndHalf = frmControls.txtParameters(10).Text 'if half of intXRnd then centers x deviation else some ray runaway
        intYRndHalf = frmControls.txtParameters(11).Text 'if half of intYRnd then centers y deviation else some ray runaway
       intDrawWidth = frmControls.txtParameters(12).Text 'ray width
     intColorInnerR = frmControls.txtParameters(13).Text 'color inside radius
     intColorInnerG = frmControls.txtParameters(14).Text 'color inside radius
     intColorInnerB = frmControls.txtParameters(15).Text 'color inside radius
     intColorOuterR = frmControls.txtParameters(16).Text 'color outside radius
     intColorOuterG = frmControls.txtParameters(17).Text 'color outside radius
     intColorOuterB = frmControls.txtParameters(18).Text 'color outside radius
           dblTimes = frmControls.txtParameters(19).Text 'How many displays, negative numbers loops always
       intSleepRays = frmControls.txtParameters(20).Text 'Delay between rays
   intSleepSegments = frmControls.txtParameters(21).Text 'Delay between segments
     lngColorOuter = RGB(intColorOuterR, intColorOuterG, intColorOuterB)
'**************************


'Limits
'********************************************
If intLowerAngle < 0 Then intLowerAngle = 0     'make sure angles within 360 degrees
If intLowerAngle > 360 Then intLowerAngle = 360
If intUpperAngle < 0 Then intUpperAngle = 0
If intUpperAngle > 360 Then intUpperAngle = 360
    If intLowerAngle > intUpperAngle Then 'make sure lower angle is smaller
    intAngle = intLowerAngle              'than upper, else swap
    intLowerAngle = intUpperAngle
    intUpperAngle = intAngle
    End If
    
    'if user selected drawwidth is positive then do it
    If intDrawWidth >= 1 Then frmScreen.DrawWidth = intDrawWidth '
    'if negative then clip at -15, cls to abs value and exit
    If intDrawWidth < -15 Then intDrawWidth = -15
        If intDrawWidth >= -15 And intDrawWidth <= 0 Then
        frmScreen.BackColor = QBColor(Abs(intDrawWidth))
        frmScreen.Cls
        Exit Sub
        End If
'********************************************

'*********************Redim****************************
            'with no fades selected
ReDim xCoordinates(intQtyRaysPerDisplay, 1, intQtySegments)
ReDim yCoordinates(intQtyRaysPerDisplay, 1, intQtySegments)
ReDim dblRadiuses(intQtyRaysPerDisplay, 1, intQtySegments)
If intQtyFades < 1 Then GoTo lblNoFades1 'No fades selected

'            with fades selected
intFadeDecrementR = (intColorInnerR / intQtyFades) - 1
intFadeDecrementG = (intColorInnerG / intQtyFades) - 1
intFadeDecrementB = (intColorInnerB / intQtyFades) - 1

ReDim xCoordinates(intQtyRaysPerDisplay, intQtyFades, intQtySegments)
ReDim yCoordinates(intQtyRaysPerDisplay, intQtyFades, intQtySegments)
ReDim dblRadiuses(intQtyRaysPerDisplay, intQtyFades, intQtySegments)
lblNoFades1:
'********************************************************************


'********Reset arrays to starting point************************
For X = 0 To intQtyRaysPerDisplay
For xx = 0 To intQtyFades
For yy = 0 To intQtySegments
xCoordinates(X, xx, yy) = intXStart
yCoordinates(X, xx, yy) = intYStart
Next yy
Next xx
Next X
'**************************************************************


'**************************************************************
'**************************************************************
'************************Main Routine**************************
'**************************************************************
'**************************************************************
dblCount = 0
frmScreen.ForeColor = RGB(intColorInnerR, intColorInnerG, intColorInnerB)

'**************************************************************
'************************Main Loop*****************************
'**************************************************************
Do While dblTimes <> dblCount
dblCount = dblCount + 1
    '**********************************************************
    '**********Make One Display At A Time Consisting***********
    '**********Of Desired Rays (Lightnings) Per Display********
    '**********************************************************
    For intRayPerDisplay = 1 To intQtyRaysPerDisplay
    If intQtyFades < 1 Then GoTo lblNoFades2 'If no fades selected skip
        '**********************************************************
        '*******First Fade Out Previous Rays From Stack Array******
        '**********************************************************
        For X = 1 To intQtyFades
        xx = intXStart: yy = intYStart
        'Blank Move
        frmScreen.Line (intXStart, intYStart)-(intXStart, intYStart), vbBlack
            '**********************************************************
            '*************Do It A Segment At A Time********************
            '**********************************************************
            For intSegments = 1 To intQtySegments                               'this ten eliminates some specs left on rim if circle
            If dblRadiuses(intRayPerDisplay, X, intSegments) > intDesiredRadius + 10 Then GoTo lblPastRadius
            frmScreen.Line -(xCoordinates(intRayPerDisplay, X, intSegments), yCoordinates(intRayPerDisplay, X, intSegments)), RGB(intColorInnerR - (X * intFadeDecrementR), intColorInnerG - (X * intFadeDecrementG), intColorInnerB - (X * intFadeDecrementB))
lblPastRadius: 'past radius, we only fade the part of rays WITHIN the user selected radius
            Sleep intSleepSegments 'user selected delay between segments
            Next intSegments
            '**********************************************************
            '*************End Do It A Segment At A Time****************
            '**********************************************************
        Sleep intSleepRays 'user selected delay between rays
        Next X
        '**********************************************************
        '*************End Fade Out Rays****************************
        '**********************************************************
lblNoFades2:
    'New Ray Processing Starts Here
    xx = intXStart: yy = intYStart 'ray start coordinates
    frmScreen.Line (intXStart, intYStart)-(intXStart, intYStart) 'blank move
    'create direction ray will go randomly between and according to user selected upper/lower angles
    intAngle = Int((intUpperAngle - intLowerAngle + 1) * Rnd + intLowerAngle)
    'reset increments and get ready to start calculating plot points
    dblTotalXinc = 0
    dblTotalYinc = 0
    dblRadius = 0
        '**********************************************************
        '*************Do It A Segment At A Time********************
        '**********************************************************
        For intSegments = 1 To intQtySegments
        'calculate the linear angle increment FROM A 0,0 CENTER
        'according to our previous random angle
        'later we will add this to the current x,y on graph
        intLinearXinc = (Cos(intAngle * Pi / 180) * 5) * intbooAngleInc
        intLinearYinc = (Sin(intAngle * Pi / 180) * 5) * intbooAngleInc
        'calculate in a little randomness for deviation of ray points
        'according to user inputted x&y rnd dev
        intRndXinc = ((Rnd * intXRnd) - intXRndHalf) 'the (intXRnd / 2)) centers the
        intRndYinc = ((Rnd * intYRnd) - intYRndHalf) 'deviation to plus or minus
        'calculate linear distance FROM A 0,0 CENTER! (radius)
        dblTotalXinc = dblTotalXinc + intLinearXinc + intRndXinc
        dblTotalYinc = dblTotalYinc + intLinearYinc + intRndYinc
        dblRadius = Sqr((dblTotalXinc ^ 2) + (dblTotalYinc ^ 2))
        'now calculate graph plot points
        xx = xx + intLinearXinc + intRndXinc
        yy = yy + intLinearYinc + intRndYinc
        'Store these new coordinates and radius in stack array
        'for later pushing stack for painting and fading
        xCoordinates(intRayPerDisplay, 0, intSegments) = xx
        yCoordinates(intRayPerDisplay, 0, intSegments) = yy
        dblRadiuses(intRayPerDisplay, 0, intSegments) = dblRadius
        If intQtyFades < 1 Then GoTo lblNoFades3 'No fades selected
            '**********************************************************
            '**Push New Coordinates and Radius into Stack Arrays*******
            '**We use these for the fades. When we push the stack******
            '**we also eliminate the tail end coordinates which were***
            '**used for the ending "darkest" fade**********************
            '**********************************************************
            For intFadeStack = intQtyFades To 1 Step -1
            'push 0 into 1, 1 into 2, 2 into 3, etc, but in reverse order
            'otherwise all elements would be the same
            xCoordinates(intRayPerDisplay, intFadeStack, intSegments) = xCoordinates(intRayPerDisplay, intFadeStack - 1, intSegments)
            yCoordinates(intRayPerDisplay, intFadeStack, intSegments) = yCoordinates(intRayPerDisplay, intFadeStack - 1, intSegments)
            dblRadiuses(intRayPerDisplay, intFadeStack, intSegments) = dblRadiuses(intRayPerDisplay, intFadeStack - 1, intSegments)
            Next intFadeStack
            '**********************************************************
            '**End Push Stac*******************************************
            '**********************************************************
lblNoFades3:
        'Have we reached our desired radius so as to change to a different color?
        If dblRadiuses(intRayPerDisplay, 0, intSegments) > intDesiredRadius Then frmScreen.ForeColor = lngColorOuter
        'ok. now we can paint our line. (one segment of one ray)
        frmScreen.Line -(xCoordinates(intRayPerDisplay, 0, intSegments), yCoordinates(intRayPerDisplay, 0, intSegments))
        frmScreen.ForeColor = RGB(intColorInnerR, intColorInnerG, intColorInnerB) ' be sure to change back to regular ray color
        Sleep intSleepSegments
        Next intSegments
        '**********************************************************
        '*************End Do It A Segment At A Time****************
        '**********************************************************
    Sleep intSleepRays 'user selected delay between rays
    DoEvents
    If booStop = True Then Exit Sub
    Next intRayPerDisplay
    '**********************************************************
    '**********End Make One Display At A Time Consisting*******
    '**********Of Desired Rays (Lightnings) Per Display********
    '**********************************************************
Loop
'**************************************************************
'************************End Main Loop*************************
'**************************************************************

'***************************************************************
'**********Finish up by fading out remaining rays***************
'***************************************************************

If intQtyFades < 1 Then Exit Sub 'if no fades selected then exit


Dim z As Integer
dblCount = 0
Do While dblCount <> intQtyFades
dblCount = dblCount + 1
    For intRayPerDisplay = 1 To intQtyRaysPerDisplay
        For X = dblCount To intQtyFades
        xx = intXStart: yy = intYStart
        'Blank Move
        frmScreen.Line (intXStart, intYStart)-(intXStart, intYStart), vbBlack
            '**********************************************************
            '*************Do It A Segment At A Time********************
            '**********************************************************
            For intSegments = 1 To intQtySegments                               'this ten eliminates some specs left on rim if circle
            If dblRadiuses(intRayPerDisplay, X, intSegments) > intDesiredRadius + 10 Then GoTo lblPastRadius2
            frmScreen.Line -(xCoordinates(intRayPerDisplay, X, intSegments), yCoordinates(intRayPerDisplay, X, intSegments)), RGB(intColorInnerR - (X * intFadeDecrementR), intColorInnerG - (X * intFadeDecrementG), intColorInnerB - (X * intFadeDecrementB))
lblPastRadius2:
            Sleep intSleepSegments
            Next intSegments
            '**********************************************************
            '*************End Do It A Segment At A Time****************
            '**********************************************************
        Sleep intSleepRays 'user selected delay between rays
        Next X
        For intSegments = 1 To intQtySegments
            For intFadeStack = intQtyFades To 1 Step -1
            'push 0 into 1, 1 into 2, 2 into 3, etc, but in reverse order
            'otherwise all elements would be the same
            xCoordinates(intRayPerDisplay, intFadeStack, intSegments) = xCoordinates(intRayPerDisplay, intFadeStack - 1, intSegments)
            yCoordinates(intRayPerDisplay, intFadeStack, intSegments) = yCoordinates(intRayPerDisplay, intFadeStack - 1, intSegments)
            dblRadiuses(intRayPerDisplay, intFadeStack, intSegments) = dblRadiuses(intRayPerDisplay, intFadeStack - 1, intSegments)
            Next intFadeStack
        Next intSegments
    Next intRayPerDisplay
Loop
'***************************************************************
'*********End Finish up by fading out remaining rays************
'***************************************************************

'**************************************************************
'*********************End Main Routine*************************
'**************************************************************
End Sub

