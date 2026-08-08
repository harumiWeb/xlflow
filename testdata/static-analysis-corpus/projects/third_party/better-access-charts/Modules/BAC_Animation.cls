VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.29.01  published: 29.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Duration As Long
Private m_Easing As chEasing


' ### Properties

Public Property Let Duration(ByVal newValue As Long)
    m_Duration = newValue
End Property
Public Property Get Duration() As Long
    Duration = m_Duration
End Property

Public Property Let Easing(ByVal newValue As chEasing)
    m_Easing = newValue
End Property
Public Property Get Easing() As chEasing
    Easing = m_Easing
End Property


' ### Generator Function

Friend Function getAnimation() As String

    With New BAC_StringBuilder
        .AddLine "    Chart.defaults.animation.duration = " & Me.Duration & ";"
        .AddLine "    Chart.defaults.animation.easing = '" & getEasingString(Me.Easing) & "';"
        getAnimation = .ToString
    End With
    
End Function

Private Function getEasingString(ByVal ThisEasing As chEasing) As String

    Select Case ThisEasing
        Case linear
            getEasingString = "linear"
        Case easeInQuad
            getEasingString = "easeInQuad"
        Case easeOutQuad
            getEasingString = "easeOutQuad"
        Case easeInOutQuad
            getEasingString = "easeInOutQuad"
        Case easeInCubic
            getEasingString = "easeInCubic"
        Case easeOutCubic
            getEasingString = "easeOutCubic"
        Case easeInOutCubic
            getEasingString = "easeInOutCubic"
        Case easeInQuart
            getEasingString = "easeInQuart"
        Case easeOutQuart
            getEasingString = "easeOutQuart"
        Case easeInOutQuart
            getEasingString = "easeInOutQuart"
        Case easeInQuint
            getEasingString = "easeInQuint"
        Case easeOutQuint
            getEasingString = "easeOutQuint"
        Case easeInOutQuint
            getEasingString = "easeInOutQuint"
        Case easeInSine
            getEasingString = "easeInSine"
        Case easeOutSine
            getEasingString = "easeOutSine"
        Case easeInOutSine
            getEasingString = "easeInOutSine"
        Case easeInExpo
            getEasingString = "easeInExpo"
        Case easeOutExpo
            getEasingString = "easeOutExpo"
        Case easeInOutExpo
            getEasingString = "easeInOutExpo"
        Case easeInCirc
            getEasingString = "easeInCirc"
        Case easeOutCirc
            getEasingString = "easeOutCirc"
        Case easeInOutCirc
            getEasingString = "easeInOutCirc"
        Case easeInElastic
            getEasingString = "easeInElastic"
        Case easeOutElastic
            getEasingString = "easeOutElastic"
        Case easeInOutElastic
            getEasingString = "easeInOutElastic"
        Case easeInBack
            getEasingString = "easeInBack"
        Case easeOutBack
            getEasingString = "easeOutBack"
        Case easeInOutBack
            getEasingString = "easeInOutBack"
        Case easeInBounce
            getEasingString = "easeInBounce"
        Case easeOutBounce
            getEasingString = "easeOutBounce"
        Case easeInOutBounce
            getEasingString = "easeInOutBounce"
        Case Else
            getEasingString = "linear"
    End Select
    
End Function
