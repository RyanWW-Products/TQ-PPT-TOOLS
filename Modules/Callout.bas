Attribute VB_Name = "Callout"

Public Function formatsh(sh As Shape, sld As slide)
Dim eff As Effect
  sh.line.Weight = 3 / 4

  With sh.Shadow
    .Blur = 4
    .Size = 100
    .Transparency = 0.6
    .OffsetX = 2.12132
    .OffsetY = 2.12132
    .Visible = True
    
  End With
  

  
  
End Function


Public Function applyAnimaation(sh As Shape, sld As slide)
    Dim eff As Effect
    Dim eff2 As Effect
  Set eff = sld.TimeLine.MainSequence.AddEffect(sh, msoAnimEffectFadedZoom, trigger:=msoAnimTriggerWithPrevious, index:=-1)
    
  Set eff2 = sld.TimeLine.MainSequence.AddEffect(sh, msoAnimEffectPathDown, trigger:=msoAnimTriggerWithPrevious, index:=-1)
  eff2.Timing.Duration = 0.5
  eff2.Timing.SmoothStart = 0
  eff2.Timing.SmoothEnd = 0
eff2.Behaviors(1).MotionEffect.path = "M -0.3888 0.21667  L 0 0 "
'With eff2.Behaviors(1).MotionEffect
'    .FromX = -600
'        .FromY = -300
'        .ToX = 0
'        .ToY = 0
'End With
End Function

Sub CalloutAnimation(control As IRibbonControl)
    Dim shp As Shape
    Dim sh As Shape
    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    On Error GoTo selectPlease
    Set shp = ActiveWindow.Selection.ShapeRange(1)
       On Error GoTo 0
  If shp.Type = msoGroup Then
  
    For Each sh In shp.GroupItems
           Call formatsh(shp, sld)
      
    Next
    Else
        Call formatsh(shp, sld)
    End If
    
    Call applyAnimaation(shp, sld)
    Exit Sub
selectPlease:
    MsgBox ("Please Select a shape first")
End Sub

