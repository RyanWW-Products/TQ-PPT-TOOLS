<# Runs native Ink conversion tests in an isolated PowerPoint presentation.
   Does not modify Trust Center settings, install add-ins, or touch user decks. #>
[CmdletBinding()]
param([switch]$MouseTest, [string]$SourcePptm)
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$outputDirectory=Join-Path $PSScriptRoot ('Output/ez-highlights-'+[Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $outputDirectory -Force)
$interop=[Reflection.Assembly]::Load('Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
Add-Type -ReferencedAssemblies $interop.Location -TypeDefinition @'
public static class EZTestRunner {
 public static int[] Coordinates(object app) {
  var ppt=(Microsoft.Office.Interop.PowerPoint._Application)app;
  var win=ppt.ActiveWindow;
  return new int[]{win.HWND,win.PointsToScreenPixelsX(100),win.PointsToScreenPixelsY(140),win.PointsToScreenPixelsX(400),win.PointsToScreenPixelsY(200),win.PointsToScreenPixelsY(240),win.PointsToScreenPixelsY(280)};
 }
 public static object Run(object app,string macro) {
  object[] args=new object[0];
  return ((Microsoft.Office.Interop.PowerPoint._Application)app).Run(macro,ref args);
 }
}
'@
$ppt=$null; $deck=$null; $owns=$false
try {
    try {$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}
    catch {$ppt=New-Object -ComObject PowerPoint.Application;$owns=$true}
    $testPath=Join-Path $outputDirectory 'EZHighlightsRegression.pptm'
    if($SourcePptm){
        Copy-Item -LiteralPath (Resolve-Path -LiteralPath $SourcePptm).Path -Destination $testPath
        $sec=$ppt.AutomationSecurity
        # Explicitly run the reviewed local release in this disposable copy only.
        # Restore the automation setting immediately; no Trust Center edits.
        try{$ppt.AutomationSecurity=1;$deck=$ppt.Presentations.Open($testPath,0,0,0)}
        finally{$ppt.AutomationSecurity=$sec}
        while($deck.Slides.Count -gt 0){$deck.Slides.Item(1).Delete()}
    }else{
        $deck=$ppt.Presentations.Add(0)
        $deck.SaveAs($testPath,25)
    }
    $deck.PageSetup.SlideWidth=720; $deck.PageSetup.SlideHeight=405
    $project=$deck.VBProject
    if($null -eq $project){throw 'Trusted VBA project access is required.'}
    $components=$project.VBComponents
    foreach($relative in @('Modules/EZHighlightEvents.cls','Modules/EZHighlightInk.bas','Modules/EZHighlights.bas','build/tests/EZHighlightsTests.bas')) {
        if($SourcePptm -and $relative.StartsWith('Modules/')){continue}
        $sourcePath=Join-Path $repoRoot $relative
        if($relative.EndsWith('.cls')){
            $text=[IO.File]::ReadAllText($sourcePath)
            $name=[regex]::Match($text,'Attribute VB_Name = "([^"]+)"').Groups[1].Value
            $body=$text.Substring($text.IndexOf('Option Explicit'))
            $body=[regex]::Replace($body,'(?m)^Attribute [^\r\n]*(?:\r?\n|$)','')
            $component=$components.Add(2)
            $component.Name=$name
            $component.CodeModule.AddFromString($body)
        }elseif($relative -eq 'Modules/EZHighlights.bas'){
            $body=[IO.File]::ReadAllText($sourcePath)
            $body=[regex]::Replace($body,'(?m)^Attribute [^\r\n]*(?:\r?\n|$)','')
            $body=$body.Replace('MsgBox ', 'EZHighlightsTests.TestMessage ')
            $component=$components.Add(1);$component.Name='EZHighlights'
            $component.CodeModule.AddFromString($body)
        }else{[void]$components.Import($sourcePath)}
    }
    $deck.Save()
    [void][EZTestRunner]::Run($ppt,'EZHighlightsRegression.pptm!EZHighlightsTests.RunAll')
    $report=Get-Content -LiteralPath (Join-Path $outputDirectory 'report.txt') -Raw
    $report -split '\r?\n' | Where-Object {$_ -match '^(FAIL|RESULT)'} | Write-Output
    if($MouseTest){
        Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class EZMouseTest {
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int L,T,R,B; }
 [DllImport("user32.dll")]public static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")]public static extern bool PrintWindow(IntPtr h,IntPtr dc,uint flags);
 [DllImport("user32.dll")]public static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int w,int height,uint flags);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetAncestor(IntPtr h,uint flags);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
 [DllImport("user32.dll")]public static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
}
'@
        [void][EZTestRunner]::Run($ppt,'EZHighlightsRegression.pptm!EZHighlightsTests.PrepareMouseDraw')
        $points=$null
        for($attempt=0;$attempt -lt 20;$attempt++){
            try{$points=[EZTestRunner]::Coordinates($ppt);break}
            catch{if($attempt -eq 19){throw};Start-Sleep -Milliseconds 150}
        }
        $handle=[EZMouseTest]::GetAncestor([IntPtr]$points[0],2)
        [void][EZMouseTest]::SetWindowPos($handle,[IntPtr](-1),0,0,0,0,0x43)
        [void][EZMouseTest]::SetForegroundWindow($handle)
        if([EZMouseTest]::GetForegroundWindow() -ne $handle){throw 'The test PowerPoint window could not take mouse focus.'}
        [void][EZMouseTest]::SetCursorPos($points[1],$points[2])
        Start-Sleep -Milliseconds 150
        [EZMouseTest]::mouse_event(2,0,0,0,[UIntPtr]::Zero)
        [EZMouseTest]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
        Start-Sleep -Milliseconds 150
        [void][EZTestRunner]::Run($ppt,'EZHighlightsRegression.pptm!EZHighlightsTests.ArmMouseDraw')
        Start-Sleep -Milliseconds 200
        $points=[EZTestRunner]::Coordinates($ppt)
        [void][EZMouseTest]::SetCursorPos($points[1],$points[2])
        [EZMouseTest]::mouse_event(2,0,0,0,[UIntPtr]::Zero)
        for($step=1;$step -le 15;$step++){
            [void][EZMouseTest]::SetCursorPos(($points[1]+($points[3]-$points[1])*$step/15),($points[2]+($points[4]-$points[2])*$step/15))
            Start-Sleep -Milliseconds 20
        }
        [EZMouseTest]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
        Start-Sleep -Milliseconds 800
        $last=$deck.Slides.Item($deck.Slides.Count)
        if(Test-Path -LiteralPath (Join-Path $outputDirectory 'ui-errors.txt')){throw (Get-Content -LiteralPath (Join-Path $outputDirectory 'ui-errors.txt') -Raw)}
        if($last.Shapes.Count -ne 1 -or $last.Shapes.Item(1).Tags.Item('EZHighlight') -ne '1'){
            Add-Type -AssemblyName System.Drawing
            $rect=New-Object EZMouseTest+Rect
            [void][EZMouseTest]::GetWindowRect($handle,[ref]$rect)
            $bitmap=New-Object Drawing.Bitmap(($rect.R-$rect.L),($rect.B-$rect.T))
            $graphics=[Drawing.Graphics]::FromImage($bitmap);$dc=$graphics.GetHdc()
            try{[void][EZMouseTest]::PrintWindow($handle,$dc,2)}finally{$graphics.ReleaseHdc($dc)}
            $bitmap.Save((Join-Path $outputDirectory 'draw-failure.png'));$graphics.Dispose();$bitmap.Dispose()
            Write-Output ('Mouse points='+($points -join ',')+' screenshot='+$outputDirectory)
            throw ('Real mouse draw did not produce one EZ Highlight; shapes='+$last.Shapes.Count)
        }
        Write-Output 'PASS | real mouse drag created an EZ Highlight'
        [void][EZTestRunner]::Run($ppt,'EZHighlightsRegression.pptm!EZHighlightsTests.ArmMouseDraw')
        [EZMouseTest]::keybd_event(27,0,0,[UIntPtr]::Zero)
        [EZMouseTest]::keybd_event(27,0,2,[UIntPtr]::Zero)
        Start-Sleep -Milliseconds 150
        [void][EZTestRunner]::Run($ppt,'EZHighlightsRegression.pptm!EZHighlightsTests.PlainRectangleMode')
        [void][EZMouseTest]::SetCursorPos($points[1],$points[5])
        [EZMouseTest]::mouse_event(2,0,0,0,[UIntPtr]::Zero)
        for($step=1;$step -le 15;$step++){
            [void][EZMouseTest]::SetCursorPos(($points[1]+($points[3]-$points[1])*$step/15),($points[5]+($points[6]-$points[5])*$step/15))
            Start-Sleep -Milliseconds 20
        }
        [EZMouseTest]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
        Start-Sleep -Milliseconds 500
        if($last.Shapes.Count -ne 2 -or $last.Shapes.Item(2).Type -ne 1 -or $last.Shapes.Item(2).Tags.Item('EZHighlight') -eq '1'){
            throw 'Escape did not cancel the pending highlight conversion.'
        }
        Write-Output 'PASS | Escape cancels highlighting; subsequent ordinary rectangle stays ordinary'
        [void][EZMouseTest]::SetWindowPos($handle,[IntPtr](-2),0,0,0,0,3)
    }
    $deck.SaveCopyAs((Join-Path $outputDirectory 'PortableHighlights.pptx'),24)
    $deck.Close(); $deck=$null
    $sec=$ppt.AutomationSecurity
    try {$ppt.AutomationSecurity=3;$deck=$ppt.Presentations.Open((Join-Path $outputDirectory 'PortableHighlights.pptx'),-1,0,0)}
    finally {$ppt.AutomationSecurity=$sec}
    $highlights=0
    foreach($slide in $deck.Slides){foreach($shape in $slide.Shapes){if($shape.Tags.Item('EZHighlight') -eq '1'){$highlights++}}}
    $expectedHighlights=9
    if($MouseTest){$expectedHighlights++}
    if($highlights -ne $expectedHighlights){throw "Save/reopen did not retain all native highlights: $highlights"}
    Write-Output "PASS | macro-free save/reopen retained $highlights highlights"
    if($report -notmatch 'RESULT \| checks=\d+ failures=0'){throw 'EZ Highlights regression failure.'}
    Write-Output ('Artifacts: '+$outputDirectory)
} finally {
    if($MouseTest -and $handle){[void][EZMouseTest]::SetWindowPos($handle,[IntPtr](-2),0,0,0,0,3)}
    if($null -ne $deck){$deck.Saved=-1;$deck.Close()}
    if($owns -and $null -ne $ppt -and $ppt.Presentations.Count -eq 0){$ppt.Quit()}
}
