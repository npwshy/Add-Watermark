#
# Add watermark to PPT deck
#
using module .\lib.office\PPTApp.psm1
using module .\lib.pwsh\stdps.psm1

Param(
[string]$File,
[string]$Fodler,
[Alias('dest')]
[string]$Destination,
[Alias('Template','Templ')]
[string]$WatermarkTemplate,
[alias('pdf')]
[switch]$SaveAsPdf,
[string]$Suffix = '-watermarked%yyyy%%MM%%dd%',
[string]$OldSuffix,
[switch]$ShowSlide,
[switch]$SendToBack,
[string]$LogFile = "logs\log.txt",
[int]$LogGenerations = 9,
[switch]$Help,
#
# Internal
#
[int]$WaitForSave = 3000
)

function showHelp() {
    write-host @"
Options:
-File <file>        ... Specifies ppt-file to be processed
-Folder <folder>    ... Specifies folder to be processed. All PPT files in the folder are processed
-Destination <path> ... Specifies destination path (file or folder)
-WatermarkTemplate <file> ... Specifies PPT file of watermark template. Only 1st slide will be used as watermark
-SaveAsPdf          ... PDF file is also created with watermark (Default: $SaveAsPdf)
-Suffix <string>    ... Adding suffix to original filename (default: $Suffix)
-OldSuffix <string> ... Specified <old suffix> will be deleted from filename if any (default: $OldSuffix)
-ShowSlide          ... Show the current slide during the processing (Default: $ShowSlide)
-SendToBack         ... Put the watermark to the back (Default: $SendToBack)
"@
}

$ErrorActionPreference = "stop"
Set-StrictMode -Version latest

if ($Help) { showHelp; exit }

$EXITCODE = 0
__RunApp ([Main]::New()) (Join-Path $PSScriptRoot $LogFile) $LogGenerations
exit $EXITCODE


class Main {
    $PPT;
    $Files;
    $DestDir;
    $Watermark;

    Run() {
        $this.Init()
        try {
            $this.Watermark.GetWatermark($this.PPT, $script:WatermarkTemplate)
            foreach ($file in $this.Files) {
                $destfp = $this.GetDestinationFilename($file)
                Copy-Item -Path $file -Destination $destfp -Force -Verbose:1
                $this.ApplyWatermarks($destfp)
            }
        } catch {
            logerror "$_"
            if ($script:StackTrace) {
                logerror $script:StackTrace
            }
        } finally {
            $this.PPT.Quit()
        }
    }

    ApplyWatermarks([string]$destfp) {
        $deck = $this.PPT.Open($destfp)
        $deck = $this.PPT.ReopenIfReadOnly($deck)
        $slideno = 0
        foreach ($slide in $this.PPT.GetAllSlides($deck)) {
            $slideno++
            log "Applying watermarks to slide: $slideno"
            $this.Watermark.Apply($slide)
            $slide = $null
        }

        $this.CoolForSave()
        $this.PPT.Save($deck)
        $this.CoolForSave()

        if ($script:SaveAsPdf) {
            $ext = Split-Path -Path $destfp -Extension
            $pdf = $destfp -replace "$ext$",'.pdf'
            log "Saving PDF file: $pdf"
            $this.PPT.SaveAs($deck, $pdf, [PPTApp]::FormatPDF)
            $this.CoolForSave()
        }
        $this.PPT.Close($deck)
        $deck = $null
    }

    CoolForSave() {
        if ($script:WaitForSave) {
            Start-Sleep -Milliseconds $script:WaitForSave
        }
    }

    [string] GetDestinationFilename([string]$srcfp) {
        $destfp = Split-Path -Path $srcfp -Leaf
        if ($script:OldSuffix) {
            $os = $script:OldSuffix
            $os = $os -replace '%yyyy%','\d{4}'
            $os = $os -replace '%MM%','\d{2}'
            $os = $os -replace '%dd%','\d{2}'
            $destfp = $destfp -replace $os,''
        }
        $ext = Split-Path -Path $destfp -Extension
        if ($script:Suffix) {
            $dt = [DateTime]::Now
            $destfp = $destfp -replace "$ext$","$($script:Suffix)$ext"
            $destfp = $destfp -replace '%yyyy%',$dt.Year
            $destfp = $destfp -replace '%MM%',$dt.Month
            $destfp = $destfp -replace '%dd%',$dt.Day
        }
        $destfp = Join-Path $this.DestDir $destfp
        log "dest=$destfp src=$srcfp"
        if ($destfp -eq $srcfp) {
            throw "Destination file is same as soure file. Use -Destination to specify another folder or -suffix to add suffix"
        }
        return $destfp
    }

    Init() {
        if ($script:File -and $script:Fodler) { throw "Both -File and -Folder specified. Use either" }
        if ($script:File) {
            $fp = [IO.Path]::GetFullPath($script:File)
            if (-not (Test-Path $fp)) {
                throw "File does not exist: $fp"
            }
            $this.Files = @($fp)
            $this.DestDir = $script:Destination ? $script:Destination : (Split-Path -parent $script:File)
        } elseif ($script:Folder) {
            $this.Files = Get-ChildItem -Filter *.pptx -Path $script:Folder |%{ $_.FullName }
            if (-not $this.Files) {
                throw "File does not exist in folder: $($script:Folder)"
            }
            $this.DestDir = $script:Destination ? $script:Destination : $script:Folder
        } else {
            throw "Neither -File nor -Folder specified"
        }
        $this.DestDir = [IO.Path]::GetFullPath($this.DestDir)

        if (-not $script:WatermarkTemplate) { throw "WatermarkTemplate not specified" }
        if (-not (Test-Path $script:WatermarkTemplate)) { throw "WatermarkTemplate not found" }

        $this.PPT = [PPTApp]::New()
        $this.PPT.Init()

        $this.Watermark = [Watermark]::New()
        $this.Watermark.Init()
    }

}

class WatermarkShape {
    $Top; $Left; $Width; $Height;
    $Rotation;
    $Font;
    $Line;
    $Fill;
    $TF_Orientation;
    $TF_HAnchor; $TF_VAnchor;
    $TF_MarginTBLR;
    $TF_TR_Text;
    $TF_TR_PF_Alignment;

    static $FontProps = 'Bold,Italic,Name,Size,Underline' -split(',')
    static $LineProps = 'DashStyle,Style,Transparency,Weight' -split(',')

    Show([string]$m) {
        log "$m Watermark: T:$($this.TF_TR_Text) $($this.Left)x$($this.Top)-$($this.Width)x$($this.Height) R:$($this.Rotation)"
    }

    WatermarkShape($s) {
        $this.Top = $s.Top
        $this.Left = $s.Left
        $this.Width = $s.Width
        $this.Height = $s.Height
        $this.Rotation = $s.Rotation
        $this.GetFontInfo($s)
        $this.GetLineInfo($s)
        $this.GetFillInfo($s)
        $this.TF_Orientation = $s.TextFrame.Orientation
        $this.TF_HAnchor = $s.TextFrame.HorizontalAnchor
        $this.TF_VAnchor = $s.TextFrame.VerticalAnchor
        $this.TF_MarginTBLR = @(
            $s.TextFrame.MarginTop,
            $s.TextFrame.MarginBottom,
            $s.TextFrame.MarginLeft,
            $s.TextFrame.MarginRight)
        $this.TF_TR_Text = $s.TextFrame.TextRange.Text
        $this.TF_TR_PF_Alignment = $s.TextFrame.TextRange.ParagraphFormat.Alignment
    }

    GetFontInfo($s) {
        $sf = $s.TextFrame.TextRange.Font
        $this.Font = @{}
        [WatermarkShape]::FontProps |%{ $this.Font.$_ = $sf.$_ }
        $this.Font.Color_RGB = $sf.Color.RGB
    }

    SetFontInfo($t) {
        $sf = $t.TextFrame.TextRange.Font
        [WatermarkShape]::FontProps |%{ $sf.$_ = $this.Font.$_ }
        $sf.Color.RGB = $this.Font.Color_RGB
    }

    GetLineInfo($s) {
        if ($s.Line.Visible -eq 0) {
            $this.Line = $null
        } else {
            $this.Line = @{}
            [WatermarkShape]::LineProps |%{ $this.Line.$_ = $s.Line.$_ }
            $this.Line.ForeColor_RGB = $s.Line.ForeColor.RGB
        }
    }

    SetLineInfo($s) {
        if ($this.Line) {
            [WatermarkShape]::LineProps |%{ $s.Line.$_ = $this.Line.$_ }
            $s.Line.ForeColor.RGB = $this.Line.ForeColor_RGB
        }
    }

    GetFillInfo($s) {
        if ($s.Fill.Visible -eq 0) {
            $this.Fill = $null
        } else {
            $this.Fill = @{}
            $this.Fill.ForeColor_RGB = $s.Fill.ForeColor.RGB
        }
    }

    SetFillInfo($s) {
        if ($this.Fill) {
            $s.Fill.ForeColor.RGB = $this.Fill.ForeColor_RGB
        }
    }

    AddToSlide($slide) {
        $tb = $slide.Shapes.AddTextbox($this.TF_Orientation, $this.Left, $this.Top, $this.Width, $this.Height)
        $tb.Rotation = $this.Rotation
        (
            $tb.TextFrame.MarginTop,
            $tb.TextFrame.MarginBottom,
            $tb.TextFrame.MarginLeft,
            $tb.TextFrame.MarginRight) = $this.TF_MarginTBLR
        $tb.TextFrame.TextRange.Text = $this.TF_TR_Text
        $tb.TextFrame.TextRange.ParagraphFormat.Alignment = $this.TF_TR_PF_Alignment
        $tb.TextFrame.Orientation = $this.TF_Orientation
        $tb.TextFrame.HorizontalAnchor = $this.TF_HAnchor
        $tb.TextFrame.VerticalAnchor = $this.TF_VAnchor

        $this.SetFontInfo($tb)
        $this.SetLineInfo($tb)
        $this.SetFillInfo($tb)

        $tb.Width = $this.Width
        $tb.Height = $this.Height
        $tb.Left = $this.Left
        $tb.Top = $this.Top

        if ($script:SendToBack) {
            $tb.ZOrder([PPTApp]::SendToBack)
            log "pushed to $([int][PPTApp]::SendToBack)"
        }
    }
}

class Watermark {
    [WatermarkShape[]]$Watermarks;

    Init() {}
    GetWatermark([PPTApp]$app, [string]$fp) {
        $this.Watermarks = @()
        $deck = $null
        $slide = $null
        try {
            log "Capturing watermark data from template: $fp"
            $deck = $app.Open($fp)
            $slide = $app.GetSlide($deck, 1)

            foreach ($shape in $slide.Shapes) {
                if (-not $shape.HasTextFrame) {
                    log "Shape (Id:$($shape.Id)) does not have text. Skipping."
                } else {
                    $this.Watermarks += ,[WatermarkShape]::New($shape)
                }
            }
        } catch {
            logerror "ERROR! $_"
        } finally {
            $app.Close($deck)
            $deck = $null
            $slide = $null
            $shape = $null
        }

        $this.ShowWatermark()
    }

    ShowWatermark() {
        $this.Watermarks |% -Begin { $c = 0 } -Process { $_.Show($c); $c++ }
    }

    Apply($slide) {
        foreach ($wm in $this.Watermarks) {
            $wm.AddToSlide($slide)
        }
    }
}