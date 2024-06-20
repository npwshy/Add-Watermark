#
# PowerPoint applicatio wrapper
#
using module ..\lib.pwsh\stdps.psm1

Set-StrictMode -Version latest

class PPTApp {
    $App;
    $IsRunning;

    static [int] $FormatPDF = 32;
    static [int] $SendToBack = 1;

    PPTApp() {
        $this.App = $null
    }

    Init() {
        $this.IsRunning = Get-Process -Name POWERPNT -ErrorAction Ignore
        $this.App = New-Object -ComObject PowerPoint.Application
    }

    [object] Open([string]$file) {
        $fp = [IO.Path]::GetFullPath($file)
        log "PowerPoint: Opening file: $fp"
        $openReadWrite = 0
        return $this.App.Presentations.Open($fp, $openReadWrite)
    }

    [object] ReopenIfReadOnly($deck) {
        if ($deck.ReadOnlyRecommended) {
            $fp = $deck.FullName
            $tmpfilename = [IO.Path]::GetRandomFileName() + (Split-Path -Extension -Path $fp)

            #--- Close the deck once and rename it
            $deck.Close()
            Rename-Item -Path $fp -NewName $tmpfilename

            #--- reopen renamed file and saveas to the original one with no readyonly flag
            $deck = $this.Open($tmpfilename)
            $deck.SaveCopyAs2($fp, 11, 0, 0) # default-format, no-embedded font, no-readonly recommended
            $deck.Close()

            #--- wait for sometime for save being done
            $s = 3;
            log "Deck is ReadOnly recommended. Waiting for copy to be saved ($s sec)"
            Start-Sleep -s $s

            $deck = $this.Open($fp)
            if ($deck.ReadOnlyRecommended) {
                throw "Cannot clear 'ReadOnlyRecommended' by saving via SaveCopyAs2(): $tmpfilename"
            }
        }
        return $deck
    }

    Save($deck) {
        $deck.Save()
        log "Deck saved: $($deck.Name)"
    }

    SaveAs($deck, $fp, $format) {
        $deck.SaveAs($fp, $format)
    }

    Close($deck) {
        $deck.Close()
    }

    Quit() {
        if ($this.App) {
            if (-not $this.IsRunning) {
                #--- powerpoint was NOT running at app launch
                $this.App.Quit()
                $this.App = $null
                Get-Process -Name POWERPNT -ErrorAction ignore |Stop-Process
            }
        }
    }

    [Object] GetAllSlides($deck) { return $deck.Slides }
    [object] GetSlide($deck, $no) { return $deck.Slides($no) }
    GotoSlide($deck, $no) {
        $deck.Window(1).View.GotoSlide($no) # 1?
    }
}