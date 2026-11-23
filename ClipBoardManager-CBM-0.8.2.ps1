<#
    .SYNOPSIS
        Clipboard history viewer with filtering, tagging and file capture.

    .DESCRIPTION
        Watches the Windows clipboard and lists every captured item. Filter by typing
        in the filter box. Right-click any item for copy / remove / colour / file
        actions, or use the Tools menu for display options.

    .EXAMPLE
        Preferred - no console window at all:
            double-click Start-ClipBoardManager.vbs (keep it beside this .ps1)

        From a command line:
            powershell -ExecutionPolicy Bypass -File <path>\ClipBoardManager-CBM-0.8.2.ps1

        The window opens in the light theme. Tools > Theme changes it while it runs.
        To start in another theme (Dark, or Auto to follow the Windows setting):
            powershell -ExecutionPolicy Bypass -File <path>\ClipBoardManager-CBM-0.8.2.ps1 -Theme Dark

        -noexit is no longer needed. The script now waits on the GUI runspace itself
        rather than relying on a console window to keep the process alive.

    .NOTES
        Keyboard shortcuts:
            F2       -> Edit the selected item
            Right    -> Expand the selected item(s): long ones open, cut-off ones wrap
            Left     -> Collapse the selected item(s)
            Drag     -> With ONE item selected, drag it to a new position
            Ctrl + A -> Select all items
            Ctrl + C -> Copy selected items
            Ctrl + R -> Remove selected items
            Ctrl + W -> Write selected items to a new file
            Ctrl + Q -> Append selected items to an existing file
            Ctrl + L -> Append selected items to the Review Log
            Ctrl + O -> Open the Review Log
            Ctrl + E -> Exit

        Review Log lives at:
            %USERPROFILE%\Documents\ClipBoardManager-CBM-ReviewLog.txt

        Authors:
            Boe Prox      - original clipboard viewer (2014)
            ChrisDee      - GitHub (2016)
            Duane Joseph  - feature development (2023 - )

    Version history
        1.0  2014-07-10  Boe Prox - initial version
        1.1  2014-07-24  Boe Prox - filter moved to TextChanged, multi-select, mouse wheel,
                                    keyboard shortcuts
        0.0.1 2023-03-13 Duane - tools menu, pause, window titling, timestamps, stay-on-top,
                                    masking, ChatGPT options
        0.5.0 2026-      Duane - line-numbered items, import, Edge/Google/Notepad actions
        0.5.1            Insert Topic divider added to the context menu
        0.5.2            Tools > Line Numbers toggle; copy no longer strips leading digits
                            when numbering is off
        0.5.3            Dialogs given owner windows so they cannot open behind the
                            always-on-top main window; VB InputBox replaced with
                            Show-InputPrompt; Options window no longer shadows $Window
        0.5.4            Per-item colours via right-click > Assign Color
        0.6.0            Masked terms changed from one concatenated string to a list;
                            both capture branches now share Test-Masked
        0.6.1            Masks store raw text via Get-ItemText instead of the formatted row
        0.6.2            Dead code purged; unreachable Edit_Menu handler removed; Enter-key
                            edit path guarded; empty Options entries dropped
        0.6.3            Item model: list holds ClipItem objects (Text / Display / Color)
                            instead of packed strings. Colour is a bound property, the two
                            capture branches collapsed into Add-ClipboardItem, the filter
                            searches raw text, and masked items no longer store the secret.
        0.6.4            Script waits on the GUI runspace instead of depending on -noexit,
                            so it can run with no console window. Closing the history
                            window now asks for confirmation when the list is not empty.
                            Added Start-ClipBoardManager.vbs for a flash-free launch.
        0.6.5            Tools > Voice Note: offline dictation via System.Speech.Recognition.
                            Accuracy was unusable in practice; replaced in 0.6.6.
        0.6.6            Voice Note rebuilt on xAI Grok speech-to-text. Records with the
                            winmm MCI recorder and posts the wav to /v1/stt.
        0.6.7            Recording format is negotiated rather than hardcoded. 0.6.6 asked
                            for 16 kHz mono and failed outright on drivers that refuse it
                            ("No wave device is installed that can record files in the
                            current format"). Now walks a ladder of formats and falls back
                            to the device default. See Test-Microphone.ps1 to diagnose.
        0.6.8            Ladder reordered from real probe results (stereo 44.1 first) and a
                            short settle delay added between attempts.
        0.7.0            AI router. Capability -> provider map (Chat, SpeechToText,
                            TextToSpeech) with one key per provider, all in memory.
                            Tools > AI Settings picks the provider per capability.
                            Ask ChatGPT became Ask AI and works with any chat provider;
                            the old magic-string settings and the retired /v1/completions
                            call are gone.
        0.7.1            Edit item restored (context menu and F2), replacing the hidden
                            textbox path that had been unreachable since 2023. Editing
                            keeps the line number, timestamp and colour.
        0.7.2            Tools > New Instance starts an independent copy in its own
                            process; closing either window does not affect the other.
        0.7.3            Long items collapse. Anything over 150 characters shows as a
                            one-line preview with a toggle arrow; click the arrow or use
                            Right / Left to expand and collapse. Expand All / Collapse All
                            on the context menu. Copy, filter, edit and save still work on
                            the full text.
        0.7.4            Right-click menu shortened. Item actions (Ask AI - TriplOOO,
                            Edge, Select All, Copy, Remove, Edit, Expand / Collapse All,
                            Insert Topic, Assign Color) moved into an Options fly-out.
        0.7.5            Open Selected in New Instance. The selected items (text, colour,
                            timestamp) are handed to a new, independent instance through a
                            temporary file in %TEMP% that the new instance deletes as soon
                            as it has read it. Items are renumbered from 1 there and stay
                            in this list too. Instances started from the app no longer
                            capture whatever was already on the clipboard at startup.
        0.7.6            Selected rows are solid Windows blue (#0078D4) with white text,
                            and stay that colour when focus moves elsewhere. With exactly
                            one item selected it can be dragged to a new position; line
                            numbers are renumbered top to bottom after the move.
        0.7.7            Consolidate Selected (Options fly-out) merges two or more items
                            into one collapsible group with an optional title. The group
                            keeps the original items, so Ungroup restores them with their
                            colours and timestamps. Groups always collapse, whatever their
                            length, and travel intact to Open Selected in New Instance.
        0.7.8            Right-click menu: Copy, Remove item, Edit item, Insert Topic and
                            Assign Color moved back out of the Options fly-out to the top
                            of the menu, where they are one click away. The fly-out keeps
                            the less-used actions.
        0.7.9            New look, built from stock WPF only (no theme libraries, still
                            one file, normal Windows title bar).
                            List: line numbers sit in their own column and always run
                            1..N; selection is a soft tint; a row colour is a tint plus a
                            marker stripe; Insert Topic rows read as headings; the scroll
                            bar is slim. Rows no longer scroll sideways - a row that does
                            not fit ends in "..." and shows its text in a tooltip, and
                            Right / Left expand and collapse any row, not only long ones.
                            A time stamp shows as month/day and time; the full stamp is
                            its tooltip. The list builds only the rows that are on
                            screen, so a long history stays quick.
                            Top and bottom: the filter box has an icon and hint text,
                            Tools is a button beside it, and a status strip shows
                            Capturing / Paused and the item counts.
                            Menus: icons, real separator lines, shortcuts in a column.
                            Colours follow the Windows light / dark setting (and high
                            contrast) as it is when the window opens; -Theme overrides.
                            Item model: ClipItem carries Numbered / Number / Suffix
                            instead of one pre-formatted Display string. That removes
                            the string splicing behind Edit item, Ungroup and Open in New
                            Instance, and line numbers can no longer repeat after Remove.
                            Not restyled yet: the Edit / Insert Topic prompts, Options,
                            AI Settings and Voice Note windows.
        0.8.0            Window icon: a clipboard, drawn by the script (no icon file),
                            on the main window and on every window it opens.
                            Menus stand out more: accent-coloured border, deeper shadow,
                            accent-coloured icons, and the item under the pointer gets a
                            tinted background with an accent bar at its left edge.
                            Alternating rows are back as a soft stripe on every other
                            row. It follows light / dark, keeps the rounded corners, and
                            sits under selection, the pointer and Assign Color tints.
                            Edit item opens a multi-line box that wraps, can be resized
                            and opens sized to the text. Enter starts a new line,
                            Ctrl+Enter saves, Esc cancels.
        0.8.1            Favorites. Every row has a star in front of its number: click
                            it to mark or unmark the row (it shows as an outline while
                            the pointer is over a row, solid gold when set). Favorite on
                            the right-click menu does the same for all selected rows,
                            and Options > Select Favorites selects every favorite on
                            show. The star button beside the filter box, or Tools > Show
                            Only Favorites, hides everything else; it works together
                            with the filter text. The status strip counts favorites.
                            A favorite is written with "* " in front of its text by
                            Save As File, Write / Append file, the Review Log, Open in
                            NOTEPAD and Open with Edge. Copy stays clean. Favorites are
                            held in memory only, like colours; they survive Edit item,
                            Consolidate / Ungroup and Open Selected in New Instance.
                            Import File does not turn "* " back into a star.
                            Window title: "Clip Board Manager 0.8.1" from the start. The
                            Good Morning / Good Afternoon / Have a good evening title
                            that replaced it at the first capture is gone. Cancel in
                            Tools > Title Window now leaves the title as it was instead
                            of blanking it.
        0.8.2            Light is the default theme; the window no longer follows the
                            Windows setting unless asked to. Tools > Theme switches
                            between Light, Dark and Follow Windows while the window is
                            open: the list, menus, tooltips and title bar change at
                            once and the history is untouched. Follow Windows reads the
                            Windows light / dark setting (or high contrast) at the
                            moment it is chosen. The choice is held in memory only; an
                            instance started from the window opens in the same theme.
                            -Theme Light / Dark / Auto on the command line sets the
                            starting theme. With no -Theme at all, Windows high
                            contrast still wins over light.
                            How: every brush is bound to a colour holder named after it
                            plus "Source", and Set-UiTheme changes the holders.

    Known gaps
        - Masking is unverified end to end. Suspect the ListBox does not select on
          right-click, leaving SelectedItems stale.

    AI features (Voice Note, Ask AI)
        - Providers: Grok (xAI), ChatGPT (OpenAI), Claude (Anthropic), Gemini (Google).
          Each capability is routed separately under Tools > AI Settings.
            Chat            all four
            Speech to text  Grok and OpenAI only. Anthropic publishes no speech API;
                            Gemini transcription needs a two-step Files upload that is
                            not wired up here.
            Text to speech  Windows offline synthesizer
        - One API key per provider, prompted on demand and held in memory only. Nothing
          is written to disk, so keys are re-entered after a restart.
        - Model IDs live in the $Script:AiProviders table - the only place they appear.
          If a call returns "unknown model", edit it there.
        - Microphone access allowed for desktop apps:
          Settings > Privacy & security > Microphone.

        Privacy note: unlike the rest of this tool, Voice Note sends data off the machine.
        Audio is recorded to a temporary wav in %TEMP%, uploaded to xAI, and the wav is
        deleted immediately afterwards - including when the request fails.
        Billing is roughly $0.10 per hour of audio.

    By design
        - Nothing is written outside the Review Log. Menu toggles, masks and the API key
          all reset on restart so the script stays portable and leaves no footprint.
          Exception: Open Selected in New Instance writes the selected items to a
          temporary file in %TEMP% for about a second while the new instance starts.
#>

# 11/7/26 - Only used when one instance starts another. A normal launch passes neither.
#   -SeedFile            temporary JSON of items to preload; deleted as soon as it is read
#   -SkipInitialCapture  do not capture whatever is already on the clipboard at startup
# 11/22/26 - optional, the theme the window opens in:
#   -Theme               Light (default), Dark, or Auto to follow the Windows light /
#                        dark setting. Tools > Theme changes it while the window is
#                        open. 11/23/26 (0.8.2) - the default used to be Auto.
param(
    [string]$SeedFile,
    [switch]$SkipInitialCapture,
    [ValidateSet('Auto','Light','Dark')]
    [string]$Theme = 'Light'
)

##Requires -Version 3.0
$ExecutionContext.SessionState.LanguageMode = 'FullLanguage' #Added 4/26/26
$Runspacehash = [hashtable]::Synchronized(@{})
$Runspacehash.Host = $Host
# 11/7/26 - the GUI runspace cannot see $PSCommandPath, so hand it the script's own
# path here. Tools > New Instance relaunches this file from it.
$Runspacehash.ScriptPath = $(if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path })
$Runspacehash.SeedFile = $SeedFile
$Runspacehash.SkipInitialCapture = [bool]$SkipInitialCapture
$Runspacehash.Theme = $Theme
# 11/23/26 (0.8.2) - whether -Theme was typed. If it was not, the window opens light
# unless Windows is in high contrast mode.
$Runspacehash.ThemeGiven = $PSBoundParameters.ContainsKey('Theme')
$Runspacehash.runspace = [RunspaceFactory]::CreateRunspace()
$Runspacehash.runspace.ApartmentState = "STA"
$Runspacehash.runspace.Open()
$Runspacehash.runspace.SessionStateProxy.SetVariable("Runspacehash",$Runspacehash)
$Runspacehash.PowerShell = {Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase}.GetPowerShell()

$Runspacehash.PowerShell.Runspace = $Runspacehash.runspace
$Runspacehash.Handle = $Runspacehash.PowerShell.AddScript({
    Function Get-ClipBoard {
        [Windows.Clipboard]::GetText()
    }

# 8/25/26 - ITEM MODEL
    # The list used to hold plain strings, which forced line number, timestamp, colour
    # and text into one string that had to be pulled apart again with regex.
    # ClipItem carries them as real properties and raises PropertyChanged, so the UI
    # reacts when a property changes (colour in particular).
    #   Text    - raw clipboard content, or the mask for a masked item
    #   Display - the formatted row (frozen at capture time until 0.7.9; see 11/22/26)
    #   Color   - the Assign Color value, empty for none
    # 11/7/26 - collapsible long items:
    #   Preview    - one-line short form, empty when the item is short enough to show whole
    #   IsExpanded - toggled by the arrow; Body and Glyph follow it
    # 11/8/26 - consolidated groups:
    #   IsGroup    - this item was made by Consolidate Selected
    #   GroupTitle - optional title shown in the collapsed row
    #   Parts      - the original items (Text / Color / Suffix / Numbered), for Ungroup
    # 11/22/26 - the row is no longer one pre-formatted string. Each thing a row shows is
    # its own property and the XAML puts each in its own column:
    #   Numbered   - this row shows a line number (the Tools > Line Numbers state when
    #                the row was added)
    #   Number     - its position in the list, kept current by Renumber, so numbers
    #                always run 1..N and never repeat
    #   Suffix     - the time stamp exactly as it was captured (" : dddd MM/dd/yyyy
    #                HH:mm:ss"), or empty. Shown compactly; the full text is the tooltip.
    #   Preview    - the one-line snippet of a long item or a group; empty when the item
    #                is short enough to show whole. IsLong follows it.
    #   Detail     - small grey note at the right of the row ("348 chars", "3 items")
    #   Icon       - icon-font character for the row type (voice, AI, link, masked)
    #   IsTopic    - an Insert Topic divider; Caption holds its title
    #   IsDim      - draw the text muted (masked rows)
    #   IsFavorite - the row is starred (0.8.1). Memory only, like Color.
    #   Body, Glyph, Meta, NumberText, TagFill, TagMark - what the columns bind to,
    #                worked out from the properties above
    #   Display    - "<number> <text><time stamp>" as one string, for anything that
    #                still wants the old formatted row. Nothing parses it any more.
    # The class name carries the version so a console session that already ran an older
    # copy of this script cannot hand this one an out-of-date ClipItem.
    if (-not ('Cbm082.ClipItem' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Cbm082
{
    public class ClipItem : INotifyPropertyChanged
    {
        // Assign Color stores a fixed pastel hex on the item. This maps that stored
        // value to { row tint, marker stripe } for the theme in use, so a colour means
        // the same thing in a light and a dark window.
        public static readonly Dictionary<string, string[]> TagColors =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

        // month, day, hour and minute out of " : dddd MM/dd/yyyy HH:mm:ss". The
        // separators are whatever the regional settings produced, hence \D.
        private static readonly Regex StampTail =
            new Regex(@"(\d{1,2})\D(\d{1,2})\D\d{4}\s+(\d{1,2})\D(\d{2})\D\d{2}\s*$");

        private string _text = "";
        private string _suffix = "";
        private string _color = "";
        private string _preview = "";
        private string _detail = "";
        private string _caption = "";
        private string _icon = "";
        private bool _numbered;
        private bool _expanded;
        private bool _topic;
        private bool _dim;
        private bool _favorite;
        private int _number;

        public DateTime Captured { get; set; }

        public bool IsGroup { get; set; }
        public string GroupTitle { get; set; }
        public object Parts { get; set; }

        public string Text {
            get { return _text; }
            set { _text = value ?? ""; OnChanged("Text"); OnChanged("Body"); OnChanged("Tip"); OnChanged("Display"); }
        }

        // ---- line number ----
        public bool Numbered {
            get { return _numbered; }
            set { _numbered = value; OnChanged("Numbered"); OnChanged("NumberText"); OnChanged("Display"); }
        }

        public int Number {
            get { return _number; }
            set {
                if (_number == value) return;
                _number = value;
                OnChanged("Number"); OnChanged("NumberText"); OnChanged("Display");
            }
        }

        public string NumberText { get { return (_numbered && _number > 0) ? _number.ToString() : ""; } }

        // Gives every item from index 'from' down its position as its number, and
        // returns the highest number any row shows (0 when no row is numbered) so the
        // caller can size the number column. Done here rather than in script because
        // Remove on a big selection calls it once per row.
        public static int Renumber(IList items, int from)
        {
            if (items == null) return 0;
            if (from < 0) from = 0;
            for (int i = from; i < items.Count; i++) {
                ClipItem c = items[i] as ClipItem;
                if (c != null) c.Number = i + 1;
            }
            // numbers grow with position, so the last numbered row has the highest one
            for (int i = items.Count - 1; i >= 0; i--) {
                ClipItem c = items[i] as ClipItem;
                if (c != null && c.Numbered) return i + 1;
            }
            return 0;
        }

        // ---- time stamp ----
        public string Suffix {
            get { return _suffix; }
            set { _suffix = value ?? ""; OnChanged("Suffix"); OnChanged("Meta"); OnChanged("MetaTip"); OnChanged("Display"); }
        }

        private string FullStamp { get { return _suffix.Trim().TrimStart(':').Trim(); } }

        private string ShortStamp {
            get {
                if (_suffix.Length == 0) return "";
                Match m = StampTail.Match(_suffix);
                if (!m.Success) return FullStamp;
                return m.Groups[1].Value + "/" + m.Groups[2].Value + " " + m.Groups[3].Value + ":" + m.Groups[4].Value;
            }
        }

        // ---- long items and groups ----
        public string Preview {
            get { return _preview; }
            set { _preview = value ?? ""; OnChanged("Preview"); OnChanged("IsLong"); OnChanged("Body"); OnChanged("Glyph"); }
        }

        public bool IsLong { get { return _preview.Length > 0; } }

        public bool IsExpanded {
            get { return _expanded; }
            set { _expanded = value; OnChanged("IsExpanded"); OnChanged("Body"); OnChanged("Glyph"); }
        }

        // ---- row type, set by Update-ClipPreview ----
        public string Detail {
            get { return _detail; }
            set { _detail = value ?? ""; OnChanged("Detail"); OnChanged("Meta"); }
        }

        public string Caption {
            get { return _caption; }
            set { _caption = value ?? ""; OnChanged("Caption"); }
        }

        public string Icon {
            get { return _icon; }
            set { _icon = value ?? ""; OnChanged("Icon"); OnChanged("Glyph"); }
        }

        public bool IsTopic {
            get { return _topic; }
            set { _topic = value; OnChanged("IsTopic"); OnChanged("Tip"); }
        }

        public bool IsDim {
            get { return _dim; }
            set { _dim = value; OnChanged("IsDim"); }
        }

        // ---- favorite (0.8.1): the star in front of the number ----
        public bool IsFavorite {
            get { return _favorite; }
            set { if (_favorite == value) return; _favorite = value; OnChanged("IsFavorite"); }
        }

        // How many of the items are favorites. Here rather than in script because the
        // status strip asks after every change to the list.
        public static int CountFavorites(IList items)
        {
            if (items == null) return 0;
            int n = 0;
            for (int i = 0; i < items.Count; i++) {
                ClipItem c = items[i] as ClipItem;
                if (c != null && c._favorite) n++;
            }
            return n;
        }

        // ---- what the row columns bind to ----
        public string Body { get { return (IsLong && !_expanded) ? _preview : _text; } }

        // chevron right / chevron down for a row that opens, otherwise the row's icon
        public string Glyph { get { return IsLong ? (_expanded ? "\uE972" : "\uE974") : _icon; } }

        public string Meta {
            get {
                string stamp = ShortStamp;
                if (_detail.Length == 0) return stamp;
                if (stamp.Length == 0) return _detail;
                return _detail + "  \u00B7  " + stamp;
            }
        }

        public string MetaTip { get { return _suffix.Length == 0 ? null : FullStamp; } }

        // tooltip for a row that is cut short; capped so a huge clip stays a tooltip
        public string Tip {
            get {
                if (_topic || _text.Length == 0) return null;
                if (_text.Length > 700) return _text.Substring(0, 700).TrimEnd() + "\u2026";
                return _text;
            }
        }

        // ---- colour ----
        public string Color {
            get { return _color; }
            set { _color = value ?? ""; OnChanged("Color"); OnChanged("HasColor"); OnChanged("TagFill"); OnChanged("TagMark"); }
        }

        public bool HasColor { get { return _color.Length > 0; } }
        public string TagFill { get { return TagValue(0); } }
        public string TagMark { get { return TagValue(1); } }

        private string TagValue(int part)
        {
            if (_color.Length == 0) return "Transparent";
            string[] pair;
            if (TagColors.TryGetValue(_color, out pair) && pair != null && pair.Length > part && !string.IsNullOrEmpty(pair[part]))
                return pair[part];
            // a colour this build has no theme entry for: no tint, stripe in the stored colour
            return part == 0 ? "Transparent" : _color;
        }

        public string Display {
            get {
                string n = NumberText;
                return (n.Length > 0 ? n + " " : "") + _text + _suffix;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string name) {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }

        // Safety net: anything that stringifies an item still gets the row text
        public override string ToString() { return Display; }
    }

    // Asks Windows to draw the title bar dark when the window is dark.
    public static class Native
    {
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        // 0.8.2 - used to have the window frame redrawn after a theme switch
        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);
    }

    // Row tooltips. A row whose text is cut short shows the whole text in a tooltip;
    // a row that already shows everything does not, so tooltips do not pop up all
    // over the list. WPF raises ToolTipOpening on the element itself only, never on
    // its parents, so a handler on the list would not see it. A class handler does:
    // it runs for every TextBlock whose tooltip is about to open, and cancels the ones
    // on row text (Tag "body" in the row template) that is not cut short. It lives
    // here rather than in script because a class handler can never be taken off again,
    // and compiled code is safe to leave registered when the script is run a second
    // time in the same console session.
    public static class RowTip
    {
        private static bool _hooked;

        public static void Hook()
        {
            if (_hooked) return;
            _hooked = true;
            System.Windows.EventManager.RegisterClassHandler(
                typeof(System.Windows.Controls.TextBlock),
                System.Windows.FrameworkElement.ToolTipOpeningEvent,
                new System.Windows.Controls.ToolTipEventHandler(OnOpening));
        }

        private static void OnOpening(object sender, System.Windows.Controls.ToolTipEventArgs e)
        {
            System.Windows.Controls.TextBlock box = sender as System.Windows.Controls.TextBlock;
            if (box == null || !"body".Equals(box.Tag as string)) return;
            try {
                if (!IsCutShort(box)) e.Handled = true;
            } catch { }
        }

        // "Cut short" means a collapsed long item or group, or a line that runs past
        // the window edge - found by measuring the same text with no width limit.
        public static bool IsCutShort(System.Windows.Controls.TextBlock box)
        {
            ClipItem item = box.DataContext as ClipItem;
            if (item == null || item.IsExpanded) return false;
            if (item.IsLong) return true;

            System.Windows.Controls.TextBlock probe = new System.Windows.Controls.TextBlock();
            probe.Text = box.Text;
            probe.FontFamily = box.FontFamily;
            probe.FontSize = box.FontSize;
            probe.FontWeight = box.FontWeight;
            probe.FontStyle = box.FontStyle;
            probe.FontStretch = box.FontStretch;
            System.Windows.Media.TextOptions.SetTextFormattingMode(probe,
                System.Windows.Media.TextOptions.GetTextFormattingMode(box));
            probe.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
            return probe.DesiredSize.Width > box.ActualWidth + 0.5;
        }
    }
}
'@ -ReferencedAssemblies PresentationCore, PresentationFramework, WindowsBase, System.Xaml
    }

# 8/25/26 - Microphone recorder.
    # Uses the MCI string interface from winmm.dll - the long-standing Windows way to
    # record a wav with no external dependency. 16 kHz mono 16-bit is what speech
    # recognition wants and keeps the upload small (~32 KB per second).
    if (-not ('CbmRecorder' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class CbmRecorder
{
    [DllImport("winmm.dll", CharSet = CharSet.Auto)]
    private static extern int mciSendString(string command, StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

    [DllImport("winmm.dll", CharSet = CharSet.Auto)]
    private static extern bool mciGetErrorString(int errorCode, StringBuilder buffer, int bufferSize);

    private static void Send(string command)
    {
        StringBuilder result = new StringBuilder(256);
        int rc = mciSendString(command, result, result.Capacity, IntPtr.Zero);
        if (rc != 0)
        {
            StringBuilder err = new StringBuilder(256);
            mciGetErrorString(rc, err, err.Capacity);
            throw new Exception(err.ToString().Trim() + "  (" + command + ")");
        }
    }

    // The requested format has to be one the driver will actually accept through the
    // legacy MCI waveform path. 16 kHz mono is ideal for speech recognition but plenty
    // of modern drivers refuse it and return "No wave device is installed that can
    // record files in the current format".
    // So negotiate: walk a ladder of formats, most-compatible first, and fall back to
    // whatever the device offers by default. Returns the format that worked.
    // Order matters. Probing a real Intel Smart Sound digital mic array showed it
    // refuses mono at 44.1 and 48 kHz but accepts stereo at both, and mono at 22.05.
    // Stereo 44.1 leads because it is the most widely accepted; 22.05 mono sits next
    // because it is a quarter the size and still well supported.
    private static readonly string[] Formats = new string[] {
        "bitspersample 16 channels 2 samplespersec 44100",
        "bitspersample 16 channels 1 samplespersec 22050",
        "bitspersample 16 channels 2 samplespersec 48000",
        "bitspersample 16 channels 1 samplespersec 44100",
        "bitspersample 16 channels 1 samplespersec 16000",
        "bitspersample 16 channels 2 samplespersec 22050",
        "bitspersample 8 channels 1 samplespersec 22050",
        null    // whatever the driver defaults to
    };

    public static string Start()
    {
        // clear anything a previous run left open
        try { Send("close cbmrec"); } catch { }

        string lastError = "no audio input device responded";

        foreach (string fmt in Formats)
        {
            try
            {
                Send("open new Type waveaudio Alias cbmrec");
                Send("set cbmrec time format ms");
                if (fmt != null) { Send("set cbmrec " + fmt); }
                Send("record cbmrec");
                return (fmt == null) ? "device default" : fmt;
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                // a refused format can leave the device half-open; start clean.
                // The brief pause matters: reopening instantly can report
                // "All wave devices that can record files in the current format
                // are in use" even though the device is merely still releasing.
                try { Send("close cbmrec"); } catch { }
                System.Threading.Thread.Sleep(120);
            }
        }

        throw new Exception(
            "None of the standard recording formats were accepted.\n\n" +
            "Last error: " + lastError + "\n\n" +
            "Check that a microphone is connected and enabled, and that " +
            "Settings > Privacy & security > Microphone allows desktop apps.");
    }

    public static void SaveAndClose(string path)
    {
        Send("stop cbmrec");
        Send("save cbmrec \"" + path + "\"");
        Send("close cbmrec");
    }

    public static void Abort()
    {
        try { Send("stop cbmrec"); } catch { }
        try { Send("close cbmrec"); } catch { }
    }
}
'@
    }

# =====================================================================
# 8/25/26 - AI ROUTER
#
# Providers differ in WHAT they can do, not only in how they are called, so this is
# a capability -> provider map rather than one global "which AI" setting:
#
#     Chat            xAI, OpenAI, Anthropic, Google
#     SpeechToText    xAI, OpenAI          (Anthropic has no speech API at all;
#                                           Google needs a two-step Files upload,
#                                           which is not wired here - see SttUrl)
#     TextToSpeech    Windows (offline, local)
#
# Keys are per PROVIDER; routing is per CAPABILITY. So speech can go to Grok while
# chat goes to Claude, each with its own key, at the same time.
#
# Nothing here is persisted. Keys live in memory for the session only, consistent
# with the rest of this tool.
#
# Adding a provider means one entry below plus a case in each of the three switch
# statements (auth header, request body, response shape). Nothing else changes.
# =====================================================================
    $Script:AiProviders = [ordered]@{
        "xai" = @{
            Display   = "Grok (xAI)"
            ChatUrl   = "https://api.x.ai/v1/chat/completions"
            ChatModel = "grok-4"
            SttUrl    = "https://api.x.ai/v1/stt"
            SttModel  = ""              # xAI's /v1/stt takes no model field
            Style     = "openai"        # OpenAI-compatible request/response shape
        }
        "openai" = @{
            Display   = "ChatGPT (OpenAI)"
            ChatUrl   = "https://api.openai.com/v1/chat/completions"
            ChatModel = "gpt-5"
            SttUrl    = "https://api.openai.com/v1/audio/transcriptions"
            SttModel  = "gpt-transcribe"
            Style     = "openai"
        }
        "anthropic" = @{
            Display   = "Claude (Anthropic)"
            ChatUrl   = "https://api.anthropic.com/v1/messages"
            ChatModel = "claude-sonnet-4-6"
            SttUrl    = ""              # no speech API; Anthropic pairs with third parties
            SttModel  = ""
            Style     = "anthropic"
        }
        "google" = @{
            Display   = "Gemini (Google)"
            ChatUrl   = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash:generateContent"
            ChatModel = "gemini-3.1-flash"
            SttUrl    = ""              # transcription requires a Files API upload first
            SttModel  = ""
            Style     = "google"
        }
    }

    # Model IDs move over time. If a call comes back "unknown model", edit the string
    # above - that is the only place any model name appears.

    Function Get-AiProvidersFor {
        param([string]$Capability)
        $out = @()
        foreach ($id in $Script:AiProviders.Keys) {
            $p = $Script:AiProviders[$id]
            if ($Capability -eq "Chat" -and -not [string]::IsNullOrEmpty($p.ChatUrl)) { $out += $id }
            if ($Capability -eq "SpeechToText" -and -not [string]::IsNullOrEmpty($p.SttUrl)) { $out += $id }
        }
        return $out
    }

    Function Get-AiKey {
        param([string]$Provider, [switch]$Force)

        if (-not $Force -and $global:AiKeys.ContainsKey($Provider) -and
            -not [string]::IsNullOrWhiteSpace($global:AiKeys[$Provider])) {
            return $global:AiKeys[$Provider]
        }

        $name = $Script:AiProviders[$Provider].Display
        $entered = Show-InputPrompt -Message "Enter your $name API key.`n`nHeld in memory for this session only - never written to disk, so it must be re-entered after a restart." -Title "API Key - $name"
        if ([string]::IsNullOrWhiteSpace($entered)) { return $null }

        $global:AiKeys[$Provider] = $entered.Trim()
        return $global:AiKeys[$Provider]
    }

    Function Invoke-AiChat {
        param([string]$Prompt)

        $providerId = $global:AiRoute["Chat"]
        $p = $Script:AiProviders[$providerId]
        $key = Get-AiKey -Provider $providerId
        if ($null -eq $key) { return $null }

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $headers = @{}
        $body = $null

        if ($p.Style -eq "openai") {
            $headers["Authorization"] = "Bearer $key"
            $body = @{ model = $p.ChatModel; messages = @(@{ role = "user"; content = $Prompt }) }
        }
        elseif ($p.Style -eq "anthropic") {
            $headers["x-api-key"] = $key
            $headers["anthropic-version"] = "2023-06-01"
            $body = @{ model = $p.ChatModel; max_tokens = 2048; messages = @(@{ role = "user"; content = $Prompt }) }
        }
        elseif ($p.Style -eq "google") {
            $headers["x-goog-api-key"] = $key
            $body = @{ contents = @(@{ parts = @(@{ text = $Prompt }) }) }
        }

        try {
            $json = $body | ConvertTo-Json -Depth 10
            $resp = Invoke-RestMethod -Uri $p.ChatUrl -Method Post -Headers $headers -Body $json -ContentType "application/json"
        } catch {
            $detail = $_.Exception.Message
            if ($null -ne $_.Exception.Response) {
                try {
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $detail = $detail + "`n`n" + $reader.ReadToEnd()
                } catch { }
            }
            # a rejected key should not linger and fail every later attempt
            if ($detail -match '401|invalid.?api.?key|unauthor') { $global:AiKeys[$providerId] = $null }
            [System.Windows.MessageBox]::Show($Window, "$($p.Display) request failed.`n`n$detail",
                "Ask AI", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $null
        }

        if ($p.Style -eq "openai")    { return $resp.choices[0].message.content }
        if ($p.Style -eq "anthropic") { return (@($resp.content) | Where-Object { $_.type -eq "text" } | Select-Object -First 1).text }
        if ($p.Style -eq "google")    { return $resp.candidates[0].content.parts[0].text }
        return $null
    }

    Function Invoke-AiTranscribe {
        param([string]$WavPath)

        $providerId = $global:AiRoute["SpeechToText"]
        $p = $Script:AiProviders[$providerId]

        if ([string]::IsNullOrEmpty($p.SttUrl)) {
            [System.Windows.MessageBox]::Show($Window,
                "$($p.Display) does not offer a speech-to-text API.`n`nChoose a different provider under Tools > AI Settings.",
                "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $null
        }

        $key = Get-AiKey -Provider $providerId
        if ($null -eq $key) { return $null }

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Add-Type -AssemblyName System.Net.Http

        $client = $null
        $form = $null
        try {
            $client = New-Object System.Net.Http.HttpClient
            $client.Timeout = [TimeSpan]::FromSeconds(180)
            $client.DefaultRequestHeaders.Authorization =
                New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $key)

            $form = New-Object System.Net.Http.MultipartFormDataContent

            # NOTE: xAI requires the file part AFTER every other field. OpenAI does not
            # care, so adding the file last satisfies both.
            if (-not [string]::IsNullOrEmpty($p.SttModel)) {
                $form.Add((New-Object System.Net.Http.StringContent($p.SttModel)), "model")
            }
            $form.Add((New-Object System.Net.Http.StringContent("en")), "language")
            if ($providerId -eq "xai") {
                $form.Add((New-Object System.Net.Http.StringContent("true")), "format")
            }

            $bytes = [System.IO.File]::ReadAllBytes($WavPath)
            $fileContent = New-Object System.Net.Http.ByteArrayContent -ArgumentList (,$bytes)
            $fileContent.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("audio/wav")
            $form.Add($fileContent, "file", "note.wav")

            $response = $client.PostAsync($p.SttUrl, $form).Result
            $bodyText = $response.Content.ReadAsStringAsync().Result

            if (-not $response.IsSuccessStatusCode) {
                if ([int]$response.StatusCode -eq 401) { $global:AiKeys[$providerId] = $null }
                [System.Windows.MessageBox]::Show($Window,
                    "$($p.Display) transcription failed.`n`n$([int]$response.StatusCode) $($response.ReasonPhrase)`n`n$bodyText",
                    "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $null
            }

            # xAI and OpenAI both return the transcript in a top-level "text" field
            return ($bodyText | ConvertFrom-Json).text

        } catch {
            [System.Windows.MessageBox]::Show($Window, "Transcription failed.`n`n$($_.Exception.Message)",
                "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $null
        } finally {
            if ($null -ne $form)   { $form.Dispose() }
            if ($null -ne $client) { $client.Dispose() }
        }
    }

    Function Invoke-AiSpeak {
        param([string]$Text)
        # Local Windows synthesizer. A cloud TTS provider would slot in here exactly
        # the way the chat and STT dispatchers above handle multiple providers.
        try {
            Add-Type -AssemblyName System.Speech -ErrorAction Stop
            $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $synth.SpeakAsync($Text) | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show($Window, "Speech playback failed.`n`n$($_.Exception.Message)",
                "Speak", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
    }

    # 11/7/26 - Collapsible long items.
    # Anything longer than CollapseThreshold gets a one-line Preview: whitespace and
    # line breaks flattened, cut to PreviewLength, with the full length noted.
    # 11/22/26 - the preview is only the snippet now. The line number and the time
    # stamp are columns of their own, so nothing is spliced around it any more, and the
    # length ("348 chars") moved to the small note at the right of the row.
    $Script:CollapseThreshold = 150
    $Script:PreviewLength = 60

    # 11/22/26 - what a masked item holds instead of the secret, and the row icons.
    # The icons are characters of the Windows icon font (Segoe Fluent Icons on Windows
    # 11, Segoe MDL2 Assets on Windows 10 - same code points in both), written as
    # numbers so this file stays plain ASCII.
    $Script:MaskText  = "*************"
    $Script:IconVoice = [string][char]0xE720
    $Script:IconAi    = [string][char]0xE99A
    $Script:IconLink  = [string][char]0xE71B
    $Script:IconLock  = [string][char]0xE72E

    # The title inside an Insert Topic divider - exactly what Insert Topic writes: 11
    # dashes, the title, 33 dashes - or an empty string when the text is not one.
    Function Get-TopicTitle {
        param([string]$Text)
        if ($Text -match '^-{11}([^\r\n]+)-{33}\z') {
            $title = $matches[1].Trim()
            if ($title.Trim('-').Trim().Length -gt 0) { return $title }
        }
        return ""
    }

    # Works out everything about a row that depends on its text: the collapsed snippet,
    # the note at the right, the icon, and whether it is a topic heading or a masked
    # item. Called whenever an item is created or its text, group or parts change.
    Function Update-ClipPreview {
        param($Item)

        $text = [string]$Item.Text
        $topicTitle = Get-TopicTitle -Text $text
        $preview = ""
        $detail = ""
        $caption = ""
        $icon = ""
        $isTopic = $false
        $isDim = $false

        if ($Item.IsGroup) {
            # 11/8/26 - a group always collapses, however short, and its preview strings
            # its items together with " / ". 11/22/26 - the title leads the snippet and
            # the item count is the note at the right. A topic divider inside the group
            # shows as its title, so its dashes do not use up the snippet.
            $parts = @($Item.Parts)
            $preview = (@($parts | ForEach-Object {
                $partText = [string]$_.Text
                $partTopic = Get-TopicTitle -Text $partText
                if ($partTopic.Length -gt 0) { $partTopic } else { ($partText -replace '\s+', ' ').Trim() }
            })) -join " / "
            if ($preview.Length -gt $Script:PreviewLength) {
                $preview = $preview.Substring(0, $Script:PreviewLength).TrimEnd() + [char]0x2026
            }
            if ($preview.Length -eq 0) { $preview = "(empty)" }
            if (-not [string]::IsNullOrWhiteSpace($Item.GroupTitle)) {
                $preview = "$($Item.GroupTitle):  $preview"
            }
            $detail = "$($parts.Count) items"
        }
        elseif ($text -ceq $Script:MaskText) {
            $icon = $Script:IconLock
            $detail = "masked"
            $isDim = $true
        }
        elseif ($topicTitle.Length -gt 0) {
            $isTopic = $true
            $caption = $topicTitle
        }
        else {
            if ($text.StartsWith("[Voice] ", [StringComparison]::Ordinal)) { $icon = $Script:IconVoice }
            elseif ($text.StartsWith("[AI] ", [StringComparison]::Ordinal)) { $icon = $Script:IconAi }
            elseif ($text.Length -le 2048 -and $text -match '^\s*https?://\S+\s*\z') { $icon = $Script:IconLink }

            if ($text.Length -gt $Script:CollapseThreshold) {
                $preview = ($text -replace '\s+', ' ').Trim()
                if ($preview.Length -gt $Script:PreviewLength) {
                    $preview = $preview.Substring(0, $Script:PreviewLength).TrimEnd()
                }
                $preview = $preview + [char]0x2026
                $detail = "{0:N0} chars" -f $text.Length
            }
        }

        $Item.Caption = $caption
        $Item.Icon = $icon
        $Item.Detail = $detail
        $Item.IsDim = $isDim
        $Item.IsTopic = $isTopic
        # 0.8.1 - a topic heading has no star, so a favorite that is edited into a
        # heading stops being one rather than staying a favorite nobody can see
        if ($isTopic -and $Item.IsFavorite) { $Item.IsFavorite = $false }
        $Item.Preview = $preview
    }

    # 11/22/26 - an item is made from its parts, not from a formatted row:
    #   -Numbered  show a line number (callers pass the Tools > Line Numbers state)
    #   -Suffix    the time stamp text, " : dddd MM/dd/yyyy HH:mm:ss", or nothing
    Function New-ClipItem {
        param([string]$Text, [bool]$Numbered = $false, [string]$Suffix = "", [string]$Color = "")
        $i = New-Object Cbm082.ClipItem
        $i.Text = $Text
        $i.Captured = Get-Date
        $i.Numbered = $Numbered
        $i.Suffix = $Suffix
        $i.Color = $Color
        Update-ClipPreview -Item $i
        return $i
    }

# 8/24/26 - Replacement for [Microsoft.VisualBasic.Interaction]::InputBox
    # The VB InputBox cannot be given an owner window, so whenever Stay On Top is
    # enabled it opens BEHIND the main window. Because it is modal, the main window
    # stops accepting input and the app appears frozen.
    # This WPF version is owned by the main window, so it always sits in front.
    # 11/22/26 - -MultiLine (Edit item uses it): the box wraps, Enter starts a new
    # line and a scroll bar appears when needed. The window can be resized and opens
    # just tall enough for the text, from 4 lines up to about 16. Ctrl+Enter is OK
    # and Esc is Cancel; the caret starts at the end of the text with nothing
    # selected, so a stray key press cannot wipe a long item. Without the switch
    # this is the same one-line prompt as before.
    Function Show-InputPrompt {
        param(
            [string]$Message,
            [string]$Title,
            [string]$DefaultText = "",
            [switch]$MultiLine
        )

        $prompt = New-Object System.Windows.Window
        $prompt.Title = $Title
        $prompt.Width = 430
        $prompt.SizeToContent = "Height"
        $prompt.ResizeMode = "NoResize"
        $prompt.ShowInTaskbar = $false
        if ($Script:AppIcon -ne $null) { $prompt.Icon = $Script:AppIcon }

        # Owning the prompt keeps it above the main window even when Topmost is set
        if ($Window -ne $null -and $Window.IsLoaded) {
            $prompt.Owner = $Window
            $prompt.WindowStartupLocation = "CenterOwner"
        } else {
            $prompt.WindowStartupLocation = "CenterScreen"
        }
        $prompt.Topmost = $true

        # 11/22/26 - three rows: message, box, buttons. The box row takes whatever
        # height is left, which is what lets the multi-line box grow with the window.
        $panel = New-Object System.Windows.Controls.Grid
        $panel.Margin = New-Object System.Windows.Thickness(14)
        foreach ($rowHeight in 'Auto', '*', 'Auto') {
            $rowDef = New-Object System.Windows.Controls.RowDefinition
            if ($rowHeight -eq '*') {
                $rowDef.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
            } else {
                $rowDef.Height = [System.Windows.GridLength]::Auto
            }
            [void]$panel.RowDefinitions.Add($rowDef)
        }

        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $Message
        $label.TextWrapping = "Wrap"
        $label.Margin = New-Object System.Windows.Thickness(0,0,0,10)
        [System.Windows.Controls.Grid]::SetRow($label, 0)
        [void]$panel.Children.Add($label)

        $box = New-Object System.Windows.Controls.TextBox
        $box.Text = $DefaultText
        $box.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        [System.Windows.Controls.Grid]::SetRow($box, 1)
        [void]$panel.Children.Add($box)

        $bottom = New-Object System.Windows.Controls.DockPanel
        $bottom.LastChildFill = $true
        [System.Windows.Controls.Grid]::SetRow($bottom, 2)

        $buttons = New-Object System.Windows.Controls.StackPanel
        $buttons.Orientation = "Horizontal"
        [System.Windows.Controls.DockPanel]::SetDock($buttons, [System.Windows.Controls.Dock]::Right)

        $ok = New-Object System.Windows.Controls.Button
        $ok.Content = "OK"
        $ok.Width = 80
        $ok.IsDefault = $true
        $ok.Margin = New-Object System.Windows.Thickness(0,0,8,0)
        $ok.Add_Click({ $prompt.DialogResult = $true })

        $cancelBtn = New-Object System.Windows.Controls.Button
        $cancelBtn.Content = "Cancel"
        $cancelBtn.Width = 80
        $cancelBtn.IsCancel = $true

        [void]$buttons.Children.Add($ok)
        [void]$buttons.Children.Add($cancelBtn)
        [void]$bottom.Children.Add($buttons)

        # left of the buttons: the key reminder, shown for the multi-line box only
        $hint = New-Object System.Windows.Controls.TextBlock
        $hint.Foreground = [System.Windows.SystemColors]::GrayTextBrush
        $hint.FontSize = 11
        $hint.VerticalAlignment = "Center"
        $hint.TextTrimming = "CharacterEllipsis"
        $hint.Margin = New-Object System.Windows.Thickness(0,0,12,0)
        [void]$bottom.Children.Add($hint)
        [void]$panel.Children.Add($bottom)

        if ($MultiLine) {
            $prompt.Width = 560
            $prompt.MinWidth = 340
            $prompt.MinHeight = 200
            $prompt.ResizeMode = "CanResizeWithGrip"

            $box.AcceptsReturn = $true
            $box.TextWrapping = "Wrap"
            $box.VerticalScrollBarVisibility = "Auto"
            $box.VerticalContentAlignment = "Top"
            $box.Padding = New-Object System.Windows.Thickness(3,2,3,2)
            # opening height: no less than about 4 lines, no more than about 16.
            # The upper limit is lifted once the window has taken its first size.
            $box.MinHeight = 76
            $box.MaxHeight = 280

            # Enter belongs to the box now, so OK is no longer the Enter button
            $ok.IsDefault = $false
            $hint.Text = "Enter: new line     Ctrl+Enter: OK     Esc: Cancel"

            $box.Add_PreviewKeyDown({
                if ($_.Key -eq [System.Windows.Input.Key]::Return -and
                    ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
                    $_.Handled = $true
                    $prompt.DialogResult = $true
                }
            })
        }

        $prompt.Content = $panel

        # Put the caret in the textbox as soon as the window paints
        $prompt.Add_ContentRendered({
            if ($MultiLine) {
                # the window sized itself to the text; from here on the user sizes it.
                # The 2 extra units absorb pixel rounding at 125% / 150% display
                # scaling, which otherwise leaves the box a fraction too short and
                # brings up a scroll bar with nothing to scroll.
                $opened = [Math]::Ceiling($prompt.ActualHeight) + 2
                $prompt.SizeToContent = "Manual"
                $prompt.Height = $opened
                $box.MaxHeight = [double]::PositiveInfinity
                [void]$box.Focus()
                $box.CaretIndex = $box.Text.Length
                $box.ScrollToEnd()
            } else {
                [void]$box.Focus()
                $box.SelectAll()
            }
        })

        if ($prompt.ShowDialog() -eq $true) { return $box.Text }
        # 11/8/26 - $null rather than "" so a caller can tell Cancel from a blank OK.
        # Existing callers test with IsNullOrWhiteSpace, which treats both the same.
        return $null
    }

# Write file to filing system........
    # the function below is exclusively for keyboard shortcut ctl + W to write file to file system using the first item selected as the default file name.
    Function Create_File_Selected_keyboard{ 
    
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        $NameFilter = "Text Files (*.txt)|*.txt"
        $Dialog.FileName = $listbox.SelectedItems[0].Text # Get the first listbox item value
        $global:firstSelectedItem = $Dialog.FileName
        $Dialog.Filter = $NameFilter
        
        if($Dialog.ShowDialog() -eq 'Ok'){
        
            @($listbox.SelectedItems) | ForEach-Object { # 
                        
                # 6/2/23
                $save_data += (Get-ExportText -Item $_) + "`n"    # 0.8.1: "* " in front of a favorite
                $save_data | Out-File $Dialog.FileName
            }

        }
    }

        # Append selected text to existing file
        # the function below is exclusively for keyboard shortcut ctl + W to write file to file system using the first item selected as the default file name.
        Function Append_File_Selected_keyboard # 
        {  #6/3/23

            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
            $Dialog = New-Object System.Windows.Forms.SaveFileDialog
            $NameFilter = "Text Files (*.txt)|*.txt"
            $Dialog.Filter = $NameFilter
            $Dialog.FileName = $global:firstSelectedItem
            
            if($Dialog.ShowDialog() -eq 'Ok'){
            
                 @($listbox.SelectedItems) | ForEach-Object {     
                    
                    $save_data = Get-ExportText -Item $_    # 0.8.1: "* " in front of a favorite
                    $save_data | Out-File $Dialog.FileName -Append
                }
            }

            #Show-AutoClosingMessageBox -Message "Done......" -Title "Info" -Timeout 3000 # ozzzy 6/3/2023

        }

# Open the Review Log 
        # this function is exclusively for keyboard shortcut ctl + O to write file to file system using the first item selected as the default file name.
        Function OpenLog_File_Selected_keyboard
        {  
            #6/4/23 - coming soon

            $scriptPath = "$env:USERPROFILE\Documents" #split-path -parent $PSCommandPath
            $global:ReviewLog = $scriptPath + "\ClipBoardManager-CBM-ReviewLog.txt"

            $notepad = New-Object -ComObject "WScript.Shell"
            $notepad.run("notepad.exe $global:ReviewLog")

        }

# Append selected text to existing file
        # the function below is exclusively for keyboard shortcut ctl + L to write file to file system Review Log / My Documents
        Function Log_File_Selected_keyboard # Log_File_Selected_keyboard
        {  

            # 6/3
            $scriptPath = "$env:USERPROFILE\Documents" #split-path -parent $PSCommandPath
            $global:ReviewLog = $scriptPath + "\ClipBoardManager-CBM-ReviewLog.txt"
         
# 3/5
$AddComment = Show-InputPrompt -Message "Insert a comment on top of log entry items. `n(or leave blank)" -Title "Insert Comment" -DefaultText $defaultText

            # Add comment to log entries....... # 6/5/23 - Issue where always goes back to filter box - hrrrrrrrrrrrr
            $AddComment = "`n$AddComment"
            $AddComment | Out-File $global:ReviewLog -Append
            
            if (-not (Test-Path -Path $global:ReviewLog)) {
                New-Item -ItemType File -Path $global:ReviewLog -Force

                @($listbox.SelectedItems) | ForEach-Object {     
                    
                    $save_data = Get-ExportText -Item $_    # 0.8.1: "* " in front of a favorite
                    $save_data | Out-File $global:ReviewLog -Append
                }

            }else
            {
                @($listbox.SelectedItems) | ForEach-Object {     
                    
                    $save_data = Get-ExportText -Item $_    # 0.8.1: "* " in front of a favorite
                    $save_data | Out-File $global:ReviewLog -Append
                }

            }

            $AddComment  = $null
            $return
                
        }

      # ozzzy 6/3/2023
      function Show-AutoClosingMessageBox 
      {
        param(
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [Parameter(Mandatory=$true)]
            [string]$Title,
            [Parameter(Mandatory=$true)]
            [int]$Timeout
        )
        [AutoClosingMessageBox]::Show($Message, $Title, $Timeout)
    }

# 6/3/23
#OpenSelfClosingForm 

# 11/22/26 (0.8.1) - the selected items as one text, one item per line. With
# -MarkFavorites a favorite gets "* " in front (see Get-ExportText).
Function Get-SelectedText {
        param([switch]$MarkFavorites)
        $lines = @($listbox.SelectedItems) | ForEach-Object {
            if ($MarkFavorites) { Get-ExportText -Item $_ } else { $_.Text }
        }
        return (($lines -join "`n").Trim())
}

# 0.8.1 - -MarkFavorites is passed by Open Selected in NOTEPAD only, because what
# Notepad shows is whatever this puts on the clipboard. Copy never passes it.
Function Set-ClipBoard {
        param([switch]$MarkFavorites)
        # 8/25/26 - items carry their raw text, so the old "^\d+\s+" strip is gone
        # along with the bug where it truncated content that began with digits.
        $contentToClipboard = Get-SelectedText -MarkFavorites:$MarkFavorites
        $Script:CopiedText = $contentToClipboard

        if ($global:bron -eq "bron") {
            [Windows.Clipboard]::SetText($contentToClipboard + "<br>")
        } else {
            [Windows.Clipboard]::SetText($contentToClipboard)
        }
}

Function Set-ClipBoardooo { # 9/22
    $copied = (@($listbox.SelectedItems) | ForEach-Object { $_.Text }) -join "`n"
    $Script:CopiedText = $copied
    [Windows.Clipboard]::SetText("ooo " + $copied.Trim())
}

Function Clear-Viewer {
        [void]$Script:ObservableCollection.Clear()
        [Windows.Clipboard]::Clear()
    }

# 8/24/26 - Per-item row colors.
# Colour lives on the item itself (ClipItem.Color) and a DataTrigger in the XAML
# applies it. No side table, no manual repainting, and a colour follows its item
# through filtering, sorting and edits.
# 11/22/26 - these four values are still what an item stores (and what travels to a
# new instance), but they are no longer painted as they are. Each one is looked up in
# the theme and drawn as a soft row tint plus a marker stripe - see Get-UiColors.
    $Script:ColorPalette = @{
        "Red"    = "#FFB3B3"
        "Green"  = "#B6E8B6"
        "Yellow" = "#FFF2A8"
        "Blue"   = "#B3D9FF"
    }

# 8/24/26 - Masking helpers.
# Test-Masked walks the term list and escapes each term before matching, so
# clipboard text containing regex metacharacters ( C:\temp\*.txt , (a|b) ) is
# compared literally instead of being treated as a pattern.
# -match is case-insensitive by default, which matches what the mask dialog promises.
# 8/24/26 - Reverse the display formatting to recover the raw clipboard text.
    # List rows are stored as "<number> <text>" and optionally " : <timestamp>".
    # Masking previously registered the WHOLE display string as the term, so with
    # line numbers on (the default) the term "42 hunter2" could never match the
    # raw clipboard text "hunter2" - masking silently never fired.
    Function Get-ItemText {
        param([string]$Item)

        $t = $Item
        if ([string]::IsNullOrEmpty($t)) { return "" }

        # strip a leading line number
        if ($t -match '^\d+\s+(.*)$') { $t = $matches[1] }

        # strip a trailing " : dddd MM/dd/yyyy HH:mm:ss" timestamp
        if ($t -match '^(.*?)\s+:\s+\S+\s+\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}\s*$') { $t = $matches[1] }

        return $t.Trim()
    }

    Function Add-MaskedTerm {
        param([string]$Term)
        if ([string]::IsNullOrWhiteSpace($Term)) { return }
        if ($global:MaskedTerms -notcontains $Term) {
            [void]$global:MaskedTerms.Add($Term)
        }
    }

    Function Test-Masked {
        param([string]$Text)
        if ($global:MaskedTerms -eq $null -or $global:MaskedTerms.Count -eq 0) { return $false }
        if ([string]::IsNullOrEmpty($Text)) { return $false }
        foreach ($term in $global:MaskedTerms) {
            if ([string]::IsNullOrWhiteSpace($term)) { continue }
            if ($Text -match [regex]::Escape($term)) { return $true }
        }
        return $false
    }

# 8/25/26 - Single entry point for adding a captured item.
    # Masked items store the MASK as their Text, never the secret, so the filter box
    # cannot be used to find a masked value.
    Function Add-ClipboardItem {
        param([string]$Text)

        $itemText = $Text
        if (Test-Masked -Text $Text) { $itemText = $Script:MaskText }

        $suffix = ""
        if ($AddTime_Menu.IsChecked) { $suffix = " : " + (Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss") }

        # 11/22/26 - no line number is worked out here any more. The row only records
        # whether it is numbered; the number itself is its position in the list.
        [void]$Script:ObservableCollection.Add((New-ClipItem -Text $itemText -Numbered ([bool]$LineNumbers_Menu.IsChecked) -Suffix $suffix))

        # Scroll to the bottom to show the latest item
        $listbox.ScrollIntoView($listbox.Items[$listbox.Items.Count-1])
    }

# 11/22/26 (0.8.1) - Favorites.
    # A favorite is just a flag on the item (ClipItem.IsFavorite); the star in the row
    # follows it through PropertyChanged. Like colours it lives in memory only.

    # The text of an item as it is written to a file, the Review Log, Notepad or Edge:
    # a favorite gets "* " in front. Copy does not use this, so the clipboard stays clean.
    Function Get-ExportText {
        param($Item)
        if ($null -eq $Item) { return "" }
        if ($Item.IsFavorite) { return "* " + [string]$Item.Text }
        return [string]$Item.Text
    }

    # Sets or clears the star on the given rows (topic headings are skipped), then
    # brings the Show Only Favorites view and the status strip up to date.
    Function Set-ItemsFavorite {
        param($Items, [bool]$Favorite)
        foreach ($i in @($Items)) {
            if ($null -ne $i -and -not $i.IsTopic) { $i.IsFavorite = $Favorite }
        }
        if ($Script:FavoritesOnly) { Update-ListFilter } else { Update-StatusStrip }
    }

    # The one place the list filter is set: the filter box text and, when it is
    # switched on, Show Only Favorites. A row has to pass both.
    $Script:FavoritesOnly = $false
    Function Update-ListFilter {
        [System.Windows.Data.CollectionViewSource]::GetDefaultView($Listbox.ItemsSource).Filter = [Predicate[Object]]{
            Try {
                if ($Script:FavoritesOnly -and -not $args[0].IsFavorite) { return $false }
                # 8/25/26 - filter the raw text, not the formatted row
                $args[0].Text -match [regex]::Escape($InputBox.Text)
            } Catch {
                $True
            }
        }

        # 11/22/26 - the strip shows "3 of 40 items" while a filter hides rows
        Update-StatusStrip
    }

    # Show Only Favorites has two switches (the star button and the Tools item); both
    # come here so they always agree.
    Function Set-FavoritesOnly {
        param([bool]$On)
        $Script:FavoritesOnly = $On
        $FavOnly_Button.IsChecked = $On
        $FavOnly_Menu.IsChecked = $On
        if ($On) { $FavOnly_Button.ToolTip = "Showing favorites only - click to show everything" }
        else     { $FavOnly_Button.ToolTip = "Show only favorites" }
        Update-ListFilter
    }

    Function Set-ItemColor {
        param([string]$ColorName)

        # 8/25/26 - colour is a property on the item now. The DataTrigger in the XAML
        # picks it up through PropertyChanged, so there is nothing to repaint by hand
        # and no side table to keep in sync.
        foreach ($item in @($listbox.SelectedItems)) {
            if ($ColorName -eq "Clear") {
                $item.Color = ""
            } else {
                $item.Color = $Script:ColorPalette[$ColorName]
            }
        }
    }
    
    # 11/22/26 - Colours.
    # 11/23/26 (0.8.2) - the window opens light. Tools > Theme, or -Theme on the
    # command line, picks Light, Dark or Auto. Auto follows the Windows light / dark
    # app setting (Settings > Personalization > Colors > Choose your mode) as it is at
    # that moment; there, high contrast wins over both and takes its colours from
    # Windows itself.
    # Every colour the window uses is in this one table. The XAML below turns each
    # entry into a brush of the same name and never states a colour of its own, so
    # changing the look means editing values here and nowhere else.
    #   Chrome / Surface        header and status strip / list background
    #   Line                    hairlines and topic rules
    #   Text / Muted / Disabled main text / numbers, icons, notes / unavailable items
    #   Accent                  topic titles, focus ring, check marks
    #   Sel* / Hover*           selected row and row under the pointer
    #   Input* / Menu*          filter box / menus and tooltips
    #   MenuEdge / MenuIcon     menu border / menu icons (MenuHotIcon: icon, check mark
    #                           and left bar of the item under the pointer)
    #   StripeBg                the soft stripe on every other row
    #   StarOn / StarSel        a favorite's star / the same on a selected row
    #   Thumb*                  scroll bar
    #   Ok / Warn               status dot: capturing / paused
    #   <Colour>Fill / Mark     Assign Color: row tint and marker stripe
    #   ShadowOp                opacity of the menu shadow
    Function Get-UiColors {
        param([string]$Mode)

        $light = @{
            Chrome = '#F6F7F9'; Surface = '#FFFFFF'; Line = '#E3E6EA'
            Text = '#1C2026'; Muted = '#667085'; Disabled = '#A3AAB8'; Accent = '#2563EB'
            SelBg = '#E3ECFF'; SelLine = '#B3CAF9'; SelText = '#1C2026'; SelMuted = '#56627A'; SelNum = '#2563EB'
            HoverBg = '#EFF1F5'; HoverLine = '#CBD1DB'
            InputBg = '#FFFFFF'; InputLine = '#CFD4DC'
            MenuBg = '#FFFFFF'; MenuLine = '#CDD2DA'; MenuHotBg = '#DFE9FF'; MenuHotText = '#1C2026'; MenuHotMuted = '#56627A'
            MenuEdge = '#4C82F0'; MenuIcon = '#2563EB'; MenuHotIcon = '#2563EB'; StripeBg = '#F2F4F8'
            Thumb = '#C3C8D1'; ThumbHot = '#99A1AE'; Ok = '#2F9E63'; Warn = '#D98A00'; ShadowOp = '0.30'
            RedFill = '#FDE8E8'; RedMark = '#E5484D'; GreenFill = '#E2F6EA'; GreenMark = '#2F9E63'
            YellowFill = '#FFF5CC'; YellowMark = '#D9A400'; BlueFill = '#E3EEFF'; BlueMark = '#3B82F6'
            StarOn = '#E6A700'; StarSel = '#E6A700'
        }
        $dark = @{
            Chrome = '#24272C'; Surface = '#1B1D21'; Line = '#33373E'
            Text = '#E7E9EC'; Muted = '#9AA3B2'; Disabled = '#6B7280'; Accent = '#7AA7FF'
            SelBg = '#25365A'; SelLine = '#3E5A97'; SelText = '#F2F4F7'; SelMuted = '#AEB8C8'; SelNum = '#7AA7FF'
            HoverBg = '#2A2E35'; HoverLine = '#4A505A'
            InputBg = '#17191C'; InputLine = '#3B4048'
            MenuBg = '#2A2D33'; MenuLine = '#3D424A'; MenuHotBg = '#2F426C'; MenuHotText = '#F2F4F7'; MenuHotMuted = '#AEB8C8'
            MenuEdge = '#5F8FEA'; MenuIcon = '#7AA7FF'; MenuHotIcon = '#9DBEFF'; StripeBg = '#23262C'
            Thumb = '#4A4F58'; ThumbHot = '#6A717C'; Ok = '#4CC38A'; Warn = '#F2B33D'; ShadowOp = '0.60'
            RedFill = '#3D2428'; RedMark = '#FF6369'; GreenFill = '#1D3A2B'; GreenMark = '#4CC38A'
            YellowFill = '#3F3614'; YellowMark = '#F2C230'; BlueFill = '#1D304D'; BlueMark = '#6CB0FF'
            StarOn = '#F5C431'; StarSel = '#F5C431'
        }

        # Auto (or anything unrecognised): ask Windows
        if ($Mode -ne 'Light' -and $Mode -ne 'Dark' -and $Mode -ne 'Contrast') {
            $Mode = 'Light'
            try {
                if ([System.Windows.SystemParameters]::HighContrast) {
                    $Mode = 'Contrast'
                } else {
                    $personalize = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -ErrorAction Stop
                    if ($personalize.AppsUseLightTheme -eq 0) { $Mode = 'Dark' }
                }
            } catch { }    # no such setting (older Windows 10, some servers): stay light
        }

        if ($Mode -eq 'Dark') {
            $colors = $dark
        } elseif ($Mode -eq 'Contrast') {
            # High contrast: nothing of ours. Backgrounds, text, selection and links all
            # come from the Windows contrast theme; only the Assign Color marker stripes
            # keep a colour, picked to suit a dark or a light contrast theme.
            $sys = [System.Windows.SystemColors]
            $window = $sys::WindowColor
            $back = $window.ToString()
            $fore = $sys::WindowTextColor.ToString()
            $pick = $sys::HighlightColor.ToString()
            $pickText = $sys::HighlightTextColor.ToString()
            if ((0.299 * $window.R + 0.587 * $window.G + 0.114 * $window.B) -lt 128) { $colors = $dark } else { $colors = $light }
            foreach ($k in 'Chrome','Surface','InputBg','MenuBg','HoverBg','StripeBg','RedFill','GreenFill','YellowFill','BlueFill') { $colors[$k] = $back }
            foreach ($k in 'Text','Muted','Line','InputLine','MenuLine','MenuEdge','MenuIcon','Thumb','ThumbHot','Ok','Warn','StarOn') { $colors[$k] = $fore }
            foreach ($k in 'SelBg','SelLine','HoverLine','MenuHotBg') { $colors[$k] = $pick }
            foreach ($k in 'SelText','SelMuted','SelNum','MenuHotText','MenuHotMuted','MenuHotIcon','StarSel') { $colors[$k] = $pickText }
            $colors.Accent = $sys::HotTrackColor.ToString()
            $colors.Disabled = $sys::GrayTextColor.ToString()
            $colors.ShadowOp = '0'
        } else {
            $colors = $light
        }

        $colors.Name = $Mode
        $colors.IsDark = ($Mode -eq 'Dark')
        return $colors
    }

    # 11/23/26 (0.8.2) - $Script:ThemeChoice is what was asked for (Light, Dark or
    # Auto); $Ui is the colour table that choice came to. With no -Theme on the command
    # line the window opens light, unless Windows is in high contrast mode.
    $Script:ThemeChoice = [string]$Runspacehash.Theme
    if (@('Light', 'Dark', 'Auto', 'Contrast') -notcontains $Script:ThemeChoice) { $Script:ThemeChoice = 'Light' }
    if (-not $Runspacehash.ThemeGiven) {
        try { if ([System.Windows.SystemParameters]::HighContrast) { $Script:ThemeChoice = 'Auto' } } catch { }
    }
    $Script:Ui = Get-UiColors -Mode $Script:ThemeChoice

    # 11/22/26 - Window icon. A WPF window with no icon of its own shows the icon of
    # the program hosting it, which here is powershell.exe. This draws a small
    # clipboard instead (blue board, dark clip, white sheet with three lines) on a
    # 16 x 16 grid and renders it at every size Windows asks for, so the title bar,
    # the taskbar and Alt+Tab each get a sharp copy rather than a stretched one.
    # The renders pass through a multi-page TIFF held in memory, because WPF only
    # picks the best-fitting size from an image that came out of a decoder. Nothing
    # is written to disk. Returns $null if anything fails; the windows then simply
    # keep the default icon.
    Function New-AppIcon {
        try {
            $brushes = New-Object System.Windows.Media.BrushConverter
            $board = $brushes.ConvertFromString('#2F6FED')
            $clip  = $brushes.ConvertFromString('#1B3A8C')
            $sheet = $brushes.ConvertFromString('#FFFFFF')
            $ink   = $brushes.ConvertFromString('#5B8DEF')

            $encoder = New-Object System.Windows.Media.Imaging.TiffBitmapEncoder
            foreach ($size in 16, 20, 24, 28, 32, 40, 48, 56, 64, 96, 128, 256) {
                $visual = New-Object System.Windows.Media.DrawingVisual
                $dc = $visual.RenderOpen()
                $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform(($size / 16.0), ($size / 16.0))))
                $dc.DrawRoundedRectangle($board, $null, (New-Object System.Windows.Rect(2, 2, 12, 13)), 2, 2)
                $dc.DrawRoundedRectangle($sheet, $null, (New-Object System.Windows.Rect(4, 5, 8, 8)), 0.5, 0.5)
                $dc.DrawRoundedRectangle($clip, $null, (New-Object System.Windows.Rect(5, 1, 6, 3)), 1, 1)
                $dc.DrawRectangle($ink, $null, (New-Object System.Windows.Rect(6, 7, 4, 1)))
                $dc.DrawRectangle($ink, $null, (New-Object System.Windows.Rect(6, 9, 4, 1)))
                $dc.DrawRectangle($ink, $null, (New-Object System.Windows.Rect(6, 11, 2, 1)))
                $dc.Pop()
                $dc.Close()
                $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($size, $size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
                $bitmap.Render($visual)
                $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            }

            # the stream stays referenced for as long as the icon is in use
            $Script:AppIconStream = New-Object System.IO.MemoryStream
            $encoder.Save($Script:AppIconStream)
            $Script:AppIconStream.Position = 0
            $decoder = New-Object System.Windows.Media.Imaging.TiffBitmapDecoder($Script:AppIconStream,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
                [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
            return $decoder.Frames[0]
        } catch {
            return $null
        }
    }

    #Build the GUI
    # 11/22/26 - New layout: header (filter box + Tools button), list, status strip.
    # This here-string is expandable on purpose: every dollar-parenthesis Ui reference
    # below is replaced by a colour from the table above before the XAML is parsed.
    # Keep any other dollar sign or backtick out of the markup. Icons are characters of
    # the Windows icon font, written as &#x....; entities so the file stays ASCII.
    [xml]$xaml = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        x:Name="Window" Title="Clip Board Manager 0.8.2" WindowStartupLocation="CenterScreen"
        Width="420" Height="540" MinWidth="260" MinHeight="200" ShowInTaskbar="True"
        Background="{DynamicResource Surface}" Foreground="{DynamicResource Text}" FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
        <Window.Resources>
            <!-- 0.8.2: the colours themselves, one holder for each brush below ("Chrome" reads
                 "ChromeSource"). Nothing draws with a holder. Each brush is bound to its
                 holder, and so is every copy of the brush that WPF makes for a style or a
                 trigger, wherever that copy ends up (a menu is not part of the window). So
                 when Set-UiTheme in the script changes the holders, everything follows. -->
            <SolidColorBrush x:Key="ChromeSource" Color="$($Ui.Chrome)"/>
            <SolidColorBrush x:Key="SurfaceSource" Color="$($Ui.Surface)"/>
            <SolidColorBrush x:Key="LineSource" Color="$($Ui.Line)"/>
            <SolidColorBrush x:Key="TextSource" Color="$($Ui.Text)"/>
            <SolidColorBrush x:Key="MutedSource" Color="$($Ui.Muted)"/>
            <SolidColorBrush x:Key="DisabledSource" Color="$($Ui.Disabled)"/>
            <SolidColorBrush x:Key="AccentSource" Color="$($Ui.Accent)"/>
            <SolidColorBrush x:Key="SelBgSource" Color="$($Ui.SelBg)"/>
            <SolidColorBrush x:Key="SelLineSource" Color="$($Ui.SelLine)"/>
            <SolidColorBrush x:Key="SelTextSource" Color="$($Ui.SelText)"/>
            <SolidColorBrush x:Key="SelMutedSource" Color="$($Ui.SelMuted)"/>
            <SolidColorBrush x:Key="SelNumSource" Color="$($Ui.SelNum)"/>
            <SolidColorBrush x:Key="HoverBgSource" Color="$($Ui.HoverBg)"/>
            <SolidColorBrush x:Key="HoverLineSource" Color="$($Ui.HoverLine)"/>
            <SolidColorBrush x:Key="InputBgSource" Color="$($Ui.InputBg)"/>
            <SolidColorBrush x:Key="InputLineSource" Color="$($Ui.InputLine)"/>
            <SolidColorBrush x:Key="MenuBgSource" Color="$($Ui.MenuBg)"/>
            <SolidColorBrush x:Key="MenuLineSource" Color="$($Ui.MenuLine)"/>
            <SolidColorBrush x:Key="MenuHotBgSource" Color="$($Ui.MenuHotBg)"/>
            <SolidColorBrush x:Key="MenuHotTextSource" Color="$($Ui.MenuHotText)"/>
            <SolidColorBrush x:Key="MenuHotMutedSource" Color="$($Ui.MenuHotMuted)"/>
            <SolidColorBrush x:Key="MenuEdgeSource" Color="$($Ui.MenuEdge)"/>
            <SolidColorBrush x:Key="MenuIconSource" Color="$($Ui.MenuIcon)"/>
            <SolidColorBrush x:Key="MenuHotIconSource" Color="$($Ui.MenuHotIcon)"/>
            <SolidColorBrush x:Key="StripeBgSource" Color="$($Ui.StripeBg)"/>
            <SolidColorBrush x:Key="StarOnSource" Color="$($Ui.StarOn)"/>
            <SolidColorBrush x:Key="StarSelSource" Color="$($Ui.StarSel)"/>
            <SolidColorBrush x:Key="ThumbSource" Color="$($Ui.Thumb)"/>
            <SolidColorBrush x:Key="ThumbHotSource" Color="$($Ui.ThumbHot)"/>
            <SolidColorBrush x:Key="OkSource" Color="$($Ui.Ok)"/>
            <SolidColorBrush x:Key="WarnSource" Color="$($Ui.Warn)"/>
            <SolidColorBrush x:Key="RedMarkSource" Color="$($Ui.RedMark)"/>
            <SolidColorBrush x:Key="GreenMarkSource" Color="$($Ui.GreenMark)"/>
            <SolidColorBrush x:Key="YellowMarkSource" Color="$($Ui.YellowMark)"/>
            <SolidColorBrush x:Key="BlueMarkSource" Color="$($Ui.BlueMark)"/>
            <SolidColorBrush x:Key="Chrome" Color="{Binding Color, Source={StaticResource ChromeSource}}"/>
            <SolidColorBrush x:Key="Surface" Color="{Binding Color, Source={StaticResource SurfaceSource}}"/>
            <SolidColorBrush x:Key="Line" Color="{Binding Color, Source={StaticResource LineSource}}"/>
            <SolidColorBrush x:Key="Text" Color="{Binding Color, Source={StaticResource TextSource}}"/>
            <SolidColorBrush x:Key="Muted" Color="{Binding Color, Source={StaticResource MutedSource}}"/>
            <SolidColorBrush x:Key="Disabled" Color="{Binding Color, Source={StaticResource DisabledSource}}"/>
            <SolidColorBrush x:Key="Accent" Color="{Binding Color, Source={StaticResource AccentSource}}"/>
            <SolidColorBrush x:Key="SelBg" Color="{Binding Color, Source={StaticResource SelBgSource}}"/>
            <SolidColorBrush x:Key="SelLine" Color="{Binding Color, Source={StaticResource SelLineSource}}"/>
            <SolidColorBrush x:Key="SelText" Color="{Binding Color, Source={StaticResource SelTextSource}}"/>
            <SolidColorBrush x:Key="SelMuted" Color="{Binding Color, Source={StaticResource SelMutedSource}}"/>
            <SolidColorBrush x:Key="SelNum" Color="{Binding Color, Source={StaticResource SelNumSource}}"/>
            <SolidColorBrush x:Key="HoverBg" Color="{Binding Color, Source={StaticResource HoverBgSource}}"/>
            <SolidColorBrush x:Key="HoverLine" Color="{Binding Color, Source={StaticResource HoverLineSource}}"/>
            <SolidColorBrush x:Key="InputBg" Color="{Binding Color, Source={StaticResource InputBgSource}}"/>
            <SolidColorBrush x:Key="InputLine" Color="{Binding Color, Source={StaticResource InputLineSource}}"/>
            <SolidColorBrush x:Key="MenuBg" Color="{Binding Color, Source={StaticResource MenuBgSource}}"/>
            <SolidColorBrush x:Key="MenuLine" Color="{Binding Color, Source={StaticResource MenuLineSource}}"/>
            <SolidColorBrush x:Key="MenuHotBg" Color="{Binding Color, Source={StaticResource MenuHotBgSource}}"/>
            <SolidColorBrush x:Key="MenuHotText" Color="{Binding Color, Source={StaticResource MenuHotTextSource}}"/>
            <SolidColorBrush x:Key="MenuHotMuted" Color="{Binding Color, Source={StaticResource MenuHotMutedSource}}"/>
            <SolidColorBrush x:Key="MenuEdge" Color="{Binding Color, Source={StaticResource MenuEdgeSource}}"/>
            <SolidColorBrush x:Key="MenuIcon" Color="{Binding Color, Source={StaticResource MenuIconSource}}"/>
            <SolidColorBrush x:Key="MenuHotIcon" Color="{Binding Color, Source={StaticResource MenuHotIconSource}}"/>
            <SolidColorBrush x:Key="StripeBg" Color="{Binding Color, Source={StaticResource StripeBgSource}}"/>
            <SolidColorBrush x:Key="StarOn" Color="{Binding Color, Source={StaticResource StarOnSource}}"/>
            <SolidColorBrush x:Key="StarSel" Color="{Binding Color, Source={StaticResource StarSelSource}}"/>
            <SolidColorBrush x:Key="Thumb" Color="{Binding Color, Source={StaticResource ThumbSource}}"/>
            <SolidColorBrush x:Key="ThumbHot" Color="{Binding Color, Source={StaticResource ThumbHotSource}}"/>
            <SolidColorBrush x:Key="Ok" Color="{Binding Color, Source={StaticResource OkSource}}"/>
            <SolidColorBrush x:Key="Warn" Color="{Binding Color, Source={StaticResource WarnSource}}"/>
            <SolidColorBrush x:Key="RedMark" Color="{Binding Color, Source={StaticResource RedMarkSource}}"/>
            <SolidColorBrush x:Key="GreenMark" Color="{Binding Color, Source={StaticResource GreenMarkSource}}"/>
            <SolidColorBrush x:Key="YellowMark" Color="{Binding Color, Source={StaticResource YellowMarkSource}}"/>
            <SolidColorBrush x:Key="BlueMark" Color="{Binding Color, Source={StaticResource BlueMarkSource}}"/>
            <!-- 0.8.2: the menu shadow works the same way; its holder carries the strength as Opacity -->
            <SolidColorBrush x:Key="ShadowSource" Color="#000000" Opacity="$($Ui.ShadowOp)"/>
            <DropShadowEffect x:Key="MenuShadow" BlurRadius="12" ShadowDepth="3" Direction="315" Color="#000000"
                              Opacity="{Binding Opacity, Source={StaticResource ShadowSource}}"/>
            <!-- Windows 11 has the first font, Windows 10 the second; same code points -->
            <FontFamily x:Key="IconFont">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>
            <!-- width of the line number column. Update-RowNumbers in the script replaces
                 this value whenever the longest number gains or loses a digit -->
            <GridLength x:Key="NumberColumn">0</GridLength>

            <!-- ===== slim scroll bar: a rounded thumb, no arrow buttons ===== -->
            <Style x:Key="SlimThumb" TargetType="Thumb">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="IsTabStop" Value="False"/>
                <Setter Property="Focusable" Value="False"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Thumb">
                            <!-- the transparent outer border keeps the whole bar width grabbable -->
                            <Border Background="Transparent" Padding="{TemplateBinding Padding}">
                                <Border x:Name="Bar" CornerRadius="3" Background="{StaticResource Thumb}"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bar" Property="Background" Value="{StaticResource ThumbHot}"/>
                                </Trigger>
                                <Trigger Property="IsDragging" Value="True">
                                    <Setter TargetName="Bar" Property="Background" Value="{StaticResource ThumbHot}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="ScrollPage" TargetType="RepeatButton">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="IsTabStop" Value="False"/>
                <Setter Property="Focusable" Value="False"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="RepeatButton">
                            <Border Background="Transparent"/>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="SlimScroll" TargetType="ScrollBar">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Width" Value="12"/>
                <Setter Property="MinWidth" Value="12"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ScrollBar">
                            <Grid Background="Transparent">
                                <Track x:Name="PART_Track" IsDirectionReversed="True">
                                    <Track.Resources>
                                        <!-- the thumb never gets shorter than half of this -->
                                        <sys:Double x:Key="{x:Static SystemParameters.VerticalScrollBarButtonHeightKey}">56</sys:Double>
                                    </Track.Resources>
                                    <Track.DecreaseRepeatButton>
                                        <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageUpCommand"/>
                                    </Track.DecreaseRepeatButton>
                                    <Track.Thumb>
                                        <Thumb Style="{StaticResource SlimThumb}" Padding="3,2,3,2"/>
                                    </Track.Thumb>
                                    <Track.IncreaseRepeatButton>
                                        <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageDownCommand"/>
                                    </Track.IncreaseRepeatButton>
                                </Track>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
                <Style.Triggers>
                    <Trigger Property="Orientation" Value="Horizontal">
                        <Setter Property="Width" Value="Auto"/>
                        <Setter Property="MinWidth" Value="0"/>
                        <Setter Property="Height" Value="12"/>
                        <Setter Property="MinHeight" Value="12"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ScrollBar">
                                    <Grid Background="Transparent">
                                        <Track x:Name="PART_Track">
                                            <Track.Resources>
                                                <sys:Double x:Key="{x:Static SystemParameters.HorizontalScrollBarButtonWidthKey}">56</sys:Double>
                                            </Track.Resources>
                                            <Track.DecreaseRepeatButton>
                                                <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageLeftCommand"/>
                                            </Track.DecreaseRepeatButton>
                                            <Track.Thumb>
                                                <Thumb Style="{StaticResource SlimThumb}" Padding="2,3,2,3"/>
                                            </Track.Thumb>
                                            <Track.IncreaseRepeatButton>
                                                <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageRightCommand"/>
                                            </Track.IncreaseRepeatButton>
                                        </Track>
                                    </Grid>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Trigger>
                </Style.Triggers>
            </Style>
            <Style TargetType="ScrollBar" BasedOn="{StaticResource SlimScroll}"/>

            <!-- ===== filter box: search icon, hint text, rounded border ===== -->
            <Style x:Key="SearchBox" TargetType="TextBox">
                <Setter Property="Height" Value="30"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="CaretBrush" Value="{StaticResource Text}"/>
                <Setter Property="SelectionBrush" Value="{StaticResource Accent}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TextBox">
                            <Border x:Name="Bd" Background="{StaticResource InputBg}" BorderBrush="{StaticResource InputLine}"
                                    BorderThickness="1" CornerRadius="6" SnapsToDevicePixels="True">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="32"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="&#xE721;" FontFamily="{StaticResource IconFont}" FontSize="13"
                                               Foreground="{StaticResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="Hint" Grid.Column="1" Margin="2,0,0,0" VerticalAlignment="Center"
                                               Text="{Binding Tag, RelativeSource={RelativeSource TemplatedParent}}"
                                               Foreground="{StaticResource Muted}" Visibility="Collapsed" IsHitTestVisible="False"/>
                                    <ScrollViewer x:Name="PART_ContentHost" Grid.Column="1" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                </Grid>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="Text" Value="">
                                    <Setter TargetName="Hint" Property="Visibility" Value="Visible"/>
                                </Trigger>
                                <Trigger Property="IsKeyboardFocused" Value="True">
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- ===== 0.8.1: the Show Only Favorites button beside the filter box =====
                 An outline star when off; a solid gold star and an accent border when on. -->
            <Style x:Key="StarToggle" TargetType="ToggleButton">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ToggleButton">
                            <Border x:Name="Bd" Width="32" Height="30" Background="Transparent" BorderBrush="{StaticResource InputLine}"
                                    BorderThickness="1" CornerRadius="6" SnapsToDevicePixels="True">
                                <TextBlock x:Name="Star" Text="&#xE734;" FontFamily="{StaticResource IconFont}" FontSize="14"
                                           Foreground="{StaticResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource HoverBg}"/>
                                </Trigger>
                                <Trigger Property="IsKeyboardFocused" Value="True">
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
                                </Trigger>
                                <Trigger Property="IsChecked" Value="True">
                                    <Setter TargetName="Star" Property="Text" Value="&#xE735;"/>
                                    <Setter TargetName="Star" Property="Foreground" Value="{StaticResource StarOn}"/>
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- ===== list row =====
                 The row draws its own background so the Windows selection colours never
                 show through. Later triggers win: the alternating stripe, a row colour,
                 then the pointer, then selection. A selected coloured row keeps its
                 marker stripe. -->
            <Style x:Key="ClipRow" TargetType="ListBoxItem">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="Margin" Value="6,0,4,1"/>
                <Setter Property="Padding" Value="0"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ListBoxItem">
                            <Border x:Name="Bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="1"
                                    CornerRadius="5" MinHeight="29" SnapsToDevicePixels="True">
                                <ContentPresenter/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <!-- 0.8.0: every other row carries a soft stripe (the list sets
                                     AlternationCount to 2); a topic heading never does -->
                                <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource StripeBg}"/>
                                </Trigger>
                                <DataTrigger Binding="{Binding IsTopic}" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="Transparent"/>
                                </DataTrigger>
                                <DataTrigger Binding="{Binding HasColor}" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{Binding TagFill}"/>
                                </DataTrigger>
                                <!-- the pointer fills a plain row and only outlines a coloured one -->
                                <MultiDataTrigger>
                                    <MultiDataTrigger.Conditions>
                                        <Condition Binding="{Binding IsMouseOver, RelativeSource={RelativeSource Self}}" Value="True"/>
                                        <Condition Binding="{Binding HasColor}" Value="False"/>
                                    </MultiDataTrigger.Conditions>
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource HoverBg}"/>
                                </MultiDataTrigger>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource HoverLine}"/>
                                </Trigger>
                                <Trigger Property="IsKeyboardFocused" Value="True">
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource HoverLine}"/>
                                </Trigger>
                                <!-- stays the same when focus moves to the filter box or a menu -->
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource SelBg}"/>
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource SelLine}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
                <Style.Triggers>
                    <!-- a topic heading gets a little air above it -->
                    <DataTrigger Binding="{Binding IsTopic}" Value="True">
                        <Setter Property="Margin" Value="6,6,4,1"/>
                    </DataTrigger>
                </Style.Triggers>
            </Style>

            <!-- Columns: marker stripe | star | line number | icon | text | note and time.
                 0.8.1: the star marks a favorite. It is always there to be clicked but
                 only drawn as an outline while the pointer is over the row, and solid
                 gold once set. It carries Tag "star"; the click handler looks for it.
                 Every row takes the number column width from the NumberColumn resource,
                 so the numbers line up, the column is exactly as wide as the longest
                 number, and it vanishes when no row is numbered. (A SharedSizeGroup
                 would do the same, but it made every add re-measure the whole list.)
                 Everything is top-aligned so the number and icon stay on the first line
                 of a row that runs to several lines. The icon cell carries Tag "toggle"
                 (the expand click looks for it) and the text cell Tag "body" (the
                 tooltip rule, RowTip, looks for it). -->
            <DataTemplate x:Key="ClipTemplate">
                <Grid Margin="0,0,8,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="7"/>
                        <ColumnDefinition Width="19"/>
                        <ColumnDefinition Width="{DynamicResource NumberColumn}"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Border Width="3" CornerRadius="1.5" Margin="2,4,0,4" HorizontalAlignment="Left" Background="{Binding TagMark}"/>
                    <TextBlock x:Name="Star" Grid.Column="1" Tag="star" Text="&#xE734;" FontFamily="{StaticResource IconFont}" FontSize="12"
                               Foreground="{StaticResource Muted}" Opacity="0" Padding="3,8,3,6" VerticalAlignment="Top" Cursor="Hand"/>
                    <TextBlock x:Name="Num" Grid.Column="2" Text="{Binding NumberText}" FontFamily="Consolas" FontSize="12"
                               Foreground="{StaticResource Muted}" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,7,7,0"/>
                    <Grid x:Name="Row" Grid.Column="3">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="22"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="Ico" Text="{Binding Glyph}" Tag="toggle" FontFamily="{StaticResource IconFont}" FontSize="12"
                                   Foreground="{StaticResource Muted}" Padding="1,8,0,6" VerticalAlignment="Top"/>
                        <TextBlock x:Name="Body" Grid.Column="1" Text="{Binding Body}" Tag="body" TextTrimming="CharacterEllipsis"
                                   VerticalAlignment="Top" Margin="0,5,0,5" ToolTip="{Binding Tip}"
                                   ToolTipService.InitialShowDelay="600" ToolTipService.ShowDuration="30000"/>
                        <TextBlock x:Name="Meta" Grid.Column="2" Text="{Binding Meta}" FontSize="11.5" Foreground="{StaticResource Muted}"
                                   VerticalAlignment="Top" Margin="12,6,0,0" ToolTip="{Binding MetaTip}"/>
                    </Grid>
                    <!-- an Insert Topic row: just the title and a rule -->
                    <DockPanel x:Name="Topic" Grid.Column="3" Visibility="Collapsed" Margin="1,0,0,0">
                        <TextBlock x:Name="Title" DockPanel.Dock="Left" Text="{Binding Caption}" FontWeight="SemiBold" FontSize="12"
                                   Foreground="{StaticResource Accent}" TextTrimming="CharacterEllipsis" VerticalAlignment="Top" Margin="0,6,0,5"/>
                        <Border Height="1" Margin="10,2,0,0" VerticalAlignment="Center" Background="{StaticResource Line}"/>
                    </DockPanel>
                </Grid>
                <DataTemplate.Triggers>
                    <!-- star: outline under the pointer, solid gold when set, none on a topic -->
                    <DataTrigger Binding="{Binding IsMouseOver, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Value="True">
                        <Setter TargetName="Star" Property="Opacity" Value="0.7"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding IsFavorite}" Value="True">
                        <Setter TargetName="Star" Property="Text" Value="&#xE735;"/>
                        <Setter TargetName="Star" Property="Foreground" Value="{StaticResource StarOn}"/>
                        <Setter TargetName="Star" Property="Opacity" Value="1"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding IsTopic}" Value="True">
                        <Setter TargetName="Row" Property="Visibility" Value="Collapsed"/>
                        <Setter TargetName="Topic" Property="Visibility" Value="Visible"/>
                        <Setter TargetName="Star" Property="Visibility" Value="Hidden"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding Meta}" Value="">
                        <Setter TargetName="Meta" Property="Visibility" Value="Collapsed"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding IsLong}" Value="True">
                        <Setter TargetName="Ico" Property="Cursor" Value="Hand"/>
                    </DataTrigger>
                    <!-- expanded: show everything, wrapped to the window width -->
                    <DataTrigger Binding="{Binding IsExpanded}" Value="True">
                        <Setter TargetName="Body" Property="TextWrapping" Value="Wrap"/>
                        <Setter TargetName="Body" Property="TextTrimming" Value="None"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding IsDim}" Value="True">
                        <Setter TargetName="Body" Property="Foreground" Value="{StaticResource Muted}"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding IsSelected, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Value="True">
                        <Setter TargetName="Num" Property="Foreground" Value="{StaticResource SelNum}"/>
                        <Setter TargetName="Ico" Property="Foreground" Value="{StaticResource SelMuted}"/>
                        <Setter TargetName="Body" Property="Foreground" Value="{StaticResource SelText}"/>
                        <Setter TargetName="Meta" Property="Foreground" Value="{StaticResource SelMuted}"/>
                        <Setter TargetName="Title" Property="Foreground" Value="{StaticResource SelNum}"/>
                        <Setter TargetName="Star" Property="Foreground" Value="{StaticResource SelMuted}"/>
                    </DataTrigger>
                    <MultiDataTrigger>
                        <MultiDataTrigger.Conditions>
                            <Condition Binding="{Binding IsSelected, RelativeSource={RelativeSource AncestorType={x:Type ListBoxItem}}}" Value="True"/>
                            <Condition Binding="{Binding IsFavorite}" Value="True"/>
                        </MultiDataTrigger.Conditions>
                        <Setter TargetName="Star" Property="Foreground" Value="{StaticResource StarSel}"/>
                    </MultiDataTrigger>
                </DataTemplate.Triggers>
            </DataTemplate>

            <!-- ===== menus: icon column, label, shortcut column, real separators ===== -->
            <Style x:Key="MenuSep" TargetType="Separator">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="Focusable" Value="False"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Separator">
                            <Border Height="1" Margin="12,5,12,5" Background="{StaticResource Line}" SnapsToDevicePixels="True"/>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- One template for every item inside a menu. The shadow only reaches down
                 and to the right, the way stock WPF menus do it, so a fly-out still
                 lines up when Windows opens it on the other side (near a screen edge,
                 or on touch PCs set for right-handed use). -->
            <Style x:Key="MI" TargetType="MenuItem">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="MenuItem">
                            <Border x:Name="Bd" Background="Transparent" CornerRadius="4" Margin="4,0,4,0" MinHeight="28" SnapsToDevicePixels="True">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="34"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto" SharedSizeGroup="MenuGesture"/>
                                        <ColumnDefinition Width="24"/>
                                    </Grid.ColumnDefinitions>
                                    <!-- 0.8.0: accent bar at the left edge of the item under the pointer -->
                                    <Border x:Name="Bar" Width="3" CornerRadius="1.5" Margin="0,7,0,7" HorizontalAlignment="Left"
                                            Background="{StaticResource MenuHotIcon}" Visibility="Collapsed"/>
                                    <ContentPresenter x:Name="Ico" ContentSource="Icon" TextElement.FontFamily="{StaticResource IconFont}"
                                                      TextElement.FontSize="14" TextElement.Foreground="{StaticResource MenuIcon}"
                                                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="Check" Text="&#xE73E;" FontFamily="{StaticResource IconFont}" FontSize="12"
                                               Foreground="{StaticResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"
                                               Visibility="Collapsed"/>
                                    <ContentPresenter Grid.Column="1" ContentSource="Header" RecognizesAccessKey="True" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="Gesture" Grid.Column="2" Text="{TemplateBinding InputGestureText}" FontSize="12"
                                               Foreground="{StaticResource Muted}" Margin="28,0,0,0" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="Arrow" Grid.Column="3" Text="&#xE76C;" FontFamily="{StaticResource IconFont}" FontSize="9"
                                               Foreground="{StaticResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center"
                                               Visibility="Collapsed"/>
                                    <Popup x:Name="PART_Popup" Placement="Right" HorizontalOffset="-2" VerticalOffset="-5"
                                           AllowsTransparency="True" Focusable="False"
                                           IsOpen="{Binding IsSubmenuOpen, RelativeSource={RelativeSource TemplatedParent}}"
                                           PopupAnimation="{DynamicResource {x:Static SystemParameters.MenuPopupAnimationKey}}">
                                        <Grid Margin="0,0,10,10">
                                            <Border CornerRadius="8" Background="{StaticResource MenuBg}" Effect="{StaticResource MenuShadow}"/>
                                            <Border CornerRadius="8" Background="{StaticResource MenuBg}" BorderBrush="{StaticResource MenuEdge}"
                                                    BorderThickness="1" Padding="0,4,0,4">
                                                <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                                                    <StackPanel IsItemsHost="True" Grid.IsSharedSizeScope="True"
                                                                KeyboardNavigation.DirectionalNavigation="Cycle"/>
                                                </ScrollViewer>
                                            </Border>
                                        </Grid>
                                    </Popup>
                                </Grid>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="Role" Value="SubmenuHeader">
                                    <Setter TargetName="Arrow" Property="Visibility" Value="Visible"/>
                                </Trigger>
                                <Trigger Property="IsChecked" Value="True">
                                    <Setter TargetName="Check" Property="Visibility" Value="Visible"/>
                                </Trigger>
                                <Trigger Property="IsHighlighted" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource MenuHotBg}"/>
                                    <Setter Property="Foreground" Value="{StaticResource MenuHotText}"/>
                                    <Setter TargetName="Bar" Property="Visibility" Value="Visible"/>
                                    <Setter TargetName="Ico" Property="TextElement.Foreground" Value="{StaticResource MenuHotIcon}"/>
                                    <Setter TargetName="Check" Property="Foreground" Value="{StaticResource MenuHotIcon}"/>
                                    <Setter TargetName="Gesture" Property="Foreground" Value="{StaticResource MenuHotMuted}"/>
                                    <Setter TargetName="Arrow" Property="Foreground" Value="{StaticResource MenuHotMuted}"/>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter Property="Foreground" Value="{StaticResource Disabled}"/>
                                    <Setter TargetName="Ico" Property="Opacity" Value="0.45"/>
                                    <Setter TargetName="Gesture" Property="Foreground" Value="{StaticResource Disabled}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- the Tools button in the header is a real top-level menu item -->
            <Style x:Key="TopMenuItem" TargetType="MenuItem">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="MenuItem">
                            <Border x:Name="Bd" Height="30" Background="Transparent" BorderBrush="{StaticResource InputLine}"
                                    BorderThickness="1" CornerRadius="6" Padding="12,0,10,0" SnapsToDevicePixels="True">
                                <Grid>
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <ContentPresenter ContentSource="Header" RecognizesAccessKey="True" VerticalAlignment="Center"/>
                                        <TextBlock Text="&#xE70D;" FontFamily="{StaticResource IconFont}" FontSize="8" Margin="8,2,0,0" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <Popup x:Name="PART_Popup" Placement="Bottom" VerticalOffset="3" AllowsTransparency="True" Focusable="False"
                                           IsOpen="{Binding IsSubmenuOpen, RelativeSource={RelativeSource TemplatedParent}}"
                                           PopupAnimation="{DynamicResource {x:Static SystemParameters.MenuPopupAnimationKey}}">
                                        <Grid Margin="0,0,10,10">
                                            <Border CornerRadius="8" Background="{StaticResource MenuBg}" Effect="{StaticResource MenuShadow}"/>
                                            <Border CornerRadius="8" Background="{StaticResource MenuBg}" BorderBrush="{StaticResource MenuEdge}"
                                                    BorderThickness="1" Padding="0,4,0,4">
                                                <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                                                    <StackPanel IsItemsHost="True" Grid.IsSharedSizeScope="True"
                                                                KeyboardNavigation.DirectionalNavigation="Cycle"/>
                                                </ScrollViewer>
                                            </Border>
                                        </Grid>
                                    </Popup>
                                </Grid>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsHighlighted" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource MenuHotBg}"/>
                                    <Setter Property="Foreground" Value="{StaticResource MenuHotText}"/>
                                </Trigger>
                                <Trigger Property="IsSubmenuOpen" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="{StaticResource MenuHotBg}"/>
                                    <Setter Property="Foreground" Value="{StaticResource MenuHotText}"/>
                                </Trigger>
                                <Trigger Property="IsKeyboardFocused" Value="True">
                                    <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="CtxMenu" TargetType="ContextMenu">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="HasDropShadow" Value="False"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="FontFamily" Value="Segoe UI"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="UseLayoutRounding" Value="True"/>
                <Setter Property="TextOptions.TextFormattingMode" Value="Display"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ContextMenu">
                            <Grid Margin="0,0,10,10">
                                <Border CornerRadius="8" Background="{StaticResource MenuBg}" Effect="{StaticResource MenuShadow}"/>
                                <Border CornerRadius="8" Background="{StaticResource MenuBg}" BorderBrush="{StaticResource MenuEdge}"
                                        BorderThickness="1" Padding="0,4,0,4">
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                                        <StackPanel IsItemsHost="True" Grid.IsSharedSizeScope="True"
                                                    KeyboardNavigation.DirectionalNavigation="Cycle"/>
                                    </ScrollViewer>
                                </Border>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- tooltips match the menus and wrap long text -->
            <Style TargetType="ToolTip">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="HasDropShadow" Value="False"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Setter Property="FontFamily" Value="Segoe UI"/>
                <Setter Property="FontSize" Value="12"/>
                <Setter Property="MaxWidth" Value="460"/>
                <Setter Property="UseLayoutRounding" Value="True"/>
                <Setter Property="TextOptions.TextFormattingMode" Value="Display"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ToolTip">
                            <Border Background="{StaticResource MenuBg}" BorderBrush="{StaticResource MenuLine}" BorderThickness="1"
                                    CornerRadius="5" Padding="9,6,9,7">
                                <TextBlock Text="{Binding Content, RelativeSource={RelativeSource TemplatedParent}}" TextWrapping="Wrap"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- header: filter box and the Tools menu button -->
            <Border Grid.Row="0" Background="{StaticResource Chrome}" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="10,8,10,8">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="InputBox" Style="{StaticResource SearchBox}" Tag="Filter clips"/>
                    <!-- 0.8.1: Show Only Favorites; Tools has the same switch as a menu item -->
                    <ToggleButton x:Name="FavOnly_Button" Grid.Column="1" Style="{StaticResource StarToggle}" Margin="8,0,0,0"
                                  VerticalAlignment="Center" ToolTip="Show only favorites"/>
                    <Menu Grid.Column="2" Margin="8,0,0,0" Background="Transparent" FontFamily="Segoe UI" FontSize="13" VerticalAlignment="Center">
                        <Menu.Template>
                            <ControlTemplate TargetType="Menu">
                                <ItemsPresenter/>
                            </ControlTemplate>
                        </Menu.Template>
                        <Menu.Resources>
                            <Style TargetType="MenuItem" BasedOn="{StaticResource MI}"/>
                            <Style x:Key="{x:Static MenuItem.SeparatorStyleKey}" TargetType="Separator" BasedOn="{StaticResource MenuSep}"/>
                            <Style TargetType="ScrollBar" BasedOn="{StaticResource SlimScroll}"/>
                        </Menu.Resources>
                        <MenuItem x:Name="FileMenu" Style="{StaticResource TopMenuItem}" Header="_Tools">
                            <MenuItem x:Name="Clear_Menu" Icon="&#xE894;" Header="_Clear"/>
                            <MenuItem x:Name="Save_Menu" Icon="&#xE74E;" Header="_Save As File"/>
                            <MenuItem x:Name="Import_Menu" Icon="&#xE8B5;" Header="_Import File"/>
                            <MenuItem x:Name="NewInstance_Menu" Icon="&#xE8A7;" Header="_New Instance"/>
                            <Separator x:Name="ZZZ"/>
                            <MenuItem x:Name="Title_Window" Icon="&#xE8AC;" Header="_Title Window"/>
                            <!-- 0.8.2: Light / Dark / Follow Windows; the one in use carries the check mark -->
                            <MenuItem x:Name="Theme_Menu" Icon="&#xE771;" Header="T_heme">
                                <MenuItem x:Name="Theme_Light" Header="_Light" IsCheckable="True"/>
                                <MenuItem x:Name="Theme_Dark" Header="_Dark" IsCheckable="True"/>
                                <MenuItem x:Name="Theme_Auto" Header="_Follow Windows" IsCheckable="True"/>
                            </MenuItem>
                            <!-- the four switches show a check mark in the icon column when on -->
                            <MenuItem x:Name="StayTop_Menu" Header="_Stay On Top" IsCheckable="True"/>
                            <MenuItem x:Name="AddTime_Menu" Header="_Add Time Stamp" IsCheckable="True"/>
                            <MenuItem x:Name="LineNumbers_Menu" Header="_Line Numbers" IsCheckable="True"/>
                            <MenuItem x:Name="Pause_Menu" Header="_Pause" IsCheckable="True"/>
                            <MenuItem x:Name="FavOnly_Menu" Header="Show Only _Favorites" IsCheckable="True"/>
                            <MenuItem x:Name="VoiceNote_Menu" Icon="&#xE720;" Header="_Voice Note"/>
                            <MenuItem x:Name="AiSettings_Menu" Icon="&#xE99A;" Header="A_I Settings"/>
                            <Separator x:Name="ZZ"/>
                            <MenuItem x:Name="Options_Menu" Icon="&#xE713;" Header="_Options"/>
                        </MenuItem>
                    </Menu>
                </Grid>
            </Border>

            <!-- The list scrolls itself now instead of sitting in an outer ScrollViewer.
                 That lets it build only the rows that are on screen (plus a page either
                 side), so a long history costs the same as a short one. ScrollUnit Pixel
                 keeps scrolling smooth when rows differ in height and lets a row taller
                 than the window be scrolled through. Sideways scrolling is off: a row is
                 as wide as the window and a line that does not fit ends in an ellipsis.
                 The scroll viewer keeps the name ListScroller; the script fetches it
                 from this template for the drag-to-reorder code. -->
            <ListBox x:Name="listbox" Grid.Row="1" ItemContainerStyle="{StaticResource ClipRow}" ItemTemplate="{StaticResource ClipTemplate}"
                     SelectionMode="Extended" AllowDrop="True" AlternationCount="2"
                     Background="Transparent" BorderThickness="0" Padding="0,5,0,5" Foreground="{StaticResource Text}"
                     VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.ScrollUnit="Pixel"
                     VirtualizingPanel.VirtualizationMode="Standard">
                <ListBox.Template>
                    <ControlTemplate TargetType="ListBox">
                        <Border Background="{TemplateBinding Background}">
                            <ScrollViewer x:Name="ListScroller" Focusable="False" Padding="{TemplateBinding Padding}" CanContentScroll="True"
                                          VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </ListBox.Template>
                <ListBox.ContextMenu>
                    <ContextMenu x:Name="ClipboardMenu" Style="{StaticResource CtxMenu}">
                        <ContextMenu.Resources>
                            <Style TargetType="MenuItem" BasedOn="{StaticResource MI}"/>
                            <Style x:Key="{x:Static MenuItem.SeparatorStyleKey}" TargetType="Separator" BasedOn="{StaticResource MenuSep}"/>
                            <Style TargetType="ScrollBar" BasedOn="{StaticResource SlimScroll}"/>
                        </ContextMenu.Resources>
                        <!-- 11/21/26 - the five most-used item actions sit at the top of the
                             right-click menu again. Inside the Options fly-out (11/7/26) they
                             were hard to find. Everything else stays where it was.
                             11/22/26 - same items in the same order; each has an icon, the
                             shortcut moved out of the label into its own column, and the
                             rows of dashes became real separators (names kept). -->
                        <MenuItem x:Name="Copy_Menu" Icon="&#xE8C8;" Header="Copy" InputGestureText="Ctrl+C"/>
                        <MenuItem x:Name="Remove_Menu" Icon="&#xE74D;" Header="Remove item" InputGestureText="Ctrl+R"/>
                        <MenuItem x:Name="Edit_Item" Icon="&#xE70F;" Header="Edit item" InputGestureText="F2"/>
                        <!-- 0.8.1: stars every selected row, or unstars them when all already are.
                             The label and icon are set each time the menu opens. -->
                        <MenuItem x:Name="Favorite_Menu" Icon="&#xE734;" Header="Favorite"/>
                        <MenuItem x:Name="Insert_Topic" Icon="&#xE8EC;" Header="Insert Topic"/>
                        <MenuItem x:Name="Color_Menu" Icon="&#xE790;" Header="Assign Color">
                            <MenuItem x:Name="Color_Red" Header="Red">
                                <MenuItem.Icon><Ellipse Width="11" Height="11" Fill="{StaticResource RedMark}"/></MenuItem.Icon>
                            </MenuItem>
                            <MenuItem x:Name="Color_Green" Header="Green">
                                <MenuItem.Icon><Ellipse Width="11" Height="11" Fill="{StaticResource GreenMark}"/></MenuItem.Icon>
                            </MenuItem>
                            <MenuItem x:Name="Color_Yellow" Header="Yellow">
                                <MenuItem.Icon><Ellipse Width="11" Height="11" Fill="{StaticResource YellowMark}"/></MenuItem.Icon>
                            </MenuItem>
                            <MenuItem x:Name="Color_Blue" Header="Blue">
                                <MenuItem.Icon><Ellipse Width="11" Height="11" Fill="{StaticResource BlueMark}"/></MenuItem.Icon>
                            </MenuItem>
                            <Separator x:Name="Color_Sep"/>
                            <MenuItem x:Name="Color_Clear" Header="Clear Color">
                                <MenuItem.Icon><Ellipse Width="11" Height="11" Stroke="{StaticResource Muted}" StrokeThickness="1"/></MenuItem.Icon>
                            </MenuItem>
                        </MenuItem>
                        <Separator x:Name="Priority_Sep"/>
                        <!-- 11/7/26 - less-used item actions live in an Options fly-out to
                             keep the right-click menu short -->
                        <MenuItem x:Name="Options_Flyout" Icon="&#xE712;" Header="Options">
                            <MenuItem x:Name="oooAI" Icon="&#xE99A;" Header="Ask AI - TriplOOO"/>
                            <MenuItem x:Name="openEdge" Icon="&#xE774;" Header="Open with Edge"/>
                            <MenuItem x:Name="SelectAll_Menu" Icon="&#xE8B3;" Header="Select All" InputGestureText="Ctrl+A"/>
                            <MenuItem x:Name="SelectFav_Menu" Icon="&#xE735;" Header="Select Favorites"/>
                            <MenuItem x:Name="Expand_All" Icon="&#xE740;" Header="Expand All"/>
                            <MenuItem x:Name="Collapse_All" Icon="&#xE73F;" Header="Collapse All"/>
                            <MenuItem x:Name="Consolidate_Menu" Icon="&#xE7B8;" Header="Consolidate Selected"/>
                            <MenuItem x:Name="Ungroup_Menu" Icon="&#xE8A9;" Header="Ungroup"/>
                        </MenuItem>
                        <Separator x:Name="zzz"/>
                        <MenuItem x:Name="AppendFile_Menu" Icon="&#xE792;" Header="Append Selected items to file" InputGestureText="Ctrl+Q"/>
                        <MenuItem x:Name="CreateFile_Menu" Icon="&#xE74E;" Header="Write file from Selected items" InputGestureText="Ctrl+W"/>
                        <MenuItem x:Name="Review_Log" Icon="&#xE8A5;" Header="Review Log" InputGestureText="Ctrl+L / Ctrl+O"/>
                        <MenuItem x:Name="Open_URLs" Icon="&#xE71B;" Header="Open Selected URLs"/>
                        <MenuItem x:Name="Open_Google_Search" Icon="&#xE721;" Header="Open Selected in Google"/>
                        <MenuItem x:Name="Open_with_NOTEPAD" Icon="&#xE7C3;" Header="Open Selected in NOTEPAD"/>
                        <MenuItem x:Name="Open_New_Instance" Icon="&#xE8A7;" Header="Open Selected in New Instance"/>
                        <MenuItem x:Name="Mask_Selected_Item" Icon="&#xE8E6;" Header="Remove ALL Selected Items from Listbox"/>
                        <Separator x:Name="zzzz"/>
                        <MenuItem x:Name="ChatGPT" Icon="&#xE99A;" Header="Ask AI"/>
                    </ContextMenu>
                </ListBox.ContextMenu>
            </ListBox>

            <!-- status strip: capture state on the left, counts on the right -->
            <Border Grid.Row="2" Background="{StaticResource Chrome}" BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0"
                    Height="27" Padding="12,0,12,0">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Ellipse x:Name="StatusDot" Width="7" Height="7" Fill="{StaticResource Ok}" VerticalAlignment="Center" Margin="0,1,0,0"/>
                        <TextBlock x:Name="StatusText" Text="Capturing" Margin="7,0,0,0" FontSize="11.5" Foreground="{StaticResource Muted}"/>
                    </StackPanel>
                    <TextBlock x:Name="CountText" Grid.Column="1" Text="0 items" FontSize="11.5" Foreground="{StaticResource Muted}"
                               HorizontalAlignment="Right" VerticalAlignment="Center" Margin="12,0,0,0" TextTrimming="CharacterEllipsis"/>
                </Grid>
            </Border>
        </Grid>

    </Window>
"@
 
 # 3/29

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $Window=[Windows.Markup.XamlReader]::Load( $reader )

    # 11/22/26 - the clipboard icon. The prompts and settings windows opened from
    # here read the same $Script:AppIcon.
    $Script:AppIcon = New-AppIcon
    if ($Script:AppIcon -ne $null) { $Window.Icon = $Script:AppIcon }

    #Connect to Controls
    $listbox = $Window.FindName('listbox')
    # 11/22/26 - the scroll viewer is part of the list's own template now, so it is
    # fetched from the template instead of being looked up by name in the window
    [void]$listbox.ApplyTemplate()
    $ListScroller = $listbox.Template.FindName('ListScroller', $listbox)
    $InputBox = $Window.FindName('InputBox')
    $Copy_Menu = $Window.FindName('Copy_Menu')
# 8/24/26 - $Edit_Menu removed; no such item exists in the XAML
    $Remove_Menu = $Window.FindName('Remove_Menu')
    $Clear_Menu = $Window.FindName('Clear_Menu')

    $Create_File_Selected = $Window.FindName('CreateFile_Menu')
    $Review_Log_Selected = $Window.FindName('Review_Log')
    $Insert_Topic = $Window.FindName('Insert_Topic')
    $Edit_Item = $Window.FindName('Edit_Item')
    $Expand_All = $Window.FindName('Expand_All')
    $Collapse_All = $Window.FindName('Collapse_All')
    $Consolidate_Menu = $Window.FindName('Consolidate_Menu')
    $Ungroup_Menu = $Window.FindName('Ungroup_Menu')
    $Color_Red = $Window.FindName('Color_Red')
    $Color_Green = $Window.FindName('Color_Green')
    $Color_Yellow = $Window.FindName('Color_Yellow')
    $Color_Blue = $Window.FindName('Color_Blue')
    $Color_Clear = $Window.FindName('Color_Clear')
    
    $oooAI = $Window.FindName('oooAI') 
    $openEdge = $Window.FindName('openEdge') 
    $Append_File_Selected = $Window.FindName('AppendFile_Menu') 
    $SelectAll_Menu = $Window.FindName('SelectAll_Menu')
    $Open_URLs = $Window.FindName('Open_URLs') 
    $GoogleSearch = $Window.FindName('Open_Google_Search') 
    $OpenNotepad = $Window.FindName('Open_with_NOTEPAD')
    $OpenNewInstance = $Window.FindName('Open_New_Instance')
    $ChatGPT = $Window.FindName('ChatGPT') 
    $MaskSelectedItem = $Window.FindName('Mask_Selected_Item') 
    $Save_Menu = $Window.FindName('Save_Menu')
    $Import_Menu = $Window.FindName('Import_Menu')
    $NewInstance_Menu = $Window.FindName('NewInstance_Menu')
    $Title_Window = $Window.FindName('Title_Window') 
    $StayTop_Menu = $Window.FindName('StayTop_Menu')    
    $AddTime_Menu = $Window.FindName('AddTime_Menu')
    $LineNumbers_Menu = $Window.FindName('LineNumbers_Menu')
    $Pause_Menu = $Window.FindName('Pause_Menu')
    $VoiceNote_Menu = $Window.FindName('VoiceNote_Menu')
    $AiSettings_Menu = $Window.FindName('AiSettings_Menu')
    $Options_Menu = $Window.FindName('Options_Menu')
    # 11/22/26 (0.8.1) - favorites
    $FavOnly_Button = $Window.FindName('FavOnly_Button')
    $FavOnly_Menu = $Window.FindName('FavOnly_Menu')
    $Favorite_Menu = $Window.FindName('Favorite_Menu')
    $SelectFav_Menu = $Window.FindName('SelectFav_Menu')
    # 11/23/26 (0.8.2) - Tools > Theme
    $Theme_Light = $Window.FindName('Theme_Light')
    $Theme_Dark = $Window.FindName('Theme_Dark')
    $Theme_Auto = $Window.FindName('Theme_Auto')
    $ClipboardMenu = $Window.FindName('ClipboardMenu')
    # 11/22/26 - status strip at the bottom of the window
    $StatusDot = $Window.FindName('StatusDot')
    $StatusText = $Window.FindName('StatusText')
    $CountText = $Window.FindName('CountText')

    # 11/22/26 - tell the item class how each Assign Color choice is drawn in this
    # theme: the stored palette value -> row tint and marker stripe
    # 11/23/26 (0.8.2) - a function now, because Set-UiTheme runs it again
    Function Update-TagColors {
        foreach ($colorName in @($Script:ColorPalette.Keys)) {
            [Cbm082.ClipItem]::TagColors[$Script:ColorPalette[$colorName]] =
                [string[]]@($Ui["${colorName}Fill"], $Ui["${colorName}Mark"])
        }
    }
    Update-TagColors

    # 11/23/26 (0.8.2) - Title bar. A dark window gets a dark title bar. Attribute 20
    # is the name Windows 10 2004 and later (and Windows 11) use; 19 is what 1809 -
    # 1909 called it. Where neither is understood the title bar simply stays light.
    # -Redraw asks Windows to repaint the frame of a window that is already showing.
    Function Set-TitleBarTheme {
        param([switch]$Redraw)
        try {
            $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $Window).Handle
            if ($hwnd -eq [IntPtr]::Zero) { return }
            $on = 0
            if ($Ui.IsDark) { $on = 1 }
            if ([Cbm082.Native]::DwmSetWindowAttribute($hwnd, 20, [ref]$on, 4) -ne 0) {
                [void][Cbm082.Native]::DwmSetWindowAttribute($hwnd, 19, [ref]$on, 4)
            }
            # 0x37 = leave size, position, order and focus alone; the frame has changed
            if ($Redraw) { [void][Cbm082.Native]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, 0, 0, 0x37) }
        } catch { }
    }

    # the theme in use carries the check mark under Tools > Theme
    Function Update-ThemeMenu {
        $Theme_Light.IsChecked = ($Script:ThemeChoice -eq 'Light')
        $Theme_Dark.IsChecked = ($Script:ThemeChoice -eq 'Dark')
        $Theme_Auto.IsChecked = ($Script:ThemeChoice -eq 'Auto')
    }
    Update-ThemeMenu

    # 11/23/26 (0.8.2) - Theme. Switches the open window to Light, Dark or Auto (follow
    # Windows). Each brush in the XAML is bound to a colour holder named after it
    # ("Chrome" reads "ChromeSource"), so giving the holders new colours recolours
    # everything at once: list, menus, tooltips, scroll bar. Rows with an Assign Color
    # tint are told to look their colours up again. Nothing is saved: the choice lasts
    # while the window is open and is handed to any instance started from it.
    Function Set-UiTheme {
        param([string]$Choice)

        if (@('Light', 'Dark', 'Auto', 'Contrast') -notcontains $Choice) { $Choice = 'Light' }
        $Script:ThemeChoice = $Choice
        $Script:Ui = Get-UiColors -Mode $Choice

        foreach ($key in @($Window.Resources.Keys)) {
            if ($key -isnot [string] -or -not $key.EndsWith('Source')) { continue }
            $name = $key.Substring(0, $key.Length - 6)
            if ($Ui.ContainsKey($name)) {
                $Window.Resources[$key].Color = [System.Windows.Media.ColorConverter]::ConvertFromString([string]$Ui[$name])
            }
        }
        $Window.Resources['ShadowSource'].Opacity = [double]$Ui.ShadowOp

        Update-TagColors
        if ($null -ne $Script:ObservableCollection) {
            foreach ($item in $Script:ObservableCollection) {
                # same value again: the row re-reads its tint and marker stripe
                if ($item.HasColor) { $item.Color = $item.Color }
            }
        }

        Set-TitleBarTheme -Redraw
        Update-ThemeMenu
    }

    # 11/22/26 - Line numbers. Gives the rows from index $From down their position as
    # their number, and keeps the number column just wide enough for the longest number
    # on show (no column at all when no row is numbered). The width is one value that
    # every row reads from the NumberColumn resource. It only changes when the longest
    # number gains or loses a digit, so this costs next to nothing on an ordinary add.
    $Script:NumberDigits = 0
    Function Update-RowNumbers {
        param([int]$From = 0)

        $highest = [Cbm082.ClipItem]::Renumber($Script:ObservableCollection, $From)

        $digits = 0
        if ($highest -gt 0) { $digits = ([string]$highest).Length }
        if ($digits -eq $Script:NumberDigits) { return }
        $Script:NumberDigits = $digits

        $width = 0
        if ($digits -gt 0) {
            # 3 before the number and 7 after it; 7 per digit is only the fallback
            $width = (7 * $digits) + 10
            try {
                # measured rather than assumed, so it is right at any display scaling.
                # Font and size must match the "Num" TextBlock in the row template.
                $probe = New-Object System.Windows.Controls.TextBlock
                $probe.Text = "8" * $digits
                $probe.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
                $probe.FontSize = 12
                [System.Windows.Media.TextOptions]::SetTextFormattingMode($probe, [System.Windows.Media.TextFormattingMode]::Display)
                $probe.Measure((New-Object System.Windows.Size([double]::PositiveInfinity, [double]::PositiveInfinity)))
                $width = [Math]::Ceiling($probe.DesiredSize.Width) + 10
            } catch { }
        }
        # a cast, not New-Object: New-Object hands back a wrapped value, and a row
        # cannot read its column width out of a wrapped GridLength
        $Window.Resources['NumberColumn'] = [System.Windows.GridLength]([double]$width)
    }

    # 11/22/26 - Status strip: capture state on the left, counts on the right.
    # Called when the list, the selection, the filter or Pause changes. It only reports,
    # so a failure in here is swallowed rather than allowed to interrupt a capture.
    Function Update-StatusStrip {
        if ($null -eq $Script:ObservableCollection) { return }
        try {
            if ($Pause_Menu.IsChecked) {
                $StatusText.Text = "Paused"
                $StatusDot.Fill = $Window.FindResource('Warn')
            } else {
                $StatusText.Text = "Capturing"
                $StatusDot.Fill = $Window.FindResource('Ok')
            }

            $total = $Script:ObservableCollection.Count
            $shown = $listbox.Items.Count           # rows the filter lets through
            $picked = $listbox.SelectedItems.Count
            $noun = $(if ($total -eq 1) { "item" } else { "items" })

            $counts = "$total $noun"
            if ($shown -ne $total) { $counts = "$shown of $total $noun" }
            # 0.8.1 - favorites are counted over the whole list, shown or not
            $stars = [Cbm082.ClipItem]::CountFavorites($Script:ObservableCollection)
            if ($stars -gt 0) { $counts += "  " + [char]0xB7 + "  $stars " + $(if ($stars -eq 1) { "favorite" } else { "favorites" }) }
            if ($picked -gt 0) { $counts += "  " + [char]0xB7 + "  $picked selected" }
            $CountText.Text = $counts
        } catch { }
    }

$global:ChatGPTSelectedItem = $null
# 8/24/26 - Masked terms are now a LIST, not one concatenated string.
# Previously "$global:HideCopiedItems += $term" glued every mask together, so masking
# a second term produced "term1term2" and silently disabled the first mask.
$global:MaskedTerms = New-Object System.Collections.ArrayList
# 8/25/26 - AI router state. Keys are per provider; routing is per capability.
$global:AiKeys  = @{}
$global:AiRoute = @{ Chat = "xai"; SpeechToText = "xai"; TextToSpeech = "windows" }
$global:tempVar = $null

$global:firstSelectedItem = $null # $firstSelectedItem

# Set a varialbe path to write to review log
$global:ReviewLog = $null #Join-Path $executingScriptDirectory $fileName

#445

$Global:Cancel = $null

$global:bron = $null # 5/12                            

Add-Type -AssemblyName System.Windows.Forms
    
$Clear_Menu.Add_Click({
    Clear-Viewer
})

$MaskSelectedItem.Add_Click(
{ 
#6/4/23

$remove = $listbox.SelectedItems[0].Text
    #"Remove all items in the list box that have the following text: `n $remove `n`n 
    #"Note: This function is meant to remove any collected passwords and does not stop future items",Confirm Action", "YesNo", "Question")

    $applyMask2Items = [System.Windows.MessageBox]::Show(
    "Remove all items in the list box that have the following text: `n`n $remove `n`nNote: This function is meant to remove any collected passwords and does not stop future items from being added.",
    "Confirm Action",
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Question

    )

    #"across multiple lines for readability."

    if ($applyMask2Items -eq "Yes")
    {

        # Assuming $ListBox is your ListBox object
        for ($i = 0; $i -lt $listbox.Items.Count; $i++) {

                if ($listbox.Items[$i] -eq $remove) { # 

                #6/4

                @($listbox.Items[$i]) | ForEach-Object {
                    
                    [void]$Script:ObservableCollection.Remove($_)
                    $i--                    

                }

                #", 'Info', 'OK', [System.Windows.Forms.MessageBoxIcon]::Information) # $index

                #text("********")

            }

        }

        # 8/24/26 - show the exact term that will be masked, so it is obvious that the
        # line number / timestamp is not part of it
        $maskTerm = Get-ItemText -Item $remove

        $applyMaskk = [System.Windows.MessageBox]::Show("This will mask (or hide = *********) all future items containing:" +
        "`n`n     $maskTerm `n`n" +
        "This option can be reset in the Options form by copying the - Reset Masked Selection option. `n`n " +
        "Note: Multiple masked items can be saved." +
        "`nNote: Is NOT case sensitive.", "Reset in Options", "YesNo", "Question")

            if ($applyMaskk -eq "Yes")
            {
                # 8/24/26 - register the RAW text, not the formatted display row,
                # otherwise the line number / timestamp is baked into the mask term
                # and it can never match incoming clipboard text.
                Add-MaskedTerm -Term $maskTerm
            }

}

    # Not Working..... Attempting to cycle through all values and mask items found as a LIKE match
    
})

# 8/25/26 - Ask AI. Was a 200-line ChatGPT-only block with its own key prompt,
    # settings dialog and a hardcoded call to the retired /v1/completions endpoint.
    # It now just asks the router, so it works with whichever provider is routed to Chat.
    $ChatGPT.Add_Click({

        $selected = @($listbox.SelectedItems)
        if ($selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show($Window, "Select one or more items first.", "Ask AI",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        $sourceText = ($selected | ForEach-Object { $_.Text }) -join "`n"

        $question = Show-InputPrompt -Title "Ask AI" -DefaultText "Explain this in one paragraph." -Message "What should $($Script:AiProviders[$global:AiRoute['Chat']].Display) do with the selected item(s)?"
        if ([string]::IsNullOrWhiteSpace($question)) { return }

        $Window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $answer = Invoke-AiChat -Prompt ($question + "`n`n" + $sourceText)
        } finally {
            $Window.Cursor = $null
        }

        if ([string]::IsNullOrWhiteSpace($answer)) { return }

        # Show the answer, editable, so it can be trimmed before it lands in the list
        $kept = Show-InputPrompt -Title "Ask AI - Response" -Message "Response. Edit if you like, then press OK to add it to the list." -DefaultText $answer.Trim()
        if ([string]::IsNullOrWhiteSpace($kept)) { return }

        $suffix = ""
        if ($AddTime_Menu.IsChecked) { $suffix = " : " + (Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss") }

        $noteText = "[AI] " + $kept.Trim()
        [void]$Script:ObservableCollection.Add((New-ClipItem -Text $noteText -Numbered ([bool]$LineNumbers_Menu.IsChecked) -Suffix $suffix))
        $listbox.ScrollIntoView($listbox.Items[$listbox.Items.Count-1])
    })

    $OpenNotepad.Add_Click({ 
        
        #
        Add-Type -AssemblyName System.Windows.Forms
        
        @($listbox.SelectedItems) | ForEach-Object #Ozzy
        {     
            if ($_.ToString() -like "System.Windows.Controls.ListBox") 
            {
                
            }else
            {

                    $notepad_data += $_.Text
               
            }

        }
            # 0.8.1 - was "Set-ClipBoard = $notepad_data"; the extra words were ignored.
            # A favorite is pasted into Notepad with "* " in front.
            Set-ClipBoard -MarkFavorites
            $notepad = New-Object -ComObject "WScript.Shell"
            $notepad.run("notepad.exe")
            Start-Sleep -s 2
            [Windows.Forms.SendKeys]::SendWait("^v")
    })

    $Open_URLs.Add_Click({ #

        @($listbox.SelectedItems) | ForEach-Object {     
            
            $urlExecute = $_.Text
            Start-Process $urlExecute
            Start-Sleep -Seconds 3
        }

    })

    $SelectAll_Menu.Add_Click({ 
        $listbox.SelectAll()
    })

    $GoogleSearch.Add_Click({ 
        
        #https://www.google.com/search?q=flags

        @($listbox.SelectedItems) | ForEach-Object {     

            $gSearchVar = $_.Text
            $gSearchVar = $gSearchVar -replace " ","%20"
            
            $google = New-Object -ComObject "WScript.Shell"
            $google.Run("https://www.google.com/search?q=$gSearchVar")

        }
        
    })

    $oooAI.Add_Click({
        
        #1/31
        Set-ClipBoardooo

    })

    #9/24
    $openEdge.Add_Click({
        
        @($listbox.SelectedItems) | ForEach-Object {
        
            $randomNumber = Get-Random
            $save_data += (Get-ExportText -Item $_) + "`n"    # 0.8.1: "* " in front of a favorite
            $save_data | Out-File "$env:TEMP\$randomNumber.txt"
            
        }

        Start-Process "msedge.exe" -ArgumentList "file://$env:TEMP\$randomNumber.txt"

    })

    $Append_File_Selected.Add_Click({
        
        #1/31
    
        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        
        $NameFilter = "Text Files (*.txt)|*.txt"
        
        $Dialog.Filter = $NameFilter
        
        if($Dialog.ShowDialog() -eq 'Ok'){
        
             @($listbox.SelectedItems) | ForEach-Object {     
                
                $save_data = Get-ExportText -Item $_    # 0.8.1: "* " in front of a favorite
                $save_data | Out-File $Dialog.FileName -Append
            }
        }
      
    })

    $Review_Log_Selected.Add_Click({ 
        #6/3/23

        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show(" The Review Log provides a single key stroke option to  capture items you may want to review later. `n`n Use ctrl + L to write seleted items to the Review Log. `n Use ctrl + O to open the Review Log. `n`n Note: Items appended to the Review Log will appear in the  order you select them.", 'Info', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

    })

    $Create_File_Selected.Add_Click({ 
    
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        $NameFilter = "Text Files (*.txt)|*.txt"
    
        $Dialog.Filter = $NameFilter
        
        if($Dialog.ShowDialog() -eq 'Ok'){
        
            @($listbox.SelectedItems) | ForEach-Object { # 
        
                $save_data += (Get-ExportText -Item $_) + "`n"    # 0.8.1: "* " in front of a favorite
                $save_data | Out-File $Dialog.FileName

            }

        }
    })


    $Remove_Menu.Add_Click({
        @($listbox.SelectedItems) | ForEach-Object {
            [void]$Script:ObservableCollection.Remove($_)

        }
    })

    # 8/25/26 - Edit item.
    # The original Edit_Menu had no XAML element behind it, so it was dead from the
    # start and was removed in 0.6.2. It is straightforward now that items carry Text
    # and Display separately: edit the raw text, then splice it back into the formatted
    # row so the line number and any timestamp survive untouched.
    # 11/22/26 - no splice any more. The line number and time stamp are separate
    # properties, so editing just replaces Text. (The splice searched the formatted row
    # for the old text and could hit the line number instead: editing "1" on row 12
    # changed the "1" of "12".)
    Function Edit-SelectedItem {

        $selected = @($listbox.SelectedItems)
        if ($selected.Count -ne 1) {
            [System.Windows.MessageBox]::Show($Window, "Select exactly one item to edit.", "Edit item",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        $item = $selected[0]
        $oldText = $item.Text

        # 11/8/26 - an edited group can no longer be split back into its items
        if ($item.IsGroup) {
            $answer = [System.Windows.MessageBox]::Show($Window,
                "This is a consolidated group.`n`nEditing it turns it into a single plain item, and Ungroup will no longer be able to restore the original items.`n`nEdit anyway?",
                "Edit item", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }

        # 11/22/26 - multi-line box: an item of several lines can be read and edited whole
        $newText = Show-InputPrompt -Title "Edit item" -Message "Edit the item text. The line number and timestamp are kept." -DefaultText $oldText -MultiLine
        if ([string]::IsNullOrWhiteSpace($newText)) { return }

        $newText = $newText.Trim()
        if ($newText -eq $oldText) { return }

        $item.IsGroup = $false
        $item.GroupTitle = $null
        $item.Parts = $null

        $item.Text = $newText

        # 11/7/26 - rebuild the collapsed form so it never shows the old text
        Update-ClipPreview -Item $item
    }

    $Edit_Item.Add_Click({ Edit-SelectedItem })

    # 11/7/26 - Expand / collapse long items
    # 11/22/26 - any row can be expanded now, not only long ones. Rows no longer
    # scroll sideways, so expanding (the text wraps) is also how a row that is cut
    # short by the window edge gets read in full. It changes nothing on a row that
    # already fits.
    Function Set-ItemsExpanded {
        param($Items, [bool]$Expanded)
        foreach ($i in @($Items)) { $i.IsExpanded = $Expanded }
    }

    $Expand_All.Add_Click({   Set-ItemsExpanded -Items $Script:ObservableCollection -Expanded $true  })
    $Collapse_All.Add_Click({ Set-ItemsExpanded -Items $Script:ObservableCollection -Expanded $false })

    # Clicking the arrow toggles that row. Handled so the click does not also change
    # the selection; clicking anywhere else on the row selects as before.
    # 11/22/26 (0.8.1) - the star in front of the number works the same way: a click on
    # it marks or unmarks that row as a favorite and leaves the selection alone.
    $listbox.Add_PreviewMouseLeftButtonDown({
        $src = $_.OriginalSource
        if ($src -is [System.Windows.Controls.TextBlock] -and $src.Tag -eq "toggle" -and $src.DataContext.IsLong) {
            $src.DataContext.IsExpanded = -not $src.DataContext.IsExpanded
            $_.Handled = $true
        }
        elseif ($src -is [System.Windows.Controls.TextBlock] -and $src.Tag -eq "star" -and $null -ne $src.DataContext) {
            Set-ItemsFavorite -Items @($src.DataContext) -Favorite (-not $src.DataContext.IsFavorite)
            $_.Handled = $true
        }
    })

    # Right expands, Left collapses the selected items. Preview, so the ListBox and
    # its scroll viewer do not consume the arrow keys first. Keys pass through
    # untouched when nothing is selected.
    $listbox.Add_PreviewKeyDown({
        $e = $_
        if ($e.Key -ne "Right" -and $e.Key -ne "Left") { return }
        $picked = @($listbox.SelectedItems)
        if ($picked.Count -eq 0) { return }
        Set-ItemsExpanded -Items $picked -Expanded ($e.Key -eq "Right")
        $e.Handled = $true
    })

    # 11/7/26 - Drag to reorder.
    # Only when exactly ONE item is selected and it is the one under the mouse. A
    # multi-selection, or a Ctrl / Shift click, never starts a drag, so extended
    # selection behaves exactly as before. Dropping on a row puts the dragged item in
    # that row's place. Line numbers are then rewritten top to bottom.
    $Script:DragItem = $null    # row pressed on, waiting to see if the mouse moves
    $Script:Dragging = $null    # item currently being dragged

    Function Get-ListItemFromSource {
        param($Source)
        $d = $Source
        while ($null -ne $d -and -not ($d -is [System.Windows.Controls.ListBoxItem])) {
            if (-not ($d -is [System.Windows.Media.Visual])) { return $null }
            $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
        }
        return $d
    }

    # Rewrites the leading "<n> " of each numbered row to match its position. Rows
    # captured with numbering off are left alone; a row only counts as numbered when
    # the number is followed by exactly its own text, so content that merely starts
    # with digits is never mistaken for a line number.
    # 11/22/26 - there is no "<n> " in the text to rewrite any more. Number is a
    # property set from the row's position, and the list renumbers itself on every
    # add, remove and move (see the CollectionChanged handler in SourceInitialized).
    # This stays as the "renumber everything now" call its existing callers expect.
    Function Update-LineNumbers {
        Update-RowNumbers -From 0
    }

    # handledEventsToo: the expand-arrow handler marks its clicks handled, and this one
    # still has to see them so a stale DragItem is cleared
    $listbox.AddHandler([System.Windows.UIElement]::PreviewMouseLeftButtonDownEvent,
        [System.Windows.Input.MouseButtonEventHandler]{
            param($s, $e)
            $Script:DragItem = $null
            if ($e.Handled) { return }
            $row = Get-ListItemFromSource -Source $e.OriginalSource
            if ($null -eq $row) { return }
            $Script:DragItem = $row.DataContext
            $Script:DragStart = $e.GetPosition($listbox)
        }, $true)

    $listbox.Add_PreviewMouseMove({
        param($s, $e)
        if ($null -eq $Script:DragItem -or $e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }

        $pos = $e.GetPosition($listbox)
        if ([Math]::Abs($pos.X - $Script:DragStart.X) -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
            [Math]::Abs($pos.Y - $Script:DragStart.Y) -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }

        $item = $Script:DragItem
        $Script:DragItem = $null    # one drag per press

        # IF and only IF a single item is selected, and it is the one being dragged
        if ($listbox.SelectedItems.Count -ne 1 -or -not [object]::ReferenceEquals($listbox.SelectedItem, $item)) { return }

        $Script:Dragging = $item
        try {
            [void][System.Windows.DragDrop]::DoDragDrop($listbox, $item, [System.Windows.DragDropEffects]::Move)
        } finally {
            $Script:Dragging = $null
        }
    })

    $listbox.Add_DragOver({
        param($s, $e)
        # only our own rows can be dropped here, not text dragged in from other apps
        if ($null -eq $Script:Dragging) {
            $e.Effects = [System.Windows.DragDropEffects]::None
        } else {
            $e.Effects = [System.Windows.DragDropEffects]::Move
            # scroll when the pointer nears the top or bottom of the visible list
            $y = $e.GetPosition($ListScroller).Y
            if ($y -lt 24) { $ListScroller.LineUp() }
            elseif ($y -gt $ListScroller.ActualHeight - 24) { $ListScroller.LineDown() }
        }
        $e.Handled = $true
    })

    $listbox.Add_Drop({
        param($s, $e)
        $item = $Script:Dragging
        if ($null -eq $item) { return }
        $e.Handled = $true

        $row = Get-ListItemFromSource -Source $e.OriginalSource
        if ($null -eq $row) { return }

        $from = $Script:ObservableCollection.IndexOf($item)
        $to = $Script:ObservableCollection.IndexOf($row.DataContext)
        if ($from -lt 0 -or $to -lt 0 -or $from -eq $to) { return }

        $Script:ObservableCollection.Move($from, $to)
        Update-LineNumbers

        $listbox.SelectedItem = $item
        $listbox.ScrollIntoView($item)
    })

    # 11/8/26 - Consolidate / Ungroup.
    # A group is an ordinary ClipItem whose Text is its items' text joined one per line,
    # so copy, save, filter, Review Log and Ask AI all just work on it. The original
    # items ride along in Parts, which is what lets Ungroup put them back exactly.

    # What Ungroup needs to rebuild one item: its text, colour, whatever followed the
    # text in its row (the timestamp), and whether it carried a line number.
    # 11/22/26 - all four are plain properties of the item now; nothing is recovered
    # from the formatted row.
    Function ConvertTo-PartRecord {
        param($Item)
        # 0.8.1 - and whether it was a favorite
        return @{ Text = $Item.Text; Color = [string]$Item.Color; Suffix = [string]$Item.Suffix; Numbered = [bool]$Item.Numbered
                  Favorite = [bool]$Item.IsFavorite }
    }

    Function Join-SelectedItems {
        $selected = @($listbox.SelectedItems | Sort-Object { $Script:ObservableCollection.IndexOf($_) })
        if ($selected.Count -lt 2) {
            [System.Windows.MessageBox]::Show($Window, "Select two or more items to consolidate.", "Consolidate",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        $title = Show-InputPrompt -Title "Consolidate" -Message "Optional title for this group of $($selected.Count) items.`nLeave blank for none."
        if ($null -eq $title) { return }    # Cancel

        # consolidating a group with other items folds its parts in, so groups never nest
        $parts = New-Object System.Collections.ArrayList
        foreach ($item in $selected) {
            if ($item.IsGroup) {
                foreach ($p in @($item.Parts)) { [void]$parts.Add($p) }
            } else {
                [void]$parts.Add((ConvertTo-PartRecord -Item $item))
            }
        }

        $at = $Script:ObservableCollection.IndexOf($selected[0])
        foreach ($item in $selected) { [void]$Script:ObservableCollection.Remove($item) }

        $text = (@($parts) | ForEach-Object { $_.Text }) -join "`n"

        $group = New-ClipItem -Text $text -Numbered ([bool]$LineNumbers_Menu.IsChecked)
        $group.IsGroup = $true
        $group.GroupTitle = $title.Trim()
        $group.Parts = $parts.ToArray()
        Update-ClipPreview -Item $group

        $Script:ObservableCollection.Insert($at, $group)
        Update-LineNumbers

        $listbox.SelectedItem = $group
        $listbox.ScrollIntoView($group)
    }

    Function Split-SelectedGroups {
        $groups = @($listbox.SelectedItems | Where-Object { $_.IsGroup } |
                    Sort-Object { $Script:ObservableCollection.IndexOf($_) })
        if ($groups.Count -eq 0) {
            [System.Windows.MessageBox]::Show($Window, "Select a consolidated group to ungroup.", "Ungroup",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        $restored = New-Object System.Collections.ArrayList
        foreach ($g in $groups) {
            $at = $Script:ObservableCollection.IndexOf($g)
            [void]$Script:ObservableCollection.Remove($g)
            foreach ($p in @($g.Parts)) {
                $restoredItem = New-ClipItem -Text ([string]$p.Text) -Numbered ([bool]$p.Numbered) -Suffix ([string]$p.Suffix) -Color ([string]$p.Color)
                $restoredItem.IsFavorite = [bool]$p.Favorite    # 0.8.1
                $Script:ObservableCollection.Insert($at, $restoredItem)
                $at++
                [void]$restored.Add($restoredItem)
            }
        }
        Update-LineNumbers

        # leave the restored items selected so it is obvious what came back
        $listbox.SelectedItems.Clear()
        foreach ($r in $restored) { [void]$listbox.SelectedItems.Add($r) }
    }

    $Consolidate_Menu.Add_Click({ Join-SelectedItems })
    $Ungroup_Menu.Add_Click({ Split-SelectedGroups })

    # 8/24/26 - Assign Color submenu
    $Color_Red.Add_Click({    Set-ItemColor -ColorName "Red"    })
    $Color_Green.Add_Click({  Set-ItemColor -ColorName "Green"  })
    $Color_Yellow.Add_Click({ Set-ItemColor -ColorName "Yellow" })
    $Color_Blue.Add_Click({   Set-ItemColor -ColorName "Blue"   })
    $Color_Clear.Add_Click({  Set-ItemColor -ColorName "Clear"  })
    
    $Copy_Menu.Add_Click({
        Set-ClipBoard

    })

    # 8/24/26 - Insert Topic
    # Prompts for a topic name and appends a divider line to the bottom of the list.
    # Does not touch the clipboard, so the polling timer will not re-capture it.
    $Insert_Topic.Add_Click({

        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $topic = Show-InputPrompt -Message "Type a Topic to insert as a divider in the list." -Title "Insert Topic"

        # Cancel or blank entry - do nothing
        if ([string]::IsNullOrWhiteSpace($topic)) { return }

        $topic = $topic.Trim()

        # 11/22/26 - the item still holds the dashed divider line, so a saved file and a
        # copy read exactly as before. Update-ClipPreview recognises the pattern and the
        # row is drawn as a heading with just the title.
        $divider = ("-" * 11) + $topic + ("-" * 33)

        [void]$Script:ObservableCollection.Add((New-ClipItem -Text $divider -Numbered ([bool]$LineNumbers_Menu.IsChecked)))

        # Scroll to the bottom to show the new topic
        $listbox.ScrollIntoView($listbox.Items[$listbox.Items.Count-1])

    })
    
    $Save_Menu.Add_Click({
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null #
        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        $NameFilter = "Text Files (*.txt)|*.txt"
        $Dialog.Filter = $NameFilter
        
        if($Dialog.ShowDialog() -eq 'Ok'){
            # 0.8.1: "* " in front of a favorite
            $save_data = @($listbox.Items | ForEach-Object { Get-ExportText -Item $_ })
            $save_data | Out-File $Dialog.FileName
        }
        
    })

$Title_Window.Add_Click({ # 

    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    # 11/22/26 (0.8.1) - Cancel, or OK on an empty box, leaves the title as it was. The
    # old test here was true for every answer, so Cancel blanked the title. The entered
    # title is kept in $global:tempVar as before, because Pause restores the title from it.
    $newTitle = Show-InputPrompt -Message "Type the new Window Title." -Title "Informational" -DefaultText $Window.Title
    if (-not [string]::IsNullOrWhiteSpace($newTitle)) {
        $global:tempVar = $newTitle.Trim()
        $Window.Title = $global:tempVar
    }
})

    $Import_Menu.Add_Click({ # 
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        
        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $NameFilter = "Text Files (*.txt)|*.txt"
        $Dialog.Filter = $NameFilter
        
        if($Dialog.ShowDialog() -eq 'Ok'){
        
            Clear-Viewer
        
            Get-Content $Dialog.FileName | ForEach-Object {
                [void]$Script:ObservableCollection.Add((New-ClipItem -Text $_ -Numbered ([bool]$LineNumbers_Menu.IsChecked)))
            }

            # Scroll to the bottom to show the latest item.
            # 11/22/26 - once, after the whole file is in, instead of after every line.
            # Each scroll makes the list lay itself out, which made a big import slow.
            if ($listbox.Items.Count -gt 0) {
                $listbox.ScrollIntoView($listbox.Items[$listbox.Items.Count-1])
            }

        }
        
    })

    # 11/7/26 - New Instance
    # Starts a second, fully separate copy of this script in its own powershell.exe
    # process. Nothing is shared with this one: closing either window leaves the other
    # running. WScript.Shell.Run with style 0 is the same no-console trick the .vbs
    # launcher uses, so no window flashes. Launches with the same PowerShell host
    # (Windows PowerShell or pwsh) that is running this copy.
    # 11/7/26 - shared by Tools > New Instance and Open Selected in New Instance.
    # Returns $true if the process was started.
    Function Start-NewInstance {
        param([string]$SeedFile)

        $target = $Runspacehash.ScriptPath
        if ([string]::IsNullOrEmpty($target) -or -not (Test-Path -LiteralPath $target)) {
            [System.Windows.MessageBox]::Show($Window,
                "Could not find this script's file on disk, so a new instance cannot be started.`n`n$target",
                "New Instance", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $false
        }

        # the new window should not grab whatever happens to be on the clipboard right now
        $cmdArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$target`" -SkipInitialCapture"
        if (-not [string]::IsNullOrEmpty($SeedFile)) { $cmdArgs += " -SeedFile `"$SeedFile`"" }
        # 11/22/26 - the theme carries over to the new window
        # 11/23/26 (0.8.2) - the one in use now, which Tools > Theme may have changed
        if (@('Light', 'Dark', 'Auto') -contains $Script:ThemeChoice) { $cmdArgs += " -Theme $($Script:ThemeChoice)" }

        try {
            $hostExe = (Get-Process -Id $PID).Path
            $shell = New-Object -ComObject "WScript.Shell"
            # 0 = hidden console, False = do not wait for it
            [void]$shell.Run("`"$hostExe`" $cmdArgs", 0, $false)
            return $true
        } catch {
            [System.Windows.MessageBox]::Show($Window, "Could not start a new instance.`n`n$($_.Exception.Message)",
                "New Instance", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $false
        }
    }

    $NewInstance_Menu.Add_Click({ [void](Start-NewInstance) })

    # 11/7/26 - Open Selected in New Instance
    # Hands the selected items to a new instance through a temporary JSON file. Each
    # item carries its raw text, colour and whatever followed the text in its row (the
    # timestamp), but not its line number, so the new list numbers from 1. The new
    # instance deletes the file as soon as it has loaded it; if the launch fails it is
    # deleted here instead. The items also stay in this list - it is a copy, not a move.
    $OpenNewInstance.Add_Click({

        $selected = @($listbox.SelectedItems)
        if ($selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show($Window, "Select one or more items first.", "Open Selected in New Instance",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        # keep the order they appear in the list, not the order they were clicked
        $selected = @($selected | Sort-Object { $Script:ObservableCollection.IndexOf($_) })

        $payload = @(foreach ($item in $selected) {
            # 11/8/26 - groups travel as groups, parts included, so Ungroup works there too
            # 11/22/26 - the time stamp is read straight from the item (Suffix)
            @{ Text = $item.Text; Color = [string]$item.Color; Suffix = [string]$item.Suffix
               IsGroup = [bool]$item.IsGroup; GroupTitle = [string]$item.GroupTitle
               Favorite = [bool]$item.IsFavorite    # 0.8.1 - the star travels too
               Parts = $(if ($item.IsGroup) { @($item.Parts) } else { $null }) }
        })

        $seedPath = Join-Path $env:TEMP ("cbm-handoff-" + [Guid]::NewGuid().ToString("N") + ".json")
        try {
            [System.IO.File]::WriteAllText($seedPath, (ConvertTo-Json -InputObject $payload -Depth 6),
                (New-Object System.Text.UTF8Encoding($false)))
        } catch {
            [System.Windows.MessageBox]::Show($Window, "Could not prepare the items for the new instance.`n`n$($_.Exception.Message)",
                "Open Selected in New Instance", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        if (-not (Start-NewInstance -SeedFile $seedPath)) {
            try { Remove-Item -LiteralPath $seedPath -Force -ErrorAction SilentlyContinue } catch { }
        }
    })

    # 11/7/26 - Load items handed over by another instance, then delete the file.
    # Runs before the clipboard timer starts, so the line numbers begin at 1.
    Function Import-SeedItems {
        param([string]$Path)

        if ([string]::IsNullOrEmpty($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
        try {
            $json = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
            # assign first: in 5.1 ConvertFrom-Json emits a JSON array as ONE object,
            # so @() around the call would wrap it instead of enumerating it
            $seedItems = ConvertFrom-Json -InputObject $json
            foreach ($s in $seedItems) {
                $seeded = New-ClipItem -Text ([string]$s.Text) -Numbered ([bool]$LineNumbers_Menu.IsChecked) -Suffix ([string]$s.Suffix) -Color ([string]$s.Color)
                $seeded.IsFavorite = [bool]$s.Favorite    # 0.8.1
                if ($s.IsGroup) {
                    $seeded.IsGroup = $true
                    $seeded.GroupTitle = $s.GroupTitle
                    $seeded.Parts = @($s.Parts)
                    Update-ClipPreview -Item $seeded
                }
                [void]$Script:ObservableCollection.Add($seeded)
            }
        } catch {
            [System.Windows.MessageBox]::Show("Could not load the items passed from the other window.`n`n$($_.Exception.Message)",
                "Open Selected in New Instance", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        } finally {
            # clipboard content - do not leave it sitting in TEMP
            try { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue } catch { }
        }
    }

    # Set Stay On Top to be checked by default
    $StayTop_Menu.IsChecked = $true
    $Window.Topmost = $true

    # 8/24/26 - Line numbering on by default. Toggling only affects items added
    # from that point on; rows already in the list keep whatever they were given.
    # 11/22/26 - still true: each row remembers whether it is numbered. What changed is
    # that a numbered row always shows its current position in the list.
    $LineNumbers_Menu.IsChecked = $true
    
    $StayTop_Menu.Add_Click({
        if ($Window.Topmost -eq $True) {
            $Window.Topmost = $False
        } else {
            $Window.Topmost = $True
        }
    })

$Options_Menu.Add_Click({

# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# 8/24/26 - renamed from $Window. It was shadowing the main window variable inside
# this handler, and without an Owner it opened behind the always-on-top main window.
$OptionsWindow = New-Object System.Windows.Window
$OptionsWindow.Title = "Clipboard Manager Options"
if ($Script:AppIcon -ne $null) { $OptionsWindow.Icon = $Script:AppIcon }    # 11/22/26
$OptionsWindow.Width = 400
$OptionsWindow.Height = 450
$OptionsWindow.Owner = $Window
$OptionsWindow.WindowStartupLocation = "CenterOwner"
$OptionsWindow.Topmost = $true

# Attempt to display encoded images

$Labell = New-Object System.Windows.Controls.Label
$Labell.Content = "Right click and copy the settings you wish to activate `r Select OK."
$Labell.Margin = New-Object System.Windows.Thickness(10)

$listboxOptions = New-Object System.Windows.Controls.ListBox
$listboxOptions.Width = 200
$listboxOptions.Height = 200
$listboxOptions.Margin = New-Object System.Windows.Thickness(10)
# 8/24/26 - dropped "Suggestions", "Ideas", "Comments", "Thoughts": placeholder labels
# from 5/12 that never had handlers, so selecting them did nothing at all.
$listboxOptions.ItemsSource = @("Reset Masked Selections", "Add <br> Copying", "Remove <br> Copying")

$ContextMenu = New-Object System.Windows.Controls.ContextMenu

$MenuItem = New-Object System.Windows.Controls.MenuItem
$MenuItem.Header = "Copy"
$MenuItem.Add_Click({
$text = $listboxOptions.SelectedItem

if ($text -eq "Reset Masked Selections") 
{ 
    
    [System.Windows.Clipboard]::SetText($text)
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show('Masked Items Reset', 'Info', 'OK', [System.Windows.Forms.MessageBoxIcon]::Information)

}

# 5/12
if ($text -eq "Add <br> Copying") 
{

    [System.Windows.Clipboard]::SetText($text)
    #"Add <br> Copying", , "Remove <br> Copying"
    $global:bron = "bron"                            
}

if ($text -eq "Remove <br> Copying") 
{

    [System.Windows.Clipboard]::SetText($text)
    #"Add <br> Copying", , "Remove <br> Copying"
    $global:bron = "broff"                            
}

# Open self closing message box form
# 6/3/23 ozzzy

Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Forms;
    using System.Threading;

    public class AutoClosingMessageBox {
        System.Threading.Timer _timeoutTimer;
        string _caption;
        AutoClosingMessageBox(string text, string caption, int timeout) {
            _caption = caption;
            _timeoutTimer = new System.Threading.Timer(OnTimerElapsed, null, timeout, Timeout.Infinite);
            using(_timeoutTimer)
                MessageBox.Show(text, caption);
        }
        public static void Show(string text, string caption, int timeout) {
            new AutoClosingMessageBox(text, caption, timeout);
        }
        void OnTimerElapsed(object state) {
            IntPtr mbWnd = FindWindow("#32770", _caption); // lpClassName is #32770 for MessageBox
            if(mbWnd != IntPtr.Zero)
                SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            _timeoutTimer.Dispose();
        }
        const int WM_CLOSE = 0x0010;
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    }
"@

# Open ChatGPT Options from main menu...............................
if ($text -eq "ChatGPT Open Options") 
{ 
    
                        # Creates listbox to right click and execute command via watching the clipboard (This is a Chat Settings dialogbox and App config dialogbox = "Coming Soon")
                        Add-Type -AssemblyName PresentationFramework

                        $WindowChatt = New-Object System.Windows.Window
                        if ($Script:AppIcon -ne $null) { $WindowChatt.Icon = $Script:AppIcon }    # 11/22/26
                        $WindowChatt.SizeGripStyle = "Hide"
                        $WindowChatt.MaximizeBox = $false

                        $WindowChatt.Title = "ChatGPT Settings Dialogbox" 
                        $WindowChatt.Width = 425
                        $WindowChatt.Height = 475
                        $WindowChatt.WindowStartupLocation = "CenterOwner"
                        $WindowChatt.Owner = $OptionsWindow
                        $WindowChatt.Topmost = $true

                        $Labelll = New-Object System.Windows.Controls.Label
                        $Labelll.Content = "Right click and copy the setting you wish to activate `r Select OK. `r`r Note: These settings can be adjusted in TOOLS - Options `r Right Click and copy the ChatGPT Option"
                        $Labelll.Margin = New-Object System.Windows.Thickness(10)
                        
                        $listboxChatGPTt = New-Object System.Windows.Controls.ListBox
                        
                        $listboxChatGPTt.Width = 300
                        $listboxChatGPTt.Height = 200
                        
                        $listboxChatGPTt.Margin = New-Object System.Windows.Thickness(10)
                        $listboxChatGPTt.ItemsSource = @("ChatGPT Change GPT API", "ChatGPT Short Response", "ChatGPT Medium Response", "ChatGPT Complete Response", "ChatGPT Open Response Notepad", "ChatGPT Vocalize Response", "Open Windows Speech Settings", "ChatGPT Reset All Settings")
                        
                        $ContextMenuu = New-Object System.Windows.Controls.ContextMenu
                        $MenuItemm = New-Object System.Windows.Controls.MenuItem
                        $MenuItemm.Header = "Copy"
                        $MenuItemm.Add_Click({
                            $textt = $listboxChatGPTt.SelectedItem
                            [System.Windows.Clipboard]::SetText($textt)
                        })
                        $ContextMenuu.Items.Add($MenuItemm)
                        $listboxChatGPTt.ContextMenu = $ContextMenuu

                        $Buttonn = New-Object System.Windows.Controls.Button
                        $Buttonn.Content = "Close"
                        $Buttonn.Margin = New-Object System.Windows.Thickness(10)

                        $Buttonn.Add_Click(
                        { $WindowChatt.Close() })

                        $Labela = New-Object System.Windows.Controls.Label
                        $Labela.Content = "Developed by Duane Joseph (2023) / ChrisDee (GitHub 2016)"
                        $Labela.Margin = New-Object System.Windows.Thickness(10)
                        
                        $StackPanell = New-Object System.Windows.Controls.StackPanel
                        $StackPanell.Orientation = [System.Windows.Controls.Orientation]::Vertical
                        $StackPanell.Margin = New-Object System.Windows.Thickness(10)

                        $StackPanell.Children.Add($Labelll)
                        $StackPanell.Children.Add($listboxChatGPTt)
                        $StackPanell.Children.Add($Buttonn)
                        $StackPanell.Children.Add($Labela)
                        
                        $WindowChatt.Content = $StackPanell
                        $WindowChatt.ShowDialog() | Out-Null

 }

})

$ContextMenu.Items.Add($MenuItem)

$SubMenu = New-Object System.Windows.Controls.MenuItem
$SubMenu.Header = "Options"
$ContextMenu.Items.Add($SubMenu)

$Item1 = New-Object System.Windows.Controls.MenuItem
$Item1.Header = "Option 1"
$Item1.Add_Click({
Write-Host "Option 1 clicked"
})
$SubMenu.Items.Add($Item1)

$Item2 = New-Object System.Windows.Controls.MenuItem
$Item2.Header = "Option 2"
$Item2.Add_Click({
Write-Host "Option 2 clicked"
})
$SubMenu.Items.Add($Item2)

$listboxOptions.ContextMenu = $ContextMenu

$Button = New-Object System.Windows.Controls.Button
$Button.Content = "Close"
$Button.Margin = New-Object System.Windows.Thickness(10)
$Button.Add_Click({
$OptionsWindow.Close()
})

$Labelc = New-Object System.Windows.Controls.Label
$Labelc.Content = "Developed by Duane Joseph (2023) / ChrisDee (GitHub 2016)"
$Labelc.Margin = New-Object System.Windows.Thickness(10)

$StackPanel = New-Object System.Windows.Controls.StackPanel
$StackPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical
$StackPanel.Margin = New-Object System.Windows.Thickness(10)
$StackPanel.Children.Add($Labell)
$StackPanel.Children.Add($listboxOptions)
$StackPanel.Children.Add($Button)
$StackPanel.Children.Add($Labelc)
$OptionsWindow.Content = $StackPanel
$OptionsWindow.ShowDialog() | Out-Null

})

# 8/25/26 - AI Settings.
    # Two independent things live here, which is the whole point of the design:
    #   1. Routing  - which provider handles each capability
    #   2. Keys     - one per provider, entered on demand, memory only
    # A provider only appears in a capability's dropdown if it actually supports it,
    # so Claude never shows up under Speech-to-text.
    $AiSettings_Menu.Add_Click({

        $aiWindow = New-Object System.Windows.Window
        $aiWindow.Title = "AI Settings"
        if ($Script:AppIcon -ne $null) { $aiWindow.Icon = $Script:AppIcon }    # 11/22/26
        $aiWindow.Width = 520
        $aiWindow.SizeToContent = "Height"
        $aiWindow.ResizeMode = "NoResize"
        $aiWindow.ShowInTaskbar = $false
        $aiWindow.Owner = $Window
        $aiWindow.WindowStartupLocation = "CenterOwner"
        $aiWindow.Topmost = $true

        $root = New-Object System.Windows.Controls.StackPanel
        $root.Margin = New-Object System.Windows.Thickness(16)

        $intro = New-Object System.Windows.Controls.TextBlock
        $intro.Text = "Each capability is routed independently, so speech and chat can use different providers with different keys."
        $intro.TextWrapping = "Wrap"
        $intro.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        [void]$root.Children.Add($intro)

        # ---------- routing ----------
        $routeHeader = New-Object System.Windows.Controls.TextBlock
        $routeHeader.Text = "Routing"
        $routeHeader.FontWeight = "Bold"
        $routeHeader.Margin = New-Object System.Windows.Thickness(0,0,0,6)
        [void]$root.Children.Add($routeHeader)

        $combos = @{}

        foreach ($cap in @("Chat","SpeechToText")) {

            $row = New-Object System.Windows.Controls.StackPanel
            $row.Orientation = "Horizontal"
            $row.Margin = New-Object System.Windows.Thickness(0,0,0,6)

            $capLabel = New-Object System.Windows.Controls.TextBlock
            $capLabel.Text = $(if ($cap -eq "Chat") { "Ask AI" } else { "Speech to text" })
            $capLabel.Width = 130
            $capLabel.VerticalAlignment = "Center"
            [void]$row.Children.Add($capLabel)

            $combo = New-Object System.Windows.Controls.ComboBox
            $combo.Width = 240
            foreach ($id in (Get-AiProvidersFor -Capability $cap)) {
                $entry = New-Object System.Windows.Controls.ComboBoxItem
                $entry.Content = $Script:AiProviders[$id].Display
                $entry.Tag = $id
                [void]$combo.Items.Add($entry)
                if ($id -eq $global:AiRoute[$cap]) { $combo.SelectedItem = $entry }
            }
            if ($null -eq $combo.SelectedItem -and $combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }
            [void]$row.Children.Add($combo)
            $combos[$cap] = $combo

            [void]$root.Children.Add($row)
        }

        # text to speech is local only for now - shown so the routing picture is complete
        $ttsRow = New-Object System.Windows.Controls.StackPanel
        $ttsRow.Orientation = "Horizontal"
        $ttsRow.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        $ttsLabel = New-Object System.Windows.Controls.TextBlock
        $ttsLabel.Text = "Text to speech"
        $ttsLabel.Width = 130
        $ttsLabel.VerticalAlignment = "Center"
        [void]$ttsRow.Children.Add($ttsLabel)
        $ttsValue = New-Object System.Windows.Controls.TextBlock
        $ttsValue.Text = "Windows (offline)"
        $ttsValue.Opacity = 0.7
        $ttsValue.VerticalAlignment = "Center"
        [void]$ttsRow.Children.Add($ttsValue)
        [void]$root.Children.Add($ttsRow)

        # ---------- keys ----------
        $keyHeader = New-Object System.Windows.Controls.TextBlock
        $keyHeader.Text = "API keys (this session only)"
        $keyHeader.FontWeight = "Bold"
        $keyHeader.Margin = New-Object System.Windows.Thickness(0,0,0,6)
        [void]$root.Children.Add($keyHeader)

        $statusLabels = @{}

        foreach ($id in $Script:AiProviders.Keys) {

            $provId = $id     # capture per iteration, otherwise every button sees the last one
            $row = New-Object System.Windows.Controls.StackPanel
            $row.Orientation = "Horizontal"
            $row.Margin = New-Object System.Windows.Thickness(0,0,0,5)

            $name = New-Object System.Windows.Controls.TextBlock
            $name.Text = $Script:AiProviders[$provId].Display
            $name.Width = 170
            $name.VerticalAlignment = "Center"
            [void]$row.Children.Add($name)

            $status = New-Object System.Windows.Controls.TextBlock
            $status.Width = 90
            $status.VerticalAlignment = "Center"
            $status.Text = $(if ($global:AiKeys.ContainsKey($provId) -and -not [string]::IsNullOrWhiteSpace($global:AiKeys[$provId])) { "key set" } else { "not set" })
            $status.Opacity = 0.7
            [void]$row.Children.Add($status)
            $statusLabels[$provId] = $status

            $setBtn = New-Object System.Windows.Controls.Button
            $setBtn.Content = "Set key"
            $setBtn.Width = 80
            $setBtn.Margin = New-Object System.Windows.Thickness(0,0,6,0)
            $setBtn.Tag = $provId
            $setBtn.Add_Click({
                $target = $this.Tag
                $null = Get-AiKey -Provider $target -Force
                $statusLabels[$target].Text = $(if ($global:AiKeys.ContainsKey($target) -and -not [string]::IsNullOrWhiteSpace($global:AiKeys[$target])) { "key set" } else { "not set" })
            })
            [void]$row.Children.Add($setBtn)

            $clearBtn = New-Object System.Windows.Controls.Button
            $clearBtn.Content = "Clear"
            $clearBtn.Width = 70
            $clearBtn.Tag = $provId
            $clearBtn.Add_Click({
                $target = $this.Tag
                $global:AiKeys[$target] = $null
                $statusLabels[$target].Text = "not set"
            })
            [void]$row.Children.Add($clearBtn)

            # note which capabilities this provider can actually serve
            $caps = @()
            if (-not [string]::IsNullOrEmpty($Script:AiProviders[$provId].ChatUrl)) { $caps += "chat" }
            if (-not [string]::IsNullOrEmpty($Script:AiProviders[$provId].SttUrl))  { $caps += "speech" }
            $capNote = New-Object System.Windows.Controls.TextBlock
            $capNote.Text = "   " + ($caps -join ", ")
            $capNote.Opacity = 0.5
            $capNote.FontSize = 11
            $capNote.VerticalAlignment = "Center"
            [void]$row.Children.Add($capNote)

            [void]$root.Children.Add($row)
        }

        $footnote = New-Object System.Windows.Controls.TextBlock
        $footnote.Text = "Claude has no speech API, so it is not offered for speech to text. Gemini transcription needs a two-step file upload and is not wired up."
        $footnote.TextWrapping = "Wrap"
        $footnote.FontSize = 11
        $footnote.Opacity = 0.6
        $footnote.Margin = New-Object System.Windows.Thickness(0,8,0,14)
        [void]$root.Children.Add($footnote)

        # ---------- buttons ----------
        $btnRow = New-Object System.Windows.Controls.StackPanel
        $btnRow.Orientation = "Horizontal"
        $btnRow.HorizontalAlignment = "Right"

        $okBtn = New-Object System.Windows.Controls.Button
        $okBtn.Content = "OK"
        $okBtn.Width = 80
        $okBtn.IsDefault = $true
        $okBtn.Margin = New-Object System.Windows.Thickness(0,0,8,0)
        $okBtn.Add_Click({ $aiWindow.DialogResult = $true })

        $cancelBtn = New-Object System.Windows.Controls.Button
        $cancelBtn.Content = "Cancel"
        $cancelBtn.Width = 80
        $cancelBtn.IsCancel = $true

        [void]$btnRow.Children.Add($okBtn)
        [void]$btnRow.Children.Add($cancelBtn)
        [void]$root.Children.Add($btnRow)

        $aiWindow.Content = $root

        if ($aiWindow.ShowDialog() -eq $true) {
            foreach ($cap in @("Chat","SpeechToText")) {
                $chosen = $combos[$cap].SelectedItem
                if ($null -ne $chosen) { $global:AiRoute[$cap] = $chosen.Tag }
            }
        }
    })

# 8/25/26 - Voice Note
    # Records from the default microphone with the Windows MCI recorder, then posts the
    # wav to xAI's speech-to-text endpoint.
    #
    #   POST https://api.x.ai/v1/stt   multipart/form-data
    #   NOTE: xAI requires the 'file' part to be added AFTER all other form fields.
    #
    # The API key is prompted for once per session and held in memory only, matching how
    # the ChatGPT key is handled. The temporary wav is deleted as soon as it is sent.
    $VoiceNote_Menu.Add_Click({

        # 8/25/26 - the router owns key prompting and provider choice now
        $sttProvider = $global:AiRoute["SpeechToText"]
        if ([string]::IsNullOrEmpty($Script:AiProviders[$sttProvider].SttUrl)) {
            [System.Windows.MessageBox]::Show($Window,
                "$($Script:AiProviders[$sttProvider].Display) has no speech-to-text API.`n`nChoose a different provider under Tools > AI Settings.",
                "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $wavPath = Join-Path $env:TEMP ("cbm-voice-" + [Guid]::NewGuid().ToString("N") + ".wav")

        # --- start recording ---
        $usedFormat = ""
        try {
            $usedFormat = [CbmRecorder]::Start()
        } catch {
            [System.Windows.MessageBox]::Show($Window,
                "Could not start the microphone.`n`n$($_.Exception.Message)",
                "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        # --- recording window ---
        $recWindow = New-Object System.Windows.Window
        $recWindow.Title = "Voice Note - Recording"
        if ($Script:AppIcon -ne $null) { $recWindow.Icon = $Script:AppIcon }    # 11/22/26
        $recWindow.Width = 380
        $recWindow.SizeToContent = "Height"
        $recWindow.ResizeMode = "NoResize"
        $recWindow.ShowInTaskbar = $false
        $recWindow.Owner = $Window
        $recWindow.WindowStartupLocation = "CenterOwner"
        $recWindow.Topmost = $true

        $recPanel = New-Object System.Windows.Controls.StackPanel
        $recPanel.Margin = New-Object System.Windows.Thickness(16)

        $recLabel = New-Object System.Windows.Controls.TextBlock
        $recLabel.Text = "Recording. Speak now."
        $recLabel.FontSize = 15
        $recLabel.Margin = New-Object System.Windows.Thickness(0,0,0,6)
        [void]$recPanel.Children.Add($recLabel)

        # which format the driver actually accepted - useful when diagnosing
        $recFormat = New-Object System.Windows.Controls.TextBlock
        $recFormat.Text = $usedFormat
        $recFormat.FontSize = 10
        $recFormat.Opacity = 0.6
        $recFormat.Margin = New-Object System.Windows.Thickness(0,0,0,4)
        [void]$recPanel.Children.Add($recFormat)

        $recTimerLabel = New-Object System.Windows.Controls.TextBlock
        $recTimerLabel.Text = "00:00"
        $recTimerLabel.FontSize = 22
        $recTimerLabel.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        [void]$recPanel.Children.Add($recTimerLabel)

        $recButtons = New-Object System.Windows.Controls.StackPanel
        $recButtons.Orientation = "Horizontal"
        $recButtons.HorizontalAlignment = "Right"

        $stopBtn = New-Object System.Windows.Controls.Button
        $stopBtn.Content = "Stop and Transcribe"
        $stopBtn.Width = 140
        $stopBtn.IsDefault = $true
        $stopBtn.Margin = New-Object System.Windows.Thickness(0,0,8,0)
        $stopBtn.Add_Click({ $recWindow.DialogResult = $true })

        $abortBtn = New-Object System.Windows.Controls.Button
        $abortBtn.Content = "Cancel"
        $abortBtn.Width = 80
        $abortBtn.IsCancel = $true

        [void]$recButtons.Children.Add($stopBtn)
        [void]$recButtons.Children.Add($abortBtn)
        [void]$recPanel.Children.Add($recButtons)
        $recWindow.Content = $recPanel

        # elapsed counter
        $started = Get-Date
        $recTimer = New-Object System.Windows.Threading.DispatcherTimer
        $recTimer.Interval = [TimeSpan]::FromSeconds(1)
        $recTimer.Add_Tick({
            $elapsed = (Get-Date) - $started
            $recTimerLabel.Text = "{0:mm\:ss}" -f $elapsed
        })
        $recTimer.Start()

        $recResult = $recWindow.ShowDialog()
        $recTimer.Stop()

        if ($recResult -ne $true) {
            [CbmRecorder]::Abort()
            return
        }

        # --- stop and write the wav ---
        try {
            [CbmRecorder]::SaveAndClose($wavPath)
        } catch {
            [CbmRecorder]::Abort()
            [System.Windows.MessageBox]::Show($Window,
                "Could not save the recording.`n`n$($_.Exception.Message)",
                "Voice Note", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        if (-not (Test-Path $wavPath)) {
            [System.Windows.MessageBox]::Show($Window, "Nothing was recorded.", "Voice Note",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        # --- transcribe ---
        $Window.Cursor = [System.Windows.Input.Cursors]::Wait
        $transcribed = ""
        $errorText = ""
        try {
            $transcribed = Invoke-AiTranscribe -WavPath $wavPath
        } finally {
            $Window.Cursor = $null
            # the recording is voice data - do not leave it sitting in TEMP
            try { Remove-Item $wavPath -Force -ErrorAction SilentlyContinue } catch { }
        }

        if ([string]::IsNullOrWhiteSpace($transcribed)) {
            [System.Windows.MessageBox]::Show($Window, "No speech was detected.", "Voice Note",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }

        # --- review before saving ---
        $noteText = Show-InputPrompt -Message "Transcribed note - edit if needed, then press OK." -Title "Voice Note" -DefaultText $transcribed.Trim()
        if ([string]::IsNullOrWhiteSpace($noteText)) { return }

        $suffix = ""
        if ($AddTime_Menu.IsChecked) { $suffix = " : " + (Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss") }

        $noteText = "[Voice] " + $noteText.Trim()

        [void]$Script:ObservableCollection.Add((New-ClipItem -Text $noteText -Numbered ([bool]$LineNumbers_Menu.IsChecked) -Suffix $suffix))
        $listbox.ScrollIntoView($listbox.Items[$listbox.Items.Count-1])
    })

    $Pause_Menu.Add_Click({ # 3/7
        if ($Pause_Menu.IsChecked -eq $False) {
            
            $Window.Title = $global:tempVar 
        } else {
            $global:tempVar = $Window.Title
            $Window.Title = "Pause - " + $global:tempVar
        }
        # 11/22/26 - the status strip says Paused / Capturing as well
        Update-StatusStrip
    })

    $Window.Add_Activated({ #6/5
    })

    $Window.Add_SourceInitialized({

        # 11/22/26 - a dark window gets a dark title bar
        # 11/23/26 (0.8.2) - the call itself moved into Set-TitleBarTheme
        if ($Ui.IsDark) { Set-TitleBarTheme }

        #Create observable collection # 12/5
        $Script:ObservableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[object]

        # 11/22/26 - Line numbers follow position. Whenever the list changes, every row
        # from the change downwards is given its position as its number, so numbers run
        # 1..N after a remove, a drag, a consolidate or an ungroup without any of those
        # having to ask. Hooked up before ItemsSource so a new row already has its
        # number when it is first drawn.
        $Script:ObservableCollection.Add_CollectionChanged({
            $change = $_    # inside the switch below, $_ is the value being switched on
            $from = 0
            switch ([string]$change.Action) {
                'Add'    { $from = $change.NewStartingIndex }
                'Remove' { $from = $change.OldStartingIndex }
                'Move'   { $from = [Math]::Min($change.OldStartingIndex, $change.NewStartingIndex) }
            }
            if ($from -lt 0) { $from = 0 }
            Update-RowNumbers -From $from
        })

        $Listbox.ItemsSource = $Script:ObservableCollection

        # 11/22/26 - the status strip counts follow the list too. Hooked up after
        # ItemsSource on purpose: handlers run in the order they were added, so this one
        # runs once the list itself has taken the change in. In the handler above the
        # list is still one change behind, and the strip would say "12 of 13 items".
        $Script:ObservableCollection.Add_CollectionChanged({ Update-StatusStrip })

        # 11/7/26 - started from another instance: preload the handed-over items and
        # treat what is already on the clipboard as seen, so it is not captured again
        Import-SeedItems -Path $Runspacehash.SeedFile
        if ($Runspacehash.SkipInitialCapture) {
            try { $Script:Previous = Get-Clipboard -Raw -ErrorAction SilentlyContinue } catch { }
        }

        Update-StatusStrip

        #Create Timer object
        $Script:timer = new-object System.Windows.Threading.DispatcherTimer 
        $timer.Interval = [TimeSpan]"0:0:.1"

# 3/14

        #Add event per tick
        $timer.Add_Tick({

            $text = $null
            
            try {
                # Try to get clipboard text with error handling
                $text = Get-Clipboard -Raw -ErrorAction SilentlyContinue
            } catch {
                # If standard Get-Clipboard fails, try alternative method
                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $text = [System.Windows.Forms.Clipboard]::GetText()
                } catch {
                    # If all methods fail, set text to empty
                    $text = ""
                }
            }
            
            # Ensure we have valid text and it's different from previous captures
            If (($Script:Previous -ne $text -AND $Script:CopiedText -ne $text) -AND $text -ne $null -AND $text.length -gt 0) {
                
                # 11/22/26 (0.8.1) - the first capture used to swap the window title for
                # "Good Morning" / "Good Afternoon" / "Have a good evening". Removed: the
                # title stays "Clip Board Manager <version>" until Tools > Title Window
                # changes it.

                if ($Pause_Menu.IsChecked -eq $False) {

                        # 8/25/26 - The old ChatGPT settings were driven by COPYING magic
                        # strings ("ChatGPT Short Response" etc) and having the timer notice
                        # them. That is gone; Tools > AI Settings owns provider and key
                        # choice now. Masking reset is the one command still worth keeping.
                        if ($text -eq "Reset Masked Selections")
                        {
                            $global:MaskedTerms.Clear()
                        }

                    # 8/25/26 - one call. The timestamp / no-timestamp split used to be two
                    # near-duplicate blocks that drifted apart; Add-ClipboardItem handles both.
                    Add-ClipboardItem -Text $text
                }

                $Script:Previous = $text
            }
             
        })
    
        $timer.Start()
        If (-NOT $timer.IsEnabled) {
            $Window.Close()
        }
    })

    # 8/25/26 - Guard against losing the whole history to a stray click on the X.
    # Uses the owner overload of MessageBox::Show so the prompt cannot open behind
    # the always-on-top main window.
    $Window.Add_Closing({
        param($eventSender, $eventArgs)

        if ($Script:ObservableCollection -ne $null -and $Script:ObservableCollection.Count -gt 0) {

            $count = $Script:ObservableCollection.Count
            $answer = [System.Windows.MessageBox]::Show(
                $Window,
                "Close the clipboard viewer?`n`n$count item(s) in the history will be lost.",
                "Confirm Exit",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)

            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
                $eventArgs.Cancel = $true
            }
        }
    })

    $Window.Add_Closed({
        $Script:timer.Stop()
        $Script:ObservableCollection.Clear()
        $Runspacehash.PowerShell.Dispose()
    })

    # 11/22/26 (0.8.1) - the filter itself moved into Update-ListFilter, which also
    # applies Show Only Favorites
    $InputBox.Add_TextChanged({ #6/5/23
        Update-ListFilter
    })

    # 0.8.1 - Show Only Favorites: the star button beside the filter box and the
    # Tools menu item are the same switch
    $FavOnly_Button.Add_Click({ Set-FavoritesOnly -On ([bool]$FavOnly_Button.IsChecked) })
    $FavOnly_Menu.Add_Click({ Set-FavoritesOnly -On ([bool]$FavOnly_Menu.IsChecked) })

    # 11/23/26 (0.8.2) - Tools > Theme
    $Theme_Light.Add_Click({ Set-UiTheme -Choice 'Light' })
    $Theme_Dark.Add_Click({ Set-UiTheme -Choice 'Dark' })
    $Theme_Auto.Add_Click({ Set-UiTheme -Choice 'Auto' })

    # 0.8.1 - right-click > Favorite: star every selected row, or unstar them when
    # every one of them already is a favorite
    $Favorite_Menu.Add_Click({
        $rows = @($listbox.SelectedItems | Where-Object { -not $_.IsTopic })
        if ($rows.Count -eq 0) { return }
        $allSet = (@($rows | Where-Object { -not $_.IsFavorite }).Count -eq 0)
        Set-ItemsFavorite -Items $rows -Favorite (-not $allSet)
    })

    # the menu item says what it is about to do
    $ClipboardMenu.Add_Opened({
        $rows = @($listbox.SelectedItems | Where-Object { -not $_.IsTopic })
        $allSet = ($rows.Count -gt 0 -and @($rows | Where-Object { -not $_.IsFavorite }).Count -eq 0)
        if ($allSet) {
            $Favorite_Menu.Header = "Remove Favorite"
            $Favorite_Menu.Icon = [string][char]0xE735
        } else {
            $Favorite_Menu.Header = "Favorite"
            $Favorite_Menu.Icon = [string][char]0xE734
        }
        $Favorite_Menu.IsEnabled = ($rows.Count -gt 0)
    })

    # 0.8.1 - Options > Select Favorites: selects every favorite the list is showing,
    # so Copy, Save, Append and the rest can be used on them
    $SelectFav_Menu.Add_Click({
        $stars = @($listbox.Items | Where-Object { $_.IsFavorite })
        $listbox.SelectedItems.Clear()
        foreach ($s in $stars) { [void]$listbox.SelectedItems.Add($s) }
        if ($stars.Count -gt 0) { $listbox.ScrollIntoView($stars[0]) }
    })

    # 11/22/26 - keep the "N selected" count in the status strip current
    $listbox.Add_SelectionChanged({ Update-StatusStrip })

    # 11/22/26 - A row that is cut short shows its text in a tooltip; a row that already
    # shows everything does not. "Cut short" means a collapsed long item or group, or a
    # line that runs past the window edge. The rule is RowTip, in the C# block near the
    # top of this script, where the reason it cannot be an ordinary handler is given.
    [Cbm082.RowTip]::Hook()

    $listbox.Add_MouseRightButtonUp({
        If ($Script:ObservableCollection.Count -gt 0) {
            $Remove_Menu.IsEnabled = $True
            $Copy_Menu.IsEnabled = $True
        } Else {
            $Remove_Menu.IsEnabled = $False
            $Copy_Menu.IsEnabled = $False
        }
    })

    $Window.Add_KeyDown({ 
        $key = $_.Key

        # 8/25/26 - F2 edits the selected item, matching Explorer's rename shortcut
        If ($key -eq "F2") { Edit-SelectedItem; return }
  
        If ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) {
            Switch ($Key) {
            "C" {
                Set-ClipBoard                                          
            }
            "R" {
                @($listbox.SelectedItems) | ForEach-Object {
                    [void]$Script:ObservableCollection.Remove($_)
                }            
            }
            "E" {
                $This.Close()
            }
            "W" { #6/2/23 Duane adding keyboard function to Write File
               
                Create_File_Selected_keyboard

            }
            "Q" { #6/3/23 Duane adding keyboard function to Append to File

                Append_File_Selected_keyboard
                # ozzzy
            }
            "L" { #6/3/23222222 Duane adding keyboard function to add selected item to review log 
                
                Log_File_Selected_keyboard #-pathh $global:ReviewLog

            }
            "O" { #6/3/23222222 Duane adding keyboard function to add selected item to review log 
                
                OpenLog_File_Selected_keyboard #

            }
            Default {$Null}
            }
        }
        # 8/25/26 - the old hidden-textbox inline edit path lived here. Edit item /
        # F2 replaces it, so the editBox and indexBox controls are gone too.
    })

    [void]$Window.ShowDialog()
    
}).BeginInvoke()

# 8/25/26 - The runspace above is started asynchronously, so this script used to fall
# straight off the end. The only thing keeping the process (and therefore the GUI)
# alive was the console window held open by -noexit - which is why closing that
# console killed the history.
# Waiting on the async handle here keeps the process alive on its own, so the script
# can be launched hidden and there is no console left to close by accident.
try {
    while (-not $Runspacehash.Handle.IsCompleted) {
        Start-Sleep -Milliseconds 250
    }
} catch {
    # the runspace disposes itself on window close; nothing useful to do here
}

