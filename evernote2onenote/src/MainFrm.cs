// Evernote2Onenote - imports Evernote notes to Onenote
// Copyright (C) 2014, 2023 - Stefan Kueng

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using Evernote2Onenote.Enums;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Windows.Forms;
using System.Xml;

using OneNote = Microsoft.Office.Interop.OneNote;

namespace Evernote2Onenote
{
    public partial class MainFrm : Form
    {
        // ===== Fields =====

        private string _evernoteNotebookPath;
        private readonly SynchronizationContext _synchronizationContext;
        private bool _cancelled;
        private SyncStep _syncStep = SyncStep.Start;
        private Microsoft.Office.Interop.OneNote.Application _onApp;

        private readonly string _xmlNewOutlineContent =
            "<one:Meta name=\"{2}\" content=\"{1}\"/>" +
            "<one:OEChildren><one:HTMLBlock><one:Data><![CDATA[{0}]]></one:Data></one:HTMLBlock>{3}</one:OEChildren>";

        private const string XmlSourceUrl =
            "<one:OE alignment=\"left\" quickStyleIndex=\"2\"><one:T><![CDATA[From &lt;<a href=\"{0}\">{0}</a>&gt; ]]></one:T></one:OE>";

        private const string XmlNewOutline =
            "<?xml version=\"1.0\"?>" +
            "<one:Page xmlns:one=\"{2}\" ID=\"{1}\" dateTime=\"{5}\">" +
            "<one:Title selected=\"partial\" lang=\"en-US\">" +
            "<one:OE creationTime=\"{5}\" lastModifiedTime=\"{5}\">" +
            "<one:T><![CDATA[{3}]]></one:T> " +
            "</one:OE>" +
            "</one:Title>{4}" +
            "<one:Outline>{0}</one:Outline></one:Page>";

        private const string Xmlns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        private string _enNotebookName = "";
        private bool _useUnfiledSection;
        private string _newnbId = "";

        private readonly string _cmdNoteBook = "";
        private DateTime _cmdDate = new DateTime(0);

        // Multi-file import state
        private string _importFolderPath = "";
        private string _logFolderPath = "";
        private List<ImportJob> _importJobs = new List<ImportJob>();
        private string _currentManifestPath = "";
        private bool _headlessMode; // true when launched via CLI args (auto-close on finish)

        // Regexes (unchanged from original)
        private readonly Regex _rxStyle = new Regex("(?<text>\\<(?:pre|code|div|span|li|ul|ol|p|td|tr|table|tbody|h1|h2|a\\s+href=\\\"[^\\\"]*\\\").)\\s*style=\\\"[^\\\"]*\\\"", RegexOptions.IgnoreCase);
        private readonly Regex _rxFontFamily = new Regex(@"font-family: \""[^\""]*\""", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdata = new Regex(@"<!\[CDATA\[<\?xml version=[""']1.0[""'][^?]*\?>", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdata2 = new Regex(@"<!\[CDATA\[<!DOCTYPE en-note \w+ ""https?://xml.evernote.com/pub/enml2.dtd"">", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdataInner = new Regex(@"\<\!\[CDATA\[(?<text>.*)\]\]\>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private readonly Regex _rxEmptyCdata = new Regex(@"<!\[CDATA\[<\?xml version=[""']1.0[""'][^?]*\?>[\s\n]+\]\]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxEmptyCdata2 = new Regex(@"<!\[CDATA\[[\s\n]+\]\]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxEmptyCdata3 = new Regex(@"<!\[CDATA\[>[\s\n]+\]\]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxBodyStart = new Regex(@"<en-note[^>/]*>", RegexOptions.IgnoreCase);
        private readonly Regex _rxBodyEnd = new Regex(@"</en-note\s*>\s*]]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxBodyEmpty = new Regex(@"<en-note[^>/]*/>\s*]]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxDate = new Regex(@"^date:(.*)$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        private readonly Regex _rxNote = new Regex("<title>(.+)</title>", RegexOptions.IgnoreCase);
        private readonly Regex _rxComment = new Regex("<!--(.+)-->", RegexOptions.IgnoreCase);
        private readonly Regex _rxDtd = new Regex(@"<!DOCTYPE en-note SYSTEM \""http:\/\/xml\.evernote\.com\/pub\/enml\d*\.dtd\"">", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _rxBrOnly = new Regex("<br( |/)", RegexOptions.IgnoreCase);

        // ===== Constructor =====

        public MainFrm(string cmdNotebook, string cmdDate, string cmdLogFolder = "")
        {
            InitializeComponent();
            LayoutBrowseRows();
            _synchronizationContext = SynchronizationContext.Current;
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            versionLabel.Text = $@"Version: {version}";

            if (cmdDate.Length > 0)
            {
                try
                {
                    _cmdDate = DateTime.Parse(cmdDate);
                }
                catch (Exception)
                {
                    MessageBox.Show($"The Datestring\n{cmdDate}\nis not valid!");
                }
            }
            try
            {
                importDatePicker.Value = _cmdDate;
            }
            catch (Exception)
            {
                importDatePicker.Value = importDatePicker.MinDate;
            }

            if (cmdNotebook.Length > 0)
            {
                if (Directory.Exists(cmdNotebook))
                {
                    // CLI headless mode: treat arg as import folder path
                    _importFolderPath = cmdNotebook;
                    _logFolderPath = string.IsNullOrEmpty(cmdLogFolder) ? cmdNotebook : cmdLogFolder;
                    _headlessMode = true;
                    txtImportFolder.Text = _importFolderPath;
                    txtLogFolder.Text = _logFolderPath;
                    txtNotebookName.Text = SanitizeName(
                        Path.GetFileName(_importFolderPath.TrimEnd('\\', '/')), 30);
                    StartSync();
                }
                else
                {
                    // Legacy: treat as a notebook name override (no auto-start without a folder)
                    _cmdNoteBook = cmdNotebook;
                }
            }
        }

        // ===== Layout =====

        private void LayoutBrowseRows()
        {
            const int rightMargin = 15;
            const int gap = 6;
            int btnLeft = ClientSize.Width - rightMargin - btnBrowseImportFolder.Width;
            btnBrowseImportFolder.Left = btnLeft;
            btnBrowseLogFolder.Left    = btnLeft;
            txtImportFolder.Width = btnLeft - txtImportFolder.Left - gap;
            txtLogFolder.Width    = btnLeft - txtLogFolder.Left    - gap;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutBrowseRows();
        }

        // ===== Event Handlers =====

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnBrowseImportFolder_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the folder containing your .enex files";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _importFolderPath = dlg.SelectedPath;
                    txtImportFolder.Text = _importFolderPath;
                    if (string.IsNullOrWhiteSpace(txtNotebookName.Text))
                        txtNotebookName.Text = SanitizeName(
                            Path.GetFileName(_importFolderPath.TrimEnd('\\', '/')), 30);
                }
            }
        }

        private void btnBrowseLogFolder_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the folder where import manifests will be saved";
                dlg.ShowNewFolderButton = true;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _logFolderPath = dlg.SelectedPath;
                    txtLogFolder.Text = _logFolderPath;
                }
            }
        }

        private void btnENEXImport_Click(object sender, EventArgs e)
        {
            if (btnENEXImport.Text == "Cancel")
            {
                _cancelled = true;
                return;
            }

            // Accept typed paths in addition to browsed ones
            _importFolderPath = txtImportFolder.Text.Trim();
            _logFolderPath    = txtLogFolder.Text.Trim();

            if (string.IsNullOrEmpty(_importFolderPath) || !Directory.Exists(_importFolderPath))
            {
                MessageBox.Show("Please select or enter a valid import folder path.");
                return;
            }
            if (string.IsNullOrEmpty(_logFolderPath))
            {
                MessageBox.Show("Please select or enter a manifest log folder path.");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtNotebookName.Text))
            {
                MessageBox.Show("Please enter a notebook name.");
                return;
            }

            StartSync();
        }

        private void homeLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://tools.stefankueng.com/Evernote2Onenote.html");
        }

        private void MainFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_onApp != null)
                _cancelled = true;
            _onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // ===== Progress =====

        private void SetInfo(string line1, string line2, int pos, int max)
        {
            var fullpos = 0;

            switch (_syncStep)
            {
                case SyncStep.ExtractNotes:      // 0-10%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) : 0;
                    break;
                case SyncStep.ParseNotes:        // 10-20%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) + 10000 : 10000;
                    break;
                case SyncStep.CalculateWhatToDo: // 30-35%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.05) + 30000 : 30000;
                    break;
                case SyncStep.ImportNotes:       // 35-100%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.65) + 35000 : 35000;
                    break;
            }

            _synchronizationContext.Send(delegate
            {
                if (line1 != null)
                    infoText1.Text = line1;
                if (line2 != null)
                    infoText2.Text = line2;
                progressIndicator.Minimum = 0;
                progressIndicator.Maximum = 100000;
                progressIndicator.Value = fullpos;
            }, null);

            if (max == 0)
                _syncStep++;
        }

        // ===== StartSync =====

        private void StartSync()
        {
            try
            {
                _onApp = new OneNote.Application();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n\n{ex}");
                return;
            }
            if (_onApp == null)
            {
                MessageBox.Show(
                    "Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n");
                return;
            }

            // Read and validate notebook name from form field
            _enNotebookName = SanitizeName(txtNotebookName.Text.Trim(), 30);
            if (string.IsNullOrEmpty(_enNotebookName))
            {
                MessageBox.Show("Please enter a notebook name.", "Notebook Name Required",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fetch existing notebooks so we can warn about name conflicts
            _onApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out var xmlHierarchy);

            if (!_headlessMode && NotebookExists(_enNotebookName, xmlHierarchy))
            {
                var result = MessageBox.Show(
                    $"A notebook named \"{_enNotebookName}\" already exists in OneNote.\n\nDo you want to import into it anyway?",
                    "Notebook Already Exists",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                    return;
            }

            // Create the single target notebook for this entire import run
            try
            {
                try
                {
                    _onApp.GetSpecialLocation(OneNote.SpecialLocation.slDefaultNotebookFolder, out _evernoteNotebookPath);
                }
                catch (Exception)
                {
                    _onApp.GetSpecialLocation(OneNote.SpecialLocation.slUnfiledNotesSection, out _evernoteNotebookPath);
                }

                _evernoteNotebookPath += "\\" + _enNotebookName;
                _onApp.OpenHierarchy(_evernoteNotebookPath, "", out var newnbId, OneNote.CreateFileType.cftNotebook);
                _onApp.GetHierarchy(newnbId, OneNote.HierarchyScope.hsPages, out _);

                var docHierarchy = new XmlDocument();
                docHierarchy.LoadXml(xmlHierarchy);
                var hierarchy = new StringBuilder();
                AppendHierarchy(docHierarchy.DocumentElement, hierarchy, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Could not create the target notebook in Onenote!\nTrying to use the \"unfiled notes\" section instead.\nThe error was:\n\n{ex}");
                try
                {
                    _onApp.GetHierarchy("", OneNote.HierarchyScope.hsPages, out _);
                    _onApp.GetSpecialLocation(OneNote.SpecialLocation.slUnfiledNotesSection, out _evernoteNotebookPath);
                    _onApp.OpenHierarchy(_evernoteNotebookPath, "", out _newnbId);
                    _onApp.GetHierarchy(_newnbId, OneNote.HierarchyScope.hsPages, out _);
                    _useUnfiledSection = true;
                }
                catch (Exception ex2)
                {
                    MessageBox.Show(
                        $"Could not create the target notebook in Onenote!\nMake sure you have the desktop version of OneNote installed, not the one from the Windows store!\n\n{ex2}");
                    return;
                }
            }

            if (importDatePicker.Value > _cmdDate)
                _cmdDate = importDatePicker.Value;

            // Check for a resumable manifest before scanning
            var resumable = ManifestManager.FindResumable(_logFolderPath);
            if (resumable != null)
            {
                var answer = MessageBox.Show(
                    $"An existing import manifest was found:\n{Path.GetFileName(resumable)}\n\nResume it?",
                    "Resume import?", MessageBoxButtons.YesNo);

                if (answer == DialogResult.Yes)
                {
                    _currentManifestPath = resumable;

                    // Mark any IN_PROGRESS entries as crashed and increment their count
                    var allEntries = ManifestManager.Load(resumable);
                    foreach (var entry in allEntries.Where(e => e.Status == ManifestStatus.InProgress))
                    {
                        entry.AttemptCount++;
                        entry.Status = entry.AttemptCount >= 3
                            ? ManifestStatus.Failed
                            : ManifestStatus.Pending;
                        ManifestManager.UpdateEntry(resumable, entry);
                    }

                    // Rebuild job list from remaining resumable entries
                    _importJobs = ManifestManager.Load(resumable)
                        .Where(e => e.Status == ManifestStatus.Pending ||
                                    (e.Status == ManifestStatus.Failed && e.AttemptCount < 3))
                        .Select(entry => new ImportJob
                        {
                            EnexPath          = Path.Combine(_importFolderPath, entry.RelativePath),
                            RelativePath      = entry.RelativePath,
                            OneNoteTargetPath = entry.OneNoteTarget,
                            SectionGroupName  = ExtractSectionGroupFromTarget(entry.OneNoteTarget),
                            SectionName       = ExtractSectionNameFromTarget(entry.OneNoteTarget)
                        })
                        .ToList();
                }
                else
                {
                    resumable = null;
                }
            }

            if (resumable == null)
            {
                _importJobs = ScanImportFolder(_importFolderPath);
                _currentManifestPath = ManifestManager.Create(_logFolderPath, _importJobs);
            }

            _cancelled = false;
            btnENEXImport.Text = "Cancel";
            MethodInvoker syncDelegate = ImportAllNotesToOnenote;
            syncDelegate.BeginInvoke(null, null);
        }

        // ===== Folder Scanning =====

        private List<ImportJob> ScanImportFolder(string rootFolder)
        {
            _syncStep = SyncStep.ScanningFolder;
            var jobs = new List<ImportJob>();

            var rootEnex = Directory.GetFiles(rootFolder, "*.enex", SearchOption.TopDirectoryOnly);
            var subfolders = Directory.GetDirectories(rootFolder)
                .Where(d => Directory.GetFiles(d, "*.enex", SearchOption.TopDirectoryOnly).Length > 0)
                .ToArray();

            string rootGroupName = null;

            foreach (var file in rootEnex)
            {
                var sec    = SanitizeName(Path.GetFileNameWithoutExtension(file), 80);
                var target = rootGroupName != null
                    ? $"{_enNotebookName}\\{rootGroupName}\\{sec}"
                    : $"{_enNotebookName}\\{sec}";

                jobs.Add(new ImportJob
                {
                    EnexPath          = file,
                    RelativePath      = Path.GetFileName(file),
                    SectionGroupName  = rootGroupName,
                    SectionName       = sec,
                    OneNoteTargetPath = target
                });
            }

            foreach (var subfolder in subfolders)
            {
                var groupName = SanitizeName(Path.GetFileName(subfolder), 80);
                foreach (var file in Directory.GetFiles(subfolder, "*.enex", SearchOption.TopDirectoryOnly))
                {
                    var sec    = SanitizeName(Path.GetFileNameWithoutExtension(file), 80);
                    var rel    = Path.GetFileName(subfolder) + "\\" + Path.GetFileName(file);
                    var target = $"{_enNotebookName}\\{groupName}\\{sec}";

                    jobs.Add(new ImportJob
                    {
                        EnexPath          = file,
                        RelativePath      = rel,
                        SectionGroupName  = groupName,
                        SectionName       = sec,
                        OneNoteTargetPath = target
                    });
                }
            }

            return jobs;
        }

        // ===== Notebook name helpers =====

        private static bool NotebookExists(string name, string xmlHierarchy)
        {
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(xmlHierarchy);
                var nsManager = new XmlNamespaceManager(doc.NameTable);
                nsManager.AddNamespace("one", Xmlns);
                foreach (XmlNode node in doc.SelectNodes("//one:Notebook", nsManager))
                {
                    if (string.Equals(node.Attributes?["name"]?.Value, name,
                            StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { /* ignore parse errors — treat as not existing */ }
            return false;
        }

        private static string SanitizeName(string name, int maxLen)
        {
            name = name
                .Replace("?", "").Replace("*", "").Replace("/", "")
                .Replace("\\", "").Replace(":", "").Replace("<", "")
                .Replace(">", "").Replace("|", "").Replace("&", "")
                .Replace("#", "").Replace("\"", "'").Replace("%", "");
            name = name.Trim('.');
            if (name.Length > maxLen)
                name = name.Substring(0, maxLen);
            return name;
        }

        // ===== OneDrive helpers =====

        private static string FindOneDrivePath()
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var path = Path.Combine(programFiles, "Microsoft OneDrive", "OneDrive.exe");
            if (!File.Exists(path))
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                path = Path.Combine(localAppData, "Microsoft", "OneDrive", "OneDrive.exe");
            }
            return File.Exists(path) ? path : null;
        }

        private static void ShutdownOneDrive()
        {
            var onedrive = FindOneDrivePath();
            if (onedrive == null) return;
            var p = new System.Diagnostics.Process();
            p.StartInfo.FileName = onedrive;
            p.StartInfo.Arguments = "/shutdown";
            p.Start();
            p.WaitForExit();
        }

        private static void StartOneDrive()
        {
            var onedrive = FindOneDrivePath();
            if (onedrive == null) return;
            var p = new System.Diagnostics.Process();
            p.StartInfo.FileName = onedrive;
            p.StartInfo.Arguments = "/background";
            p.Start();
        }

        // ===== Main import orchestrator (background thread) =====

        private void ImportAllNotesToOnenote()
        {
            _syncStep = SyncStep.Start;
            ShutdownOneDrive();

            int totalJobs = _importJobs.Count;
            int jobIndex  = 0;

            _syncStep = SyncStep.ImportNotes;

            foreach (var job in _importJobs)
            {
                if (_cancelled) break;

                if (!File.Exists(job.EnexPath))
                {
                    UpdateManifestStatus(job, ManifestStatus.Failed, 99);
                    jobIndex++;
                    continue;
                }

                SetInfo($"Processing file {jobIndex + 1} of {totalJobs}: {job.RelativePath}",
                        "", jobIndex, totalJobs);

                UpdateManifestStatus(job, ManifestStatus.InProgress);

                bool success = false;
                try
                {
                    var sectionId = _useUnfiledSection
                        ? _newnbId
                        : GetOrCreateSection(job);

                    if (string.IsNullOrEmpty(sectionId))
                        throw new InvalidOperationException(
                            $"Could not get or create OneNote section for '{job.SectionName}'.");

                    var notes = ParseNotes(job.EnexPath);
                    ImportNotesToOneNoteSection(notes, job.EnexPath, sectionId);
                    success = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to import '{job.RelativePath}':\n{ex}");
                }

                if (success)
                {
                    UpdateManifestStatus(job, ManifestStatus.Complete);
                }
                else
                {
                    int attempts = GetCurrentAttemptCount(job) + 1;
                    UpdateManifestStatus(job, attempts >= 3
                        ? ManifestStatus.Failed
                        : ManifestStatus.Pending, attempts);
                }

                jobIndex++;

                // Give OneNote breathing room between files
                if (!_cancelled && jobIndex < totalJobs)
                    Thread.Sleep(2000);
            }

            _onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            StartOneDrive();

            _synchronizationContext.Send(delegate
            {
                btnENEXImport.Text = "Import ENEX Folder";
                infoText1.Text = _cancelled ? "Operation cancelled" : "Finished";
                infoText2.Text = "";
                progressIndicator.Minimum = 0;
                progressIndicator.Maximum = 100000;
                progressIndicator.Value = 0;
            }, null);

            if (_headlessMode)
            {
                _synchronizationContext.Send(delegate { Close(); }, null);
            }
        }

        // ===== Section creation =====

        private string GetOrCreateSection(ImportJob job)
        {
            var sectionId = "";
            try
            {
                if (!string.IsNullOrEmpty(job.SectionGroupName))
                {
                    // Create the section group (idempotent — harmless if already exists)
                    var groupPath = _evernoteNotebookPath + "\\" + job.SectionGroupName;
                    _onApp.OpenHierarchy(groupPath, "", out _, OneNote.CreateFileType.cftFolder);
                    Thread.Sleep(300);

                    // Create the section inside the group
                    var sectionPath = groupPath + "\\" + job.SectionName + ".one";
                    _onApp.OpenHierarchy(sectionPath, "", out sectionId, OneNote.CreateFileType.cftSection);
                }
                else
                {
                    var sectionPath = _evernoteNotebookPath + "\\" + job.SectionName + ".one";
                    _onApp.OpenHierarchy(sectionPath, "", out sectionId, OneNote.CreateFileType.cftSection);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetOrCreateSection failed for '{job.SectionName}': {ex.Message}");
            }
            return sectionId;
        }

        // ===== Note parsing (unchanged from original) =====

        private List<Note> ParseNotes(string exportFile)
        {
            _syncStep = SyncStep.ParseNotes;
            var noteList = new List<Note>();
            if (_cancelled)
                return noteList;

            var xtrInput = new XmlTextReader(exportFile);
            var xmltext = "";
            try
            {
                while (xtrInput.Read())
                {
                    while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                    {
                        if (_cancelled) break;

                        var xmlDocItem = new XmlDocument();
                        xmltext = SanitizeXml(xtrInput.ReadOuterXml());
                        xmlDocItem.LoadXml(xmltext);
                        var node = xmlDocItem.FirstChild;
                        node = node.FirstChild;

                        var note = new Note
                        {
                            Title = HttpUtility.HtmlDecode(node.InnerText)
                        };
                        noteList.Add(note);
                    }
                }
                xtrInput.Close();
            }
            catch (XmlException ex)
            {
                var notename = "";
                if (xmltext.Length > 0)
                {
                    var notematch = _rxNote.Match(xmltext);
                    if (notematch.Groups.Count == 2)
                        notename = notematch.Groups[1].ToString();
                }
                string tempfilepathDir = string.Empty;
                if (xmltext.Length > 0)
                    tempfilepathDir = ZipFailedNote(xmltext);

                MessageBox.Show(notename.Length > 0
                    ? $"Error parsing the note \"{notename}\" in notebook \"{_enNotebookName}\",\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues"
                    : $"Error parsing the notebook \"{_enNotebookName}\"\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues");
            }

            return noteList;
        }

        private static string ZipFailedNote(string xmltext)
        {
            var temppath = Path.GetTempPath() + "\\ev2on";
            var tempfilepathDir = temppath + "\\failedNotes";
            try
            {
                Directory.CreateDirectory(tempfilepathDir);
                var noteName    = "note-" + Guid.NewGuid().ToString();
                var tempfilepath = tempfilepathDir + "\\" + noteName + ".enex";
                xmltext = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><!DOCTYPE en-export SYSTEM \"http://xml.evernote.com/pub/evernote-export4.dtd\"><en-export>" + xmltext + "</en-export>";
                File.WriteAllText(tempfilepath, xmltext);
                var zipFilePath = tempfilepath + ".zip";
                using (ZipArchive zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
                {
                    zip.CreateEntryFromFile(tempfilepath, noteName + ".enex");
                }
                File.Delete(tempfilepath);
            }
            catch (Exception)
            {
                // ignored
            }
            return tempfilepathDir;
        }

        // ===== Per-section note importer =====

        private void ImportNotesToOneNoteSection(List<Note> notesEvernote, string exportFile, string sectionId)
        {
            _syncStep = SyncStep.CalculateWhatToDo;
            var uploadcount = notesEvernote.Count;

            var temppath = Path.GetTempPath() + "\\ev2on";
            Directory.CreateDirectory(temppath);

            _syncStep = SyncStep.ImportNotes;
            var counter = 0;
            var xmltext = "";

            try
            {
                var xtrInput = new XmlTextReader(exportFile);
                while (xtrInput.Read())
                {
                    while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                    {
                        if (_cancelled) break;

                        var xmlDocItem = new XmlDocument();
                        xmltext = SanitizeXml(xtrInput.ReadOuterXml());
                        xmlDocItem.LoadXml(xmltext);
                        var node = xmlDocItem.FirstChild;
                        node = node.FirstChild;

                        var note = new Note
                        {
                            Title = HttpUtility.HtmlDecode(node.InnerText).Replace("&nbsp;", " ")
                        };
                        if (note.Title.StartsWith("=?"))
                            note.Title = Rfc2047Decoder.Parse(note.Title);

                        var contentElements = xmlDocItem.GetElementsByTagName("content");
                        if (contentElements.Count > 0)
                            node = contentElements[0];
                        note.Content = node.InnerXml;
                        if (note.Content.StartsWith("=?"))
                            note.Content = Rfc2047Decoder.Parse(note.Content);

                        var atts = xmlDocItem.GetElementsByTagName("resource");
                        foreach (XmlNode xmln in atts)
                        {
                            var attachment = new Attachment
                            {
                                Base64Data = xmln.FirstChild.InnerText
                            };
                            var data = Convert.FromBase64String(xmln.FirstChild.InnerText);
                            var hash = new System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(data);
                            attachment.Hash = BitConverter.ToString(hash).Replace("-", string.Empty).ToLower();

                            var fns = xmlDocItem.GetElementsByTagName("file-name");
                            if (fns.Count > note.Attachments.Count)
                            {
                                attachment.FileName = HttpUtility.HtmlDecode(fns.Item(note.Attachments.Count).InnerText);
                                if (attachment.FileName.StartsWith("=?"))
                                    attachment.FileName = Rfc2047Decoder.Parse(attachment.FileName);
                                var invalid = new string(Path.GetInvalidFileNameChars());
                                foreach (var c in invalid)
                                    attachment.FileName = attachment.FileName.Replace(c.ToString(), "");
                                attachment.FileName = System.Security.SecurityElement.Escape(attachment.FileName);
                            }

                            var mimes = xmlDocItem.GetElementsByTagName("mime");
                            if (mimes.Count > note.Attachments.Count)
                                attachment.ContentType = HttpUtility.HtmlDecode(mimes.Item(note.Attachments.Count).InnerText);

                            note.Attachments.Add(attachment);
                        }

                        var tagslist = xmlDocItem.GetElementsByTagName("tag");
                        foreach (XmlNode n in tagslist)
                            note.Tags.Add(HttpUtility.HtmlDecode(n.InnerText));

                        var datelist = xmlDocItem.GetElementsByTagName("created");
                        foreach (XmlNode n in datelist)
                        {
                            if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture,
                                DateTimeStyles.AdjustToUniversal, out var dateCreated))
                                note.Date = dateCreated;
                        }
                        if (modifiedDateCheckbox.Checked)
                        {
                            var datelist2 = xmlDocItem.GetElementsByTagName("updated");
                            foreach (XmlNode n in datelist2)
                            {
                                if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture,
                                    DateTimeStyles.AdjustToUniversal, out var dateUpdated))
                                    note.Date = dateUpdated;
                            }
                        }

                        note.SourceUrl = "";
                        var sourceurl = xmlDocItem.GetElementsByTagName("source-url");
                        foreach (XmlNode n in sourceurl)
                        {
                            try
                            {
                                if (n.InnerText.StartsWith("file://")) continue;
                                if (n.InnerText.StartsWith("en-cache://")) continue;
                                note.SourceUrl = n.InnerText;
                            }
                            catch (FormatException) { }
                        }

                        if (_cmdDate > note.Date)
                            continue;

                        SetInfo(null, $"importing note ({counter + 1} of {uploadcount}) : \"{note.Title}\"",
                                counter++, uploadcount);

                        var htmlBody = note.Content;

                        // Process attachments into temp files
                        var tempfiles = new List<string>();
                        var xmlAttachments = "";
                        foreach (var attachment in note.Attachments)
                        {
                            var tempfilepath = temppath + "\\" + attachment.Hash;
                            var data = Convert.FromBase64String(attachment.Base64Data);
                            using (Stream fs = new FileStream(tempfilepath, FileMode.Create))
                                fs.Write(data, 0, data.Length);
                            tempfiles.Add(tempfilepath);

                            var rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*/>", RegexOptions.IgnoreCase);
                            var rxMatch = rx.Match(htmlBody);
                            if (attachment.ContentType != null && attachment.ContentType.Contains("image") && rxMatch.Success)
                            {
                                htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                            }
                            else
                            {
                                rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*></en-media>", RegexOptions.IgnoreCase);
                                if (attachment.ContentType != null && attachment.ContentType.Contains("image") && rx.Match(htmlBody).Success)
                                {
                                    htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(attachment.FileName))
                                    {
                                        if (!attachment.ContentType.Contains("image") || attachment.FileName != "proxy.php")
                                            xmlAttachments += $"<one:InsertedFile pathSource=\"{tempfilepath}\" preferredName=\"{attachment.FileName}\" />";
                                    }
                                    else
                                    {
                                        xmlAttachments += $"<one:InsertedFile pathSource=\"{tempfilepath}\" preferredName=\"{attachment.Hash}\" />";
                                    }
                                }
                            }
                        }
                        note.Attachments.Clear();

                        // HTML cleanup
                        htmlBody = _rxFontFamily.Replace(htmlBody, string.Empty);
                        htmlBody = _rxBrOnly.Replace(htmlBody, "&nbsp;$&");
                        htmlBody = _rxStyle.Replace(htmlBody, delegate(Match m)
                        {
                            if (m.Value.Contains("--en-codeblock:true;"))
                                return m.Result("<br><br>${text}") + "style=\"background-color:#B0B0B0; font-family: Consolas, Courier New, monospace; font-size: 15px;\"";
                            return m.Result("$&");
                        });
                        htmlBody = htmlBody.Replace("<pre>", "<br><br><pre style=\"font-family: Consolas, Courier New, monospace; font-size: 15px; background-color:#B0B0B0;\">");
                        htmlBody = htmlBody.Replace("</pre>", "</pre><br><br>");
                        htmlBody = _rxComment.Replace(htmlBody, string.Empty);
                        htmlBody = _rxEmptyCdata.Replace(htmlBody, string.Empty);
                        htmlBody = _rxEmptyCdata2.Replace(htmlBody, string.Empty);
                        htmlBody = _rxEmptyCdata3.Replace(htmlBody, string.Empty);
                        htmlBody = _rxCdata.Replace(htmlBody, string.Empty);
                        htmlBody = _rxCdata2.Replace(htmlBody, string.Empty);
                        htmlBody = _rxDtd.Replace(htmlBody, string.Empty);
                        htmlBody = _rxBodyStart.Replace(htmlBody, "<body>");
                        htmlBody = _rxBodyEnd.Replace(htmlBody, "</body>");
                        htmlBody = _rxBodyEmpty.Replace(htmlBody, "<body></body>");
                        htmlBody = htmlBody.Trim();
                        htmlBody = @"<!DOCTYPE html><head></head>" + htmlBody;

                        // Escape < and > inside <pre> blocks
                        var rxPre = new Regex(@"<pre\b[^>]*?>(.+)</pre>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        foreach (Match match in rxPre.Matches(htmlBody))
                        {
                            var fullPreSection = match.ToString();
                            var innerSection = match.Groups[1].ToString();
                            var innerCorrected = innerSection.Replace("<", "&lt;").Replace(">", "&gt;");
                            fullPreSection = fullPreSection.Replace(innerSection, innerCorrected);
                            htmlBody = rxPre.Replace(htmlBody, fullPreSection);
                        }

                        var emailBody = htmlBody;
                        emailBody = _rxDate.Replace(emailBody, "Date: " + note.Date.ToString("ddd, dd MMM yyyy HH:mm:ss K"));
                        emailBody = emailBody.Replace("&apos;", "'");
                        emailBody = emailBody.Replace("‘", "'");
                        emailBody = _rxCdataInner.Replace(emailBody, "&lt;![CDATA[${text}]]&gt;");
                        emailBody = emailBody.Replace("’", "'");

                        // Inject tags as visible text at the top of the note
                        if (note.Tags.Count > 0)
                        {
                            var tagLine = $"<p><strong>Tags:</strong> {string.Join(", ", note.Tags)}</p>";
                            int bodyIdx = emailBody.IndexOf("<body>", StringComparison.OrdinalIgnoreCase);
                            if (bodyIdx >= 0)
                                emailBody = emailBody.Insert(bodyIdx + 6, tagLine);
                            else
                                emailBody = tagLine + emailBody;
                        }

                        if (!TryImportNoteWithRetry(sectionId, note, emailBody, xmlAttachments))
                        {
                            var tempDir = ZipFailedNote(xmltext);
                            MessageBox.Show(
                                $"Note: {note.Title}\nFailed to import after 3 attempts.\n\nA copy of the note is saved in {tempDir}.\nIf you want to help fix the problem, please consider creating an issue: https://github.com/stefankueng/EvImSync/issues");
                        }

                        foreach (var p in tempfiles)
                            File.Delete(p);
                    }
                }
                xtrInput.Close();
            }
            catch (XmlException ex)
            {
                var notename = "";
                if (xmltext.Length > 0)
                {
                    var notematch = _rxNote.Match(xmltext);
                    if (notematch.Groups.Count == 2)
                        notename = notematch.Groups[1].ToString();
                }
                string tempfilepathDir = string.Empty;
                if (xmltext.Length > 0)
                    tempfilepathDir = ZipFailedNote(xmltext);

                MessageBox.Show(notename.Length > 0
                    ? $"Error parsing the note \"{notename}\" in notebook \"{_enNotebookName}\",\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues"
                    : $"Error parsing the notebook \"{_enNotebookName}\"\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception importing notes:\n{ex}");
            }
        }

        // ===== Per-note COM retry wrapper =====

        private bool TryImportNoteWithRetry(string sectionId, Note note, string emailBody, string xmlAttachments)
        {
            int[] delays = { 1000, 2000, 4000 };
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    _onApp.CreateNewPage(sectionId, out var pageId, OneNote.NewPageStyle.npsBlankPageWithTitle);
                    Thread.Sleep(500);

                    var outlineId      = new Random().Next();
                    var xmlSource      = string.Format(XmlSourceUrl, note.SourceUrl);
                    var outlineContent = string.Format(
                        _xmlNewOutlineContent,
                        emailBody,
                        outlineId,
                        System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"),
                        note.SourceUrl.Length > 0 ? xmlSource : "");
                    var xml = string.Format(
                        XmlNewOutline,
                        outlineContent,
                        pageId,
                        Xmlns,
                        System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"),
                        xmlAttachments,
                        note.Date.ToString("yyyy'-'MM'-'ddTHH':'mm':'ss'Z'"));

                    _onApp.UpdatePageContent(xml, DateTime.MinValue, OneNote.XMLSchema.xs2013, true);
                    Thread.Sleep(500);
                    _onApp.SyncHierarchy(pageId);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Note '{note.Title}' attempt {attempt + 1} failed: {ex.Message}");
                    if (attempt < 2)
                        Thread.Sleep(delays[attempt]);
                }
            }
            return false;
        }

        // ===== Manifest helpers =====

        private void UpdateManifestStatus(ImportJob job, ManifestStatus status, int attempts = -1)
        {
            if (string.IsNullOrEmpty(_currentManifestPath)) return;
            var entries = ManifestManager.Load(_currentManifestPath);
            var entry   = entries.FirstOrDefault(e => e.RelativePath == job.RelativePath);
            if (entry == null) return;
            entry.Status = status;
            if (attempts >= 0) entry.AttemptCount = attempts;
            ManifestManager.UpdateEntry(_currentManifestPath, entry);
        }

        private int GetCurrentAttemptCount(ImportJob job)
        {
            if (string.IsNullOrEmpty(_currentManifestPath)) return 0;
            var entries = ManifestManager.Load(_currentManifestPath);
            return entries.FirstOrDefault(e => e.RelativePath == job.RelativePath)?.AttemptCount ?? 0;
        }

        // ===== Target path helpers =====

        private static string ExtractSectionGroupFromTarget(string target)
        {
            var parts = target.Split('\\');
            return parts.Length == 3 ? parts[1] : null;
        }

        private static string ExtractSectionNameFromTarget(string target)
        {
            var parts = target.Split('\\');
            return parts.Length >= 2 ? parts[parts.Length - 1] : target;
        }

        // ===== Hierarchy helper (unchanged from original) =====

        private void AppendHierarchy(XmlNode xml, StringBuilder str, int level)
        {
            if (xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section" || xml.Name == "one:Page")
            {
                if (xml.Attributes != null)
                {
                    var id = xml.Attributes != null && xml.LocalName == "Section" && xml.Attributes["path"].Value == _evernoteNotebookPath
                        ? "UnfiledNotes"
                        : xml.Attributes["ID"].Value;
                    var name = HttpUtility.HtmlEncode(xml.Attributes["name"].Value);
                    if (str.Length > 0)
                        str.Append("\n");
                    str.Append($"{level.ToString()} {xml.LocalName} {id} {name}");
                }
            }
            if (xml.Name == "one:Notebooks" || xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section")
            {
                foreach (XmlNode child in xml.ChildNodes)
                {
                    int nextLevel = xml.Name == "one:Notebooks" ? level : level + 1;
                    AppendHierarchy(child, str, nextLevel);
                }
            }
        }

        // ===== XML sanitization helpers (unchanged from original) =====

        private string SanitizeXml(string text)
        {
            var rxtitle = new Regex("<note><title>(.+)</title>", RegexOptions.IgnoreCase);
            var match = rxtitle.Match(text);
            if (match.Groups.Count == 2)
            {
                var title = match.Groups[1].ToString();
                title = title.Replace("&", "&amp;");
                title = title.Replace("\"", "&quot;");
                title = title.Replace("'", "&apos;");
                title = title.Replace("‘", "&apos;");
                title = title.Replace("<", "&lt;");
                title = title.Replace(">", "&gt;");
                title = title.Replace("@", "&#64;");
                text = rxtitle.Replace(text, "<note><title>" + title + "</title>");
            }

            var rxauthor = new Regex("<author>(.+)</author>", RegexOptions.IgnoreCase);
            var authormatch = rxauthor.Match(text);
            if (match.Groups.Count == 2)
            {
                var author = authormatch.Groups[1].ToString();
                author = author.Replace("&", "&amp;");
                author = author.Replace("\"", "&quot;");
                author = author.Replace("'", "&apos;");
                author = author.Replace("‘", "&apos;");
                author = author.Replace("<", "&lt;");
                author = author.Replace(">", "&gt;");
                author = author.Replace("@", "&#64;");
                text = rxauthor.Replace(text, "<author>" + author + "</author>");
            }

            var rxfilename = new Regex("<file-name>(.+)</file-name>", RegexOptions.IgnoreCase);
            if (match.Groups.Count == 2)
            {
                MatchEvaluator myEvaluator = FilenameMatchEvaluator;
                text = rxfilename.Replace(text, myEvaluator);
            }

            var rxSrcUrl = new Regex("<source-url>(.+)</source-url>", RegexOptions.IgnoreCase);
            if (match.Groups.Count == 2)
            {
                MatchEvaluator myEvaluator = FilenameMatchEvaluator;
                text = rxSrcUrl.Replace(text, myEvaluator);
            }

            return text;
        }

        private string FilenameMatchEvaluator(Match m)
        {
            var filename = m.Groups[1].ToString();
            filename = filename.Replace("&nbsp;", " ");
            var invalid = new string(Path.GetInvalidFileNameChars());
            foreach (var c in invalid)
                filename = filename.Replace(c.ToString(), "");
            filename = System.Security.SecurityElement.Escape(filename);
            return "<file-name>" + filename + "</file-name>";
        }

        private void lblNotebookName_Click(object sender, EventArgs e)
        {

        }
    }
}
