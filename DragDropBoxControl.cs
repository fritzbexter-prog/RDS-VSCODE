// ============================================================
// Dynamics NAV 2018 Control Add-in: Drag & Drop Box
// ============================================================
// Deploy: Copy DLL to RoleTailored Client\Add-ins\ folder
//
// C/AL Integration:
//   - Event FileDropped: wird ausgeloest wenn Dateien gedroppt werden
//   - GetDroppedFilePath(): Ersten Dateipfad abrufen (nach Event)
//   - GetAllDroppedFilePaths(): Alle Pfade semicolon-getrennt (nach Event)
//   - SetDisplayText(text): Anzeigetext setzen
//   - SetAllowedExtensions(".pdf;.xlsx;.csv"): Filter setzen
//   - Outlook Emails: Emails koennen direkt aus Outlook gezogen werden
//     und werden als .msg Dateien im Temp-Verzeichnis gespeichert
//   - Outlook Anhaenge: Anhaenge koennen direkt aus Outlook gezogen werden
//     und behalten ihre originale Dateiendung
// ============================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Microsoft.Dynamics.Framework.UI.Extensibility;
using Microsoft.Dynamics.Framework.UI.Extensibility.WinForms;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace DragDropAddIn
{
    // -------------------------------------------------------
    // 1) Interface: definiert die Events und Methoden,
    //    die NAV (C/AL) sehen kann
    // -------------------------------------------------------
    [ControlAddInExport("DragDropBox")]
    public interface IDragDropBox
    {
        [ApplicationVisible]
        event ApplicationEventHandler FileDropped;

        [ApplicationVisible]
        void SetDisplayText(string text);

        [ApplicationVisible]
        void SetAllowedExtensions(string extensions);

        [ApplicationVisible]
        string GetDroppedFilePath();

        [ApplicationVisible]
        string GetAllDroppedFilePaths();
    }

    // -------------------------------------------------------
    // 2) WinForms Control: die eigentliche Drag & Drop Box
    // -------------------------------------------------------
    [ControlAddInExport("DragDropBox")]
    public class DragDropBoxControl : StringControlAddInBase, IDragDropBox
    {
        // Felder
        private Panel _dropPanel;
        private Label _dropLabel;
        private string _displayText = "Dateien hier hineinziehen...";
        private string _allowedExtensions = "";
        private string _droppedFilePaths = "";
        private bool _isDragOver = false;

        // Farben
        private readonly Color _normalBackColor = Color.FromArgb(240, 245, 250);
        private readonly Color _normalBorderColor = Color.FromArgb(180, 190, 200);
        private readonly Color _hoverBackColor = Color.FromArgb(220, 235, 250);
        private readonly Color _hoverBorderColor = Color.FromArgb(0, 120, 215);
        private readonly Color _successBackColor = Color.FromArgb(220, 245, 220);
        private readonly Color _errorBackColor = Color.FromArgb(255, 230, 230);

        // Events
        public event ApplicationEventHandler FileDropped;

        // -------------------------------------------------------
        // Control erstellen
        // -------------------------------------------------------
        protected override Control CreateControl()
        {
            _dropPanel = new Panel();
            _dropPanel.Width = 400;
            _dropPanel.Height = 150;
            _dropPanel.BackColor = _normalBackColor;
            _dropPanel.BorderStyle = BorderStyle.None;
            _dropPanel.Padding = new Padding(10);
            _dropPanel.AllowDrop = true;
            _dropPanel.Cursor = Cursors.Hand;

            _dropLabel = new Label();
            _dropLabel.Text = _displayText;
            _dropLabel.Font = new Font("Segoe UI", 11f, FontStyle.Regular);
            _dropLabel.ForeColor = Color.FromArgb(100, 110, 120);
            _dropLabel.TextAlign = ContentAlignment.MiddleCenter;
            _dropLabel.Dock = DockStyle.Fill;
            _dropLabel.AutoSize = false;
            _dropLabel.AllowDrop = true;

            _dropPanel.DragEnter += new DragEventHandler(OnDragEnter);
            _dropPanel.DragLeave += new EventHandler(OnDragLeave);
            _dropPanel.DragDrop += new DragEventHandler(OnDragDrop);
            _dropPanel.Paint += new PaintEventHandler(OnPanelPaint);

            _dropLabel.DragEnter += new DragEventHandler(OnDragEnter);
            _dropLabel.DragLeave += new EventHandler(OnDragLeave);
            _dropLabel.DragDrop += new DragEventHandler(OnDragDrop);

            _dropPanel.Click += new EventHandler(OnClick);
            _dropLabel.Click += new EventHandler(OnClick);

            _dropPanel.Controls.Add(_dropLabel);

            return _dropPanel;
        }

        // -------------------------------------------------------
        // Drag Enter: visuelles Feedback
        // -------------------------------------------------------
        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)
                || e.Data.GetDataPresent("FileGroupDescriptorW")
                || e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effect = DragDropEffects.Copy;
                _isDragOver = true;
                _dropPanel.BackColor = _hoverBackColor;
                _dropLabel.Text = "Loslassen zum Ablegen...";
                _dropLabel.ForeColor = Color.FromArgb(0, 90, 170);
                _dropPanel.Invalidate();
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        // -------------------------------------------------------
        // Drag Leave: Zustand zuruecksetzen
        // -------------------------------------------------------
        private void OnDragLeave(object sender, EventArgs e)
        {
            _isDragOver = false;
            _dropPanel.BackColor = _normalBackColor;
            _dropLabel.Text = _displayText;
            _dropLabel.ForeColor = Color.FromArgb(100, 110, 120);
            _dropPanel.Invalidate();
        }

        // -------------------------------------------------------
        // Drop: Dateien verarbeiten
        // -------------------------------------------------------
        private void OnDragDrop(object sender, DragEventArgs e)
        {
            _isDragOver = false;

            // Fall 1: Normale Dateien (FileDrop)
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    ProcessFiles(files);
                }
                return;
            }

            // Fall 2: Outlook-Emails/Anhaenge (FileGroupDescriptorW + FileContents)
            if (e.Data.GetDataPresent("FileGroupDescriptorW")
                || e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                string[] files = ExtractOutlookFiles(e.Data);
                if (files != null && files.Length > 0)
                {
                    ProcessFiles(files);
                }
                else
                {
                    _dropPanel.BackColor = _errorBackColor;
                    _dropLabel.Text = "Fehler beim Verarbeiten der Outlook-Email!";
                    _dropLabel.ForeColor = Color.Red;
                    _dropPanel.Invalidate();
                }
                return;
            }
        }

        // -------------------------------------------------------
        // Outlook-Emails/Anhaenge extrahieren (OLE FileGroupDescriptorW)
        // -------------------------------------------------------
        private string[] ExtractOutlookFiles(System.Windows.Forms.IDataObject data)
        {
            try
            {
                object rawDescriptor = data.GetData("FileGroupDescriptorW");
                MemoryStream descriptorStream = rawDescriptor as MemoryStream;
                if (descriptorStream == null)
                {
                    rawDescriptor = data.GetData("FileGroupDescriptor");
                    descriptorStream = rawDescriptor as MemoryStream;
                    if (descriptorStream == null)
                        return null;
                }

                byte[] descriptorBytes = descriptorStream.ToArray();

                // Erstes DWORD = Anzahl der Dateien
                int fileCount = BitConverter.ToInt32(descriptorBytes, 0);
                if (fileCount == 0)
                    return null;

                // Temp-Verzeichnis fuer Outlook-Emails
                string tempDir = Path.Combine(Path.GetTempPath(), "DragDropBox_Outlook");
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                System.Collections.Generic.List<string> savedFiles =
                    new System.Collections.Generic.List<string>();

                // FILEDESCRIPTORW ist 592 Bytes gross (4 + n * 592)
                int descriptorSize = 592;

                for (int i = 0; i < fileCount; i++)
                {
                    // Dateiname ab Offset 72 im FILEDESCRIPTORW (Unicode, 260 chars = 520 Bytes)
                    int baseOffset = 4 + (i * descriptorSize);
                    int nameOffset = baseOffset + 72;

                    if (nameOffset + 520 > descriptorBytes.Length)
                        continue;

                    string fileName = System.Text.Encoding.Unicode.GetString(
                        descriptorBytes, nameOffset, 520);
                    fileName = fileName.TrimEnd('\0');

                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    // Nur .msg anfuegen wenn gar keine Endung vorhanden
                    if (string.IsNullOrEmpty(Path.GetExtension(fileName)))
                    {
                        fileName = fileName + ".msg";
                    }

                    // Ungueltige Zeichen im Dateinamen ersetzen
                    char[] invalidChars = Path.GetInvalidFileNameChars();
                    for (int c = 0; c < invalidChars.Length; c++)
                    {
                        fileName = fileName.Replace(invalidChars[c], '_');
                    }

                    // FileContents fuer diesen Index lesen
                    MemoryStream contentStream = GetFileContentsStream(data, i);
                    if (contentStream == null)
                        continue;

                    // Datei speichern
                    string filePath = Path.Combine(tempDir, fileName);

                    // Bei Namenskollision Nummer anfuegen
                    if (File.Exists(filePath))
                    {
                        string nameOnly = Path.GetFileNameWithoutExtension(fileName);
                        string ext = Path.GetExtension(fileName);
                        int counter = 1;
                        do
                        {
                            filePath = Path.Combine(tempDir,
                                nameOnly + "_" + counter.ToString() + ext);
                            counter++;
                        } while (File.Exists(filePath));
                    }

                    byte[] contentBytes = contentStream.ToArray();
                    File.WriteAllBytes(filePath, contentBytes);
                    savedFiles.Add(filePath);
                }

                return savedFiles.ToArray();
            }
            catch
            {
                return null;
            }
        }

        // -------------------------------------------------------
        // FileContents-Stream fuer einen bestimmten Index holen
        // Umgeht den WinForms DataObject-Wrapper und greift direkt
        // auf das native COM IDataObject zu. Probiert verschiedene
        // TYMED-Kombinationen (IStorage, IStream, HGlobal).
        // -------------------------------------------------------
        private MemoryStream GetFileContentsStream(System.Windows.Forms.IDataObject data, int index)
        {
            System.Runtime.InteropServices.ComTypes.IDataObject comData = GetNativeComDataObject(data);

            short cfFileContents = (short)RegisterClipboardFormat("FileContents");

            // TYMED_ISTORAGE = 8 (nicht in .NET TYMED enum definiert)
            TYMED TYMED_ISTORAGE = (TYMED)8;

            // Verschiedene Kombinationen versuchen
            int[] tryLindex = new int[] { index, -1 };
            TYMED[] tryTymed = new TYMED[] {
                TYMED_ISTORAGE,
                TYMED.TYMED_ISTREAM | TYMED_ISTORAGE,
                TYMED.TYMED_ISTREAM,
                TYMED.TYMED_HGLOBAL,
                TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL
            };

            for (int li = 0; li < tryLindex.Length; li++)
            {
                for (int ti = 0; ti < tryTymed.Length; ti++)
                {
                    try
                    {
                        FORMATETC formatEtc = new FORMATETC();
                        formatEtc.cfFormat = cfFileContents;
                        formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
                        formatEtc.lindex = tryLindex[li];
                        formatEtc.ptd = IntPtr.Zero;
                        formatEtc.tymed = tryTymed[ti];

                        STGMEDIUM stgMedium = new STGMEDIUM();
                        comData.GetData(ref formatEtc, out stgMedium);

                        MemoryStream result = null;

                        try
                        {
                            if (stgMedium.tymed == TYMED.TYMED_ISTREAM)
                            {
                                IStream iStream = (IStream)Marshal.GetObjectForIUnknown(stgMedium.unionmember);
                                result = ReadIStream(iStream);
                                Marshal.ReleaseComObject(iStream);
                            }
                            else if (stgMedium.tymed == TYMED.TYMED_HGLOBAL)
                            {
                                result = ReadHGlobal(stgMedium.unionmember);
                            }
                            else if (stgMedium.tymed == TYMED_ISTORAGE)
                            {
                                result = ReadIStorage(stgMedium.unionmember);
                            }
                        }
                        finally
                        {
                            ReleaseStgMedium(ref stgMedium);
                        }

                        if (result != null && result.Length > 0)
                            return result;
                    }
                    catch
                    {
                        // Naechste Kombination versuchen
                    }
                }
            }

            return null;
        }

        // -------------------------------------------------------
        // Natives COM IDataObject aus WinForms-Wrapper extrahieren
        // WinForms DataObject leitet lindex nicht korrekt weiter
        // -------------------------------------------------------
        private System.Runtime.InteropServices.ComTypes.IDataObject GetNativeComDataObject(
            System.Windows.Forms.IDataObject data)
        {
            System.Reflection.FieldInfo innerDataInfo = data.GetType().GetField("innerData",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            if (innerDataInfo != null)
            {
                object innerData = innerDataInfo.GetValue(data);
                if (innerData != null)
                {
                    System.Reflection.FieldInfo oleInnerInfo = innerData.GetType().GetField("innerData",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                    if (oleInnerInfo != null)
                    {
                        object oleInner = oleInnerInfo.GetValue(innerData);
                        if (oleInner is System.Runtime.InteropServices.ComTypes.IDataObject)
                            return (System.Runtime.InteropServices.ComTypes.IDataObject)oleInner;
                    }

                    if (innerData is System.Runtime.InteropServices.ComTypes.IDataObject)
                        return (System.Runtime.InteropServices.ComTypes.IDataObject)innerData;
                }
            }

            return (System.Runtime.InteropServices.ComTypes.IDataObject)data;
        }

        // -------------------------------------------------------
        // IStream in MemoryStream lesen
        // -------------------------------------------------------
        private MemoryStream ReadIStream(IStream iStream)
        {
            MemoryStream result = new MemoryStream();
            byte[] buffer = new byte[8192];
            IntPtr bytesReadPtr = Marshal.AllocCoTaskMem(sizeof(long));

            try
            {
                while (true)
                {
                    Marshal.WriteInt64(bytesReadPtr, 0);
                    iStream.Read(buffer, buffer.Length, bytesReadPtr);
                    int bytesRead = (int)Marshal.ReadInt64(bytesReadPtr);
                    if (bytesRead <= 0)
                        break;
                    result.Write(buffer, 0, bytesRead);
                }
            }
            finally
            {
                Marshal.FreeCoTaskMem(bytesReadPtr);
            }

            return result;
        }

        // -------------------------------------------------------
        // HGLOBAL in MemoryStream lesen
        // -------------------------------------------------------
        private MemoryStream ReadHGlobal(IntPtr hGlobal)
        {
            IntPtr ptr = GlobalLock(hGlobal);
            try
            {
                int size = (int)GlobalSize(hGlobal);
                byte[] data = new byte[size];
                Marshal.Copy(ptr, data, 0, size);
                return new MemoryStream(data);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }
        }

        // -------------------------------------------------------
        // IStorage in MemoryStream lesen (fuer Outlook .msg)
        // Speichert IStorage ueber StgCreateDocfile in Temp-Datei,
        // liest diese zurueck als MemoryStream
        // -------------------------------------------------------
        private MemoryStream ReadIStorage(IntPtr pStoragePtr)
        {
            IStorageNative srcStorage = null;
            IStorageNative dstStorage = null;
            string tempFile = null;

            try
            {
                srcStorage = (IStorageNative)Marshal.GetObjectForIUnknown(pStoragePtr);

                tempFile = Path.Combine(Path.GetTempPath(),
                    "DragDropBox_" + Guid.NewGuid().ToString("N") + ".msg");

                int hr = StgCreateDocfile(tempFile,
                    STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
                    0, out dstStorage);

                if (hr != 0 || dstStorage == null)
                    return null;

                srcStorage.CopyTo(0, IntPtr.Zero, IntPtr.Zero, dstStorage);
                dstStorage.Commit(0);
                Marshal.ReleaseComObject(dstStorage);
                dstStorage = null;

                byte[] fileBytes = File.ReadAllBytes(tempFile);
                return new MemoryStream(fileBytes);
            }
            catch
            {
                return null;
            }
            finally
            {
                if (dstStorage != null)
                    Marshal.ReleaseComObject(dstStorage);
                if (srcStorage != null)
                    Marshal.ReleaseComObject(srcStorage);
                try { if (tempFile != null && File.Exists(tempFile)) File.Delete(tempFile); }
                catch { }
            }
        }

        // STGM-Konstanten
        private const int STGM_CREATE = 0x00001000;
        private const int STGM_READWRITE = 0x00000002;
        private const int STGM_SHARE_EXCLUSIVE = 0x00000010;

        // -------------------------------------------------------
        // Minimale IStorage COM-Interface Definition
        // -------------------------------------------------------
        [ComImport]
        [Guid("0000000b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IStorageNative
        {
            [PreserveSig]
            int CreateStream(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                uint grfMode, uint reserved1, uint reserved2,
                out IntPtr ppstm);

            [PreserveSig]
            int OpenStream(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                IntPtr reserved1, uint grfMode, uint reserved2,
                out IntPtr ppstm);

            [PreserveSig]
            int CreateStorage(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                uint grfMode, uint reserved1, uint reserved2,
                out IntPtr ppstg);

            [PreserveSig]
            int OpenStorage(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                IntPtr pstgPriority, uint grfMode,
                IntPtr snbExclude, uint reserved,
                out IntPtr ppstg);

            void CopyTo(
                uint ciidExclude,
                IntPtr rgiidExclude,
                IntPtr snbExclude,
                IStorageNative pstgDest);

            void MoveElementTo(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                IntPtr pstgDest,
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName,
                uint grfFlags);

            void Commit(uint grfCommitFlags);
        }

        // -------------------------------------------------------
        // P/Invoke Deklarationen
        // -------------------------------------------------------
        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern UIntPtr GlobalSize(IntPtr hMem);

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM pmedium);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern uint RegisterClipboardFormat(string lpszFormat);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int StgCreateDocfile(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            int grfMode,
            int reserved,
            out IStorageNative ppstgOpen);

        // -------------------------------------------------------
        // Click: Datei-Dialog als Fallback
        // -------------------------------------------------------
        private void OnClick(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Datei auswaehlen";
                ofd.Multiselect = true;

                if (!string.IsNullOrEmpty(_allowedExtensions))
                {
                    string filter = BuildFileFilter(_allowedExtensions);
                    ofd.Filter = filter;
                }
                else
                {
                    ofd.Filter = "Alle Dateien (*.*)|*.*";
                }

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    ProcessFiles(ofd.FileNames);
                }
            }
        }

        // -------------------------------------------------------
        // Dateien verarbeiten (gemeinsam fuer Drop und Click)
        // -------------------------------------------------------
        private void ProcessFiles(string[] files)
        {
            if (!string.IsNullOrEmpty(_allowedExtensions))
            {
                string[] allowed = _allowedExtensions
                    .ToLower()
                    .Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                System.Collections.Generic.List<string> filtered =
                    new System.Collections.Generic.List<string>();

                for (int i = 0; i < files.Length; i++)
                {
                    string ext = Path.GetExtension(files[i]).ToLower();
                    for (int j = 0; j < allowed.Length; j++)
                    {
                        if (ext == allowed[j].Trim())
                        {
                            filtered.Add(files[i]);
                            break;
                        }
                    }
                }

                if (filtered.Count == 0)
                {
                    _dropPanel.BackColor = _errorBackColor;
                    _dropLabel.Text = "Dateityp nicht erlaubt! (" + _allowedExtensions + ")";
                    _dropLabel.ForeColor = Color.Red;
                    _dropPanel.Invalidate();
                    return;
                }

                files = filtered.ToArray();
            }

            _droppedFilePaths = string.Join(";", files);

            _dropPanel.BackColor = _successBackColor;
            _dropLabel.ForeColor = Color.FromArgb(0, 130, 60);

            if (files.Length == 1)
            {
                _dropLabel.Text = "OK: " + Path.GetFileName(files[0]);
            }
            else
            {
                _dropLabel.Text = "OK: " + files.Length.ToString() + " Dateien abgelegt";
            }

            _dropPanel.Invalidate();

            // Event an NAV senden
            if (FileDropped != null)
            {
                FileDropped();
            }
        }

        // -------------------------------------------------------
        // Gestrichelte Border zeichnen
        // -------------------------------------------------------
        private void OnPanelPaint(object sender, PaintEventArgs e)
        {
            Color borderColor = _isDragOver ? _hoverBorderColor : _normalBorderColor;
            float borderWidth = _isDragOver ? 2.5f : 1.5f;

            using (Pen pen = new Pen(borderColor, borderWidth))
            {
                pen.DashStyle = DashStyle.Dash;
                pen.DashPattern = new float[] { 6, 4 };

                Rectangle rect = new Rectangle(
                    1, 1,
                    _dropPanel.Width - 3,
                    _dropPanel.Height - 3
                );

                int radius = 8;
                using (GraphicsPath path = RoundedRect(rect, radius))
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    e.Graphics.DrawPath(pen, path);
                }
            }
        }

        private GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private string BuildFileFilter(string extensions)
        {
            string[] exts = extensions.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string allExts = "";
            string parts = "";

            for (int i = 0; i < exts.Length; i++)
            {
                string ext = exts[i].Trim();
                if (!ext.StartsWith("."))
                    ext = "." + ext;

                string wildcard = "*" + ext;

                if (i > 0)
                {
                    allExts = allExts + ";";
                    parts = parts + "|";
                }
                allExts = allExts + wildcard;
                parts = parts + ext.ToUpper().Substring(1) + " (" + wildcard + ")|" + wildcard;
            }

            return "Erlaubte Dateien (" + allExts + ")|" + allExts + "|" + parts + "|Alle Dateien (*.*)|*.*";
        }

        // -------------------------------------------------------
        // Interface-Methoden fuer C/AL
        // -------------------------------------------------------
        public void SetDisplayText(string text)
        {
            _displayText = text;
            if (_dropLabel != null)
            {
                _dropLabel.Text = text;
            }
        }

        public void SetAllowedExtensions(string extensions)
        {
            _allowedExtensions = extensions ?? "";
        }

        public string GetDroppedFilePath()
        {
            if (string.IsNullOrEmpty(_droppedFilePaths))
                return "";

            string[] parts = _droppedFilePaths.Split(';');
            return parts[0];
        }

        public string GetAllDroppedFilePaths()
        {
            return _droppedFilePaths ?? "";
        }
    }
}
