// Microsoft.Office.Interop.Word

using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;

using Task = System.Threading.Tasks.Task;
using Microsoft.VisualBasic.Logging;

namespace Word_Doc_Beautifier
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // btnRefresh.PerformClick();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            headerHistoryFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HeaderHistory.txt");
            LoadListBoxFromFile(lbHeaderHistory, headerHistoryFile);

            btnRefresh.PerformClick();
        }

        private ApplicationEvents4_Event ae;
        private bool IsWordQuitting = false;

        private void OnWordQuit()
        {
            UpdateStatus("Word is Quitting");
            wordApp = null;
            IsWordQuitting = true;
            StartDetectWordAppThread();
        }

        private void StartDetectWordAppThread()
        {

            if (wordApp is not null)
            {
                wordApp.DocumentChange += OnDocumentChange;
                ae = (ApplicationEvents4_Event)wordApp;
                ae.Quit += OnWordQuit;
            }
            else
            {
                // start a new thread to detect any new word app
                UpdateStatus("Starting DetctWordAppThread");
                Task task = Task.Run(() => DetectWordApp());
            }
        }

        private void DetectWordApp()
        {

            if (IsWordQuitting)
            {   // If word is quitting, let's wait until completely quitting
                wordApp = Marshal.GetActiveObject("Word.Application") as Application;
                while (wordApp is not null)
                {
                    Thread.Sleep(700);
                    wordApp = Marshal.GetActiveObject("Word.Application") as Application;
                }
                IsWordQuitting = false;
            }

            while (wordApp is null)
            {
                Thread.Sleep(1000);
                //UpdateStatus("No Word App detected");
                wordApp = Marshal.GetActiveObject("Word.Application") as Application;
            }

            UpdateStatus("Fond Word App, Exiting detecting thread");
            UpdateStatus("Refresh Main window");
            TriggerRefresh();
        }

        private void UninstallWordAppEventHandler()
        {
            if (wordApp is not null)
            {
                UpdateStatus("Uninstall Word Event Handler");
                wordApp.DocumentChange -= OnDocumentChange;
                ae = (ApplicationEvents4_Event)wordApp;
                ae.Quit -= OnWordQuit;
                ae.DocumentBeforeClose -= OnDocumentBeforeClose;
            }
        }
        private void InstallWordAppEventHandler()
        {
            if (wordApp is not null)
            {
                UpdateStatus("Install Word Event Handler");
                wordApp.DocumentChange += OnDocumentChange;
                ae = (ApplicationEvents4_Event)wordApp;
                ae.Quit += OnWordQuit;
                ae.DocumentBeforeClose += OnDocumentBeforeClose;
            }
        }

        private void OnDocumentBeforeClose(Document doc, ref bool Cancel)
        {
            //TriggerRefresh();
        }

        private void TriggerRefresh()
        {
            if (btnRefresh.InvokeRequired)
            {
                btnRefresh.Invoke(new Action(TriggerButtonClick));
            }
            else
            {
                btnRefresh.PerformClick();
            }
        }

        private void btnPageSetNarrorwMarginFooter_Click(object sender, EventArgs e)
        {
            wordApp.Run("Page_SetNarrowMarginAndFooter");
        }

        

        private void TriggerButtonClick()
        {
            //UpdateStatus("Receiving Document Change Event");
            btnRefresh.PerformClick(); // 触发按钮的 Click 事件
        }
        private void OnDocumentChange()
        {
            // MessageBox.Show("Get DocumentChange Event");
            UpdateStatus("Get Document Change Event");
            try
            {
                if ( wordDoc.Name == wordApp.ActiveDocument.Name)
                {
                    return;
                }
            }
            catch
            {
                // wordDoc = wordApp.ActiveDocument;
                // Go on as ActiveDocument changed
                StartDetectWordAppThread();
                return;
            }

            TriggerRefresh();
        }

        private void btnMultiChoice_FindAndConvertNumberIndex_Click(object sender, EventArgs e)
        {
            wordApp.Run("MultiChoice_FindAndConvertNumberIndex");
        }

        private void btnFindAndFormat4HChoices_Click(object sender, EventArgs e)
        {
            wordApp.Run("MultiChoice_FindAndFormat4HChoices");
        }

        private void btnConvertToDocx_Click(object sender, EventArgs e)
        {
            try
            {
                wordApp.Run("ConvertFileToDocx");
                btnRefresh.PerformClick();
            }
            catch
            {
                lblWordDoc.Text = "Must open a word file first!";
            }
        }

        private void btnSaveAsPDF_Click(object sender, EventArgs e)
        {
            wordApp.Run("SaveAsPDF");
        }

        private void lblWordDoc_Click(object sender, EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnSetHeader_Click(object sender, EventArgs e)
        {
            wordApp.Run("SetHeader", tbHeader.Text);
            lbHeaderHistory.Items.Insert(0, tbHeader.Text);
            SaveListBoxToFile(lbHeaderHistory, headerHistoryFile);
        }


        private void SaveListBoxToFile(ListBox listBox, string filePath)
        {
            // 创建文件流并写入内容
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                foreach (var item in listBox.Items)
                {
                    sw.WriteLine(item.ToString());
                }
            }
        }

        private void LoadListBoxFromFile(ListBox listBox, string filePath)
        {
            // 创建文件流并写入内容
            using (StreamReader sr = new StreamReader(filePath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    listBox.Items.Add(line);
                }
            }
        }

        private Application? wordApp;
        private Document? wordDoc;
        private String headerHistoryFile;

        private void lbHeaderHistory_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbHeader.Text = Convert.ToString(lbHeaderHistory.SelectedItem);
        }

        private void Form1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void lbHeaderHistory_DoubleClick(object sender, EventArgs e)
        {
            lbHeaderHistory.Items.Remove(lbHeaderHistory.SelectedItem);
        }

        private void lblLocalPath_Click(object sender, EventArgs e)
        {
            string filePath;
            filePath = lblLocalPath.Text;
            Process.Start("explorer.exe", $"/select,\"{filePath}\"");
        }

        private void tbFileName_TextChanged(object sender, EventArgs e)
        {
        }

        private void btnRename_Click(object sender, EventArgs e)
        {
            string oldFileName = lblLocalPath.Text;
            string filePath = Path.GetDirectoryName(oldFileName);


            string newFileName = Path.Combine(filePath, tbFileName.Text);

            if (oldFileName == newFileName)
            {
                return;
            }

            wordDoc.Close(true);
            File.Move(oldFileName, newFileName);

            // Wait until new file exists
            int time_out = 1;
            while (!File.Exists(newFileName))
            {
                while (time_out < 10)
                {
                    Thread.Sleep(1000);
                    time_out++;
                }

                break;
            }

            if (time_out >= 10)
                return;

            wordApp.Documents.Open(newFileName);
            btnRefresh.PerformClick();
        }

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            string oldFileName = lblLocalPath.Text;
            string filePath = Path.GetDirectoryName(oldFileName);


            string newFileName = Path.Combine(filePath, tbFileName.Text);

            if (oldFileName == newFileName)
            {
                return;
            }

            wordDoc.Close(true);
            File.Copy(oldFileName, newFileName);

            // Wait until new file exists
            int time_out = 1;
            while (!File.Exists(newFileName))
            {
                while (time_out < 10)
                {
                    Thread.Sleep(1000);
                    time_out++;
                }

                break;
            }

            if (time_out >= 10)
                return;

            wordApp.Documents.Open(newFileName);
            btnRefresh.PerformClick();
        }

        private void OnWordAppNotFound()
        {
            lblWordDoc.Text = "Must open a word file first!";
            lblCompatibility.Text = "Word Compatibility: Unknown";
        }
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            wordApp = Marshal.GetActiveObject("Word.Application") as Application;
            if (wordApp is null) {
                UpdateStatus("Failed to get word app");
                OnWordAppNotFound();
                StartDetectWordAppThread();
                return;
            }

            if (wordApp.Visible == false)
            {
                wordApp.Visible = true;
            }

            try
            {
                wordDoc = wordApp.ActiveDocument;
            }
            catch
            {
                UpdateStatus("Exception: No document is open");

                return;
            }

            if (wordDoc is null)
            {
                UpdateStatus("Not found active document");
                OnWordAppNotFound();
                return;
            }

            UninstallWordAppEventHandler();

            lblWordDoc.Text = wordDoc.FullName;
            lblCompatibility.Text = "Word Compatibility: " + Convert.ToString(wordDoc.CompatibilityMode);
            string fileHeader =
                wordDoc.Sections.First.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Trim('\r', '\n');

            if (fileHeader.Length == 0)
            {
                tbHeader.Text = fileHeader;
            }

            lblLocalPath.Text = wordApp.Run("Docx_GetLocalPath", wordDoc.FullName);
            tbFileName.Text = Path.GetFileName(lblLocalPath.Text);

            UpdateDocumentsListBox(lbDocuments);

            InstallWordAppEventHandler();
        }

        private void btnCloseAllWordApps_Click(object sender, EventArgs e)
        {
            try
            {
                while (true)
                {
                    Application word = Marshal.GetActiveObject("Word.Application") as Application;
                    word.Quit();
                }
            }
            catch
            {
                UpdateStatus("All Word Applications Killed");
            }
        }

        private void UpdateStatus(string msg)
        {
            string? log = null;
            log = Convert.ToString(DateTime.Now);

            string? fn = null;
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);
            if (sf is not null) {
                fn = sf.GetMethod().Name;
                log = log + " " + "[" + fn + "]";
            }

            log = log + " " + msg;

            if (tbLog.InvokeRequired)
            {
                tbLog.Invoke( (MethodInvoker) delegate()
                {
                    lblStatus.Text = log;
                    tbLog.AppendText(log + Environment.NewLine);
                } );
            }
            else
            {
                lblStatus.Text = log;
                tbLog.AppendText(log + Environment.NewLine);
            }
        }


        private void lbDocuments_DoubleClick(object sender, EventArgs e)
        {

        }

        private void ActivateDocument(string doc)
        {
            // 获取所有打开的 Word 文档窗口
            Word.Windows wordWindows = wordApp.Windows;
            wordApp.WindowState = WdWindowState.wdWindowStateMinimize;
            for (int i = 1; i <= wordWindows.Count; i++)
            {
                if (doc == wordWindows[i].Document.Name)
                {
                    wordWindows[i].WindowState = WdWindowState.wdWindowStateNormal;
                    wordWindows[i].Activate();
                }
                else // 其他窗口
                {
                    wordWindows[i].WindowState = WdWindowState.wdWindowStateMinimize;
                }
            }

             UpdateStatus("Activate: " + doc);
        }

        private void lbDocuments_Click(object sender, EventArgs e)
        {
            if (lbDocuments.Items.Count == 0)
            {
                return;
            }

            string s = Convert.ToString(lbDocuments.SelectedItem);
            ActivateDocument(s);
        }

        private void FillDocumentsListBox(ListBox listDocx)
        {
            lbDocuments.Items.Clear();
            foreach (Document d in wordApp.Documents)
            {
                lbDocuments.Items.Add(d.Name);
            }
        }
        private void UpdateDocumentsListBox(ListBox listDocs)
        {
            FillDocumentsListBox(listDocs);

            // Try to select current document
            try
            {
                int idx = listDocs.FindStringExact(wordApp.ActiveDocument.Name);
                if (idx != ListBox.NoMatches)
                {
                    listDocs.SetSelected(idx, true);
                }
            }
            catch
            {
                UpdateStatus("UpdateDocument Exception");
            }
        }

        private void lbDocuments_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
