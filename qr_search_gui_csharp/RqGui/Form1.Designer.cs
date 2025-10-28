namespace RqGui;

partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.menuStrip1 = new MenuStrip();
        this.toolsToolStripMenuItem = new ToolStripMenuItem();
        this.settingsToolStripMenuItem = new ToolStripMenuItem();
        this.clearHistoryToolStripMenuItem = new ToolStripMenuItem();
        this.btnSearchHistory = new Button();
        this.historyContextMenu = new ContextMenuStrip();
        this.txtSearchPath = new TextBox();
        this.btnBrowse = new Button();
        this.txtSearchTerm = new TextBox();
        this.btnSearch = new Button();
        this.grpFilters = new GroupBox();
        this.chkGlob = new CheckBox();
        this.chkRegex = new CheckBox();
        this.chkCaseSensitive = new CheckBox();
        this.chkIncludeHidden = new CheckBox();
        this.chkExtensions = new CheckBox();
        this.txtExtensions = new TextBox();
        this.chkFileType = new CheckBox();
        this.txtFileType = new TextBox();
        this.chkSize = new CheckBox();
        this.txtSize = new TextBox();
        this.chkDateAfter = new CheckBox();
        this.dtpDateAfter = new DateTimePicker();
        this.chkDateBefore = new CheckBox();
        this.dtpDateBefore = new DateTimePicker();
        this.chkMaxDepth = new CheckBox();
        this.txtMaxDepth = new TextBox();
        this.chkThreads = new CheckBox();
        this.txtThreads = new TextBox();
        this.chkTimeout = new CheckBox();
        this.txtTimeout = new TextBox();
        this.lvResults = new ListView();
        this.statusStrip1 = new StatusStrip();
        this.statusLabel = new ToolStripStatusLabel();
        
        Label lblDirectory = new Label();
        Label lblSearchTerm = new Label();
        
        this.statusStrip1.SuspendLayout();
        this.menuStrip1.SuspendLayout();
        this.grpFilters.SuspendLayout();
        this.SuspendLayout();
        
        // menuStrip1
        this.menuStrip1.Items.AddRange(new ToolStripItem[] { this.toolsToolStripMenuItem });
        this.menuStrip1.Location = new System.Drawing.Point(0, 0);
        this.menuStrip1.Name = "menuStrip1";
        this.menuStrip1.Size = new System.Drawing.Size(1200, 24);
        
        // toolsToolStripMenuItem
        this.toolsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.settingsToolStripMenuItem, this.clearHistoryToolStripMenuItem });
        this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
        this.toolsToolStripMenuItem.Text = "&Tools";
        
        // settingsToolStripMenuItem
        this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
        this.settingsToolStripMenuItem.Text = "&Settings...";
        this.settingsToolStripMenuItem.Click += new EventHandler(this.settingsToolStripMenuItem_Click);
        
        // clearHistoryToolStripMenuItem
        this.clearHistoryToolStripMenuItem.Name = "clearHistoryToolStripMenuItem";
        this.clearHistoryToolStripMenuItem.Text = "Clear Search &History";
        this.clearHistoryToolStripMenuItem.Click += new EventHandler(this.clearHistoryToolStripMenuItem_Click);
        
        // btnSearchHistory
        this.btnSearchHistory.Location = new System.Drawing.Point(80, 32);
        this.btnSearchHistory.Name = "btnSearchHistory";
        this.btnSearchHistory.Size = new System.Drawing.Size(140, 25);
        this.btnSearchHistory.Text = "Recent Searches ▼";
        this.btnSearchHistory.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        this.btnSearchHistory.Click += new EventHandler(this.btnSearchHistory_Click);
        
        // historyContextMenu
        this.historyContextMenu.Name = "historyContextMenu";
        this.historyContextMenu.Size = new System.Drawing.Size(180, 70);
        
        // lblDirectory
        lblDirectory.AutoSize = true;
        lblDirectory.Location = new System.Drawing.Point(12, 35);
        lblDirectory.Name = "lblDirectory";
        lblDirectory.Text = "Directory:";
        
        // txtSearchPath
        this.txtSearchPath.Location = new System.Drawing.Point(230, 32);
        this.txtSearchPath.Name = "txtSearchPath";
        this.txtSearchPath.Size = new System.Drawing.Size(750, 23);
        this.txtSearchPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        this.txtSearchPath.KeyDown += new KeyEventHandler(this.txtSearchPath_KeyDown);
        
        // btnBrowse
        this.btnBrowse.Location = new System.Drawing.Point(990, 30);
        this.btnBrowse.Name = "btnBrowse";
        this.btnBrowse.Size = new System.Drawing.Size(85, 27);
        this.btnBrowse.Text = "Browse...";
        this.btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        this.btnBrowse.Click += new EventHandler(this.btnBrowse_Click);
        
        // lblSearchTerm
        lblSearchTerm.AutoSize = true;
        lblSearchTerm.Location = new System.Drawing.Point(12, 68);
        lblSearchTerm.Name = "lblSearchTerm";
        lblSearchTerm.Text = "Search For:";
        
        // txtSearchTerm
        this.txtSearchTerm.Location = new System.Drawing.Point(80, 65);
        this.txtSearchTerm.Name = "txtSearchTerm";
        this.txtSearchTerm.Size = new System.Drawing.Size(900, 23);
        this.txtSearchTerm.Text = "*";
        this.txtSearchTerm.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        this.txtSearchTerm.KeyDown += new KeyEventHandler(this.txtSearchTerm_KeyDown);
        
        // btnSearch
        this.btnSearch.Location = new System.Drawing.Point(990, 63);
        this.btnSearch.Name = "btnSearch";
        this.btnSearch.Size = new System.Drawing.Size(85, 27);
        this.btnSearch.Text = "Search";
        this.btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        this.btnSearch.Click += new EventHandler(this.btnSearch_Click);
        
        // grpFilters
        this.grpFilters.Location = new System.Drawing.Point(12, 100);
        this.grpFilters.Name = "grpFilters";
        this.grpFilters.Size = new System.Drawing.Size(1160, 175);
        this.grpFilters.Text = "Filters";
        this.grpFilters.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        
        // ROW 1: Pattern options
        // chkGlob
        this.chkGlob.AutoSize = true;
        this.chkGlob.Location = new System.Drawing.Point(15, 25);
        this.chkGlob.Name = "chkGlob";
        this.chkGlob.Text = "Glob Patterns";
        
        // chkRegex
        this.chkRegex.AutoSize = true;
        this.chkRegex.Location = new System.Drawing.Point(130, 25);
        this.chkRegex.Name = "chkRegex";
        this.chkRegex.Text = "Use Regex";
        
        // chkCaseSensitive
        this.chkCaseSensitive.AutoSize = true;
        this.chkCaseSensitive.Location = new System.Drawing.Point(230, 25);
        this.chkCaseSensitive.Name = "chkCaseSensitive";
        this.chkCaseSensitive.Text = "Case Sensitive";
        
        // chkIncludeHidden
        this.chkIncludeHidden.AutoSize = true;
        this.chkIncludeHidden.Location = new System.Drawing.Point(360, 25);
        this.chkIncludeHidden.Name = "chkIncludeHidden";
        this.chkIncludeHidden.Text = "Include Hidden";
        
        // ROW 2: Extensions and File Type
        // chkExtensions
        this.chkExtensions.AutoSize = false;
        this.chkExtensions.Location = new System.Drawing.Point(15, 65);
        this.chkExtensions.Name = "chkExtensions";
        this.chkExtensions.Size = new System.Drawing.Size(100, 20);
        this.chkExtensions.Text = "Extensions:";
        
        // txtExtensions
        this.txtExtensions.Location = new System.Drawing.Point(120, 62);
        this.txtExtensions.Name = "txtExtensions";
        this.txtExtensions.Size = new System.Drawing.Size(140, 23);
        this.txtExtensions.PlaceholderText = ".txt,.pdf";
        
        // chkFileType
        this.chkFileType.AutoSize = false;
        this.chkFileType.Location = new System.Drawing.Point(280, 65);
        this.chkFileType.Name = "chkFileType";
        this.chkFileType.Size = new System.Drawing.Size(65, 20);
        this.chkFileType.Text = "Type:";
        
        // txtFileType
        this.txtFileType.Location = new System.Drawing.Point(350, 62);
        this.txtFileType.Name = "txtFileType";
        this.txtFileType.Size = new System.Drawing.Size(120, 23);
        this.txtFileType.PlaceholderText = "image";
        
        // chkSize
        this.chkSize.AutoSize = false;
        this.chkSize.Location = new System.Drawing.Point(490, 65);
        this.chkSize.Name = "chkSize";
        this.chkSize.Size = new System.Drawing.Size(60, 20);
        this.chkSize.Text = "Size:";
        
        // txtSize
        this.txtSize.Location = new System.Drawing.Point(555, 62);
        this.txtSize.Name = "txtSize";
        this.txtSize.Size = new System.Drawing.Size(120, 23);
        this.txtSize.PlaceholderText = "+1MB";
        
        // chkMaxDepth
        this.chkMaxDepth.AutoSize = false;
        this.chkMaxDepth.Location = new System.Drawing.Point(695, 65);
        this.chkMaxDepth.Name = "chkMaxDepth";
        this.chkMaxDepth.Size = new System.Drawing.Size(70, 20);
        this.chkMaxDepth.Text = "Depth:";
        
        // txtMaxDepth
        this.txtMaxDepth.Location = new System.Drawing.Point(770, 62);
        this.txtMaxDepth.Name = "txtMaxDepth";
        this.txtMaxDepth.Size = new System.Drawing.Size(60, 23);
        this.txtMaxDepth.PlaceholderText = "3";
        
        // ROW 3: Date filters
        // chkDateAfter
        this.chkDateAfter.AutoSize = false;
        this.chkDateAfter.Location = new System.Drawing.Point(15, 105);
        this.chkDateAfter.Name = "chkDateAfter";
        this.chkDateAfter.Size = new System.Drawing.Size(60, 20);
        this.chkDateAfter.Text = "After:";
        
        // dtpDateAfter
        this.dtpDateAfter.Location = new System.Drawing.Point(80, 103);
        this.dtpDateAfter.Name = "dtpDateAfter";
        this.dtpDateAfter.Size = new System.Drawing.Size(165, 23);
        this.dtpDateAfter.Format = DateTimePickerFormat.Short;
        
        // chkDateBefore
        this.chkDateBefore.AutoSize = false;
        this.chkDateBefore.Location = new System.Drawing.Point(265, 105);
        this.chkDateBefore.Name = "chkDateBefore";
        this.chkDateBefore.Size = new System.Drawing.Size(70, 20);
        this.chkDateBefore.Text = "Before:";
        
        // dtpDateBefore
        this.dtpDateBefore.Location = new System.Drawing.Point(340, 103);
        this.dtpDateBefore.Name = "dtpDateBefore";
        this.dtpDateBefore.Size = new System.Drawing.Size(165, 23);
        this.dtpDateBefore.Format = DateTimePickerFormat.Short;
        
        // ROW 4: Performance options
        // chkThreads
        this.chkThreads.AutoSize = false;
        this.chkThreads.Location = new System.Drawing.Point(525, 105);
        this.chkThreads.Name = "chkThreads";
        this.chkThreads.Size = new System.Drawing.Size(80, 20);
        this.chkThreads.Text = "Threads:";
        
        // txtThreads
        this.txtThreads.Location = new System.Drawing.Point(610, 103);
        this.txtThreads.Name = "txtThreads";
        this.txtThreads.Size = new System.Drawing.Size(70, 23);
        this.txtThreads.PlaceholderText = "8";
        
        // chkTimeout
        this.chkTimeout.AutoSize = false;
        this.chkTimeout.Location = new System.Drawing.Point(700, 105);
        this.chkTimeout.Name = "chkTimeout";
        this.chkTimeout.Size = new System.Drawing.Size(110, 20);
        this.chkTimeout.Text = "Timeout (ms):";
        
        // txtTimeout
        this.txtTimeout.Location = new System.Drawing.Point(815, 103);
        this.txtTimeout.Name = "txtTimeout";
        this.txtTimeout.Size = new System.Drawing.Size(90, 23);
        this.txtTimeout.PlaceholderText = "30000";
        
        // Add all controls to grpFilters
        this.grpFilters.Controls.Add(this.chkGlob);
        this.grpFilters.Controls.Add(this.chkRegex);
        this.grpFilters.Controls.Add(this.chkCaseSensitive);
        this.grpFilters.Controls.Add(this.chkIncludeHidden);
        this.grpFilters.Controls.Add(this.chkExtensions);
        this.grpFilters.Controls.Add(this.txtExtensions);
        this.grpFilters.Controls.Add(this.chkFileType);
        this.grpFilters.Controls.Add(this.txtFileType);
        this.grpFilters.Controls.Add(this.chkSize);
        this.grpFilters.Controls.Add(this.txtSize);
        this.grpFilters.Controls.Add(this.chkMaxDepth);
        this.grpFilters.Controls.Add(this.txtMaxDepth);
        this.grpFilters.Controls.Add(this.chkDateAfter);
        this.grpFilters.Controls.Add(this.dtpDateAfter);
        this.grpFilters.Controls.Add(this.chkDateBefore);
        this.grpFilters.Controls.Add(this.dtpDateBefore);
        this.grpFilters.Controls.Add(this.chkThreads);
        this.grpFilters.Controls.Add(this.txtThreads);
        this.grpFilters.Controls.Add(this.chkTimeout);
        this.grpFilters.Controls.Add(this.txtTimeout);
        
        // lvResults
        this.lvResults.Columns.AddRange(new ColumnHeader[] {
            new ColumnHeader() { Text = "Name", Width = 230 },
            new ColumnHeader() { Text = "Path", Width = 500 },
            new ColumnHeader() { Text = "Size", Width = 90 },
            new ColumnHeader() { Text = "Modified", Width = 140 },
            new ColumnHeader() { Text = "Created", Width = 140 }
        });
        this.lvResults.FullRowSelect = true;
        this.lvResults.GridLines = true;
        this.lvResults.Location = new System.Drawing.Point(12, 285);
        this.lvResults.Name = "lvResults";
        this.lvResults.Size = new System.Drawing.Size(1160, 320);
        this.lvResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.lvResults.View = View.Details;
        this.lvResults.DoubleClick += new EventHandler(this.lvResults_DoubleClick);
        
        // statusStrip1
        this.statusStrip1.Items.AddRange(new ToolStripItem[] { this.statusLabel });
        this.statusStrip1.Location = new System.Drawing.Point(0, 615);
        this.statusStrip1.Name = "statusStrip1";
        this.statusStrip1.Size = new System.Drawing.Size(1184, 22);
        
        // statusLabel
        this.statusLabel.Name = "statusLabel";
        this.statusLabel.Text = "Ready";
        
        // Form1
        this.ClientSize = new System.Drawing.Size(1184, 637);
        this.Controls.Add(this.statusStrip1);
        this.Controls.Add(this.lvResults);
        this.Controls.Add(this.grpFilters);
        this.Controls.Add(this.btnSearch);
        this.Controls.Add(this.txtSearchTerm);
        this.Controls.Add(lblSearchTerm);
        this.Controls.Add(this.btnBrowse);
        this.Controls.Add(this.txtSearchPath);
        this.Controls.Add(this.btnSearchHistory);
        this.Controls.Add(lblDirectory);
        this.Controls.Add(this.menuStrip1);
        this.MainMenuStrip = this.menuStrip1;
        this.Name = "Form1";
        this.Text = "RQ File Search";
        this.MinimumSize = new System.Drawing.Size(1200, 600);
        this.AcceptButton = this.btnSearch;
        this.KeyPreview = true;
        
        this.statusStrip1.ResumeLayout(false);
        this.menuStrip1.ResumeLayout(false);
        this.grpFilters.ResumeLayout(false);
        this.grpFilters.PerformLayout();
        this.ResumeLayout(false);
        this.PerformLayout();
    }

    private MenuStrip menuStrip1;
    private ToolStripMenuItem toolsToolStripMenuItem;
    private ToolStripMenuItem settingsToolStripMenuItem;
    private ToolStripMenuItem clearHistoryToolStripMenuItem;
    private Button btnSearchHistory;
    private ContextMenuStrip historyContextMenu;
    private TextBox txtSearchPath;
    private Button btnBrowse;
    private TextBox txtSearchTerm;
    private Button btnSearch;
    private GroupBox grpFilters;
    private CheckBox chkGlob;
    private CheckBox chkRegex;
    private CheckBox chkCaseSensitive;
    private CheckBox chkIncludeHidden;
    private CheckBox chkExtensions;
    private TextBox txtExtensions;
    private CheckBox chkFileType;
    private TextBox txtFileType;
    private CheckBox chkSize;
    private TextBox txtSize;
    private CheckBox chkDateAfter;
    private DateTimePicker dtpDateAfter;
    private CheckBox chkDateBefore;
    private DateTimePicker dtpDateBefore;
    private CheckBox chkMaxDepth;
    private TextBox txtMaxDepth;
    private CheckBox chkThreads;
    private TextBox txtThreads;
    private CheckBox chkTimeout;
    private TextBox txtTimeout;
    private ListView lvResults;
    private StatusStrip statusStrip1;
    private ToolStripStatusLabel statusLabel;
}
