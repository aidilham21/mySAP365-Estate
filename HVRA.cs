using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

//-- DEVEXPRESS
using DevExpress.Data;
using DevExpress.Data.Filtering;
using DevExpress.Xpo;
using DevExpress.Xpo.DB;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Grid;

namespace HVRA
{
    public partial class frmHVRA : Syb3.Form.RB1.Syb3FormRB_1cn_SR
    {
        private DataTable dtReport = null;
        private string activeConnString = string.Empty;

        // Kontrol UI Visual Dashboard Floating Window (Bisa di-drag, gedein, kecilin)
        private XtraForm frmDashboard = null;
        private bool dashboardPositionInitialized = false;
        private ComboBoxEdit cmbDashboardScope = null;
        private ComboBoxEdit cmbDashboardAfdeling = null;
        private ComboBoxEdit cmbDashboardBlok = null;
        private DateEdit dtpDashboardTanggal = null;
        private LabelControl lblValHarian = null;
        private LabelControl lblValBulanan = null;
        private LabelControl lblValTahunan = null;
        private GridControl gridBreakdown = null;
        private GridView gvBreakdown = null;
        private PanelControl pnlChart = null;
        private SimpleButton btnViewChart = null;
        private SimpleButton btnViewGrid = null;
        private DataTable dtBreakdownData = null;
        private BarButtonItem btnRibbonDashboard = null;

        public frmHVRA()
        {
            InitializeComponent();
        }

        public frmHVRA(Syb3init1.mySyb3Init param_AppInit1, string[] param_App2, Object[] param_App3)
            : base(param_AppInit1)
        {
            InitializeComponent();
            _Syb3InitPub1 = param_AppInit1;
            _AppParam3 = param_App3;
        }

        private void frmHVRA_Load(object sender, EventArgs e)
        {
            if (this.DesignMode) return;
            Application.Idle += new EventHandler(frmHVRA_Idle);

            Syb3FormRB_1cn_reLoad1(sender, e, "EO", null, false, false, _Syb3InitPub1, _Syb3InitLoc1);

            _Syb3InitLoc1[_Syb3InitPub1.vfirstField] = "ID_PANEN";

            if (_5Connect_BUnitDB != null)
            {
                RepositoryItemComboBox repoUnit = new RepositoryItemComboBox();
                repoUnit.Items.Add("ESTATE");
                repoUnit.TextEditStyle = TextEditStyles.DisableTextEditor;
                _5Connect_BUnitDB.Edit = repoUnit;
                _5Connect_BUnitDB.EditValue = "ESTATE";
                _5Connect_BUnitDB.Enabled = false;
            }

            if (_Rb1cn_xtraTabControl1 != null)
            {
                _Rb1cn_xtraTabControl1.ShowTabHeader = DevExpress.Utils.DefaultBoolean.False;
                _Rb1cn_xtraTabControl1.SelectedTabPage = _Rb1cn_xtraTabPage1;
            }

            if (_gridViewSR1 != null)
            {
                _gridViewSR1.IndicatorWidth = 45;
                _gridViewSR1.OptionsBehavior.Editable = false;
                _gridViewSR1.OptionsBehavior.ReadOnly = true;
                _gridViewSR1.OptionsView.ShowGroupPanel = true;
                _gridViewSR1.OptionsView.ShowFooter = true;
                _gridViewSR1.OptionsView.RowAutoHeight = false;
                _gridViewSR1.RowHeight = 24;
            }

            InitDashboardSidebar();
            InitRibbonDashboardButton();
        }

        private void frmHVRA_Shown(object sender, EventArgs e)
        {
            Syb3DBload(sender, "idil_estate2026", _tblConsolBA);
            _Syb3InitPub1._mySyb3Db.Syb3FocusLastPositionSR(sender, _gridViewSR1, _Syb3InitPub1, _Syb3InitLoc1);
            _Syb3InitPub1.loadWait1 = false;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Syb3DBload(object sender, string vdbName, DataTable tblConsol1)
        {
            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStart(sender, _AppParam3, _gridViewSR1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repoMarque, this, _Syb3InitPub1, _Syb3InitLoc1);

            string targetDatabase = "idil_estate2026";
            var dsTemp = _Syb3InitPub1._mySyb3Db.Syb3IDDict1(targetDatabase, false, _Syb3InitPub1, _Syb3InitLoc1);
            var propConn = dsTemp?.GetType().GetProperty("ConnectionString");
            activeConnString = propConn != null ? propConn.GetValue(dsTemp) as string : string.Empty;

            if (string.IsNullOrEmpty(activeConnString))
            {
                throw new InvalidOperationException("ConnectionString database idil_estate2026 tidak ditemukan.");
            }

            string sqlQuery = "SELECT ID_PANEN, TANGGAL_PANEN, NIK_MANDOR, NAMA_MANDOR, BLOK_AREA, NAMA_BLOK, KODE_AFDELING, NAMA_AFDELING, TAHUN_TANAM, TOTAL_JANJANG FROM dbo.hvd1 ORDER BY TANGGAL_PANEN DESC, ID_PANEN DESC";

            dtReport = new DataTable();
            using (SqlConnection conn = new SqlConnection(activeConnString))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(sqlQuery, conn))
                {
                    da.Fill(dtReport);
                }
            }

            _dvSr1 = dtReport.DefaultView;
            _gridControlSR1.DataSource = _dvSr1;

            ConfigureGridColumnsHVDA();
            PopulateDashboardFilterOptions();
            CalculateDashboardMetrics();

            _Syb3InitPub1._mySyb3Layout.Syb3RefreshStatusBar2(sender, _Rb1cn_barStatSheet, _Rb1cn_barStatDB, _Rb1cn_barStatRec,
                _dvSr1.Count, _Rb1cn_barStatPt, _Syb3InitLoc1, _Syb3InitPub1);

            if (_gridViewSR1 != null)
            {
                _gridViewSR1.IndicatorWidth = 45;
            }

            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStop(sender, _gridViewSR1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repo, this, _Syb3InitPub1, _Syb3InitLoc1);
        }

        private void ConfigureGridColumnsHVDA()
        {
            BandedGridView view = _gridViewSR1 as BandedGridView;
            if (view == null) return;

            view.BeginUpdate();
            view.IndicatorWidth = 45;
            view.Bands.Clear();
            view.Columns.Clear();
            view.GroupCount = 0;

            AddColumnWithBand(view, "A", "ID_PANEN", "ID Panen", 110, DevExpress.Utils.HorzAlignment.Center, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "B", "TANGGAL_PANEN", "Tanggal Panen", 105, DevExpress.Utils.HorzAlignment.Center, "dd-MMM-yyyy", DevExpress.Utils.FormatType.DateTime);
            AddColumnWithBand(view, "C", "NIK_MANDOR", "NIK Mandor", 85, DevExpress.Utils.HorzAlignment.Center, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "D", "NAMA_MANDOR", "Nama Mandor", 140, DevExpress.Utils.HorzAlignment.Near, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "E", "BLOK_AREA", "Kode Blok", 80, DevExpress.Utils.HorzAlignment.Center, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "F", "NAMA_BLOK", "Nama Blok", 100, DevExpress.Utils.HorzAlignment.Near, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "G", "KODE_AFDELING", "Kode Afdeling", 90, DevExpress.Utils.HorzAlignment.Center, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "H", "NAMA_AFDELING", "Nama Afdeling", 110, DevExpress.Utils.HorzAlignment.Near, null, DevExpress.Utils.FormatType.None);
            AddColumnWithBand(view, "I", "TAHUN_TANAM", "Tahun Tanam", 85, DevExpress.Utils.HorzAlignment.Center, null, DevExpress.Utils.FormatType.None);

            GridBand bTotal = new GridBand();
            bTotal.Caption = "J";
            view.Bands.Add(bTotal);

            BandedGridColumn colTotal = view.Columns.AddField("TOTAL_JANJANG");
            colTotal.Caption = "Total Janjang";
            colTotal.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            colTotal.DisplayFormat.FormatString = "{0:#,##0}";
            colTotal.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            colTotal.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            colTotal.AppearanceHeader.BackColor = Color.FromArgb(254, 216, 177);
            colTotal.Width = 105;
            colTotal.Visible = true;
            colTotal.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            colTotal.SummaryItem.DisplayFormat = "{0:#,##0}";
            bTotal.Columns.Add(colTotal);

            view.EndUpdate();
            _gridControlSR1.RefreshDataSource();
            view.RefreshData();
        }

        private void AddColumnWithBand(BandedGridView view, string bandCaption, string fieldName, string colCaption, int width, DevExpress.Utils.HorzAlignment align, string formatString, DevExpress.Utils.FormatType formatType)
        {
            GridBand band = new GridBand();
            band.Caption = bandCaption;
            view.Bands.Add(band);

            BandedGridColumn col = view.Columns.AddField(fieldName);
            col.Caption = colCaption;
            col.AppearanceCell.TextOptions.HAlignment = align;
            col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            col.Width = width;
            col.Visible = true;

            if (!string.IsNullOrEmpty(formatString))
            {
                col.DisplayFormat.FormatType = formatType;
                col.DisplayFormat.FormatString = formatString;
            }

            band.Columns.Add(col);
        }

        // =========================================================================
        // --- FITUR VISUAL DASHBOARD WINDOW ---
        // =========================================================================

        private Image CreateDashboardIcon(int size)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                int pad = size >= 32 ? 2 : 1;
                int gap = size >= 32 ? 2 : 1;
                int tileW = (size - (pad * 2) - gap) / 2;
                int tileH = tileW;

                int x1 = pad;
                int x2 = pad + tileW + gap;
                int y1 = pad;
                int y2 = pad + tileH + gap;

                using (SolidBrush br = new SolidBrush(ColorTranslator.FromHtml("#20539B"))) { g.FillRectangle(br, x1, y1, tileW, tileH); }
                using (SolidBrush br = new SolidBrush(Color.FromArgb(39, 174, 96))) { g.FillRectangle(br, x2, y1, tileW, tileH); }
                using (SolidBrush br = new SolidBrush(Color.FromArgb(230, 126, 34))) { g.FillRectangle(br, x1, y2, tileW, tileH); }
                using (SolidBrush br = new SolidBrush(Color.FromArgb(41, 128, 185))) { g.FillRectangle(br, x2, y2, tileW, tileH); }

                using (Pen pen = new Pen(Color.White, size >= 32 ? 1.5f : 1f))
                {
                    int xStart = x2 + 2;
                    int yBase = y1 + tileH - 2;
                    g.DrawLine(pen, xStart + 2, yBase, xStart + 2, yBase - (size >= 32 ? 6 : 3));
                    g.DrawLine(pen, xStart + (size >= 32 ? 6 : 4), yBase, xStart + (size >= 32 ? 6 : 4), yBase - (size >= 32 ? 9 : 5));
                }
            }
            return bmp;
        }

        private void InitRibbonDashboardButton()
        {
            if (this.Ribbon == null) return;

            btnRibbonDashboard = new BarButtonItem(this.Ribbon.Manager, "\u00A0\u00A0\u00A0Visual\u00A0\u00A0\u00A0\n\u00A0\u00A0\u00A0Dashboard\u00A0\u00A0\u00A0");
            btnRibbonDashboard.RibbonStyle = RibbonItemStyles.Large | RibbonItemStyles.SmallWithText;
            btnRibbonDashboard.LargeGlyph = CreateDashboardIcon(32);
            btnRibbonDashboard.Glyph = CreateDashboardIcon(16);
            btnRibbonDashboard.ItemClick += (s, ev) => ToggleDashboardSidebar();

            RibbonPage viewPage = this.Ribbon.Pages.Cast<RibbonPage>()
                .FirstOrDefault(p => p.Text.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0
                                  || p.Name.IndexOf("Page7", StringComparison.OrdinalIgnoreCase) >= 0
                                  || p.Name.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0
                                  || p.Text.IndexOf("Tampilan", StringComparison.OrdinalIgnoreCase) >= 0);

            if (viewPage == null && this.Ribbon.Pages.Count >= 7) viewPage = this.Ribbon.Pages[6];
            else if (viewPage == null && this.Ribbon.Pages.Count > 0) viewPage = this.Ribbon.Pages[this.Ribbon.Pages.Count - 1];

            if (viewPage != null)
            {
                RibbonPageGroup targetGroup = viewPage.Groups.Cast<RibbonPageGroup>().FirstOrDefault(g => g.Text.IndexOf("Dashboard", StringComparison.OrdinalIgnoreCase) >= 0);

                if (targetGroup == null)
                {
                    targetGroup = new RibbonPageGroup("Dashboard");
                    targetGroup.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
                    viewPage.Groups.Add(targetGroup);
                }
                else
                {
                    targetGroup.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
                }

                targetGroup.ItemLinks.Add(btnRibbonDashboard);
            }
        }

        private void ToggleDashboardSidebar()
        {
            if (frmDashboard == null || frmDashboard.IsDisposed) InitDashboardSidebar();

            if (frmDashboard.Visible) frmDashboard.Hide();
            else
            {
                if (!dashboardPositionInitialized)
                {
                    try
                    {
                        Point parentScreenLoc = this.PointToScreen(Point.Empty);
                        int initX = parentScreenLoc.X + this.Width - frmDashboard.Width - 25;
                        int initY = parentScreenLoc.Y + 115;
                        frmDashboard.Location = new Point(Math.Max(parentScreenLoc.X, initX), Math.Max(parentScreenLoc.Y, initY));
                        dashboardPositionInitialized = true;
                    }
                    catch { frmDashboard.StartPosition = FormStartPosition.CenterParent; }
                }

                frmDashboard.Show(this);
                frmDashboard.BringToFront();
                CalculateDashboardMetrics();
            }
        }

        private void InitDashboardSidebar()
        {
            frmDashboard = new XtraForm();
            frmDashboard.Text = "Visual Dashboard Panen - mySAP365";
            frmDashboard.Size = new Size(420, 780);
            frmDashboard.MinimumSize = new Size(380, 560);
            frmDashboard.FormBorderStyle = FormBorderStyle.Sizable;
            frmDashboard.MaximizeBox = true;
            frmDashboard.MinimizeBox = false;
            frmDashboard.ShowInTaskbar = false;
            frmDashboard.StartPosition = FormStartPosition.Manual;
            frmDashboard.Appearance.BackColor = Color.FromArgb(245, 247, 250);
            frmDashboard.Appearance.Options.UseBackColor = true;

            frmDashboard.FormClosing += (s, ev) =>
            {
                ev.Cancel = true;
                frmDashboard.Hide();
            };

            GroupControl grpFilter = new GroupControl();
            grpFilter.Text = "Parameter Analisis";
            grpFilter.Location = new Point(12, 12);
            grpFilter.Size = new Size(380, 150);
            grpFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            LabelControl lblScope = new LabelControl { Text = "Cakupan:", Location = new Point(12, 30) };
            cmbDashboardScope = new ComboBoxEdit { Location = new Point(115, 27), Width = 250 };
            cmbDashboardScope.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cmbDashboardScope.Properties.Items.AddRange(new string[] { "Seluruh Afdeling", "Per Afdeling", "Per Blok" });
            cmbDashboardScope.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
            cmbDashboardScope.SelectedIndex = 0;
            cmbDashboardScope.SelectedIndexChanged += (s, ev) =>
            {
                cmbDashboardAfdeling.Enabled = cmbDashboardScope.SelectedIndex >= 1;
                cmbDashboardBlok.Enabled = cmbDashboardScope.SelectedIndex == 2;
                CalculateDashboardMetrics();
            };

            LabelControl lblAfd = new LabelControl { Text = "Afdeling:", Location = new Point(12, 60) };
            cmbDashboardAfdeling = new ComboBoxEdit { Location = new Point(115, 57), Width = 250, Enabled = false };
            cmbDashboardAfdeling.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cmbDashboardAfdeling.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
            cmbDashboardAfdeling.SelectedIndexChanged += (s, ev) => CalculateDashboardMetrics();

            LabelControl lblBlk = new LabelControl { Text = "Blok:", Location = new Point(12, 90) };
            cmbDashboardBlok = new ComboBoxEdit { Location = new Point(115, 87), Width = 250, Enabled = false };
            cmbDashboardBlok.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cmbDashboardBlok.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
            cmbDashboardBlok.SelectedIndexChanged += (s, ev) => CalculateDashboardMetrics();

            LabelControl lblTgl = new LabelControl { Text = "Tanggal Acuan:", Location = new Point(12, 120) };
            dtpDashboardTanggal = new DateEdit { Location = new Point(115, 117), Width = 250 };
            dtpDashboardTanggal.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            dtpDashboardTanggal.EditValue = DateTime.Today;
            dtpDashboardTanggal.EditValueChanged += (s, ev) => CalculateDashboardMetrics();

            grpFilter.Controls.AddRange(new Control[] { lblScope, cmbDashboardScope, lblAfd, cmbDashboardAfdeling, lblBlk, cmbDashboardBlok, lblTgl, dtpDashboardTanggal });
            frmDashboard.Controls.Add(grpFilter);

            PanelControl cardHarian = CreateKPICard("PANEN HARIAN", Color.FromArgb(39, 174, 96), 172, out lblValHarian);
            PanelControl cardBulanan = CreateKPICard("PANEN BULANAN", ColorTranslator.FromHtml("#20539B"), 254, out lblValBulanan);
            PanelControl cardTahunan = CreateKPICard("PANEN TAHUNAN", Color.FromArgb(230, 126, 34), 336, out lblValTahunan);

            frmDashboard.Controls.AddRange(new Control[] { cardHarian, cardBulanan, cardTahunan });

            GroupControl grpVisual = new GroupControl();
            grpVisual.Text = "Distribusi & Grafik Hasil Panen";
            grpVisual.Location = new Point(12, 420);
            grpVisual.Size = new Size(380, 305);
            grpVisual.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            PanelControl pnlNavMode = new PanelControl();
            pnlNavMode.Dock = DockStyle.Top;
            pnlNavMode.Height = 34;
            pnlNavMode.BorderStyle = BorderStyles.NoBorder;
            pnlNavMode.Appearance.BackColor = Color.FromArgb(240, 243, 246);
            pnlNavMode.Appearance.Options.UseBackColor = true;

            btnViewChart = new SimpleButton { Text = "📊 Grafik Batang", Location = new Point(4, 4), Size = new Size(130, 26) };
            btnViewGrid = new SimpleButton { Text = "📋 Tabel Rincian", Location = new Point(140, 4), Size = new Size(130, 26) };

            btnViewChart.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
            btnViewChart.Appearance.Options.UseFont = true;
            btnViewGrid.Appearance.Font = new Font("Segoe UI", 8.25F);
            btnViewGrid.Appearance.Options.UseFont = true;

            pnlNavMode.Controls.Add(btnViewChart);
            pnlNavMode.Controls.Add(btnViewGrid);
            grpVisual.Controls.Add(pnlNavMode);

            pnlChart = new PanelControl();
            pnlChart.Dock = DockStyle.Fill;
            pnlChart.BorderStyle = BorderStyles.NoBorder;
            pnlChart.Appearance.BackColor = Color.White;
            pnlChart.Appearance.Options.UseBackColor = true;
            pnlChart.Paint += new PaintEventHandler(pnlChart_Paint);
            pnlChart.Resize += (s, ev) => pnlChart.Invalidate();

            gridBreakdown = new GridControl { Dock = DockStyle.Fill, Visible = false };
            gvBreakdown = new GridView(gridBreakdown);
            gvBreakdown.OptionsView.ShowGroupPanel = false;
            gvBreakdown.OptionsView.ShowIndicator = false;
            gvBreakdown.OptionsBehavior.Editable = false;
            gvBreakdown.OptionsView.ShowFooter = true;
            gridBreakdown.MainView = gvBreakdown;

            grpVisual.Controls.Add(pnlChart);
            grpVisual.Controls.Add(gridBreakdown);

            btnViewChart.Click += (s, ev) =>
            {
                pnlChart.Visible = true;
                pnlChart.BringToFront();
                gridBreakdown.Visible = false;
                btnViewChart.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
                btnViewGrid.Appearance.Font = new Font("Segoe UI", 8.25F);
                pnlChart.Invalidate();
            };

            btnViewGrid.Click += (s, ev) =>
            {
                gridBreakdown.Visible = true;
                gridBreakdown.BringToFront();
                pnlChart.Visible = false;
                btnViewGrid.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
                btnViewChart.Appearance.Font = new Font("Segoe UI", 8.25F);
            };

            frmDashboard.Controls.Add(grpVisual);
        }

        private void pnlChart_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            PanelControl pnl = sender as PanelControl;
            if (pnl == null) return;

            int w = pnl.ClientSize.Width;
            int h = pnl.ClientSize.Height;

            g.Clear(Color.White);

            if (dtBreakdownData == null || dtBreakdownData.Rows.Count == 0)
            {
                using (Font fontMsg = new Font("Segoe UI", 9F, FontStyle.Italic))
                using (SolidBrush brMsg = new SolidBrush(Color.Gray))
                {
                    string msg = "Tidak ada data panen pada filter ini.";
                    SizeF sz = g.MeasureString(msg, fontMsg);
                    g.DrawString(msg, fontMsg, brMsg, (w - sz.Width) / 2, (h - sz.Height) / 2);
                }
                return;
            }

            int padLeft = 50;
            int padRight = 20;
            int padTop = 30;
            int padBottom = 38;

            int chartW = w - padLeft - padRight;
            int chartH = h - padTop - padBottom;
            if (chartW <= 30 || chartH <= 30) return;

            decimal maxVal = 0;
            foreach (DataRow r in dtBreakdownData.Rows)
            {
                decimal v = r["Janjang"] != DBNull.Value ? Convert.ToDecimal(r["Janjang"]) : 0;
                if (v > maxVal) maxVal = v;
            }

            if (maxVal == 0) maxVal = 100;
            else
            {
                double mag = Math.Pow(10, Math.Floor(Math.Log10((double)maxVal)));
                maxVal = (decimal)(Math.Ceiling((double)maxVal / mag) * mag);
            }

            using (Pen penGrid = new Pen(Color.FromArgb(235, 238, 242), 1f))
            using (Pen penAxis = new Pen(Color.FromArgb(180, 190, 200), 1.2f))
            using (Font fontAxis = new Font("Segoe UI", 7.5F))
            using (SolidBrush brAxis = new SolidBrush(Color.FromArgb(120, 130, 140)))
            {
                int steps = 4;
                for (int s = 0; s <= steps; s++)
                {
                    int yLine = padTop + chartH - (int)((s / (double)steps) * chartH);
                    decimal valStep = (maxVal / steps) * s;
                    if (s > 0)
                    {
                        g.DrawLine(penGrid, padLeft, yLine, padLeft + chartW, yLine);
                    }
                    string lbl = valStep >= 1000 ? $"{(valStep / 1000):0.#}k" : $"{valStep:#,##0}";
                    SizeF sz = g.MeasureString(lbl, fontAxis);
                    g.DrawString(lbl, fontAxis, brAxis, padLeft - sz.Width - 5, yLine - sz.Height / 2);
                }
                g.DrawLine(penAxis, padLeft, padTop + chartH, padLeft + chartW, padTop + chartH);
            }

            int count = dtBreakdownData.Rows.Count;
            float slotW = (float)chartW / count;
            int barW = (int)Math.Min(45, Math.Max(12, slotW * 0.6f));

            Color colStart = ColorTranslator.FromHtml("#20539B");
            Color colEnd = Color.FromArgb(65, 135, 195);

            using (Font fontVal = new Font("Segoe UI", 7.5F, FontStyle.Bold))
            using (Font fontKey = new Font("Segoe UI", 8F))
            using (SolidBrush brKey = new SolidBrush(Color.FromArgb(50, 60, 70)))
            using (SolidBrush brVal = new SolidBrush(colStart))
            {
                for (int i = 0; i < count; i++)
                {
                    DataRow r = dtBreakdownData.Rows[i];
                    string key = r["Entitas"]?.ToString() ?? "";
                    decimal val = r["Janjang"] != DBNull.Value ? Convert.ToDecimal(r["Janjang"]) : 0;

                    int barH = (int)((val / maxVal) * chartH);
                    int x = padLeft + (int)(i * slotW + (slotW - barW) / 2);
                    int y = padTop + chartH - barH;

                    if (barH > 0)
                    {
                        Rectangle rectBar = new Rectangle(x, y, barW, barH);
                        using (LinearGradientBrush lgb = new LinearGradientBrush(rectBar, colStart, colEnd, LinearGradientMode.Vertical))
                        {
                            g.FillRectangle(lgb, rectBar);
                        }
                    }

                    if (val > 0)
                    {
                        string valText = $"{val:#,##0}";
                        SizeF szVal = g.MeasureString(valText, fontVal);
                        g.DrawString(valText, fontVal, brVal, x + (barW - szVal.Width) / 2, y - szVal.Height - 2);
                    }

                    SizeF szKey = g.MeasureString(key, fontKey);
                    g.DrawString(key, fontKey, brKey, x + (barW - szKey.Width) / 2, padTop + chartH + 5);
                }
            }
        }

        private PanelControl CreateKPICard(string title, Color accentColor, int topPos, out LabelControl lblValue)
        {
            PanelControl card = new PanelControl();
            card.Location = new Point(12, topPos);
            card.Size = new Size(380, 74);
            card.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.BorderStyle = BorderStyles.Simple;
            card.Appearance.BackColor = Color.White;
            card.Appearance.Options.UseBackColor = true;

            Panel pnlBar = new Panel { Dock = DockStyle.Left, Width = 6, BackColor = accentColor };
            card.Controls.Add(pnlBar);

            LabelControl lblTitle = new LabelControl
            {
                Text = title,
                Location = new Point(16, 10),
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                ForeColor = Color.Gray
            };
            card.Controls.Add(lblTitle);

            lblValue = new LabelControl
            {
                Text = "0 Janjang",
                Location = new Point(16, 30),
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = accentColor
            };
            card.Controls.Add(lblValue);

            return card;
        }

        private void PopulateDashboardFilterOptions()
        {
            if (dtReport == null || dtReport.Rows.Count == 0) return;

            cmbDashboardAfdeling.Properties.Items.Clear();
            var afdelings = dtReport.AsEnumerable()
                .Select(r => r["KODE_AFDELING"]?.ToString())
                .Where(s => !string.IsNullOrEmpty(s))
                .Distinct()
                .OrderBy(s => s);
            foreach (var a in afdelings) cmbDashboardAfdeling.Properties.Items.Add(a);
            if (cmbDashboardAfdeling.Properties.Items.Count > 0) cmbDashboardAfdeling.SelectedIndex = 0;

            cmbDashboardBlok.Properties.Items.Clear();
            var bloks = dtReport.AsEnumerable()
                .Select(r => r["BLOK_AREA"]?.ToString())
                .Where(s => !string.IsNullOrEmpty(s))
                .Distinct()
                .OrderBy(s => s);
            foreach (var b in bloks) cmbDashboardBlok.Properties.Items.Add(b);
            if (cmbDashboardBlok.Properties.Items.Count > 0) cmbDashboardBlok.SelectedIndex = 0;

            if (dtpDashboardTanggal != null && dtReport.Columns.Contains("TANGGAL_PANEN"))
            {
                var maxDate = dtReport.AsEnumerable()
                    .Where(r => r["TANGGAL_PANEN"] != DBNull.Value)
                    .Select(r => Convert.ToDateTime(r["TANGGAL_PANEN"]))
                    .DefaultIfEmpty(DateTime.Today)
                    .Max();
                dtpDashboardTanggal.EditValue = maxDate;
            }
        }

        private void CalculateDashboardMetrics()
        {
            if (dtReport == null || dtpDashboardTanggal == null || dtpDashboardTanggal.EditValue == null) return;

            DateTime targetDate = Convert.ToDateTime(dtpDashboardTanggal.EditValue).Date;
            int scopeIndex = cmbDashboardScope != null ? cmbDashboardScope.SelectedIndex : 0;
            string selAfd = cmbDashboardAfdeling?.EditValue?.ToString() ?? "";
            string selBlk = cmbDashboardBlok?.EditValue?.ToString() ?? "";

            var baseRows = dtReport.AsEnumerable().Where(r => r["TANGGAL_PANEN"] != DBNull.Value && r["TOTAL_JANJANG"] != DBNull.Value);

            if (scopeIndex == 1 && !string.IsNullOrEmpty(selAfd))
            {
                baseRows = baseRows.Where(r => (r["KODE_AFDELING"]?.ToString() ?? "").Equals(selAfd, StringComparison.OrdinalIgnoreCase));
            }
            else if (scopeIndex == 2 && !string.IsNullOrEmpty(selBlk))
            {
                baseRows = baseRows.Where(r => (r["BLOK_AREA"]?.ToString() ?? "").Equals(selBlk, StringComparison.OrdinalIgnoreCase));
            }

            var cachedRows = baseRows.ToList();

            decimal harian = cachedRows.Where(r => Convert.ToDateTime(r["TANGGAL_PANEN"]).Date == targetDate).Sum(r => Convert.ToDecimal(r["TOTAL_JANJANG"]));
            decimal bulanan = cachedRows.Where(r => Convert.ToDateTime(r["TANGGAL_PANEN"]).Year == targetDate.Year && Convert.ToDateTime(r["TANGGAL_PANEN"]).Month == targetDate.Month).Sum(r => Convert.ToDecimal(r["TOTAL_JANJANG"]));
            decimal tahunan = cachedRows.Where(r => Convert.ToDateTime(r["TANGGAL_PANEN"]).Year == targetDate.Year).Sum(r => Convert.ToDecimal(r["TOTAL_JANJANG"]));

            if (lblValHarian != null) lblValHarian.Text = $"{harian:#,##0} JJG";
            if (lblValBulanan != null) lblValBulanan.Text = $"{bulanan:#,##0} JJG";
            if (lblValTahunan != null) lblValTahunan.Text = $"{tahunan:#,##0} JJG";

            DataTable dtDist = new DataTable();
            dtDist.Columns.Add("Entitas", typeof(string));
            dtDist.Columns.Add("Janjang", typeof(decimal));

            if (scopeIndex == 0)
            {
                var dist = cachedRows
                    .Where(r => Convert.ToDateTime(r["TANGGAL_PANEN"]).Year == targetDate.Year && Convert.ToDateTime(r["TANGGAL_PANEN"]).Month == targetDate.Month)
                    .GroupBy(r => r["KODE_AFDELING"]?.ToString() ?? "N/A")
                    .Select(g => new { Key = g.Key, Total = g.Sum(x => Convert.ToDecimal(x["TOTAL_JANJANG"])) })
                    .OrderByDescending(x => x.Total);
                foreach (var d in dist) dtDist.Rows.Add(d.Key, d.Total);
            }
            else
            {
                var dist = cachedRows
                    .Where(r => Convert.ToDateTime(r["TANGGAL_PANEN"]).Year == targetDate.Year && Convert.ToDateTime(r["TANGGAL_PANEN"]).Month == targetDate.Month)
                    .GroupBy(r => r["BLOK_AREA"]?.ToString() ?? "N/A")
                    .Select(g => new { Key = g.Key, Total = g.Sum(x => Convert.ToDecimal(x["TOTAL_JANJANG"])) })
                    .OrderByDescending(x => x.Total);
                foreach (var d in dist) dtDist.Rows.Add(d.Key, d.Total);
            }

            dtBreakdownData = dtDist;

            if (gridBreakdown != null)
            {
                gridBreakdown.DataSource = dtDist;
                if (gvBreakdown.Columns.Count >= 2)
                {
                    gvBreakdown.Columns[0].Caption = scopeIndex == 0 ? "Afdeling" : "Blok Area";
                    gvBreakdown.Columns[1].Caption = "Total Janjang (Bulan Ini)";
                    gvBreakdown.Columns[1].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    gvBreakdown.Columns[1].DisplayFormat.FormatString = "{0:#,##0}";
                    gvBreakdown.Columns[1].Summary.Clear();
                    gvBreakdown.Columns[1].Summary.Add(DevExpress.Data.SummaryItemType.Sum, "Janjang", "{0:#,##0}");
                }
            }

            if (pnlChart != null) pnlChart.Invalidate();
        }

        private void frmHVRA_Idle(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
        }

        private void frmHVRA_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (frmDashboard != null && !frmDashboard.IsDisposed)
            {
                frmDashboard.Dispose();
                frmDashboard = null;
            }
            this.Dispose(true);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void _Realtimer1_Tick(object sender, EventArgs e) { Syb3DBload(sender, "idil_estate2026", _tblConsolBA); }
        private void _btnQOpen_ItemClick(object sender, ItemClickEventArgs e) { Syb3DBload(sender, "idil_estate2026", _tblConsolBA); }
        private void _5Connect_btnRefresh_ItemClick(object sender, ItemClickEventArgs e) { Syb3DBload(sender, "idil_estate2026", _tblConsolBA); }
        private void _gridViewSR1_RowCellStyle(object sender, RowCellStyleEventArgs e) { }
        private void _gridViewSR1_DoubleClick(object sender, EventArgs e) { }
        private void _gridViewSR1_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e) { }
        private void _gridViewSR1_CustomSummaryExists(object sender, CustomSummaryExistEventArgs e) { }
        private void _gridViewSR1_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e) { }
        private void _gridViewSR1_CustomRowCellEdit(object sender, CustomRowCellEditEventArgs e) { }
        private void _6Report_btnRGroup_repo_EditValueChanging(object sender, ChangingEventArgs e) { }
        private void _4Funct_btnHRM9_EditValueChanged(object sender, EventArgs e) { }
        private void Syb3Repo_Load(object sender, Boolean reload1) { }
    }
}