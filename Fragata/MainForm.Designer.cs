namespace Fragata
{
    partial class MainForm
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.btnBack = new System.Windows.Forms.ToolStripButton();
            this.btnNext = new System.Windows.Forms.ToolStripButton();
            this.btnRefresh = new System.Windows.Forms.ToolStripButton();
            this.txtURL = new System.Windows.Forms.ToolStripTextBox();
            this.btnFullscreen = new System.Windows.Forms.ToolStripButton();
            this.btnBorderless = new System.Windows.Forms.ToolStripButton();
            this.btnLocation = new System.Windows.Forms.ToolStripButton();
            this.browser = new System.Windows.Forms.WebBrowser();
            this.toolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip
            // 
            this.toolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnBack,
            this.btnNext,
            this.btnRefresh,
            this.txtURL,
            this.btnFullscreen,
            this.btnBorderless,
            this.btnLocation});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Padding = new System.Windows.Forms.Padding(3, 2, 1, 2);
            this.toolStrip.Size = new System.Drawing.Size(1008, 27);
            this.toolStrip.TabIndex = 0;
            this.toolStrip.Text = "toolStrip";
            // 
            // btnBack
            // 
            this.btnBack.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnBack.Enabled = false;
            this.btnBack.Image = ((System.Drawing.Image)(resources.GetObject("btnBack.Image")));
            this.btnBack.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(23, 20);
            this.btnBack.Text = "Ir a la página anterior";
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnNext
            // 
            this.btnNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnNext.Enabled = false;
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(23, 20);
            this.btnNext.Text = "Ir a la página siguiente";
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(23, 20);
            this.btnRefresh.Text = "Recargar la página";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // txtURL
            // 
            this.txtURL.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.txtURL.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(500, 23);
            this.txtURL.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtURL_KeyDown);
            // 
            // btnFullscreen
            // 
            this.btnFullscreen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnFullscreen.Image = ((System.Drawing.Image)(resources.GetObject("btnFullscreen.Image")));
            this.btnFullscreen.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnFullscreen.Name = "btnFullscreen";
            this.btnFullscreen.Size = new System.Drawing.Size(23, 20);
            this.btnFullscreen.Text = "Ver en pantalla completa";
            this.btnFullscreen.Click += new System.EventHandler(this.btnFullscreen_Click);
            // 
            // btnBorderless
            // 
            this.btnBorderless.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnBorderless.Image = ((System.Drawing.Image)(resources.GetObject("btnBorderless.Image")));
            this.btnBorderless.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnBorderless.Name = "btnBorderless";
            this.btnBorderless.Size = new System.Drawing.Size(23, 20);
            this.btnBorderless.Text = "Ver ventana sin bordes";
            this.btnBorderless.Click += new System.EventHandler(this.btnBorderless_Click);
            // 
            // btnLocation
            // 
            this.btnLocation.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnLocation.Image = ((System.Drawing.Image)(resources.GetObject("btnLocation.Image")));
            this.btnLocation.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnLocation.Name = "btnLocation";
            this.btnLocation.Size = new System.Drawing.Size(23, 20);
            this.btnLocation.Text = "Obtener información de la ventana";
            this.btnLocation.Click += new System.EventHandler(this.btnLocation_Click);
            // 
            // browser
            // 
            this.browser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.browser.Location = new System.Drawing.Point(0, 27);
            this.browser.MinimumSize = new System.Drawing.Size(20, 20);
            this.browser.Name = "browser";
            this.browser.Size = new System.Drawing.Size(1008, 702);
            this.browser.TabIndex = 1;
            this.browser.TabStop = false;
            this.browser.Url = new System.Uri("", System.UriKind.Relative);
            this.browser.WebBrowserShortcutsEnabled = false;
            this.browser.CanGoBackChanged += new System.EventHandler(this.browser_CanGoBackChanged);
            this.browser.CanGoForwardChanged += new System.EventHandler(this.browser_CanGoForwardChanged);
            this.browser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.browser_DocumentCompleted);
            this.browser.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.browser_Navigated);
            this.browser.GotFocus += new System.EventHandler(this.browser_GotFocus);
            this.browser.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.browser_PreviewKeyDown);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1008, 729);
            this.Controls.Add(this.browser);
            this.Controls.Add(this.toolStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Fragata";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ToolStripButton btnBack;
        private System.Windows.Forms.ToolStripButton btnNext;
        private System.Windows.Forms.ToolStripButton btnRefresh;
        private System.Windows.Forms.ToolStripTextBox txtURL;
        private System.Windows.Forms.ToolStripButton btnFullscreen;
        private System.Windows.Forms.ToolStripButton btnBorderless;
        private System.Windows.Forms.ToolStripButton btnLocation;
        private System.Windows.Forms.WebBrowser browser;
    }
}

