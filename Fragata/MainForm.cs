using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace Fragata
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class MainForm : Form
    {
        private bool windowMaximized = false;
        private XmlDocument menuXML;
        private CuadroAtencion cuadro;

        public MainForm()
        {
            InitializeComponent();
            ActivateFeatures();
        }

        private void ActivateFeatures()
        {
            const string BROWSER_EMULATION_KEY = @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
            const string GPU_RENDERING_KEY = @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_GPU_RENDERING";
            String appname = Process.GetCurrentProcess().ProcessName + ".exe";
            const int browserEmulationMode = 11001;
            const int GPURenderingMode = 1;

            RegistryKey browserEmulationKey =
                Registry.CurrentUser.OpenSubKey(BROWSER_EMULATION_KEY, RegistryKeyPermissionCheck.ReadWriteSubTree) ??
                Registry.CurrentUser.CreateSubKey(BROWSER_EMULATION_KEY);

            RegistryKey browserRenderingKey =
                Registry.CurrentUser.OpenSubKey(GPU_RENDERING_KEY, RegistryKeyPermissionCheck.ReadWriteSubTree) ??
                Registry.CurrentUser.CreateSubKey(GPU_RENDERING_KEY);

            if (browserEmulationKey != null)
            {
                browserEmulationKey.SetValue(appname, browserEmulationMode, RegistryValueKind.DWord);
                browserEmulationKey.Close();
            }

            if (browserRenderingKey != null)
            {
                browserRenderingKey.SetValue(appname, GPURenderingMode, RegistryValueKind.DWord);
                browserRenderingKey.Close();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string pattern = @"^\d{1,},\d{1,}$";
            String[] arguments = Environment.GetCommandLineArgs();
            for (int i = 0; i < arguments.Length; i++)
            {
                switch (arguments[i])
                {
                    case "-k":
                        browser_Fullscreen();
                        break;
                    case "-b":
                        browser_BorderLess();
                        break;
                    case "-url":
                        if (i < arguments.Length - 1 && arguments[i + 1] != "" && !arguments[i + 1].StartsWith("-"))
                        {
                            browser.Navigate(arguments[i + 1], false);
                        }
                        break;
                    case "-l":
                        if (i < arguments.Length - 1 && arguments[i + 1] != "" && !arguments[i + 1].StartsWith("-"))
                        {
                            if (Regex.IsMatch(arguments[i + 1], pattern))
                            {
                                this.Location = new Point(Convert.ToInt32(arguments[i + 1].Split(',')[0]), Convert.ToInt32(arguments[i + 1].Split(',')[1]));
                            }
                        }
                        break;
                    case "-s":
                        if (i < arguments.Length - 1 && arguments[i + 1] != "" && !arguments[i + 1].StartsWith("-"))
                        {
                            if (Regex.IsMatch(arguments[i + 1], pattern))
                            {
                                this.Size = new Size(Convert.ToInt32(arguments[i + 1].Split(',')[0]), Convert.ToInt32(arguments[i + 1].Split(',')[1]));
                            }
                        }
                        break;
                }
            }

            browser.ObjectForScripting = this;
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11 && e.Shift)
            {
                browser_BorderLess();
            }
            else if (e.KeyCode == Keys.F11)
            {
                browser_Fullscreen();
            }
            else if (e.KeyCode == Keys.F5)
            {
                menuXML = null;
                if (cuadro != null)
                {
                    cuadro.Dispose();
                }
                //browser.Navigate(browser.Url);
                browser.Refresh(WebBrowserRefreshOption.Completely);
            }
        }

        private void browser_Fullscreen()
        {
            if (this.WindowState == FormWindowState.Maximized && this.FormBorderStyle == FormBorderStyle.None)
            {
                toolStrip.Visible = true;
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = windowMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            }
            else
            {
                windowMaximized = this.WindowState == FormWindowState.Maximized ? true : false;
                toolStrip.Visible = false;
                this.WindowState = FormWindowState.Normal;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
            }
        }

        private void browser_BorderLess()
        {
            if (this.FormBorderStyle == FormBorderStyle.None)
            {
                toolStrip.Visible = true;
                this.FormBorderStyle = FormBorderStyle.Sizable;
            }
            else
            {
                toolStrip.Visible = false;
                this.FormBorderStyle = FormBorderStyle.None;
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            browser.GoBack();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            browser.GoForward();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            menuXML = null;
            if (cuadro != null)
            {
                cuadro.Dispose();
            }
            browser.Refresh(WebBrowserRefreshOption.Completely);
        }

        private void btnFullscreen_Click(object sender, EventArgs e)
        {
            browser_Fullscreen();
        }

        private void btnBorderless_Click(object sender, EventArgs e)
        {
            browser_BorderLess();
        }

        private void btnLocation_Click(object sender, EventArgs e)
        {
            string wsize = this.Size.Width + "," + this.Size.Height;
            string wpos = this.Location.X + "," + this.Location.Y;
            MessageBox.Show("Tamaño:\t" + wsize + "\nPosición:\t" + wpos, "Información");
        }

        private void txtURL_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control && e.Shift)
                {
                    txtURL.Text = "www." + txtURL.Text + ".org";                    
                }
                else if (e.Control)
                {
                    txtURL.Text = "www." + txtURL.Text + ".com";
                }
                else if (e.Shift)
                {
                    txtURL.Text = "www." + txtURL.Text + ".net";
                }                
                menuXML = null;
                if (cuadro != null)
                {
                    cuadro.Dispose();
                }
                browser.Navigate(txtURL.Text, false);
                browser.Focus();
            }
            else if (e.KeyCode == Keys.F11 && e.Shift)
            {
                browser_BorderLess();
            }
            else if (e.KeyCode == Keys.F11)
            {
                browser_Fullscreen();
            }
            else if (e.KeyCode == Keys.F5)
            {
                menuXML = null;
                if (cuadro != null)
                {
                    cuadro.Dispose();
                }
                //browser.Navigate(browser.Url);
                browser.Refresh(WebBrowserRefreshOption.Completely);
            }
        }

        private void browser_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.F11 && e.Shift)
            {
                browser_BorderLess();
            }
            else if (e.KeyCode == Keys.F11)
            {
                browser_Fullscreen();
            }
            else if (e.KeyCode == Keys.F5)
            {
                menuXML = null;
                if (cuadro != null)
                {
                    cuadro.Dispose();
                }
                //browser.Navigate(browser.Url);
                browser.Refresh(WebBrowserRefreshOption.Completely);
            }
        }

        private void browser_CanGoBackChanged(object sender, EventArgs e)
        {
            btnBack.Enabled = browser.CanGoBack;
        }

        private void browser_CanGoForwardChanged(object sender, EventArgs e)
        {
            btnNext.Enabled = browser.CanGoForward;
        }        

        private void browser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            txtURL.Text = browser.Url.ToString();
        }

        private void browser_GotFocus(object sender, EventArgs e)
        {
            this.Focus();
        }

        private void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url == browser.Url)
            {
                Console.WriteLine("Carga completa: {0}", e.Url);
            }
        }

        // -------------------------------------------//
        // Inicio de la lógica del cuadro de atención. //
        // -------------------------------------------//

        private void LoadMenu()
        {
            if (menuXML == null)
            {
                string url = browser.Url.ToString();
                if (!url.EndsWith("/")) url += "/";

                menuXML = new XmlDocument();
                try
                {
                    menuXML.Load(browser.Url.ToString() + "menuCuadro.xml");
                    Console.WriteLine("Menú XML cargado.");
                }
                catch (Exception ex)
                {
                    menuXML = null;

                    String dirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    if (!dirPath.EndsWith("\\")) dirPath += "\\";

                    using (StreamWriter writer = new StreamWriter(dirPath + "Fragata.log", true, Encoding.Default))
                    {
                        writer.WriteLine(DateTime.Now + "El archivo menuCuadro.xml no está disponible : " + ex.Message);
                    }
                }
                finally
                {
                    if (menuXML != null)
                    {
                        //browser.ObjectForScripting = this;
                        cuadro = new CuadroAtencion(menuXML);
                        cuadro.runScript = delegate(string method, object[] args)
                        {
                            return browser.Document.InvokeScript(method, args);
                        };
                        cuadro.StartRecognition(true);
                    }
                }
            }
        }

        private void UnloadMenu()
        {
            menuXML = null;
            if (cuadro != null)
            {
                cuadro.Dispose();
            }
            //browser.ObjectForScripting = null;
        }

        /// <summary>
        /// Métodos públicos disponibles mediante window.external
        /// </summary>

        public void startRecognition(bool start)
        {
            if (cuadro != null && menuXML != null)
            {
                cuadro.StartRecognition(start);
            }
            else
            {
                LoadMenu();
            }
        }

        public void startDocRecognition(bool start)
        {
            if (cuadro != null)
            {
                cuadro.StartDocRecognition(start);
            }
        }

        public void restartTimerIdle()
        {
            if (cuadro != null)
            {
                cuadro.RestartTimerIdle();
            }
        }

        public void getDocComplete()
        {
            if (cuadro != null)
            {
                cuadro.GetDocComplete();
            }
        }

        public void evalLiquidation(string result)
        {
            if (cuadro != null)
            {
                cuadro.EvalLiquidation(result);
            }
        }

    }
}
