using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Synthesis;
using Microsoft.Speech.Recognition;
using System.Globalization;
using System.Collections.Generic;
using System.IO;

namespace FragataLite
{
    public delegate object RunScript(string method, object[] args);

    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class MainForm : Form
    {
        private bool windowMaximized = false;
        private SpeechSynthesizer synthesizer;
        private SpeechRecognitionEngine recognizer;
        private RunScript returnRecognition;
        private string returnMethod;

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

            returnRecognition = delegate(string method, object[] args)
            {
                return browser.Document.InvokeScript(method, args);
            };
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
                if (recognizer != null)
                    recognizer.Dispose();

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
            if (recognizer != null)
                recognizer.Dispose();

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

                if (recognizer != null)
                    recognizer.Dispose();

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
                if (recognizer != null)
                    recognizer.Dispose();

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
                if (recognizer != null)
                    recognizer.Dispose();

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
            if (recognizer != null)
                recognizer.Dispose();

            if (e.Url == browser.Url)
            {
                Console.WriteLine("Carga completa: {0}", e.Url);
            }
        }

        private void SpeechRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result != null && e.Result.Text != null)
            {
                if (e.Result.Grammar.Name != "sysFragata")
                {
                    Console.WriteLine("{0} : {1} : {2}", e.Result.Text, e.Result.Grammar.Name, e.Result.Confidence.ToString());
                    returnRecognition(returnMethod, new object[3] { e.Result.Text, e.Result.Grammar.Name, e.Result.Confidence });
                }
                else
                {                    
                    switch (e.Result.Text)
                    {
                        case "Cerrar aplicación":
                            this.Close();
                            break;
                        case "Modo pantalla completa":
                            browser_Fullscreen();
                            break;
                        case "Recargar página":
                            browser.Refresh(WebBrowserRefreshOption.Completely);
                            break;
                    }

                    Console.WriteLine(e.Result.Text);
                }
            }
        }

        /// <summary>
        /// Métodos públicos disponibles mediante window.external
        /// </summary>

        public string textToSpeech(string text)
        {
            MemoryStream stream = new MemoryStream();
            using (synthesizer = new SpeechSynthesizer())
            {
                foreach (InstalledVoice voice in synthesizer.GetInstalledVoices())
                {
                    if (voice.VoiceInfo.Name.Contains("Sabina"))
                    {
                        synthesizer.SelectVoice(voice.VoiceInfo.Name);
                    }
                }

                synthesizer.Rate = 1;
                synthesizer.SetOutputToAudioStream(stream, new SpeechAudioFormatInfo(44100, AudioBitsPerSample.Sixteen, AudioChannel.Stereo));
                synthesizer.Speak(text);
                synthesizer.SetOutputToNull();
            }

            stream.Position = 0;

            return Convert.ToBase64String(stream.ToArray());
        }
        
        public void initRecognizer(string callback)
        {
            string[] commands = { "Cerrar aplicación", "Modo pantalla completa", "Recargar página" };
            GrammarBuilder gb = new GrammarBuilder(new Choices(commands));
            gb.Culture = new CultureInfo("es-MX");

            Grammar sysgrammar = new Grammar(gb);
            sysgrammar.Name = "sysFragata";

            recognizer = new SpeechRecognitionEngine("SR_MS_es-MX_TELE_11.0");
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognizedHandler);
            recognizer.SetInputToDefaultAudioDevice();
            recognizer.UnloadAllGrammars();
            recognizer.LoadGrammar(sysgrammar);

            returnMethod = callback;
        }
        
        public void startRecognition(bool start)
        {
            if (recognizer != null)
            {
                if (start)
                {
                    recognizer.RecognizeAsync(RecognizeMode.Multiple);
                }
                else
                {
                    recognizer.RecognizeAsyncCancel();
                }
            }
        }

        public void addChoices(string grammarName, string words)
        {
            GrammarBuilder gb = new GrammarBuilder(new Choices(words.Split(',')));
            gb.Culture = new CultureInfo("es-MX");
            Grammar grammar = new Grammar(gb);
            grammar.Name = grammarName;

            recognizer.LoadGrammar(grammar);

            Console.WriteLine("{0} : {1}", grammarName, words);
        }

        public void loadGrammar(string name, string words)
        {
            Choices choices = new Choices(words.Split(' '));
            GrammarBuilder gb = new GrammarBuilder(new GrammarBuilder(choices), 1, words.Split(' ').Length);
            gb.Culture = new CultureInfo("es-MX");

            Grammar grammar = new Grammar(gb);
            grammar.Name = name;

            recognizer.LoadGrammar(grammar);

            Console.WriteLine("{0} : {1}", name, words);
        }

        public void unloadGrammar()
        {
            List<Grammar> grammars = new List<Grammar>(recognizer.Grammars);
            foreach (Grammar g in grammars)
            {
                if (g.Name != "sysFragata")
                    recognizer.UnloadGrammar(g);
            }
        }
    }
}
