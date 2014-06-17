using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Fragata
{
    public delegate object RunScript(string method, object[] args);
    
    public class CuadroAtencion : IDisposable
    {
        private Boolean disposed;
        private XmlDocument menuXML;
        private XmlNode stepNode;
        private XmlNode optionNode;
        private XmlNodeList optionNodes;
        private XmlNodeList choiceNodes;
        private SpeechRecognitionEngine recognizer;
        private Recognizer recogNumber;
        private SpeechSynthesizer synth;
        private Timer timerIdle;
        private int timerCount = 0;
        public RunScript runScript;

        public CuadroAtencion(XmlDocument menu)
        {
            menuXML = menu;
            stepNode = menuXML.DocumentElement.SelectSingleNode("//step[@id='standby']");
            optionNodes = stepNode.SelectNodes(".//option");

            if (recognizer != null) 
                recognizer.Dispose();

            if (recogNumber != null)
                recogNumber.Dispose();

            recognizer = new SpeechRecognitionEngine("SR_MS_es-MX_TELE_11.0");
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognizedHandler);
            recognizer.SetInputToDefaultAudioDevice();

            timerIdle = new Timer();
            timerIdle.Tick += new EventHandler(RestartApp);
            timerIdle.Interval = 30000;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (this.recognizer != null)
                        this.recognizer.Dispose();

                    if (this.recogNumber != null)
                        this.recogNumber.Dispose();

                    if (this.timerIdle != null)
                    {
                        this.timerCount = 0;
                        this.timerIdle.Dispose();
                    }
                }
            }
            this.disposed = true;
        }

        ~CuadroAtencion()
        {
            this.Dispose(false);
        }

        private void SpeechRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence > 0.60)
            {
                String dirPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!dirPath.EndsWith("\\")) dirPath += "\\";

                using (StreamWriter writer = new StreamWriter(dirPath + "Fragata.log", true, Encoding.Default))
                {
                    writer.WriteLine(e.Result.Text + " : " + e.Result.Confidence + " : " + DateTime.Now);
                }

                timerIdle.Stop();
                timerCount = 0;

                foreach (XmlNode node in optionNodes)
                {
                    choiceNodes = node.SelectNodes(".//choice");
                    foreach (XmlNode choice in choiceNodes)
                    {
                        if (e.Result.Text == choice.InnerText)
                        {
                            recognizer.RecognizeAsyncCancel();
                            runScript("eval", new object[1] { "(function (){ $('.micOn').addClass('micOff').removeClass('micOn'); })()" });
                            optionNode = node;
                            executeAction(optionNode.Attributes["action"].InnerText);
                            return;
                        }
                    }
                }
            }
        }

        private void RestartApp(Object myObject, EventArgs myEventArgs)
        {
            object[] args;
            timerIdle.Stop();

            if (stepNode.Attributes["id"].InnerText == "menu")
            {
                if (recognizer != null)
                    recognizer.RecognizeAsyncCancel();

                args = new object[2];
                args[0] = false;
                args[1] = runScript("eval", new object[1] { "(function (){ window.external.startRecognition(true); })" });

                runScript("openDoors", args);
                stepNode = menuXML.DocumentElement.SelectSingleNode("//step[@id='standby']");
            }
            else
            {
                timerCount++;
                if (timerCount > 1)
                {
                    timerCount = 0;

                    if (recognizer != null)
                        recognizer.RecognizeAsyncCancel();

                    if (recogNumber != null)
                    {
                        recogNumber.startRecognition(false);
                    }

                    args = new object[2];
                    args[0] = false;
                    args[1] = runScript("eval", new object[1] { "(function (){ window.external.startRecognition(true); })" });
                    runScript("openDoors", args);
                    stepNode = menuXML.DocumentElement.SelectSingleNode("//step[@id='standby']");
                }
                else
                {
                    args = new object[2];
                    args[0] = textToSpeech(stepNode.Attributes["idleMessage"].InnerText);
                    args[1] = runScript("eval", new object[1] { "(function (){ window.external.restartTimerIdle(); })" });
                    runScript("speakSophia", args);
                }
            }
        }

        private void executeAction(string action)
        {
            XmlNodeList messageNodes;
            Random rnd = new Random();
            object[] args;
            string message;
            string message_alt;
            string script;

            stepNode = menuXML.DocumentElement.SelectSingleNode("//step[@id='" + optionNode.Attributes["nextStep"].InnerText + "']");
            messageNodes = optionNode.SelectNodes(".//message[@type='goodbye']");
            message = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));

            switch (action)
            {
                case "showMenu":
                    string options = "";
                    string idxs = "";

                    messageNodes = stepNode.SelectNodes(".//message[@type='greeting']");
                    message = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));
                    optionNodes = stepNode.SelectNodes(".//option");

                    foreach (XmlNode node in optionNodes)
                    {
                        options += node.Attributes["label"].InnerText + "|";
                        idxs += node.Attributes["id"].InnerText + "|";
                    }
                    options = options.Remove(options.Length - 1);
                    idxs = idxs.Remove(idxs.Length - 1);

                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        showOptions('" + options + "','" + idxs + @"', function () {
                                            window.external.startRecognition(true);
                                            $('.micOff').addClass('micOn').removeClass('micOff');
                                            $('#tone')[0].play();
                                            window.external.restartTimerIdle();
                                        });
                                    });
                                })";

                    if (optionNode.Attributes["id"].InnerText == "init")
                    {
                        args = new object[1];
                        args[0] = runScript("eval", new object[1] { script });
                        runScript("showAvatar", args);
                    }
                    else
                    {
                        args = new object[2];
                        args[0] = optionNode.Attributes["id"].InnerText;
                        args[1] = runScript("eval", new object[1] { script });
                        runScript("selectOption", args);
                    }
                    break;
                case "showNews":
                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    messageNodes = optionNode.SelectNodes(".//message[@type='alt_goodbye']");
                    message_alt = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));

                    script = @"(function () {
                                    requestNews('" + optionNode.Attributes["params"].InnerText + "','" + message + "','" + message_alt + @"',function () { window.external.startRecognition(true); });
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectOption", args);

                    break;
                case "changeArtist":
                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    messageNodes = optionNode.SelectNodes(".//message[@type='alt_goodbye']");
                    message_alt = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));
                    
                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        showArtists(function () {
                                            window.external.startRecognition(true);
                                            $('.micOff').addClass('micOn').removeClass('micOff');
                                            $('#tone')[0].play();
                                            window.external.restartTimerIdle();
                                        });
                                    }); 
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectOption", args);

                    break;
                case "selectArtist":
                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        openDoors(false, function () {
                                            window.external.startRecognition(true);
                                            showArt(true);
                                        });
                                    }); 
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectArtist", args);

                    break;
                case "showSchedule":
                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    messageNodes = optionNode.SelectNodes(".//message[@type='alt_goodbye']");
                    message_alt = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));

                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        showSophia(false, function () {
                                            showSchedule(function (){
                                                openDoors(false, function () {
                                                    window.external.startRecognition(true)
                                                });
                                            });
                                        });
                                    });
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectOption", args);
                    break;
                case "getDoc":
                    recogNumber = new Recognizer("numeric", delegate(string value)
                    {
                        object[] param = new object[1];
                        runScript("eval", new object[1] { "(function (){ docInput.setValue('" + value + "'); })()" });
                    });

                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    message = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));
                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        getDoc(function (){
                                            window.external.getDocComplete();
                                        });
                                        window.external.startDocRecognition(true);
                                        $('.micOff').addClass('micOn').removeClass('micOff');
                                        $('#tone')[0].play();
                                        window.external.restartTimerIdle();
                                    }); 
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectOption", args);
                    break;
                case "requestStatus":
                    args = new object[1];
                    args[0] = runScript("eval", new object[1] { "(function () { window.external.startRecognition(true); })" });
                    runScript("requestStatus", args);
                    break;
                case "requestInterview":
                    args = new object[1];
                    args[0] = runScript("eval", new object[1] { "(function () { window.external.startRecognition(true); })" });
                    runScript("requestInterview", args);
                    break;
                case "requestInfoUCI":
                    args = new object[2];
                    args[0] = optionNode.Attributes["id"].InnerText;

                    messageNodes = optionNode.SelectNodes(".//message[@type='alt_goodbye']");
                    message_alt = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));

                    script = @"(function () {
                                    requestInfoUCI('" + message + "','" + message_alt + @"',function () { window.external.startRecognition(true); });
                                })";
                    args[1] = runScript("eval", new object[1] { script });

                    runScript("selectOption", args);
                    break;
                case "requestLiquidation":
                    args = new object[1];
                    script = @"(function (result) {
                                    window.external.evalLiquidation(result);
                                })";
                    args[0] = runScript("eval", new object[1] { script });
                    runScript("requestLiquidation", args);
                    break;
                case "requestPayment":
                    args = new object[1];

                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        window.external.startRecognition(true);
                                        $('.micOff').addClass('micOn').removeClass('micOff');
                                        $('#tone')[0].play();
                                        window.external.restartTimerIdle();
                                    });
                                })";
                    args[0] = runScript("eval", new object[1] { script });
                    runScript("requestPayment", args);
                    break;
                case "requestPaymentMethod":
                    args = new object[2];

                    script = @"(function () {
                                    speakSophia('" + message + @"', function () {
                                        window.external.startRecognition(true);
                                        $('.micOff').addClass('micOn').removeClass('micOff');
                                        $('#tone')[0].play();
                                        window.external.restartTimerIdle();
                                    });
                                })";
                    args[0] = optionNode.Attributes["id"].InnerText;
                    args[1] = runScript("eval", new object[1] { script });
                    runScript("requestPaymentMethod", args);
                    break;
                case "requestTicket":
                    args = new object[4];

                    script = @"(function () {
                                    setTimeout(function () {
                                        showSophia(false,function () {
                                            openDoors(false, function () {
                                                window.external.startRecognition(true)
                                            });
                                        });
                                    }, 5000);
                                })";
                    messageNodes = optionNode.SelectNodes(".//message[@type='alt_goodbye']");
                    message_alt = textToSpeech(messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting()));

                    args[0] = optionNode.Attributes["id"].InnerText;
                    args[1] = message;
                    args[2] = message_alt;
                    args[3] = runScript("eval", new object[1] { script });
                    runScript("requestTicket", args);

                    break;
                case "requestAdvice":
                    args = new object[2];
                    
                    script = @"(function () {
                                    speakSophia('" + message + @"', function (){
                                        setTimeout(function () {
                                            showSophia(false,function () {
                                                openDoors(false, function () {
                                                    window.external.startRecognition(true)
                                                });
                                            });
                                        }, 5000);
                                    });
                                })";
                    args[0] = messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting());
                    args[1] = runScript("eval", new object[1] { script });
                    runScript("requestAdvice", args);
                    break;
                case "noDebt":
                    args = new object[2];
                    
                    script = @"(function () {
                                    speakSophia('" + message + @"', function (){
                                        setTimeout(function () {
                                            showSophia(false,function () {
                                                openDoors(false, function () {
                                                    window.external.startRecognition(true)
                                                });
                                            });
                                        }, 5000);
                                    });
                                })";
                    args[0] = messageNodes[rnd.Next(messageNodes.Count)].InnerText.Replace("{welcome_greeting}", getWelcomeGreeting());
                    args[1] = runScript("eval", new object[1] { script });
                    runScript("displayMessage", args);
                    break;
                case "cancelTransaction":
                    args = new object[1];

                    script = @"(function () {
                                    speakSophia('" + message + @"', function (){
                                        showSophia(false, function () {
                                            openDoors(false, function () {
                                                window.external.startRecognition(true)
                                            });
                                        });
                                    });
                                })";
                    args[0] = runScript("eval", new object[1] { script });
                    runScript("cancelTransaction", args);

                    break;
                case "showCarrousel":
                    args = new object[1];

                    script = @"";
                    args[0] = runScript("eval", new object[1] { script });
                    runScript("showCarrousel", args);

                    break;
            }
        }

        private string getWelcomeGreeting()
        {
            string message;

            if (DateTime.Now.Hour < 12)
                message = "Buenos días";
            else if ((DateTime.Now.Hour == 18 && DateTime.Now.Minute > 30) || DateTime.Now.Hour > 18)
            {
                message = "Buenas noches";
            }
            else
                message = "Buenas tardes";

            return message;
        }

        private string textToSpeech(string text)
        {
            MemoryStream stream = new MemoryStream();
            using (synth = new SpeechSynthesizer())
            {
                foreach (InstalledVoice voice in synth.GetInstalledVoices())
                {
                    if (voice.VoiceInfo.Name.Contains("Sabina"))
                    {
                        synth.SelectVoice(voice.VoiceInfo.Name);
                    }
                }
                synth.Rate = 1;
                synth.SetOutputToAudioStream(stream, new SpeechAudioFormatInfo(44100, AudioBitsPerSample.Sixteen, AudioChannel.Stereo));
                synth.Speak(text);
                synth.SetOutputToNull();
            }

            stream.Position = 0;

            return Convert.ToBase64String(stream.ToArray());
        }
        
        public void StartRecognition(bool start)
        {
            if (start)
            {
                ArrayList optionList = new ArrayList();
                optionNodes = stepNode.SelectNodes(".//option");

                for (int i = 0; i < optionNodes.Count; i++)
                {
                    choiceNodes = optionNodes[i].SelectNodes(".//choice");
                    for (int n = 0; n < choiceNodes.Count; n++)
                    {
                        optionList.Add(choiceNodes[n].InnerText);
                    }
                }

                GrammarBuilder gb_spa = new GrammarBuilder(new Choices((String[])optionList.ToArray(typeof(string))));
                gb_spa.Culture = new CultureInfo("es-MX");
                Grammar grammar = new Grammar(gb_spa);

                recognizer.UnloadAllGrammars();
                recognizer.LoadGrammar(grammar);
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
            }
            else
            {
                recognizer.RecognizeAsyncCancel();
            }
        }

        public void StartDocRecognition(bool start)
        {
            recogNumber.startRecognition(start);

            if (start)
            {
                runScript("eval", new object[1] { "(function (){ $('.micOff').addClass('micOn').removeClass('micOff'); $('#tone')[0].play(); })()" });
            }
            else
            {
                runScript("eval", new object[1] { "(function (){ $('.micOn').addClass('micOff').removeClass('micOn'); })()" });
            }
        }

        public void RestartTimerIdle()
        {
            timerIdle.Stop();
            timerIdle.Start();
        }

        public void GetDocComplete()
        {
            recogNumber.startRecognition(false);
            runScript("eval", new object[1] { "(function (){ $('.micOn').addClass('micOff').removeClass('micOn'); })()" });

            timerIdle.Stop();
            
            optionNode = stepNode.SelectSingleNode(".//options/option[1]");
            executeAction(optionNode.Attributes["action"].InnerText);
        }

        public void EvalLiquidation(string result)
        {
            optionNode = stepNode.SelectSingleNode(".//options/option[@id='" + result + "']");
            executeAction(optionNode.Attributes["action"].InnerText);
        }
    }
}
