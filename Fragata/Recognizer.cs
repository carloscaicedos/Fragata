using Microsoft.Speech.Recognition;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace Fragata
{
    public class Recognizer : IDisposable
    {
        private struct Level
        {
            public int lower;
            public int upper;
            public Level(int i, int s)
            {
                lower = i;
                upper = s;
            }
        };

        private Boolean disposed;
        private SpeechRecognitionEngine recognizer;
        private Dictionary<string, int> dictNumbers;
        private Level[] levels;
        private Action<string> complete;

        public Recognizer(string _type, Action<string> _complete)
        {
            recognizer = new SpeechRecognitionEngine("SR_MS_es-ES_TELE_11.0");
            //recognizer = new SpeechRecognitionEngine("SR_MS_es-MX_TELE_11.0");

            complete = _complete;
            setGrammar(_type);
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
                    this.recognizer.Dispose();
                }
            }
            this.disposed = true;
        }

        ~Recognizer()
        {
            this.Dispose(false);
        }

        private void setGrammar(string type)
        {
            switch (type)
            {
                case "numeric":
                    string[] numbersNames = { "cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve",
                                     "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete",
                                     "dieciocho", "diecinueve", "veinte", "veinti", "treinta", "cuarenta", "cincuenta",
                                     "sesenta", "setenta", "ochenta", "noventa", "cien", "ciento", "doscientos", "trescientos",
                                     "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos","novecientos",
                                     "mil", "millón", "millones", "treinta y", "cuarenta y", "cincuenta y", "sesenta y", "setenta y",
                                     "ochenta y", "noventa y", };

                    levels = new Level[7];
                    levels[0] = new Level(0, 0);
                    levels[1] = new Level(1, 9);
                    levels[2] = new Level(10, 19);
                    levels[3] = new Level(20, 90);
                    levels[4] = new Level(100, 900);
                    levels[5] = new Level(1000, 1000);
                    levels[6] = new Level(1000000, 1000000);

                    dictNumbers = new Dictionary<string, int>();

                    for (int i = 0; i < 21; i++)
                        dictNumbers.Add(numbersNames[i], i);

                    dictNumbers.Add("un", 1);
                    dictNumbers.Add("veinti", 20);
                    dictNumbers.Add("treinta", 30);
                    dictNumbers.Add("cuarenta", 40);
                    dictNumbers.Add("cincuenta", 50);
                    dictNumbers.Add("setenta", 70);
                    dictNumbers.Add("sesenta", 60);
                    dictNumbers.Add("ochenta", 80);
                    dictNumbers.Add("noventa", 90);
                    dictNumbers.Add("cien", 100);
                    dictNumbers.Add("ciento", 100);
                    dictNumbers.Add("doscientos", 200);
                    dictNumbers.Add("trescientos", 300);
                    dictNumbers.Add("cuatrocientos", 400);
                    dictNumbers.Add("quinientos", 500);
                    dictNumbers.Add("seiscientos", 600);
                    dictNumbers.Add("setecientos", 700);
                    dictNumbers.Add("ochocientos", 800);
                    dictNumbers.Add("novecientos", 900);
                    dictNumbers.Add("mil", 1000);
                    dictNumbers.Add("millón", 1000000);
                    dictNumbers.Add("millones", 1000000);
                    dictNumbers.Add("y", -1);

                    Choices ordinals = new Choices(numbersNames);
                    GrammarBuilder gb = new GrammarBuilder(new GrammarBuilder(ordinals), 0, 50);
                    //gb.Culture = new CultureInfo("es-MX");
                    gb.Culture = new CultureInfo("es-ES");
                    Grammar grammar;
                    grammar = new Grammar(gb);
                    grammar.Name = "numeric";
                    grammar.Priority = 0;

                    Choices gchoices = new Choices();
                    gchoices.Add("Aceptar");
                    gchoices.Add("Borrar");

                    GrammarBuilder gbuilder = new GrammarBuilder(gchoices);
                    //gbuilder.Culture = new CultureInfo("es-MX");
                    gbuilder.Culture = new CultureInfo("es-ES");
                    Grammar global = new Grammar(gbuilder);
                    global.Name = "global";
                    global.Priority = 100;

                    recognizer.UnloadAllGrammars();
                    recognizer.LoadGrammar(grammar);
                    recognizer.LoadGrammar(global);
                    recognizer.SetInputToDefaultAudioDevice();
                    recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechNumbersRecognizedHandler);

                    break;
            }
        }

        public void startRecognition(bool start)
        {
            if (start)
            {
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
            }
            else
            {
                if (recognizer != null)
                    recognizer.RecognizeAsyncCancel();
            }
        }

        private void SpeechNumbersRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result != null && e.Result.Text != null)
            {
                Console.WriteLine(e.Result.Text);
                switch (e.Result.Grammar.Name)
                {
                    case "global":
                        complete(e.Result.Text);
                        break;
                    case "numeric":
                        complete(TextToNumbers(e.Result.Words));
                        break;
                }
            }
        }

        private string TextToNumbers(ReadOnlyCollection<RecognizedWordUnit> words)
        {
            string output = "";
            int number = -1;
            int curr_level = -1;
            int prev_level = -1;
            int[] numbers = new int[words.Count];

            for (int i = 0; i < words.Count; i++)
                numbers[i] = dictNumbers[words[i].Text.ToLower()];

            for (int i = 0; i < numbers.Length; i++)
            {
                if (numbers[i] >= 0)
                {
                    if (curr_level < 0)
                    {
                        number = numbers[i];
                        curr_level = GetLevel(numbers[i]);
                    }
                    else
                    {
                        prev_level = curr_level;
                        curr_level = GetLevel(numbers[i]);

                        if (curr_level < prev_level && prev_level != 0 && curr_level != 0 && (prev_level > 3 || (prev_level == 3 && curr_level != 2)))
                        {
                            if (i > 0 && curr_level == 1 && prev_level == 3 && numbers[i - 1] != -1 && words[i - 1].Text != "veinti")
                            {
                                output += number;
                                number = numbers[i];
                            }
                            else
                            {
                                number += numbers[i];
                            }
                        }
                        else if (curr_level > prev_level && curr_level > 4)
                        {
                            if (numbers[i] == 1000000)
                            {
                                for (int j = i + 1; j < numbers.Length; j++)
                                    if (numbers[j] == 1000)
                                    {
                                        numbers[i] /= 1000;
                                        break;
                                    }
                            }
                            number *= numbers[i];
                        }
                        else
                        {
                            output += number;
                            number = numbers[i];
                        }
                    }
                }
            }

            if (number >= 0)
                output += number;
            return output;
        }

        private int GetLevel(int x)
        {
            for (int i = 0; i < levels.Length; i++)
                if (levels[i].lower <= x && x <= levels[i].upper)
                    return i;

            return -1;
        }
    }
}
