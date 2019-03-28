using System;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using System.Globalization;
using Microsoft.Speech.Recognition;
using System.Reflection;


namespace SkidanVoiceHelper
{
    public partial class Form1 : Form
    {
        private CultureInfo _culture;
        private SpeechRecognitionEngine _sre;


        public Form1()
        {
            InitializeComponent();
            notifyIcon1 = new NotifyIcon();

            notifyIcon1.Visible = true;
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetAutorunValue(true, Assembly.GetExecutingAssembly().Location);

            try
            {
                _culture = new CultureInfo("ru-RU");

                _sre = new SpeechRecognitionEngine(_culture);

                // Setup event handlers
                _sre.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(sr_SpeechDetected);
                _sre.RecognizeCompleted += new EventHandler<RecognizeCompletedEventArgs>(sr_RecognizeCompleted);
                _sre.SpeechHypothesized += new EventHandler<SpeechHypothesizedEventArgs>(sr_SpeechHypothesized);
                _sre.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(sr_SpeechRecognitionRejected);
                _sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sr_SpeechRecognized);

                // select input source
                _sre.SetInputToDefaultAudioDevice();

                // load grammar
                _sre.LoadGrammar(CreateSampleGrammar1());
                _sre.LoadGrammar(CreateSampleGrammar2());

                // start recognition
                _sre.RecognizeAsync(RecognizeMode.Multiple);

            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }
        
        private Choices CreateSampleChoices()
        {
            var val1 = new SemanticResultValue("калькулятор", "calc");
            var val2 = new SemanticResultValue("проводник", "explorer");
            var val3 = new SemanticResultValue("блокнот", "notepad");
            var val4 = new SemanticResultValue("пэйнт", "mspaint");
            var va15 = new SemanticResultValue("командную строку", "cmd");
            var va16 = new SemanticResultValue("яндекс", "https://yandex.ru");
            var va17 = new SemanticResultValue("гугл", "www.google.ru");
            var va18 = new SemanticResultValue("вконтакте", "https://vk.com/");
            var va19 = new SemanticResultValue("однокласники", "https://ok.ru");
            var va20 = new SemanticResultValue("ютуп", "www.youtube.com");
            var va21 = new SemanticResultValue("ворд", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk");
            var va22 = new SemanticResultValue("эксэль", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk");
            var va23 = new SemanticResultValue("поврпоинт", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk");
            var va24 = new SemanticResultValue("эксес", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Access 2010.lnk");
            var va25 = new SemanticResultValue("подключение к удаленному рабочему столу", "mstsc");
            var va26 = new SemanticResultValue("проигрыватель", "wmplayer");
            var va27 = new SemanticResultValue("таблицу символов", "charmap");
            var va28 = new SemanticResultValue("экранную лупу", "Magnify");
            var va29 = new SemanticResultValue("диспетчер задач", "Taskmgr");
            var va30 = new SemanticResultValue("редактор реестра", "regedit");
            var va31 = new SemanticResultValue("визуальную студию", "devenv");
            var va32 = new SemanticResultValue("матрицу", @"C:\Users\Acer\Desktop\Матрица\матрица.bat");
            var va33 = new SemanticResultValue("доту", @"steam://rungameid/570");

            return new Choices(val1, val2, val3, val4, va15, va16, va17, va18, va19, va20, va21, va22, va23, va24, va25, va26, va27, va28, va29, va30, va31, va32, va33);
        }
        
        private Grammar CreateSampleGrammar1()
        {
            var programs = CreateSampleChoices();

            var grammarBuilder = new GrammarBuilder("запустить", SubsetMatchingMode.SubsequenceContentRequired);
            grammarBuilder.Culture = _culture;
            grammarBuilder.Append(new SemanticResultKey("start", programs));

            return new Grammar(grammarBuilder);
            
            
        }

        private Grammar CreateSampleGrammar2()
        {
            var programs = CreateSampleChoices();

            var grammarBuilder = new GrammarBuilder("закрыть", SubsetMatchingMode.SubsequenceContentRequired);
            grammarBuilder.Culture = _culture;
            grammarBuilder.Append(new SemanticResultKey("close", programs));

            return new Grammar(grammarBuilder);

        }
       

        private void sr_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            AppendLine("Speech Recognition Rejected: " + e.Result.Text);
        }

        private void sr_SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            AppendLine("Speech Hypothesized: " + e.Result.Text + " (" + e.Result.Confidence + ")");
        }

        private void sr_RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            AppendLine("Recognize Completed: " + e.Result.Text);
        }

        private void sr_SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {
            AppendLine("Speech Detected: audio pos " + e.AudioPosition);
        }

        private void sr_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            AppendLine("\t" + "Speech Recognized:");

            AppendLine(e.Result.Text + " (" + e.Result.Confidence + ")");

            if (e.Result.Confidence < 0.1f)
                return;

            for (var i = 0; i < e.Result.Alternates.Count; ++i)
            {
                AppendLine("\t" + "Alternate: " + e.Result.Alternates[i].Text + " (" + e.Result.Alternates[i].Confidence + ")");
            }

            for (var i = 0; i < e.Result.Words.Count; ++i)
            {
                AppendLine("\t" + "Word: " + e.Result.Words[i].Text + " (" + e.Result.Words[i].Confidence + ")");

                if (e.Result.Words[i].Confidence < 0.1f)
                    return;
            }

            foreach (var s in e.Result.Semantics)
            {
                var program = (string)s.Value.Value;

                switch (s.Key)
                {
                    case "start":
                        Process.Start(program);
                        break;
                    case "close":
                        var p = Process.GetProcessesByName(program);
                        if (p.Length > 0)
                            p[0].Kill();
                        break;
                }
            }
        }

        private void AppendLine(string text)
        {
            richTextBox1.AppendText(text + Environment.NewLine);
            richTextBox1.ScrollToCaret();
        }



        
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private bool closing = true;

       

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = closing;
            this.Hide();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            closing = false;
            Application.Exit();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Для того чтобы начать работать с программой необходимо знать 2 ключевых слова «Запустить» для запуска какой либо программы и «Закрыть» для закрытия программы. Например «Запустить Яндекс» или «Закрыть Яндекс».", "О программе");
        }
//автозагрузка
        public bool SetAutorunValue(bool autorun, string path)
        {
            const string name = "Вова";
            string ExePath = path;
            RegistryKey reg;
            reg = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run\\");
            try
            {
                if (autorun)
                    reg.SetValue(name, ExePath);
                else
                    reg.DeleteValue(name);

                reg.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }

        private void circularProgressBar1_Click(object sender, EventArgs e)
        {

        }


        //
    }
}
        