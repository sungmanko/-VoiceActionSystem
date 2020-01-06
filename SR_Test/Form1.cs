using System;
using System.Globalization;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;

namespace SR_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void initRS()
        {
            try
            {
                SpeechRecognitionEngine sre = new SpeechRecognitionEngine(new CultureInfo("ja-JP"));

                Grammar g = new Grammar("input.xml");
                sre.LoadGrammar(g);

                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception e)
            {
                label1.Text = "init RS Error : " + e.ToString();
            }
        }

        SpeechSynthesizer tts;

        public void initTTS()
        {
            try
            {
                tts = new SpeechSynthesizer();
                tts.SelectVoice("Microsoft Server Speech Text to Speech Voice (ja-JP, Haruka)");
                tts.SetOutputToDefaultAudioDevice();
                tts.Volume = 100;
            }
            catch (Exception e)
            {
                label1.Text = "init TTS Error : " + e.ToString();
            }
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            label1.Text = e.Result.Text;

            switch (e.Result.Text)
            {
                case "お疲れ様でした":
                    tts.SpeakAsync("お疲れ様でした。またのご利用を心よりお待ちしております。");
                    Application.Exit();
                    break;

                case "おはよう":
                    tts.SpeakAsync("おはようございます。今日もいちにち頑張りましょう。");
                    break;

                case "電卓 起動":
                    tts.SpeakAsync("電卓を起動します。");
                    doProgram("c:\\windows\\system32\\calc.exe", "");
                    break;

                case "電卓 終了":
                    {
                        tts.SpeakAsync("電卓を終了します。");
                        closeProcess("calc");
                        break;
                    }

                case "メモ長 起動":
                    {
                        tts.SpeakAsync("ノートパットを起動します。");
                        doProgram("c:\\windows\\system32\\notepad.exe", "");
                        break;
                    }

                case "メモ長 終了":
                    {
                        tts.SpeakAsync("ノートパットを終了します。");
                        closeProcess("notepad");
                        break;
                    }

                case "FAQ":
                    {
                        tts.SpeakAsync("コンタクトのエフエーキューページを起動します。");
                        Process.Start("https://support.rakuten-card.jp/faq/show/43543?site_domain=contact");
                        break;
                    }
            }
        }

        /// <summary>
        /// プロセス実行
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="arg"></param>
        private static void doProgram(string filename, string arg)
        {
            ProcessStartInfo psi;
            if (arg.Length != 0)
                psi = new ProcessStartInfo(filename, arg);
            else
                psi = new ProcessStartInfo(filename);
            Process.Start(psi);
        }

        /// <summary>
        /// プロセス終了
        /// </summary>
        /// <param name="filename"></param>
        private static void closeProcess(string filename)
        {
            Process[] myProcesses;
            // Returns array containing all instances of Notepad.
            myProcesses = Process.GetProcessesByName(filename);
            foreach (Process myProcess in myProcesses)
            {
                myProcess.CloseMainWindow();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            initRS();
            initTTS();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            initRS2();
            initTTS();
        }

        public void initRS2()
        {
            try
            {
                SpeechRecognitionEngine sre = new SpeechRecognitionEngine(new CultureInfo("ja-JP"));

                Grammar g = new Grammar("input.xml");
                sre.LoadGrammar(g);

                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized2);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception e)
            {
                label1.Text = "init RS Error : " + e.ToString();
            }
        }

        void sre_SpeechRecognized2(object sender, SpeechRecognizedEventArgs e)
        {
            textBox1.Text = e.Result.Text;
        }
    }
}
