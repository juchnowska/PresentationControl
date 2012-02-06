using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Threading;
using System.IO;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Research.Kinect.Audio;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.AudioFormat;

namespace PresentationControl
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Thread t;
        private const string RecognizerId = "SR_MS_en-US_Kinect_10.0";
        bool shouldListen = true;
        
        private string wordStart, wordStop, wordPowerpoint, wordFullscreen, wordExit, wordNext, wordPrevious, wordHide, wordShow;
        private KinectAudioSource source;
        private SpeechRecognitionEngine sre;
        private Stream stream;
        bool isControlling;
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += new RoutedEventHandler(Window_Loaded);
            this.Closed += new EventHandler(Window_Closed);
        }

        private void CaptureAudio() 
        {
            this.source = new KinectAudioSource();
            this.source.FeatureMode = true;
            this.source.NoiseSuppression = true;
            this.source.AutomaticGainControl = false;
            this.source.SystemMode = SystemMode.OptibeamArrayOnly;
            RecognizerInfo ri = SpeechRecognitionEngine.InstalledRecognizers().
                Where(r => r.Id == RecognizerId).FirstOrDefault();
            if (ri == null)
            {
                return;
            }
            this.sre = new SpeechRecognitionEngine(ri.Id);

            var words = new Choices();
            
            words.Add(this.wordPowerpoint);
            words.Add(this.wordStart);
            words.Add(this.wordStop);
            words.Add(this.wordNext);
            words.Add(this.wordPrevious);
            words.Add(this.wordFullscreen);
            words.Add(this.wordExit);
            words.Add(this.wordHide);
            words.Add(this.wordShow);

            var gb = new GrammarBuilder();
            
            gb.Culture = ri.Culture;
            gb.Append(words);
            var g = new Grammar(gb);
            sre.LoadGrammar(g);
            this.sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
            this.stream = this.source.Start();
            this.sre.SetInputToAudioStream(this.stream, new SpeechAudioFormatInfo(
                EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            
            while (shouldListen)
            {
                sre.Recognize();
            }
            
        }

        private void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //jeżeli jest odpowiednio dokładnie rozpoznane słowo
            if (e.Result.Confidence > 0.97)
            {
                if (e.Result.Text == this.wordStart)
                {
                    this.StartListening();
                }

                if (this.isControlling)
                {

                    //zdobycie nazwy procesu aktywnej aplikacji
                    string appProcessName = ProcessHandling.GetActiveWindowProcessName();
                    //rozpoznawanie nawet gdy okno powerpointa nie jest aktywne
                    if (e.Result.Text == this.wordPowerpoint)
                    {
                        this.Dispatcher.BeginInvoke((Action)delegate
                        {
                            this.textBlock2.Text = "Recognized command: " + this.wordPowerpoint;
                        });
                        
                        if (appProcessName.CompareTo("POWERPNT") != 0)
                        {
                            ProcessHandling.SetActiveWindow("POWERPNT");
                        }
                    }
                    else if (e.Result.Text == this.wordHide)
                    {
                        this.Dispatcher.BeginInvoke((Action)delegate
                        {
                            this.textBlock2.Text = "Recognized command: " + this.wordHide;
                            this.WindowState = System.Windows.WindowState.Minimized;
                        });
                    }
                    else if (e.Result.Text == this.wordShow)
                    {
                        this.Dispatcher.BeginInvoke((Action)delegate
                        {
                            this.textBlock2.Text = "Recognized command: " + this.wordShow;
                            this.WindowState = System.Windows.WindowState.Normal;
                        });
                    }
                    else if (e.Result.Text == this.wordStop)
                    {
                        this.StopListening();
                    }
                    //aktywne okno to powerpoint
                    if (appProcessName.CompareTo("POWERPNT") == 0)
                    {
                        if (this.isControlling == true)
                        {
                            if(e.Result.Text == this.wordNext)
                            {
                                CommandRecognized(this.wordNext, "{Right}");
                            }
                            else if(e.Result.Text == this.wordPrevious)
                            {
                                CommandRecognized(this.wordPrevious, "{Left}");
                            }
                            else if(e.Result.Text == this.wordFullscreen)
                            {
                                CommandRecognized(this.wordFullscreen, "{F5}");
                            }
                            else if(e.Result.Text == this.wordExit)
                            {
                                CommandRecognized(this.wordExit, "{ESC}");
                            }
                            else
                            {
                                this.Dispatcher.BeginInvoke((Action)delegate
                                {
                                    textBlock2.Text = "";
                                });
                            }
                        }
                    }
                }
            }
        }

        private void StartListening()
        {
            this.isControlling = true;
            this.Dispatcher.BeginInvoke((Action)delegate
            {

               this.textBlock1.Text = "Kinect is listening.\nSay \"" + this.wordStop + "\" to stop listening.";
        
            });  
        }

        private void StopListening()
        {
            this.isControlling = false;
            this.Dispatcher.BeginInvoke((Action)delegate
            {

                this.textBlock1.Text = "Kinect is not listening.\nSay \"" + this.wordStart + "\" to start listening.";

            });  
        }

        private void CommandRecognized(string word, string key)
        {
            this.Dispatcher.BeginInvoke((Action)delegate
            {
                this.textBlock2.Text = "Recognized Command: " + word;
                System.Windows.Forms.SendKeys.SendWait(key);
                System.Windows.Forms.SendKeys.Flush();
            });
        }
         
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
            this.textBox1.Text = this.wordStart = "start";
            this.textBox2.Text = this.wordStop = "stop";
            this.textBox3.Text = this.wordPowerpoint = "powerpoint";
            this.textBox4.Text = this.wordFullscreen = "full screen";
            this.textBox5.Text = this.wordExit = "exit";
            this.textBox6.Text = this.wordNext = "next";
            this.textBox7.Text = this.wordPrevious = "previous";
            this.textBox8.Text = this.wordHide = "hide";
            this.textBox9.Text = this.wordShow = "show";
            this.StopListening();
            StartSpeechRecognition();
        }

        private void StartSpeechRecognition()
        {
            t = new Thread(new ThreadStart(CaptureAudio));
            t.SetApartmentState(ApartmentState.MTA);
            t.Start();
        }

        private void ApplySettings(object sender, RoutedEventArgs e)
        {
            t.Abort();
            if (!this.ApplyWord(ref this.textBox1, ref this.wordStart)) return;
            if (!this.ApplyWord(ref this.textBox2, ref this.wordStop)) return;
            if (!this.ApplyWord(ref this.textBox3, ref this.wordPowerpoint)) return;
            if (!this.ApplyWord(ref this.textBox4, ref this.wordFullscreen)) return;
            if (!this.ApplyWord(ref this.textBox5, ref this.wordExit)) return;
            if (!this.ApplyWord(ref this.textBox6, ref this.wordNext)) return;
            if (!this.ApplyWord(ref this.textBox7, ref this.wordPrevious)) return;
            if (!this.ApplyWord(ref this.textBox8, ref this.wordHide)) return;
            if (!this.ApplyWord(ref this.textBox9, ref this.wordShow)) return;
            if (this.isControlling)
            {
                this.StartListening();
            }
            else
            {
                this.StopListening();
            }
            StartSpeechRecognition();
            MessageBox.Show("Changes applied.");
        }

        private bool ApplyWord(ref TextBox textBox, ref string word)
        {
            if (textBox.Text.Length == 0)
            {
                MessageBox.Show("Word " + word + " cannot be empty!");
                return false;
            }
            else
            {

                word = textBox.Text;
                return true;
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            this.isControlling = false;
        }
    }
}

