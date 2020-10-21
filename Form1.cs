using AudioSwitcher.AudioApi.CoreAudio;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Media;
using Payloads;
using System.IO;
using System.Threading;
using System.Net;
using NAudio;
using NAudio.Wave;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using System.Security.Cryptography;
using Telegram.Bot;
using System.Net.NetworkInformation;

namespace RemotePresentationManager
{

    /*
     * This project is made by Adryzz and it is under GNU Gpl v3 or newer version (at your own choice)
     * You can find the license at (https://github.com/adryzz/RemotePresentationManager/LICENSE)
     * This project IS open source and will remain open-source.
     * Feel free to take some code snippets from this project (the code here needs to be improved), (link the source of that code in your project)
     * You can find this project and compiled binaries (x86 and x64) here: https://github.com/adryzz/RemotePresentationManager
     * Hope this will be helpful
     * Feel free to contact me for anything related to this at my discord (Adryzz#7264)
     * (do NOT remove or edit this comment in any way if you are distributing this file or part of it)
     */

    public partial class Form1 : Form
    {
        public bool Connected = false;
        public bool HideWindow = false;//this hides the window when it is shown
        public SoundPlayer Player;
        public WaveOut WaveOutDevice = new WaveOut();
        public AudioFileReader AudioFileReader;
        public CoreAudioDevice DefaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;
        public Process cmd = new Process();
        public StreamWriter CmdStream;
        public int remaining = 5;//remaining tries
        public string title = "";//title of all messageboxes
        public bool loop = false;//audio loop
        string RegistryKey = @"HKEY_CURRENT_USER\Software\RemotePresentationManager\Passwords";
        string ValueName = "Key";
        public string CurrentPass = null;
        public TelegramServer TelegramBot;
        public SerialServer Serial;
        public bool IsCPURunning = false;
        public int CPUPercentage = 100;

        public Form1(string[] args)
        {
            InitializeComponent();
            WaveOutDevice.PlaybackStopped += (s, e) =>
            {
                if (loop && e.Exception == null)
                {
                    WaveOutDevice.Play();
                }
            };
            if (args.Length > 0)
                if (args[0].Equals("WEBONLY"))
                    HideWindow = true;

            string token = GetBotToken();
            if (token != null)
                TelegramBot = new TelegramServer(token, GetPasswordHash());
        }

        private string GetPasswordHash()
        {
            try
            {
                object o = Registry.GetValue(RegistryKey, ValueName, null);

                if (o == null)
                {
                    return null;
                }
                else
                {
                    return o.ToString();
                }
            }
            catch
            {
                return null;
            }
        }

        private string GetBotToken()
        {
            if (File.Exists("token.txt"))
            {
                return File.ReadAllText("token.txt");
            }
            return null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (string s in SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(s);
            }
            try
            {
                object o = Registry.GetValue(RegistryKey, ValueName, null);
                if (o == null)
                {

                }
                else
                {
                    CurrentPass = o.ToString();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(this, ex.Message, "RPMPasswordSet", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
            }

            if (Environment.GetCommandLineArgs().Length > 1 && !Environment.GetCommandLineArgs()[1].Equals("WEBONLY"))
            {
                Serial = new SerialServer();
            }
            InitializeCmd();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (HideWindow)
            {
                Hide();
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            Serial = new SerialServer(comboBox1.SelectedItem.ToString());
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Connected)
            {
                //MessageBox.Show("You cannot change the current Port. Restart the program");
            }
        }

        private void InitializeCmd()
        {
            try
            {
                cmd = new Process();
                cmd.StartInfo.FileName = @"C:\Windows\System32\cmd.exe";
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardError = true;
                cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmd.StartInfo.ErrorDialog = false;
                cmd.OutputDataReceived += new DataReceivedEventHandler((s, e) =>
                {
                    if (Serial != null)
                    if (Serial.pass)
                    {
                        Serial.Port.WriteLine("VCMD: " + e.Data);
                    }
                    if (TelegramBot != null)
                    if (TelegramBot.Authenticated)
                    {
                        TelegramBot.SendMessage("VCMD: " + e.Data);
                    }
                });
                cmd.ErrorDataReceived += new DataReceivedEventHandler((s, e) =>
                {
                    if (Serial != null)
                    if (Serial.pass)
                    {
                        Serial.Port.WriteLine("VCMD: " + e.Data);
                    }
                    if (TelegramBot != null)
                    if (TelegramBot.Authenticated)
                    {
                        TelegramBot.SendMessage("VCMD: " + e.Data);
                    }
                });
                cmd.Exited += new EventHandler((s, e) =>
                {
                    if (Serial != null)
                    if (Serial.pass)
                    {
                        Serial.Port.WriteLine("VCMD: The CMD has exited");
                    }
                    if (TelegramBot != null)
                    if (TelegramBot.Authenticated)
                    {
                        TelegramBot.SendMessage("VCMD: The CMD has exited");
                    }
                    InitializeCmd();
                    if (Serial.pass)
                    {
                        Serial.Port.WriteLine("VCMD: Process restarted");
                    }
                    if (TelegramBot.Authenticated)
                    {
                        TelegramBot.SendMessage("VCMD: Process restarted");
                    }
                });
                bool r = cmd.Start();
                cmd.BeginOutputReadLine();
                cmd.BeginErrorReadLine();
                CmdStream = cmd.StandardInput;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

        }

        
        protected override void OnClosing(CancelEventArgs e)
        {
            if (Connected)
            {
                e.Cancel = true;
                Hide();
            }
            else
            {
                e.Cancel = false;
            }
        }
    }
}