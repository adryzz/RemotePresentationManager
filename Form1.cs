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

namespace RemotePresentationManager
{
    public partial class Form1 : Form
    {
        SerialPort port;
        Boolean connected = false;
        System.Media.SoundPlayer player;
        CoreAudioDevice defaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;
        int window = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (string s in SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(s);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            try
            {
                port = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
                port.Open();
                port.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
                if (port.IsOpen)
                {
                    port.Write("Ready");
                    button1.Enabled = false;
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Status: " + port.PortName + " online");
                    comboBox1.Enabled = false;
                    checkBox1.Enabled = false;
                    checkBox2.Enabled = false;
                    checkBox3.Enabled = false;
                    checkBox4.Enabled = false;
                    checkBox5.Enabled = false;
                    checkBox6.Enabled = false;
                    checkBox7.Enabled = false;
                    checkBox8.Enabled = false;
                    checkBox9.Enabled = false;
                    checkBox10.Enabled = false;
                    connected = true;
                }
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MessageBox.Show("You cannot change the current port. Restart the program");
        }

        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            CheckFunctions(port.ReadExisting());
        }

        private void CheckFunctions(String data)
        {
            if (data.ToString().Equals("AT"))
            {
                Console.WriteLine("Sending current status...");
                port.Write("Online!");
                port.Write(port.NewLine + "PORT INFO: Current port: " + port.PortName + port.NewLine + "Baud rate: " + port.BaudRate);

            }
            else if (data.ToString().Equals("SHOW"))
            {
                window = 2;
                Console.WriteLine("Show window");
                port.Write("Command received!");
            }
            else if (data.ToString().Contains("PLAY"))
            {
                PlaySound(data);
                Console.WriteLine("Play sound");
                port.Write("Command received!");
            }
            else if (data.ToString().Equals("MUTE"))
            {
                defaultPlaybackDevice.Mute(true);
                Console.WriteLine("Mute");
                port.Write("Command received!");
            }
            else if (data.ToString().Equals("UNMUTE"))
            {
                defaultPlaybackDevice.Mute(false);
                Console.WriteLine("Unmute");
                port.Write("Command received!");
            }
            else if (data.ToString().Contains("VOLUME"))
            {
                AdjustVolume(data);
                Console.WriteLine("Adjust volume");
                port.Write("Command received!");
            }
            else if (data.ToString().Equals("HIDE"))
            {
                window = 1;
                Console.WriteLine("Hide window");
                port.Write("Command received!");
            }
            else if (data.ToString().Equals("CLOSE"))
            {
                window = 3;
                Console.WriteLine("Exit");
                port.Write("Command received! Exiting...");
            }
            else if (data.ToString().Equals("ESC"))
            {
                Console.WriteLine("ESC key");
                port.Write("Command received!");
                Esc();
            }
            else if (data.ToString().Equals("F5"))
            {
                Console.WriteLine("F5 key");
                port.Write("Command received!");
                F5();
            }
            else if (data.ToString().Equals("LEFT"))
            {
                Console.WriteLine("LEFT ARROW key");
                port.Write("Command received!");
                LeftArrow();
            }
            else if (data.ToString().Equals("RIGHT"))
            {
                Console.WriteLine("RIGHT ARROW key");
                port.Write("Command received!");
                RightArrow();
            }
            else if (data.ToString().Equals("ALTF4"))
            {
                Console.WriteLine("ALT+F4 KEYSTROKE");
                port.Write("Command received!");
                Altf4();
            }
            else if (data.ToString().Equals("SHUTDOWN"))
            {
                Console.WriteLine("SHUTDOWN COMMAND");
                port.Write("Command received!");
                Shutdown();
            }
            else if (data.ToString().Equals("REBOOT"))
            {
                Console.WriteLine("REBOOT COMMAND");
                port.Write("Command received!");
                Reboot();
            }
            else if (data.ToString().Equals("CRASH"))
            {
                Console.WriteLine("CRASH COMMAND");
                port.Write("Command received!");
                Crash();
            }
            else if (data.ToString().Equals("EXPLORER"))
            {
                Console.WriteLine("KILL EXPLORER");
                port.Write("Command received!");
                Explorer();
            }
            else if (data.ToString().Contains("CMD"))
            {
                Console.WriteLine("CMD COMMAND");
                port.Write("Command received!");
                Cmd(data);
            }
            else
            {
                Console.WriteLine(data.ToString() + " not recognized as a command");
                port.Write(data.ToString() + " not recognized as a command");
            }

        }
        protected override void OnClosing(CancelEventArgs e)
        {
            if (connected)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                e.Cancel = false;
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (window == 1)
            {
                this.Hide();
                window = 0;
            }
            else if (window == 2)
            {
                this.Show();
                window = 0;
            }
            else if (window == 3)
            {
                connected = false;
                Close();
            }
        }

        private void Esc()
        {
            if (checkBox1.Checked)
            {

            }
            else
            {
                port.Write("ESC is disabled.");
            }
        }

        private void F5()
        {
            if (checkBox2.Checked)
            {

            }
            else
            {
                port.Write("F5 is disabled.");
            }
        }

        private void Altf4()
        {
            if (checkBox5.Checked)
            {

            }
            else
            {
                port.Write("ALT+F4 is disabled.");
            }
        }

        private void LeftArrow()
        {
            if (checkBox3.Checked)
            {

            }
            else
            {
                port.Write("LEFT is disabled.");
            }
        }

        private void RightArrow()
        {
            if (checkBox4.Checked)
            {

            }
            else
            {
                port.Write("RIGHT is disabled.");
            }
        }

        private void Shutdown()
        {
            if (checkBox6.Checked)
            {
                Process.Start("shutdown", "/f /s /t 0");
            }
            else
            {
                port.Write("SHUTDOWN is disabled.");
            }
        }

        private void Reboot()
        {
            if (checkBox7.Checked)
            {
                Process.Start("shutdown", "/f /r /t 0");
            }
            else
            {
                port.Write("REBOOT is disabled.");
            }
        }

        void Crash()
        {
            if (checkBox8.Checked)
            {
                Boolean t1;
                uint t2;
                RtlAdjustPrivilege(19, true, false, out t1);
                NtRaiseHardError(0xc0000022, 0, 0, IntPtr.Zero, 6, out t2);
            }
            else
            {
                port.Write("CRASH is disabled.");
            }
        }

        private void Explorer()
        {
            if (checkBox9.Checked)
            {
                Process.Start("taskkill", "/f /im explorer.exe");
            }
            else
            {
                port.Write("KILLING EXPLORER is disabled.");
            }
        }

        private void Cmd(String NewData)
        {
            var replacements = new[] { new { Find = "CMD ", Replace = "" }, };

            foreach (var set in replacements)
            {
                NewData = NewData.Replace(set.Find, set.Replace);
            }
            if (checkBox10.Checked)
            {
                Process.Start("cmd", "/c start " + NewData);
                port.Write("CMD is not implemented.");
            }
            else
            {
                port.Write("CMD is disabled.");
            }
        }

        private void PlaySound(String NewData)
        {
            var replacements = new[]{ new{Find="PLAY ",Replace=""},};

            foreach (var set in replacements)
            {
                NewData = NewData.Replace(set.Find, set.Replace);
            }

            if (NewData.Equals("ERROR"))
            {
                player = new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Hardware Fail.wav");
                player.Play();
            }
            else if (NewData.Equals("BACKGROUND"))
            {
                player = new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Background.wav");
                player.Play();
            }
            else if (NewData.Equals("FOREGROUND"))
            {
                player = new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Foreground.wav");
                player.Play();
            }
            else if (NewData.Equals("DEVICEIN"))
            {
                player = new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Hardware Insert.wav");
                player.Play();
            }
            else if (NewData.Equals("DEVICEOUT"))
            {
                player = new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Hardware Remove.wav");
                player.Play();
            }
            else
            {
                port.Write("No corresponding sound!");
            }
        }

        private void AdjustVolume(String NewData)
        {
            var replacements = new[]{ new{Find="VOLUME ",Replace=""}};

            foreach (var set in replacements)
            {
                NewData = NewData.Replace(set.Find, set.Replace);

                defaultPlaybackDevice.Volume = int.Parse(NewData);
            }
        }
        [DllImport("ntdll.dll")]
        public static extern uint RtlAdjustPrivilege(int Privilege, bool bEnablePrivilege, bool IsThreadPrivilege, out bool PreviousValue);

        [DllImport("ntdll.dll")]
        public static extern uint NtRaiseHardError(uint ErrorStatus, uint NumberOfParameters, uint UnicodeStringParameterMask, IntPtr Parameters, uint ValidResponseOption, out uint Response);

        
    }
}
