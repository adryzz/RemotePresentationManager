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
        SerialPort Port;
        bool Connected = false;
        bool Hide = false;//this hides the window when it is shown
        SoundPlayer Player;
        WaveOut WaveOutDevice = new WaveOut();
        AudioFileReader AudioFileReader;
        CoreAudioDevice DefaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;
        Process cmd = new Process();
        StreamWriter CmdStream;
        bool pass = false;//Password protection
        int remaining = 5;//remaining tries
        string title = "";//title of all messageboxes
        bool loop = false;//audio loop
        string RegistryKey = @"HKEY_CURRENT_USER\Software\RemotePresentationManager\Passwords";
        string ValueName = "Key";
        string CurrentPass = null;
        bool MouseSoftLock = false;//if the mouse is softlocked
        Point MouseLockPoint = new Point(200, 200);//the point t which the pointer wiil be softlocked
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
            {
                Hide = true;
                Port = new SerialPort(args[0], 9600, Parity.None, 8, StopBits.One);
                Port.Open();
                Port.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
                if (Port.IsOpen)
                {
                    Port.WriteLine("Insert the password");
                    button1.Enabled = false;
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Status: " + Port.PortName + " online");
                    Connected = true;
                    InitializeCmd();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (string s in SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(s);
            }
            string current = null;
            try
            {
                current = Registry.GetValue(RegistryKey, ValueName, null).ToString();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(this, ex.Message, "RPMPasswordSet", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
            }
            if (current == null)
            {
                
            }
            else
            {
                CurrentPass = current;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (Hide)
            {
                Hide();
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            try
            {
                Port = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
                Port.Open();
                Port.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
                if (Port.IsOpen)
                {
                    Port.WriteLine("Insert the password");
                    button1.Enabled = false;
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Status: " + Port.PortName + " online");
                    Connected = true;
                    InitializeCmd();
                }
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Connected)
            {
                //MessageBox.Show("You cannot change the current Port. Restart the program");
            }
        }

        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                CheckFunctions(Port.ReadExisting());
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Port.WriteLine(ex.GetType() + ": " + ex.Message);
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
                    if (pass)
                    {
                        Port.WriteLine("VCMD: " + e.Data);
                    }
                });
                cmd.ErrorDataReceived += new DataReceivedEventHandler((s, e) =>
                {
                    Port.WriteLine("VCMD: " + e.Data);
                });
                cmd.Exited += new EventHandler((s, e) =>
                {
                    Port.WriteLine("VCMD: The CMD has exited");
                    InitializeCmd();
                    Port.WriteLine("VCMD: Process restarted");
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

        private void CheckFunctions(string data)
        {
            data = data.Replace("\n", string.Empty);
            data = data.Replace("\r", string.Empty);
            data = data.Replace("\t", string.Empty);

            if (pass)
            {
                if (data.Equals("AT"))
                {
                    Console.WriteLine("Sending current status...");
                    Port.WriteLine("Online!");
                    Port.WriteLine(Port.NewLine + "PORT INFO: Current Port: " + Port.PortName + Port.NewLine + "Baud rate: " + Port.BaudRate + Port.NewLine + "Using RPM " + Application.ProductVersion + " By Adryzz\n(Adryzz#7264)");

                }
                else if(data.Equals("LOCK"))
                {
                    Port.WriteLine("Command received! Locking...");
                    remaining = 5;//reset remaining tries
                    pass = false;//lock
                    Port.WriteLine("Insert the password");
                    Console.WriteLine("Lock system");
                }
                else if (data.Equals("HELP"))
                {
                    Help();
                }
                else if (data.Equals("SHOW"))
                {
                    Port.WriteLine("Command received!");
                    Invoke(new WindowStateDelegate(SetWindowState), true);
                    Console.WriteLine("Show window");
                }
                else if (data.Contains("SOFTLOCK"))
                {
                    Port.WriteLine("Command received!");
                    SoftLock(data);
                    Console.WriteLine("Softlock");
                }
                else if (data.Contains("PLAYSOUND "))
                {
                    Port.WriteLine("Command received!");
                    PlaySound(data);
                    Console.WriteLine("Play sound file");
                }
                else if (data.Contains("PLAYURL "))
                {
                    Port.WriteLine("Command received!");
                    PlayMp3FromUrl(data);
                    Console.WriteLine("Play url");
                }
                else if (data.Contains("PLAY "))
                {
                    Port.WriteLine("Command received!");
                    Play(data);
                    Console.WriteLine("Play sound");
                }
                else if (data.Contains("URLIMAGE "))
                {
                    Port.WriteLine("Command received!");
                    UrlImage(data);
                    Console.WriteLine("Url image");
                }
                else if (data.Equals("MUTE"))
                {
                    Port.WriteLine("Command received!");
                    Mute();
                    Console.WriteLine("Mute");
                }
                else if (data.Equals("STOP"))
                {
                    Port.WriteLine("Command received!");
                    Stop();
                    Console.WriteLine("Stop sound file player");
                }
                else if (data.Equals("PAUSE"))
                {
                    Port.WriteLine("Command received!");
                    Pause();
                    Console.WriteLine("Pause sound file player");
                }
                else if (data.Equals("RESUME"))
                {
                    Port.WriteLine("Command received!");
                    Resume();
                    Console.WriteLine("Resume sound file player");
                }
                else if (data.Equals("UNMUTE"))
                {
                    Port.WriteLine("Command received!");
                    UnMute();
                    Console.WriteLine("Unmute");
                }
                else if (data.Contains("VOLUME "))
                {
                    Port.WriteLine("Command received!");
                    AdjustVolume(data);
                    Console.WriteLine("Adjust volume");
                }
                else if (data.Contains("CLIP "))
                {
                    Port.WriteLine("Command received!");
                    ClipBoard(data);
                    Console.WriteLine("Edit clipboard");
                }
                else if (data.Contains("KEY "))
                {
                    Port.WriteLine("Command received!");
                    Key(data);
                    Console.WriteLine("Send custom key");
                }
                else if (data.Contains("LOOP "))
                {
                    Port.WriteLine("Command received!");
                    Loop(data);
                    Console.WriteLine("Set loop mode");
                }
                else if (data.Contains("SAY "))
                {
                    Port.WriteLine("Command received!");
                    Say(data);
                    Console.WriteLine("Say text");
                }
                else if (data.Contains("IMG "))
                {
                    Port.WriteLine("Command received!");
                    Draw(data);
                    Console.WriteLine("Draw bitmap");
                }
                else if (data.Contains("ROTATE "))
                {
                    Port.WriteLine("Command received!");
                    Rotate(data);
                    Console.WriteLine("Rotate screen");
                }
                else if (data.Contains("MOUSE "))
                {
                    Port.WriteLine("Command received!");
                    Mouse(data);
                    Console.WriteLine("Set mouse position");
                }
                else if (data.Contains("QMSG "))
                {
                    Port.WriteLine("Command received!");
                    QMsg(data);
                    Console.WriteLine("Question MessageBox");
                }
                else if (data.Contains("MSG "))
                {
                    Port.WriteLine("Command received!");
                    Msg(data);
                    Console.WriteLine("MessageBox");
                }
                else if (data.Contains("TITLE "))
                {
                    Port.WriteLine("Command received!");
                    Title(data);
                    Console.WriteLine("Set Title");
                }
                else if (data.Equals("MSGLOOP"))
                {
                    Port.WriteLine("Command received!");
                    MsgLoop();
                    Console.WriteLine("Loop MessageBox");
                }
                else if (data.Equals("HOOK"))
                {
                    Port.WriteLine("Command received!");
                    Hook();
                    Console.WriteLine("Hook keys");
                }
                else if (data.Equals("UNHOOK"))
                {
                    Port.WriteLine("Command received!");
                    UnHook();
                    Console.WriteLine("Unhook keys");
                }
                else if (data.Equals("HIDE"))
                {
                    Port.WriteLine("Command received!");
                    Invoke(new WindowStateDelegate(SetWindowState), false);
                    Console.WriteLine("Hide window");
                }
                else if (data.Equals("CLOSE"))
                {
                    Port.WriteLine("Command received! Exiting...");
                    Invoke(new ExitDelegate(Exit));
                    Console.WriteLine("Exit");
                }
                else if (data.Equals("ESC"))
                {
                    Port.WriteLine("Command received!");
                    Esc();
                    Console.WriteLine("ESC key");
                }
                else if (data.Equals("F5"))
                {
                    Port.WriteLine("Command received!");
                    F5();
                    Console.WriteLine("F5 key");
                }
                else if (data.Equals("LEFT"))
                {
                    Port.WriteLine("Command received!");
                    LeftArrow();
                    Console.WriteLine("LEFT ARROW key");
                }
                else if (data.Equals("RIGHT"))
                {
                    Port.WriteLine("Command received!");
                    RightArrow();
                    Console.WriteLine("RIGHT ARROW key");
                }
                else if (data.Equals("UP"))
                {
                    Port.WriteLine("Command received!");
                    UpArrow();
                    Console.WriteLine("UP ARROW key");
                }
                else if (data.Equals("DOWN"))
                {
                    Port.WriteLine("Command received!");
                    DownArrow();
                    Console.WriteLine("Down ARROW key");
                }
                else if (data.Equals("ALTF4"))
                {
                    Port.WriteLine("Command received!");
                    Altf4();
                    Console.WriteLine("ALT+F4 KEYSTROKE");
                }
                else if (data.Equals("SHUTDOWN"))
                {
                    Port.WriteLine("Command received!");
                    Shutdown();
                    Console.WriteLine("SHUTDOWN COMMAND");
                }
                else if (data.Equals("REBOOT"))
                {
                    Port.WriteLine("Command received!");
                    Reboot();
                    Console.WriteLine("REBOOT COMMAND");
                }
                else if (data.Equals("CRASH"))
                {
                    Port.WriteLine("Command received!");
                    Crash();
                    Console.WriteLine("CRASH COMMAND");
                }
                else if (data.Equals("EXPLORER"))
                {
                    Port.WriteLine("Command received!");
                    Explorer();
                    Console.WriteLine("KILL EXPLORER");
                }
                else if (data.Equals("STARTUP"))
                {
                    Port.WriteLine("Command received!");
                    Startup();
                    Console.WriteLine("SETUP STARTUP RUN");
                }
                else if (data.Contains("VCMD "))
                {
                    Port.WriteLine("Command received!");
                    VCmd(data);
                    Console.WriteLine("VCMD COMMAND");
                }
                else if (data.Contains("CMD "))
                {
                    Port.WriteLine("Command received!");
                    Cmd(data);
                    Console.WriteLine("CMD COMMAND");
                }
                else
                {
                    Console.WriteLine(data + " not recognized as a command");
                    Port.WriteLine(data + " not recognized as a command");
                }
            }
            else
            {
                Password(data);
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


        private void Password(string data)//it is reccommended to change the default password, even if it is secure (use my tool called RPMPasswordSet)
        {
            if (remaining == 0)
            {
                Shutdown();
            }
            if (CurrentPass == null)
            {
                if (data.Equals("dbe6a4b729ff"))//the password
                {
                    pass = true;
                    Port.WriteLine("Access granted");
                }
                else
                {
                    remaining--;
                    Port.WriteLine("Wrong password. you have " + remaining + " tries until the system shuts down");
                }
            }
            else
            {
                SHA512 sha = new SHA512Managed();
                string hash = Encoding.UTF8.GetString(sha.ComputeHash(Encoding.UTF8.GetBytes(data)));//calculate the hash of the typed password
                if (hash.Equals(CurrentPass))//the hash of the current password
                {
                    pass = true;
                    Port.WriteLine("Access granted");
                }
                else
                {
                    remaining--;
                    Port.WriteLine("Wrong password. you have " + remaining + " tries until the system shuts down");
                }
            }
        }

        #region Delegates and stuff

        private delegate void WindowStateDelegate(bool visible);

        private void SetWindowState(bool visible)
        {
            if (visible)
            {
                Show();
            }
            else
            {
                Hide();
            }
        }

        private delegate void SetClipboardDelegate(string data);

        private void SetClipboard(string data)
        {
            Clipboard.SetText(data);
        }

        private delegate void ExitDelegate();

        private void Exit()
        {
            if (AudioFileReader != null)
            {
                AudioFileReader.Dispose();
            }
            if (WaveOutDevice != null)
            {
                WaveOutDevice.Dispose();
            }
            Port.Close();
            Port.Dispose();
            Connected = false;
            Application.Exit();
        }

        #endregion

        #region funcs
        private void Esc()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{ESC}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void F5()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{F5}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void Altf4()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("%{F4}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void LeftArrow()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{LEFT}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void RightArrow()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{RIGHT}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void UpArrow()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{UP}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void DownArrow()
        {
            if (checkBox1.Checked)
            {
                SendKeys.SendWait("{DOWN}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void Mouse(string data)
        {
            try
            {
                data = data.Remove(0, 6);//MOUSE X200Y300
                int x = Int32.Parse(data.Remove(data.IndexOf('Y')).Remove(0, 1));
                int y = Int32.Parse(data.Remove(0, data.IndexOf('Y')).Remove(0, 1));
                Cursor.Position = new Point(x, y);
            }
            catch (Exception)
            {
                Port.WriteLine("The argument was not in the correct format - Correct format: MOUSE X<num>Y<num>");
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
                Port.WriteLine("SHUTDOWN is disabled.");
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
                Port.WriteLine("REBOOT is disabled.");
            }
        }

        private void Crash()
        {
            if (checkBox8.Checked)
            {
                StaticPayloads.Crash();
            }
            else
            {
                Port.WriteLine("CRASH is disabled.");
            }
        }

        private void Explorer()
        {
            Process.Start("taskkill", "/f /im explorer.exe");
        }

        private void Cmd(string data)
        {
            data = data.Remove(0, 4);
            if (checkBox10.Checked)
            {
                Process.Start("cmd", "/c start " + data);
            }
            else
            {
                Port.WriteLine("CMD is disabled.");
            }
        }

        private void VCmd(string data)
        {
            data = data.Remove(0, 5);
            if (checkBox10.Checked)
            {
                CmdStream.WriteLine(data);
                CmdStream.Flush();
                CmdStream = cmd.StandardInput;
            }
            else
            {
                Port.WriteLine("CMD is disabled.");
            }
        }

        private void Play(string data)
        {
            if (checkBox3.Checked)
            {
                data = data.Remove(0, 5);

                if (data.Equals("ERROR"))
                {
                    Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Fail.wav");
                    Player.Play();
                }
                else if (data.Equals("BACKGROUND"))
                {
                    Player = new SoundPlayer(@"C:\Windows\Media\Windows Background.wav");
                    Player.Play();
                }
                else if (data.Equals("FOREGROUND"))
                {
                    Player = new SoundPlayer(@"C:\Windows\Media\Windows Foreground.wav");
                    Player.Play();
                }
                else if (data.Equals("DEVICEIN"))
                {
                    Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Insert.wav");
                    Player.Play();
                }
                else if (data.Equals("DEVICEOUT"))
                {
                    Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Remove.wav");
                    Player.Play();
                }
                else
                {
                    Port.WriteLine("No corresponding sound!");
                }
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void PlaySound(string data)
        {
            if (checkBox3.Checked)
            {
                data = data.Remove(0, 10);
                if (WaveOutDevice.PlaybackState == PlaybackState.Playing || WaveOutDevice.PlaybackState == PlaybackState.Paused)
                {
                    WaveOutDevice.Stop();
                }
                if (AudioFileReader != null)
                {
                    AudioFileReader.Dispose();
                }
                AudioFileReader = new AudioFileReader(data);
                WaveOutDevice.Init(AudioFileReader);
                WaveOutDevice.Play();
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void Stop()
        {
            WaveOutDevice.Stop();
        }

        private void Pause()
        {
            WaveOutDevice.Pause();
        }

        private void Resume()
        {
            WaveOutDevice.Resume();
        }

        private void Loop(string data)
        {
            data = data.Remove(0, 5);
            if (data.Equals("ON"))
            {
                loop = true;
                Port.WriteLine("Loop mode set!");
            }
            else if (data.Equals("OFF"))
            {
                loop = false;
                Port.WriteLine("Loop mode set!");
            }
            else
            {
                Port.WriteLine("This command accepts only ON and OFF as arguments");
            }
        }

        private void Key(string data)
        {
            data = data.Remove(0, 4);

            if (data.Equals("ENTER"))
            {
                data = "{ENTER}";
            } else if (data.Equals("CANC"))
            {
                data = "{DEL}";
            }
            else if (data.Equals("TAB"))
            {
                data = "{TAB}";
            }
            else if (data.Equals("HELP"))
            {
                data = "{HELP}";
            }
            else if (data.Equals("CTRLV"))
            {
                data = "^V";
            }
            else if (data.Equals("TAB"))
            {
                data = "{TAB}";
            }
            else if (data.Equals("SHIFTAB"))
            {
                data = "+{TAB}";
            }
            else if (data.Equals("PLAYPAUSE"))
            {
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("PLAY"))
            {
                keybd_event(VK_MEDIA_PLAY, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PLAY, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("PAUSE"))
            {
                keybd_event(VK_MEDIA_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PAUSE, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("STOP"))
            {
                keybd_event(VK_MEDIA_STOP, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_STOP, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("NEXT"))
            {
                keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("PREV"))
            {
                keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("REW"))
            {
                keybd_event(VK_MEDIA_REWIND, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_REWIND, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (data.Equals("FFW"))
            {
                keybd_event(VK_MEDIA_FAST_FORWARD, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_FAST_FORWARD, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }

            SendKeys.SendWait(data);
        }

        private void AdjustVolume(string data)
        {
            if (checkBox3.Checked)
            {
                data = data.Remove(0, 7);
                DefaultPlaybackDevice.Volume = int.Parse(data);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void Mute()
        {
            if (checkBox3.Checked)
            {
                DefaultPlaybackDevice.Mute(true);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
            
        }

        private void UnMute()
        {
            if (checkBox3.Checked)
            {
                DefaultPlaybackDevice.Mute(false);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void Msg(string data)
        {
            if (checkBox2.Checked)
            {
                data = data.Remove(0, 4);
                new Thread(() => {
                    MessageBox.Show(new Form { TopMost = true }, data, title);
                    Port.WriteLine("MSG: The user pressed OK");
                }).Start();
            }
            else
            {
                Port.WriteLine("Message Boxes are disabled");
            }
        }

        private void QMsg(string data)
        {
            if (checkBox2.Checked)
            {
                data = data.Remove(0, 5);
                new Thread(() => {
                    DialogResult res = MessageBox.Show(new Form { TopMost = true }, data, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        Port.WriteLine("QMSG: The user pressed YES");
                    }
                    else
                    {
                        Port.WriteLine("QMSG: The user pressed NO");
                    }
                }).Start();
            }
            else
            {
                Port.WriteLine("Message Boxes are disabled");
            }
        }

        private void MsgLoop()
        {
            if (checkBox2.Checked)
            {
                new Thread(() => {
                    string t1 = "Are you";
                    string t2 = " sure you want to do this?";
                    string t = "";
                    while (MessageBox.Show(new Form { TopMost = true }, t1 + t + t2, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        t += " really";
                    }
                    
                }).Start();
            }
            else
            {
                Port.WriteLine("Message Boxes are disabled");
            }
        }

        private void Title(string data)
        {
            title = data.Remove(0, 6);
        }

        private void ClipBoard(string data)
        {
            if (checkBox4.Checked)
            {
                data = data.Remove(0, 5);
                Invoke(new SetClipboardDelegate(SetClipboard), data);
            }
            else
            {
                Port.WriteLine("Clipboard edit is disabled");
            }
        }

        private void Say(string data)
        {
            if (checkBox3.Checked)
            {
                data = data.Remove(0, 4);
                StaticPayloads.Say(data);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void Draw(string data)
        {
            if (checkBox5.Checked)
            {
                data = data.Remove(0, 7);
                StaticPayloads.DrawBitmapToScreen((Bitmap)Bitmap.FromFile(data));
            }
            else
            {
                Port.WriteLine("Images are disabled.");
            }
        }

        private void Rotate(string data)
        {
            if (checkBox5.Checked)
            {
                data = data.Remove(0, 7);//ROTATE

                if (data.Equals("0"))
                {
                    StaticPayloads.Rotate(0, StaticPayloads.Orientations.DEGREES_CW_0);
                }
                else if (data.Equals("90"))
                {
                    StaticPayloads.Rotate(0, StaticPayloads.Orientations.DEGREES_CW_90);
                }
                else if (data.Equals("180"))
                {
                    StaticPayloads.Rotate(0, StaticPayloads.Orientations.DEGREES_CW_180);
                }
                else if (data.Equals("270"))
                {
                    StaticPayloads.Rotate(0, StaticPayloads.Orientations.DEGREES_CW_270);
                }
                else
                {
                    Port.WriteLine("No corresponding rotation! This command accepts only 0, 90, 180 and 270 as arguments");
                }
            }
            else
            {
                Port.WriteLine("Display is disabled.");
            }
        }

        private void Hook()
        {
            if (checkBox1.Checked)
            {
                StaticPayloads.KeyboardHook();
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void UnHook()
        {
            if (checkBox1.Checked)
            {
                StaticPayloads.ReleaseKeyboardHook();
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        private void Startup()
        {
            string exe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RemotePresentationManager", "RemotePresentationManager.exe");
            string bat = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Start Menu\Programs\Startup\RPM.bat");
            Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RemotePresentationManager"));
            if (File.Exists(exe))
            {
                File.Delete(exe);
            }
            File.Copy(Application.ExecutablePath, exe);
            File.WriteAllText(bat, String.Format("start {0} {1}", exe, Port.PortName));
        }

        private void UrlImage(string url)
        {
            if (checkBox5.Checked)
            {
                url = url.Remove(0, 9);//URLIMAGE
                WebClient client = new WebClient();
                client.OpenReadCompleted += (s, e) =>
                {
                    byte[] imageBytes = new byte[e.Result.Length];
                    e.Result.Read(imageBytes, 0, imageBytes.Length);

                    // Now you can use the returned stream to set the image source too
                    StaticPayloads.DrawBitmapToScreen((Bitmap)Image.FromStream(e.Result));
                };
                client.OpenReadAsync(new Uri(url));
            }
            else
            {
                Port.WriteLine("Images are disabled.");
            }
        }

        private void Help()
        {
            Port.WriteLine("List of all commands\n\nAT - Show connection info\nLOCK - Re-lock the program as at startup\nHELP - Show this help\nSHOW - Shows the main window\nHIDE - Hides the main window\nCLOSE - Quits the program\nKEY <some text> - Types that text/Presses keys\nMOUSE X<num>Y<num> - Sets mouse position\nCLIP <some text> - Copies that text into the clipboard\nMUTE - Mutes volume\nUNMUTE - Unmutes volume\nVOLUME <0-100> - Set volume percentage\nPLAY <FOREGROUND - BACKGROUND - ERROR - DEVICEIN - DEVICEOUT> - Plays the sound\nSAY <some text> - Says the text\nUP/DOWN/LEFT/RIGHT/F5/ALTF4/ESC - the key(s)\nIMG <path of an image> - Draws the image\nURLIMAGE <url> - Draws an image on the Internet (experimental)\nPLAYSOUND <path of a sound file> - Plays the sound file\nSTOP - Stops the sound player\nPAUSE - Pauses the sound player\nRESUME - Resumes the sound player after pause\nPLAYURL <url> - Plays a sound file in an URL\nROTATE <0-90-180-270> - Rotates the screen\nMSG <some text> - Shows a message box\nSHUTDOWN/REBOOT - Shuts down/Reboots the system\nCRASH - Real BSOD\nEXPLORER - Kills explorer.exe\nCMD <command> - Runs a command\nHOOK/UNHOOK - Blocks shortcuts\nSTARTUP - Setups Run at startup");
        }

        private void PlayMp3FromUrl(string url)
        {
            if (checkBox3.Checked)
            {
                    url = url.Remove(0, 8);
                    new Thread(() => 
                    {
                        try
                        {
                        if (WaveOutDevice.PlaybackState == PlaybackState.Playing || WaveOutDevice.PlaybackState == PlaybackState.Paused)
                        {
                            WaveOutDevice.Stop();
                        }
                        using (Stream ms = new MemoryStream())
                        {
                            using (Stream stream = WebRequest.Create(url)
                                .GetResponse().GetResponseStream())
                            {
                                byte[] buffer = new byte[32768];
                                int read;
                                while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                                {
                                    ms.Write(buffer, 0, read);
                                }
                            }

                            ms.Position = 0;
                            using (WaveStream blockAlignedStream =
                                new BlockAlignReductionStream(
                                    WaveFormatConversionStream.CreatePcmStream(
                                        new Mp3FileReader(ms))))
                            {
                                WaveOutDevice.Init(blockAlignedStream);
                                WaveOutDevice.Play();
                                while (WaveOutDevice.PlaybackState == PlaybackState.Playing)
                                {
                                    Thread.Sleep(100);
                                }
                            }
                        }
                        }
                        catch (Exception ex)
                        {
                            Port.WriteLine(ex.GetType() + ": " + ex.Message);
                        }
                    }).Start();
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        private void SoftLock(string data)
        {
            if (checkBox9.Checked)
            {
                data = data.Remove(0, 9);
                if (data.Equals("ENABLE"))
                {
                    timer1.Start();
                }
                else if (data.Equals("DISABLE"))
                {
                    timer1.Stop();
                }
                else if (data.Equals("MOUSE ON"))
                {
                    MouseSoftLock = true;
                }
                else if (data.Equals("MOUSE OFF"))
                {
                    MouseSoftLock = false;
                }
                else if (data.Contains("MOUSE SET "))
                {
                    data = data.Remove(0, 9);
                    int x = int.Parse(data.Substring(1, data.IndexOf('Y')));
                    int y = int.Parse(data.Substring(data.IndexOf('Y')));
                    MouseLockPoint = new Point(x, y);
                }
                else if (data.Contains("SET TIME "))
                {
                    data = data.Remove(0, 9);
                    try
                    {
                        timer1.Stop();
                        timer1.Interval = int.Parse(data);
                        timer1.Start();
                    }
                    catch (Exception)
                    {
                        Port.WriteLine("Unable to set timer interval");
                    }
                }
            
            }
            else
            {
                Port.Write("Softlocks are disabled!");
            }
        }

        #endregion

        #region special keys dont touch here
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);
        const int VK_MEDIA_NEXT_TRACK = 0xB0;
        const int VK_MEDIA_PREV_TRACK = 0xB1;
        const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        const int VK_MEDIA_PLAY = 0xFA;
        const int VK_MEDIA_PAUSE = 0x13;
        const int VK_MEDIA_STOP = 0xB2;
        const int VK_MEDIA_FAST_FORWARD = 0x31;
        const int VK_MEDIA_REWIND = 0x32;
        const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        #endregion

        #region Softlocks
        private void Timer1_Tick(object sender, EventArgs e)//all softlocks
        {
            if (MouseSoftLock)
            {
                Cursor.Position = MouseLockPoint;
            }
        }

        #endregion

    }
}