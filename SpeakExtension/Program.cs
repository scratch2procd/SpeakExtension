using System;
using System.Collections.Generic;
using System.Speech.Synthesis;
using System.Net;
using System.Text;

namespace SpeakExtension
{
    /// <summary>
    /// Scratch Extension to convert text to speech using new blocks;
    /// Speak %s
    /// Speak %s and wait
    /// Set Volume %n
    /// Change Volume by %n
    /// Set Rate %n
    /// Change Rate by %n
    /// and reporters
    /// Volume  [0.100]
    /// Rate    [-10..10]
    /// Gender  ["male, female", "other"]
    /// </summary>
    class Program
    {
        static readonly HttpListener listener = new HttpListener(); // handles the http connection with Scratch
        static SpeechToText s2Text = new SpeechToText(); // does the speech synthesis
        static int port = 8080; // default port is 8080. If changed then s2e file needs to reflect this
        static void Main(string[] args)
        {
            // accept -p=<int> to change port number from default 8080
            foreach(string arg in args)
            {
                if (arg.StartsWith("-p="))
                {
                    int newport = 0;
                    if (int.TryParse(arg.Substring(3),out newport))
                    {
                        port = newport;
                    }
                }
            }
            listener.Prefixes.Add(string.Format("http://+:{0}/", port));
            listener.Start();
            listener.BeginGetContext(new AsyncCallback(ListenerCallback),listener);
            // If get access denied then need to run program as Administrator or give URL admin priveledges with
            // netsh http add urlacl url=http://+:8080/MyUri user=DOMAIN\user
            Console.WriteLine("Speak Extension (c) 2014 Procd");
            Console.WriteLine(String.Format("Listening on port {0}",port));
            Console.WriteLine("Press return to exit.");
            Console.ReadLine();
            listener.Close();
        }
        // Flash cross domain policy that Scratch needs.
        static string crossdomainpolicy = @"<cross-domain-policy><allow-access-from domain=""*"" to-ports=""{0}""/></cross-domain-policy>\0";

        public static void ListenerCallback(IAsyncResult result)
        {
            HttpListener listener = (HttpListener)result.AsyncState;
            // Call EndGetContext to complete the asynchronous operation.
            HttpListenerContext context = listener.EndGetContext(result);
            // start listening for another request
            listener.BeginGetContext(new AsyncCallback(ListenerCallback), listener);
            // carry on with response
            HttpListenerRequest request = context.Request;
            string responseString = "";
            string msg = request.RawUrl;
            switch (msg)
            {
                // Basic Scratch requests
                case "/crossdomain.xml":
                    responseString = string.Format(crossdomainpolicy, port);
                    break;
                case "/poll":
                    responseString = s2Text.Poll();
                    break;
                case "/reset_all":
                    s2Text.Reset();
                    break;
                // Scratch command requests
                default:
                        string decoded = Uri.UnescapeDataString(msg);
                        string[] tokens = decoded.Split('/'); //tokens[0] will be "" for "/..../...../", so ignore
                        if (tokens.Length > 1)
                        {
                            switch (tokens[1])
                            {
                                case "speak":
                                    s2Text.Speak(tokens[2]);
                                    break;
                                case "speakwait":
                                    s2Text.SpeakAndWait(tokens[3], tokens[2]);
                                    break;
                                case "setvolume":
                                    {
                                        int vol = 0;
                                        if (int.TryParse(tokens[2],out vol))
                                        {
                                            s2Text.Volume = vol;
                                        }
                                    }
                                    break;
                                case "changevolume":
                                    {
                                        int vol = 0;
                                        if (int.TryParse(tokens[2], out vol))
                                        {
                                            s2Text.Volume += vol;
                                        }
                                    }
                                    break;
                                case "setrate":
                                    {
                                        int rate = 0;
                                        if (int.TryParse(tokens[2], out rate))
                                        {
                                            s2Text.Rate = rate;
                                        }
                                    }
                                    break;
                                case "changerate":
                                    {
                                        int rate = 0;
                                        if (int.TryParse(tokens[2], out rate))
                                        {
                                            s2Text.Rate += rate;
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                    break;

            }
            // Obtain a response object.
            HttpListenerResponse response = context.Response;
            response.StatusCode = (int)HttpStatusCode.OK;
            // Construct a response. 
            byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
            // Get a response stream and write the response to it.
            response.ContentLength64 = buffer.Length;
            System.IO.Stream output = response.OutputStream;
            output.Write(buffer, 0, buffer.Length);
            // You must close the output stream.
            output.Close();
        }
    }

    public class SpeechToText
    {
        class TextID
        {
            public string Text { get; set; } // Text to speak
            public string ID { get; set; } // If wait block then get ID as well
            public TextID(string text, string id)
            {
                this.Text = text;
                this.ID = id;
            }
        }
        Queue<TextID> SpeechQ = new Queue<TextID>();// Queue all the text to speak and playback one at a time
        Object Qlock = new Object(); // Multithreaded lock
        SpeechSynthesizer reader = new SpeechSynthesizer(); // This does the speech synthesis
        public SpeechToText()
        {
            reader.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(reader_SpeakCompleted);
        }
        public void reader_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            lock (Qlock)
            {
                if (SpeechQ.Count > 0)
                {
                    SpeechQ.Dequeue();// remove finished speech
                    if (SpeechQ.Count > 0)
                    {
                        TextID tid = SpeechQ.Peek(); // Set next speech going
                        reader.SpeakAsync(tid.Text);
                    }                    
                }
            }
        }
        public void Speak(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }
            TextID tid = new TextID(text,null);
            lock (Qlock)
            {
                SpeechQ.Enqueue(tid);
                if (SpeechQ.Count == 1)
                {
                    // can start speaking
                    reader.SpeakAsync(text);
                }
            }
        }

        public void SpeakAndWait(string text, string id)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }
            TextID tid = new TextID(text, id);
            lock (Qlock)
            {
                SpeechQ.Enqueue(tid);
                if (SpeechQ.Count == 1)
                {
                    // can start speaking
                    reader.SpeakAsync(text);
                }
            }
        }

        public void Reset()
        {
            lock (Qlock)
            {
                SpeechQ.Clear();
                reader.SpeakAsyncCancelAll();
            }
        }
        public int Volume
        {
            get
            {
                return reader.Volume;//0..100
            }
            set
            {
                reader.Volume = value;
            }
        }
        public int Rate
        {
            get
            {
                return reader.Rate;//-10..10
            }
            set
            {
                reader.Rate = value;
            }
        }
        public string Gender()
        {
            switch (reader.Voice.Gender)
            {
                case VoiceGender.Female:
                    return "female";
                case VoiceGender.Male:
                    return "male";
                default:
                    return "other";
            }
        }
        public string Poll()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("speechgender {0}",Gender()));
            sb.Append((char)0xA);// new line char
            sb.Append(string.Format("speechrate {0}", Rate));
            sb.Append((char)0xA);// new line char
            sb.Append(string.Format("speechvolume {0}", Volume));
            sb.Append((char)0xA);// new line char
            sb.Append(getBusyItems());
            return sb.ToString();
        }
        string getBusyItems()
        {
            StringBuilder sb = new StringBuilder("_busy");
            bool busy = false;
            lock (Qlock)
            {
                foreach (TextID tid in SpeechQ)
                {
                    if (tid.ID != null)
                    {
                        busy = true;
                        sb.AppendFormat(" {0}", tid.ID);
                    }
                }
            }
            if (busy)
            {
                sb.Append((char)0xA);// new line char
                return sb.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
