using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace EnableTTSVoices
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("============================================================================================================");
            Console.WriteLine("=========   EnableTTSVoices   ==============================================================================");
            Console.WriteLine("============================================================================================================");
            Console.WriteLine();
            Console.WriteLine("This application will copy registry keys that exist for Microsoft Speech OneCore Voices across to regular Speech Voices to enable use in System.Speech applications (ie. vatSys).");
            Console.WriteLine("This method is documented here: https://www.ghacks.net/2018/08/11/unlock-all-windows-10-tts-voices-system-wide-to-get-more-of-them/");
            Console.WriteLine();

            Console.WriteLine("Open Speech Settings to install additional voices? (y/n):");
            string reply = Console.ReadLine();
            if (reply == "y" || reply == "yes" || reply == "open" || reply == "1")
            {
                Process.Start("ms-settings:speech");
                Console.WriteLine("Press any key to continue:");
                Console.ReadKey();
                Console.WriteLine();
            }
            else
                Console.WriteLine();

            string oneCorePath = @"SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\";
            string speechPathx64 = @"SOFTWARE\Microsoft\Speech\Voices\Tokens\";
            string speechPathx86 = @"SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens\";

            var corekey = Registry.LocalMachine.OpenSubKey(oneCorePath);
            var voices = corekey.GetSubKeyNames();

            var installed_voices = Registry.LocalMachine.OpenSubKey(speechPathx64)?.GetSubKeyNames();

            Console.WriteLine("Found " + voices.Length.ToString() + " OneCore voices:");
            Console.WriteLine(voices.Aggregate((a, b) => a + Environment.NewLine + b));

            Console.WriteLine();

            Console.WriteLine("Copying Keys to enable as regular voices...");
            foreach (string k in voices)
            {
                if (k.Substring(0, 10) != "MSTTS_V110" || k.Count(c=>c == '_') < 3)//MSTTS_V110_enUS_ZiraM
                {
                    Console.WriteLine(k + " is not a MSTTS voice, skipping.");
                    continue;
                }

                string name = k.Split('_')?[3];//ZiraM
                if (name == null)
                {
                    Console.WriteLine("Could not resolve voice name for " + k + ", skipping.");
                    continue;
                }
                name = name.ToLowerInvariant().Substring(0, name.Length - 1);//Zira

                if (installed_voices.Any(v => v.ToLowerInvariant().Contains(name)))//TTS_MS_EN-US_ZIRA_11.0
                {
                    Console.WriteLine("Key already exists for " + CultureInfo.InvariantCulture.TextInfo.ToTitleCase(name) +", skipping.");
                    continue;//voice already exists
                }

                if(Registry.LocalMachine.OpenSubKey(speechPathx86 + k) == null)
                {
                    CopyKey(oneCorePath + k, speechPathx86 + k);
                }
                if(Registry.LocalMachine.OpenSubKey(speechPathx64 + k) == null)
                {
                    CopyKey(oneCorePath + k, speechPathx64 + k);
                }
                Console.WriteLine("Copied " + k);
            }
            Console.WriteLine();
            Console.WriteLine("Done. Press any key to exit:");
            Console.ReadKey();
        }

        private static void CopyKey(string sourcePath, string destinationPath)
        {
            var key = Registry.LocalMachine.OpenSubKey(sourcePath);
            var dest_key = Registry.LocalMachine.CreateSubKey(destinationPath, true);

            foreach (var p in key.GetSubKeyNames())
            {
                CopyKey(sourcePath + @"\" + p, destinationPath + @"\" + p);
            }
            foreach(var val in key.GetValueNames())
            {
                dest_key.SetValue(val, key.GetValue(val));
            }
        }
    }
}
