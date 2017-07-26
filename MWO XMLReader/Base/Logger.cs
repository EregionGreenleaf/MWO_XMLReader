using ERwin_CA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MWO_XMLReader
{/// <summary>
/// Manage the logging system.
/// </summary>
    static class Logger
    {
        /// <summary>
        /// Spaces for each log level
        /// </summary>
        private const string DEF_SPACE = "   ";
        private const string EMPTY = "";
        private static string FileName;
        private static FileInfo FileInfos;
        //private static StreamWriter StrWr;
        private static string FileNameStream;

        public static void Initialize(string fileName)
        {
            Timer.SetFirstTime(DateTime.Now);
            FileName = fileName;
            FileInfos = new FileInfo(FileName);
            FileNameStream = FileInfos.DirectoryName + 
                             @"\" +
                             Path.GetFileNameWithoutExtension(FileInfos.FullName) + 
                             "_" +
                             Timer.GetTimestampDay(DateTime.Now) + 
                             ".txt";

            if (!Directory.Exists(FileInfos.DirectoryName))
            {
                Directory.CreateDirectory(FileInfos.DirectoryName);
            }
            //StrWr = File.AppendText(FileNameStream);
        }

        /// <summary>
        /// Prints on console
        /// </summary>
        /// <param name="text"></param>
        public static void PrintC(string text, string type = EMPTY)
        {
            string line = Timer.GetTimestampPrecision(DateTime.Now) + DEF_SPACE + type + text;
            Console.WriteLine(line);
        }

        /// <summary>
        /// Prints both on console and log file
        /// </summary>
        /// <param name="text"></param>
        /// <param name="level"></param>
        public static void PrintLC(string text, int level = 1, string type = EMPTY)
        {
            if (!(level > ConfigFile.LOG_LEVEL))
            {
                string spaces = string.Empty;
                string line = Timer.GetTimestampPrecision(DateTime.Now);
                if (!(level >= 0 && level <= 10))
                    level = 1;
                for (int x = 0; x < level; x++)
                {
                    //line = line + DEF_SPACE;
                    spaces = spaces + DEF_SPACE;
                }
                line += spaces;
                // Check if multiline
                if(!text.Contains("\r\n") && !text.Contains("\n") && !text.Contains(Environment.NewLine))
                    line = line + type + text;
                else
                {
                    string[] arrayText = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                    spaces += "            ";
                    line = line + type + arrayText[0];
                    text = arrayText[0];
                    if (arrayText.Count() > 0)
                    {
                        for(int i = 1; i < arrayText.Count(); i++)
                        {
                            line = line + Environment.NewLine + spaces + arrayText[i];
                        }
                    }
                }
                Console.WriteLine(line);
                using (StreamWriter StrWr = File.AppendText(FileNameStream))
                {
                    StrWr.WriteLine(line);
                    StrWr.Dispose();
                }
            }
        }

        /// <summary>
        /// Prints on the specified file.
        /// Can print with timestamp.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="text"></param>
        /// <param name="timestamp"></param>
        /// <param name="type"></param>
        public static void PrintF(string fileName, string text, bool timestamp = false, string type = EMPTY)
        {
            string line = (timestamp ? (Timer.GetTimestampPrecision(DateTime.Now) + DEF_SPACE) : string.Empty) 
                          + type 
                          + text;
            FileInfo file = new FileInfo(fileName);
            DirectoryInfo dir = new DirectoryInfo(file.DirectoryName);
            if (dir.Exists)
                using (StreamWriter StrWr = File.AppendText(fileName))
                {
                    StrWr.WriteLine(line);
                    StrWr.Close();
                    StrWr.Dispose();
                }
        }

    }
}
