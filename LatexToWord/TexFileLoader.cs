using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LatexToWord
{
    public class TexFileLoader
    {
        string fileName;
        public char[] separators = new char[] { '\r', '\n' };

        string[] ignoreCmd;

        public TexFileLoader(string fn)
        {
            if (!fn.EndsWith(".tex", StringComparison.OrdinalIgnoreCase))
            {
                fn = fn + ".tex";
            }

            fileName = fn;
            ignoreCmd = new string[] { };
        }

        public TexFileLoader(string fn, string[] cmds) : this(fn) 
        {
            SetIgnoreCommand(cmds);
        }

        public void SetIgnoreCommand(string[] cmds)
        {
            ignoreCmd = cmds.Select(cmds => cmds.Trim()).ToArray();
        }

        public string LoadAllText()
        {
            string texText = File.ReadAllText(fileName);
            return texText;
        }

        public string[] LoadAllParagraphs()
        {
            string text = LoadAllText();
            return text.Split(separators);
        }

        public string[] LoadAllParagraphsWithExpansion()
        {
            string[] paragraphs = LoadAllParagraphs();
            List<string> plist = new List<string>();
            foreach (string paragraph in paragraphs)
            {
                if (checkInclude(paragraph))
                {
                    string subfilename = extFilenameInInput(paragraph);
                    Console.WriteLine("extending " + subfilename + "\r\n");

                    TexFileLoader loader = new TexFileLoader(subfilename, ignoreCmd);
                    string[] inner = loader.LoadAllParagraphsWithExpansion();
                    plist.AddRange(inner);
                } else if (checkIgnoreCommand(paragraph)) {
                    continue;
                }
                else {
                    plist.Add(paragraph);
                }
            }
            return plist.ToArray();
        }

        private bool checkInclude(string tmp)
        {
            if (tmp.Trim().StartsWith("\\input")) return true;
            // if (tmp.Trim().StartsWith("\\include")) return true;
            return false;
        }

        private string extFilenameInInput(string tmp)
        {
            int left = tmp.IndexOf("{");
            int right = tmp.IndexOf("}");
            return  tmp.Substring( left + 1, right - left - 1);
        }


        private bool checkIgnoreCommand(string tmp)
        {
            string para = tmp.Trim();
            foreach (var cmd in ignoreCmd)
            {
                if (cmd == "") continue;
                if (para.StartsWith(cmd)) return true;
            }
            return false;
        }


    }
}
