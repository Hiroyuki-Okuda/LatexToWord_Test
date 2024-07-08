

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LatexToWord;


static List<string[]> SplitArray(string[] array, string delimiter)
{
    List<string[]> result = new List<string[]>();
    List<string> temp = new List<string>();
    bool multi_delim = false;

    foreach (var item in array)
    {
        if (item == delimiter)
        {
            if (multi_delim)
            {
                continue;
            }
            else
            {
                multi_delim = true;
                if (temp.Count > 0)
                {
                    result.Add(temp.ToArray());
                    temp.Clear();
                }
            }
        }
        else
        {
            multi_delim = false;
            temp.Add(item);
        }
    }

    // 最後の部分配列を追加
    if (temp.Count > 0)
    {
        result.Add(temp.ToArray());
    }

    return result;
}

string[] ignoreCmd = new string[] { };

// string[] args = Environment.GetCommandLineArgs();
if (args.Length < 1)
{
    Console.WriteLine("Input file must be directed at first option");
    return;
}
if (args.Length >= 2)
{
    if(File.Exists(args[1]))
        ignoreCmd = File.ReadAllLines(args[1]);
}

string fileName = args[0];

Console.WriteLine("Target file=" + fileName);

// load latex file. change directory tentatively
string old_path = Directory.GetCurrentDirectory();
Directory.SetCurrentDirectory(Path.GetDirectoryName(fileName));
fileName = Path.GetFileName(fileName);

TexFileLoader tfl0 = new TexFileLoader(fileName, ignoreCmd);
string[] sepText = tfl0.LoadAllParagraphsWithExpansion();

// roleback current directly
Directory.SetCurrentDirectory(old_path);

// make paragraph
List<string[]> paragraphs = SplitArray(sepText, "");

Console.WriteLine("Statistics:");
Console.WriteLine("#paragraphs:"+paragraphs.Count.ToString());
Console.WriteLine("#total lines:" + paragraphs.Sum(o=>o.Length).ToString());
Console.WriteLine("#total lines in tex:" + sepText.Length.ToString());


string outputFP = "text.docx";
// make Word document
using (SimpleWordDocument wordDocument = new SimpleWordDocument(outputFP))
{
    //wordDocument.AddParagraph("moji");
    //wordDocument.AddMathParagraph("a=b+c");

    //var run = wordDocument.BeginParagraph();
    //wordDocument.AddText(run, "moji2");
    //wordDocument.AddMath(run, "d=e+f");
    //wordDocument.EndParagraph();

    // wordDocument.AddParagraph(sepText);
    foreach (string[] paragraph in paragraphs)
    {
        wordDocument.AddParagraphs(paragraph);
    }

}



