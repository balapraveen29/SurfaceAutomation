using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SurfaceAutomation
{
    public class clsPdfParser
    {
        private static int _noOfCharsToKeep = 15;

        public bool ExtractText(string inFileName, string outFileName)
        {
            StreamWriter outFile = null;
            try
            {
                PdfReader reader = new PdfReader(inFileName);
                outFile = new StreamWriter(outFileName, false, System.Text.Encoding.UTF8);
                Console.WriteLine("Processing: ");
                int totalLen = 68;
                float charUnit = ((float)totalLen) / (float)reader.NumberOfPages;
                int totalWritten = 0;
                float curUnit = 0;

                for(int page = 1; page <= reader.NumberOfPages; page++)
                {
                    outFile.Write(ExtractTextFromPDFBytes(reader.GetPageContent(page)) + "");

                    if(charUnit >= 1.0f)
                    {
                        for(int i = 0; i<(int)charUnit; i++)
                        {
                            Console.WriteLine("#");
                            totalWritten++;
                        }
                    }
                    else
                    {
                        curUnit += charUnit;
                        if(curUnit >= 1.0f)
                        {
                            for(int i=0;i<(int)charUnit; i++)
                            {
                                Console.WriteLine("#");
                                totalWritten++;
                            }
                            curUnit = 0;
                        }
                    }
                }
                if(totalWritten < totalLen)
                {
                    for(int i = 0; i < (totalLen - totalWritten); i++)
                    {
                        Console.WriteLine("#");
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if(outFile != null)
                outFile.Close();
            }
        }

        public string ExtractTextFromPDFBytes(byte[] input)
        {
            if (input == null || input.Length == 0) return "";
            try
            {
                string resultString = "";

                bool inTextObject = false;
                bool nextLiteral = false;
                int bracketDepth = 0;
                char[] previousCharacters = new char[_noOfCharsToKeep];
                for (int j = 0; j < _noOfCharsToKeep; j++) previousCharacters[j] = ' ';

                for(int i =0; i < input.Length; i++)
                {
                    char c = (char)input[i];
                    if (input[i] == 213)
                        c = "'".ToCharArray()[0];
                    if(inTextObject)
                    {
                        if(bracketDepth == 0)
                        {
                            if(CheckToken(new string[] {"TD", "Td"}, previousCharacters))
                            {
                                resultString = "\n\r";
                            }
                            else
                            {
                                if(CheckToken(new string[] {"'", "T*", "\""}, previousCharacters))
                                {
                                    resultString += "\n";
                                }
                                else
                                {
                                    if(CheckToken(new string[] {"Tj"}, previousCharacters))
                                    {
                                        resultString += " ";
                                    }
                                }
                            }
                        }
                        if(bracketDepth == 0 && CheckToken(new string[] {"ET"}, previousCharacters))
                        {
                            intTextObject = false;
                            resultString += " ";
                        }
                        else
                        {
                            if((c == ')') && (bracketDepth == 1) && (!nextLiteral))
                            {
                                bracketDepth = 0;
                            }
                            else
                            {
                                if(bracketDepth == 1)
                                {
                                    if(c == '\\' && !nextLiteral)
                                    {
                                        resultString += c.ToString();
                                        nextLiteral = true;
                                    }
                                    else
                                    {
                                        if(((c > ' ') && (c <= '~')) || ((c >= 128) && (c < 255)))
                                        {
                                            resultString += c.ToString();
                                        }
                                        nextLiteral = false;
                                    }
                                }
                            }
                        }
                    }
                    for (int j = 0; j < _noOfCharsToKeep - 1; j++)
                    {
                        previousCharacters[j] = previousCharacters[j + 1];
                    }
                    previousCharacters[_noOfCharsToKeep - 1] = c;
                    if (!inTextObject && CheckToken(new string[] { "BT" }, previousCharacters))
                    {
                        inTextObject = true;
                    }
                }
                return CleanupContent(resultString);
            }
            catch
            {
                return "";
            }
        }
        private string CleanupContent(string text)
        {
            string[] patterns = { @"\\\(", @"\\\)", @"\\226", @"\\222", @"\\223", @"\\224", @"\\340", @"\\342", @"\\344", @"\\300", @"\\302", @"\\304", @"\\351", @"\\350", @"\\352", @"\\353", @"\\311", @"\\310", @"\\312", @"\\313", @"\\362", @"\\364", @"\\366", @"\\322", @"\\324", @"\\326", @"\\354", @"\\356", @"\\357", @"\\314", @"\\316", @"\\317", @"\\347", @"\\307", @"\\371", @"\\373", @"\\374", @"\\331", @"\\333", @"\\334", @"\\256", @"\\231", @"\\253", @"\\273", @"\\251", @"\\221" };
            string[] replace = { "(", ")", "-", "'", "\"", "à", "â", "ä", "è", "ë", "ê", "é", "ì", "í", "î", "ï", "ò", "ó", "ô", "ö", "ù", "ú", "û", "ü", "ç", "À", "Â", "Ä", "È", "É", "Ê", "Ë", "Ì", "Í", "Î", "Ï", "Ò", "Ó", "Õ", "Ö", "Ù", "Ú", "Û", "Ü", "Ç", "»", "«", "©", "™", "®" };

            for(int i = 0; i < patterns.Length; i++)
            {
                string regExPattern = patterns[i];
                Regex regex = new Regex(regExPattern, RegexOptions.IgnoreCase);
                text = regex.Replace(text, replace[i]);
            }
            return text;
        }
        private bool CheckToken(string[] tokens, char[] recent)
        {
            foreach (string token in tokens)
            {
                if ((recent[_noOfCharsToKeep - 3] == token[0]) && (recent[_noOfCharsToKeep - 2] == token[1]) && 
                    ((recent[_noOfCharsToKeep - 1] == ' ') || (recent[_noOfCharsToKeep - 1] == 0x0d) || (recent[_noOfCharsToKeep - 1] == 0x0a)) &&
                    ((recent[_noOfCharsToKeep - 4] == ' ') || (recent[_noOfCharsToKeep - 4] == 0x0d) || (recent[_noOfCharsToKeep - 4] == 0x0a)))
                {
                    return true;
                }

            }
            return false;
        }
    }
}
