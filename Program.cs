using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CommandLine;
using CommandLine.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

namespace LineCount
{
    class Program
    {
        public class Options
        {
            [Value(0, HelpText = "file to filter")]
            public string filePath { get; set; } = ".";

            [Option('m', "match", Required = true, HelpText = "match pattern eg: new|org, new|new")]
            public string matchPatterns { get; set; } = "";

            [Option('b', "prev-module", Required = true, HelpText = "module name before keyword, eg: dtsre or Runtime")]
            public string prevModule { get; set; } = "dtsre";

            [Option('o', "output", Required = false, HelpText = "output as a .xlsx file")]
            public string output { get; set; } = "";
            
            [Option('d', "debug", Required = false, HelpText = "stop at the end")]
            public bool bDebug { get; set; } = false;

            [Option('a', "append", Required = false, HelpText = "write with append mode")]
            public bool bAppend { get; set; } = false;
            
        }
        class Message
        {
            public int LineCount { get; set; }
            public int TokenCount { get; set; }
            public List<FileStart> FileStarts { get; set; }
            public List<string> Code { get; set; }
        }

        class FileStart
        {
            public int LineNumber { get; set; }
            public string FilePath { get; set; }
        }

        static void Filter(string filePath, HashSet<string> filterPatterns, string outputPath, string prevModule)
        {
            var messages = new List<Message>();
            bool bOutputFile = (outputPath.Length > 0);

            string linePattern = @"Found a (\d+) line \((\d+) tokens\) duplication in the following files:";
            string startLinePattern = @"Starting at line (\d+) of (.*)";

            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    Message currentMessage = null;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (Regex.IsMatch(line, linePattern))
                        {
                            //filter the message.
                            if (currentMessage != null)
                            {
                                //Console.WriteLine("===============");

                                bool findPattern(Message message, string pattern)
                                {
                                    foreach (var fileStart in message.FileStarts)
                                    {
                                        if (fileStart.FilePath.Contains(pattern))
                                            return true;
                                    }
                                    return false;
                                };

                                //both flags should be appear
                                bool findAllPatterns(Message message, HashSet<string> patterns)
                                {
                                    bool valid = true;
                                    foreach (var pattern in patterns)
                                    {
                                        valid = valid && (findPattern(currentMessage, pattern));
                                        if (!valid)
                                        {
                                            return false;
                                        }
                                    }
                                    return valid;
                                }

                                bool findAnyPatterns(Message message, HashSet<string> patterns)
                                {
                                    foreach (var pattern in patterns)
                                    {
                                        if (findPattern(currentMessage, pattern))
                                            return true;
                                    }
                                    return false;
                                }

                                //if no filter pattern, all will be kept.
                                bool bKeepAll = (filterPatterns.Count == 0);
                                if (!bKeepAll) //have filter
                                {
                                    bKeepAll = findAllPatterns(currentMessage, filterPatterns);
                                }
                                if (bKeepAll)
                                {
                                    messages.Add(currentMessage);
                                }
                            }

                            var match = Regex.Match(line, linePattern);
                            currentMessage = new Message
                            {
                                LineCount = int.Parse(match.Groups[1].Value),
                                TokenCount = int.Parse(match.Groups[2].Value),
                                FileStarts = new List<FileStart>(),
                                Code = new List<string>()
                            };
                        }
                        else if (Regex.IsMatch(line, startLinePattern))
                        {
                            var match = Regex.Match(line, startLinePattern);
                            currentMessage.FileStarts.Add(new FileStart
                            {
                                LineNumber = int.Parse(match.Groups[1].Value),
                                FilePath = match.Groups[2].Value
                            });
                        }
                        else if(!line.Equals("====================================================================="))
                        {
                            if (currentMessage.Code.Count == 0 && line.Trim().Length == 0) //first blank line
                                continue;
                            currentMessage.Code.Add(line);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading file: {ex.Message}");
            }

            
            int totallines = 0;
            int except = 0;
            foreach(Message msg in messages)
            {
                totallines += msg.LineCount;
                foreach (FileStart fs in msg.FileStarts)
                {
                    if (fs.FilePath.Contains("VulkanStaticStates.h"))
                    {
                        except++;
                        Console.WriteLine($"==========> findExcept: {except}");
                    }
                }
            }
            Console.WriteLine($"==========> total messages: {messages.Count}");
            Console.WriteLine($"==========> total lines: {totallines}");

            if (bOutputFile)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                var fileInfo = new FileInfo(outputPath);

                string worksheetName = "new";

                try
                {
                    if (!fileInfo.Exists)
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add(worksheetName);
                            WriteExcel(worksheet, messages, prevModule, false);
                            //foreach (var sheet in package.Workbook.Worksheets)
                            //{
                            //    sheet.Calculate();
                            //}
                            package.SaveAs(fileInfo);
                        }
                    }
                    else
                    {
                        CloseSpecificExcelFile(outputPath);
                        using (ExcelPackage package = new ExcelPackage(fileInfo))
                        {
                            // 检查是否存在同名工作表
                            var worksheet = package.Workbook.Worksheets[worksheetName];
                            if (worksheet == null)
                            {
                                worksheet = package.Workbook.Worksheets.Add(worksheetName);
                            }

                            WriteExcel(worksheet, messages, prevModule, g_bAppend);

                            //foreach (var sheet in package.Workbook.Worksheets)
                            //{
                            //    sheet.Calculate();
                            //}
                            package.Save();
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("exception:" + e.ToString());
                }
            }
            else
            { 
                foreach (var message in messages)
                {
                    Console.WriteLine($"Found a {message.LineCount} line ({message.TokenCount} tokens) duplication in the following files:");
                    foreach (var fileStart in message.FileStarts)
                    {
                        Console.WriteLine($"Starting at line {fileStart.LineNumber} of {fileStart.FilePath}");
                    }
                    Console.WriteLine();
                }
            }
        }

        static void WriteExcel(ExcelWorksheet worksheet, List<Message> messages, string prevModule = "dtsre", bool appendMode = false)
        {
            //=============清除原来的数据===================
            if (worksheet.Dimension != null && !appendMode)
            {
                worksheet.DeleteRow(1, worksheet.Dimension.End.Row);
            }

            const int CL_LINE = 1;
            const int CL_TOKEN = 2;
            const int CL_OCC = 3;
            const int CL_MODULE = 4;
            const int CL_FILE = 5;
            const int CL_START = 6;
            const int CL_END = 7;
            const int CL_CODE = 8;
            const int CL_FIX = 9;
            const int CL_HASH = 10;

            int row = worksheet.Dimension != null ? (worksheet.Dimension.Rows + 1) : 1;
            if (row == 1)
            {
                worksheet.Cells[row, CL_LINE].Value = "lines";
                worksheet.Cells[row, CL_TOKEN].Value = "token";
                worksheet.Cells[row, CL_OCC].Value = "occurrence";
                worksheet.Cells[row, CL_MODULE].Value = "module";
                worksheet.Cells[row, CL_FILE].Value = "file";
                worksheet.Cells[row, CL_START].Value = "start\n(fist filter)";
                worksheet.Cells[row, CL_END].Value = "end\n(fist filter)";
                worksheet.Cells[row, CL_CODE].Value = "code";
                worksheet.Cells[row, CL_HASH].Value = "code hash";
                worksheet.Cells[row, CL_FIX].Value = "fix";

                worksheet.Row(row).Height = 30;
                worksheet.Row(row).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(row).Style.WrapText = true;

                worksheet.Column(CL_LINE).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_TOKEN).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_OCC).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_MODULE).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_FILE).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_START).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_END).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_FIX).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_HASH).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(CL_CODE).Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                worksheet.Column(CL_MODULE).Width = 15;
                worksheet.Column(CL_START).Width = 12;
                worksheet.Column(CL_END).Width = 12;
                worksheet.Column(CL_FILE).Width = 30;
                worksheet.Column(CL_CODE).Width = 120;

                ++row;
            }
            //
            int DBGMaxLineNum = 0;
            int DBGMaxLineAtRow = 2;
            string DBGMaxFile = "";
            //

            foreach (var message in messages)
            {
                worksheet.Cells[row, CL_LINE].Value = message.LineCount;
                worksheet.Cells[row, CL_TOKEN].Value = message.TokenCount;
                worksheet.Cells[row, CL_OCC].Value = message.FileStarts.Count;

                int lineNum = -1;
                string startFile = "";
                string module = "";
                string file = "";

                var groupedData1 = message.FileStarts.GroupBy(x => x.FilePath);
                var groupedData = groupedData1.ToDictionary(g => g.Key, g => g.Select(x => x.LineNumber)).ToList();

                foreach (var group in groupedData)
                {
                    if (group.Value.Count() > 1)
                    {
                        startFile += $"[{string.Join(",", group.Value)}]    {group.Key}\n";
                    }
                    else
                    {
                        startFile += $"{group.Value.First()}    {group.Key}\n";
                    }
                }

                foreach (FileStart fs in message.FileStarts)
                {
                    //if (fs.FilePath.Contains("VulkanStaticStates.h"))
                    //{
                    //    Console.WriteLine($"=======> except at line: {fs.LineNumber}, {fs.FilePath}");
                    //}
                    if (lineNum == -1 && fs.FilePath.Contains(firstFilter))
                    {
                        lineNum = fs.LineNumber;
                        module = ExtractModuleString(fs.FilePath, prevModule, firstFilter);
                        file = ExtractFileString(fs.FilePath, module);
                    }
                }
                worksheet.Cells[row, CL_FILE].Value = file;
                worksheet.Cells[row, CL_START].Value = lineNum;
                worksheet.Cells[row, CL_END].Value = lineNum + message.LineCount;
                worksheet.Cells[row, CL_MODULE].Value = module;


                string rawCode = "";
                string code = "\n";
                foreach (string codeLine in message.Code)
                {
                    code += $"{lineNum:D4}    {codeLine.Replace("\t", "    ")}\n";
                    rawCode += codeLine;
                    lineNum++;
                }

                if (startFile.Length > DBGMaxLineNum)
                {
                    DBGMaxLineNum = startFile.Length;
                    DBGMaxLineAtRow = row;
                    DBGMaxFile = startFile;
                }

                worksheet.Cells[row, CL_CODE].Value = (startFile + code).TrimEnd('\n');
                worksheet.Cells[row, CL_CODE].Style.WrapText = true; // Enable wrap text for the code cell
                worksheet.Row(row).Height = 200;

                worksheet.Cells[row, CL_HASH].Value = GenerateShortHash(rawCode);
                row++;
            }
            Console.WriteLine($"========>total row: {row}");

            //===============重新排序=====================
            // 将数据加载到一个 List 中进行排序
            var totalRow = worksheet.Dimension.Rows;

            var data = new List<(int lines, int token, int occ, string module, string file, int start, int end, string code, string hash)>(totalRow);
            for (row = 2; row <= totalRow; row++)
            {
                data.Add((
                    lines: Convert.ToInt32(worksheet.Cells[row, CL_LINE].Value),
                    token: Convert.ToInt32(worksheet.Cells[row, CL_TOKEN].Value),
                    occ: Convert.ToInt32(worksheet.Cells[row, CL_OCC].Value),
                    module: worksheet.Cells[row, CL_MODULE].Value.ToString(),
                    file: worksheet.Cells[row, CL_FILE].Value.ToString(),
                    start: Convert.ToInt32(worksheet.Cells[row, CL_START].Value),
                    end: Convert.ToInt32(worksheet.Cells[row, CL_END].Value),
                    code: worksheet.Cells[row, CL_CODE].Value.ToString(),
                    hash: worksheet.Cells[row, CL_HASH].Value.ToString()
                ));
            }

            // 根据 module, file, start 排序
            var sortedData = data.OrderBy(x => x.module).ThenBy(x => x.file).ThenBy(x => x.start).ToList();

            // 将排序后的数据写回到工作表
            int startRow = 2;
            foreach (var item in sortedData)
            {
                worksheet.Cells[startRow, CL_LINE].Value = item.lines;
                worksheet.Cells[startRow, CL_TOKEN].Value = item.token;
                worksheet.Cells[startRow, CL_OCC].Value = item.occ;
                worksheet.Cells[startRow, CL_MODULE].Value = item.module;
                worksheet.Cells[startRow, CL_FILE].Value = item.file;
                worksheet.Cells[startRow, CL_START].Value = item.start;
                worksheet.Cells[startRow, CL_END].Value = item.end;
                worksheet.Cells[startRow, CL_CODE].Value = item.code;

                string fixFormuler = $"=IF(ISERROR(VLOOKUP(J{startRow},base!J:J,1,FALSE)),1,0)";
                worksheet.Cells[startRow, CL_FIX].Formula = fixFormuler;

                worksheet.Cells[startRow, CL_HASH].Value = item.hash;
                startRow++;
            }

            worksheet.Calculate();
            worksheet.View.FreezePanes(2, 1);
        }

        static void CloseSpecificExcelFile(string filePath)
        {
            ExcelInterop.Application excelApp = null;
            try
            {
                excelApp = (ExcelInterop.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                // 如果Excel未运行，直接返回
                return;
            }

            if (excelApp != null)
            {
                foreach (ExcelInterop.Workbook workbook in excelApp.Workbooks)
                {
                    if (workbook.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Close(false);
                        break;
                    }
                }
                Marshal.ReleaseComObject(excelApp);
            }
        }

        public static string GenerateShortHash(string input)
        {
            // 使用SHA256生成哈希值
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));

                // 将哈希值转换为短字符串
                StringBuilder hashBuilder = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    hashBuilder.Append(hashBytes[i].ToString("X2"));
                }

                // 返回短的哈希字符串，可以截取一部分
                return hashBuilder.ToString().Substring(0, 8); // 这里截取前8个字符
            }
        }
        static string ExtractModuleString(string path, string primaryKeyword, string secondaryKeyword)
        {
            // 尝试根据主要关键字提取
            string result = ExtractAfterKeyword(path, primaryKeyword);

            // 如果主要关键字未找到，则尝试根据次要关键字提取
            if (result == null)
            {
                result = ExtractAfterKeyword(path, secondaryKeyword);
            }

            return result ?? "";
        }
        static string ExtractFileString(string path, string module)
        {
            //F:\\newengine\dtsre\RenderEngine\src\xxx.cpp
            int keywordIndex = path.IndexOf(module);
            string result = "";
            if (keywordIndex != -1)
            {
                int startIndex = keywordIndex + module.Length + 1;
                startIndex = path.IndexOf('\\', startIndex) + 1;

                result = path.Substring(startIndex);
                //result = result.Replace("Private\\", "");
                //result = result.Replace("Public\\", "");
            }

            return result ?? "";
        }
        static string ExtractAfterKeyword(string path, string keyword)
        {
            int keywordIndex = path.IndexOf(keyword);

            if (keywordIndex != -1)
            {
                int startIndex = keywordIndex + keyword.Length + 1;
                int endIndex = path.IndexOf('\\', startIndex);

                if (endIndex != -1)
                {
                    return path.Substring(startIndex, endIndex - startIndex);
                }
                else
                {
                    return path.Substring(startIndex);
                }
            }

            return null; // 返回 null 如果关键字未找到
        }
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(Run)
                .WithNotParsed(HandleParseError);
            if (bDebug)
            {
                Console.ReadLine();
            }
        }
        static void HandleParseError(IEnumerable<Error> errs)
        {
            
        }

        static bool g_bAppend = false;
        static string g_prevModule = "dtsre";
        static void Run(Options option)
        {
            bool bDebug = option.bDebug;
            string filePath = Path.GetFullPath(option.filePath);

            g_bAppend = option.bAppend;
            g_prevModule = option.prevModule;

            //如果代码有两个来源，A和B
            //--- -- 无pattern，保留所有。
            //A|A -- A only
            //A|B -- contains B only
            //B|B -- contains B only
            //Test();
            string[] rawPatterns = option.matchPatterns.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            string outputFile = option.output;
            if (!outputFile.EndsWith(".xlsx"))
            {
                outputFile+=(".xlsx");
            }
            outputFile = Path.GetFullPath(outputFile);

            HashSet<string> matchPatterns = new HashSet<string>();
            foreach (string pattern in rawPatterns)
            {
                matchPatterns.Add(pattern);
            }
            if (rawPatterns.Length > 0)
                firstFilter = rawPatterns[0];

            Filter(filePath, matchPatterns, outputFile, g_prevModule);
        }

        static bool bDebug = false;
        static string firstFilter = "";
        static void Test()
        {
            string[] patternArray =
            {
                "vulkanengine|vulkanengine",
                "vulkanengine|org",
                "",
                "org|org",
                "joooe"
            };
            foreach (string patternStr in patternArray)
            {
                string[] rawPatterns = patternStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                HashSet<string> matchPatterns = new HashSet<string>();
                foreach (string pattern in rawPatterns)
                {
                    matchPatterns.Add(pattern);
                }
                Console.WriteLine("======={0}========", patternStr);
                Console.WriteLine(matchPatterns.Count());
                foreach (string pattern in matchPatterns)
                {
                    Console.WriteLine(pattern);
                }
                Console.WriteLine("==");
            }
        }

    }
}
