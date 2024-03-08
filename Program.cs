using System.Text;
using ExcelDataReader;

namespace Xlsx2Json
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
#if NET5_0_OR_GREATER
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
            if (args.Length < 2)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"LogError:参数错误");
                Console.WriteLine($"LogError:程序被迫退出，请修正错误后重试");
                Console.ForegroundColor = ConsoleColor.White;
                Environment.Exit(0);
                return;
            }
            var excelPath = args[0];
            var savePath = args[1];
            
            using (var stream = File.Open(excelPath,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    var sheet = dataSet.Tables[0];
                    var columns = sheet.Columns.Count;
                    var rows = sheet.Rows.Count;

                    var dic = new Dictionary<string, StringBuilder>();
                    for (int col = 1; col < columns; col++)
                    {
                        // 描述
                        var mulKey = sheet.Rows[0][col].ToString();
                        // 语种
                        var mulLang = sheet.Rows[1][col].ToString();
                        // 占位
                        var mulNull = sheet.Rows[2][col].ToString();
                        
                        if (mulLang == null)
                        {
                            return;
                        }
                        
                        var stringBuilder = new StringBuilder();
                        stringBuilder.AppendLine("{");
                        for (int row = 3; row < rows; row++)
                        {
                            var key = sheet.Rows[row][0].ToString();
                            var target = sheet.Rows[row][col];
                            stringBuilder.AppendLine($"\t\"{key}\":\"{target}\",");
                        }

                        stringBuilder.AppendLine("}");
                        dic.TryAdd(mulLang,stringBuilder);
                    }

                    // 文件写出
                    foreach (var stringBuilder in dic)
                    {
                        try
                        {
                            var saveFilePath = Path.Combine(savePath, $"{stringBuilder.Key}.json");
                            var dicName = Path.GetDirectoryName(saveFilePath);
                            if (!Directory.Exists(dicName))
                            {
                                Directory.CreateDirectory(dicName);
                            }

                            StreamWriter writer = new StreamWriter(saveFilePath, false, new UTF8Encoding(false));
                            writer.Write(stringBuilder.Value);
                            writer.Flush();
                            writer.Close();
                        }
                        catch (Exception e)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"LogException:{e}");
                            Console.WriteLine($"LogException:程序被迫退出，请修正错误后重试");
                            Console.ForegroundColor = ConsoleColor.White;
                            Environment.Exit(0);
                        }
                    }

                }
            }
        }
    }
}