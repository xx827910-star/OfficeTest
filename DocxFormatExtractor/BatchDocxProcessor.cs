using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocxFormatExtractor
{
    internal static class BatchDocxProcessor
    {
        public static void Run(string inputDirectory, string outputDirectory)
        {
            if (string.IsNullOrWhiteSpace(inputDirectory))
            {
                Console.WriteLine("批处理输入目录不能为空");
                return;
            }

            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"输入目录不存在: {inputDirectory}");
                return;
            }

            var docFiles = Directory.GetFiles(inputDirectory, "*.docx", SearchOption.TopDirectoryOnly);

            if (docFiles.Length == 0)
            {
                Console.WriteLine($"目录中未找到 docx 文件: {inputDirectory}");
                return;
            }

            Directory.CreateDirectory(outputDirectory);

            int successCount = 0;
            int failureCount = 0;
            var failureMessages = new List<string>();

            Console.WriteLine("批处理任务开始...");
            Console.WriteLine($"输入目录: {inputDirectory}");
            Console.WriteLine($"输出目录: {outputDirectory}");
            Console.WriteLine($"待处理文件数: {docFiles.Length}");
            Console.WriteLine("输出格式: 仅 JSON");
            Console.WriteLine();

            foreach (string docPath in docFiles.OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    string baseName = Path.GetFileNameWithoutExtension(docPath) + "_format_output";
                    var result = EnhancedProgram.ProcessDocument(docPath, outputDirectory, "json", baseName);

                    successCount++;
                    Console.WriteLine($"✅ 成功: {docPath}");
                    if (!string.IsNullOrEmpty(result.JsonOutputPath))
                    {
                        Console.WriteLine($"   JSON 输出 -> {result.JsonOutputPath}");
                    }
                }
                catch (Exception ex)
                {
                    failureCount++;
                    string message = $"❌ 失败: {docPath} - {ex.Message}";
                    failureMessages.Add(message);
                    Console.WriteLine(message);
                }
            }

            Console.WriteLine();
            Console.WriteLine($"批处理完成，总计 {docFiles.Length} 个文件，成功 {successCount} 个，失败 {failureCount} 个。");

            if (failureMessages.Count > 0)
            {
                Console.WriteLine("失败详情：");
                foreach (string message in failureMessages)
                {
                    Console.WriteLine(message);
                }
            }
        }
    }
}
