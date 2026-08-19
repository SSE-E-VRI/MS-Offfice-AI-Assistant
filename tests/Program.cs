using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Tests
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            Console.WriteLine("=================================================");
            Console.WriteLine("MS Office AI Assistant - Standalone Test Suite   ");
            Console.WriteLine("=================================================");
            Console.WriteLine();

            int passed = 0;
            int failed = 0;
            var suites = new List<Tuple<string, Action>>
            {
                Tuple.Create<string, Action>("WordDocumentContextBuilderTests", WordDocumentContextBuilderTests.RunAll),
                Tuple.Create<string, Action>("SpreadsheetActionParserTests", SpreadsheetActionParserTests.RunAll),
                Tuple.Create<string, Action>("PowerPointActionParserTests", PowerPointActionParserTests.RunAll),
                Tuple.Create<string, Action>("ActionAuditStoreTests", ActionAuditStoreTests.RunAll),
                Tuple.Create<string, Action>("ChatMessageSerializationTests", ChatMessageSerializationTests.RunAll)
            };

            foreach (var suite in suites)
            {
                Console.Write("Running " + suite.Item1 + "... ");
                try
                {
                    suite.Item2();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("[PASS]");
                    Console.ResetColor();
                    passed++;
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[FAIL]");
                    Console.WriteLine("  Error: " + ex.Message);
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine("  Inner: " + ex.InnerException.Message);
                    }
                    Console.ResetColor();
                    failed++;
                }
            }

            Console.WriteLine();
            Console.WriteLine("=================================================");
            if (failed == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(string.Format("ALL TEST SUITES PASSED ({0}/{1})", passed, passed + failed));
                Console.ResetColor();
                return 0;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("TEST FAILURES OCCURRED: {0} passed, {1} failed", passed, failed));
                Console.ResetColor();
                return 1;
            }
        }
    }
}
