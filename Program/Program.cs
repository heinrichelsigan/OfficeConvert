﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OfficeConvert;
using System.Reflection;
using static System.Net.Mime.MediaTypeNames;

namespace Program
{
    public class Program
    {
        private static String inputFile;
        private static String outputFile;
        private static Hashtable options;

        private static void log(String message)
        {
            if (!options.ContainsKey("verbose")) return;
            Console.WriteLine(message);
        }

        private static void init()
        {
            options = new Hashtable();

            inputFile = null;
            outputFile = null;
        }

        public static void Main(string[] args)
        {
            init();

            for (int i = 0; i < args.Length; i++)
            {
                String arg = args[i];
                if (arg.Substring(0, 2) == "--")
                {
                    String[] values = arg.Split('=');
                    String option = values[0].Substring(2);
                    if (arg.Contains("="))
                    {
                        options[option] = values[1];
                    }
                    else
                    {
                        options[option] = true;
                    }
                }
                else if (arg.Substring(0, 1) == "-")
                {
                    String[] values = arg.Split('=');
                    String option = values[0].Substring(1);
                    if (arg.Contains("="))
                    {
                        options[option] = values[1];
                    }
                    else
                    {
                        options[option] = true;
                    }
                }
                else if (inputFile == null)
                {
                    inputFile = arg;
                }
                else if (outputFile == null)
                {
                    outputFile = arg;
                }
            }

            if (options.ContainsKey("v") || options.ContainsKey("version"))
            {
                Console.WriteLine("Version: " + Assembly.GetExecutingAssembly().GetName().Version);
                Environment.Exit(0);
                return;
            }


            if (inputFile == null || outputFile == null)
            {
                Console.WriteLine("No input File or no output file");
                Environment.Exit(1);
            }

            Regex fileMatch = new Regex(@"\.(((ppt|pps|pot|do[ct]|xls|xlt)[xm]?)|od[cpt]|rtf|csv|vsd[xm]?|pub|msg|eml|vcf|ics|mpp|svg|txt|html?)$", RegexOptions.IgnoreCase);
            Match extMatch = fileMatch.Match(inputFile);
            if (!extMatch.Success)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher or Visio");
                Environment.Exit(1);
            }

            if (!outputFile.ToLower().EndsWith(".pdf"))
            {
                Console.WriteLine($"Output file {outputFile} will not be accepted. It must end with \".pdf\"!");
                Environment.Exit(2);
            }

            String extname = extMatch.Groups[1].ToString().ToLower();

            log("Input: " + inputFile);
            log("Output: " + outputFile);
            try
            {

                switch (extname)
                {
                    case "rtf":
                    case "odt":
                    case "doc":
                    case "dot":
                    case "docx":
                    case "dotx":
                    case "docm":
                    case "dotm":
                        // Word
                        new WordConverter().Convert(inputFile, outputFile);
                        break;
                    case "csv":
                    case "odc":
                    case "xls":
                    case "xlsx":
                    case "xlt":
                    case "xltx":
                    case "xlsm":
                    case "xltm":
                        // Excel
                        new ExcelConverter().Convert(inputFile, outputFile);
                        break;
                    case "odp":
                    case "ppt":
                    case "pptx":
                    case "pptm":
                    case "pot":
                    case "potm":
                    case "potx":
                    case "pps":
                    case "ppsx":
                    case "ppsm":
                        // Powerpoint
                        new PowerPointConverter().Convert(inputFile, outputFile);
                        break;
                    case "msg":
                    case "eml":
                        new OutlookConverter().Convert(inputFile, outputFile);
                        break;
                    case "vsd":
                    case "vsdx":
                        new VisioConverter().Convert(inputFile, outputFile);
                        break;
                }
            }
            catch (ConvertException e)
            {
                Console.WriteLine(e.Message);
                Environment.Exit(1);
            }


        }
    }
}
