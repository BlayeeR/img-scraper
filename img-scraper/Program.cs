﻿using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.AccessControl;
using OfficeOpenXml;

namespace img_scraper
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Wklej adres strony:");
            string url, folder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\images\";
            List<string> exts = new List<string>()
            {
                ".jpg", ".png"
            };
            try
            {
                if (Directory.Exists(folder) && !CheckFolderPermission(folder))
                {
                    throw new UnauthorizedAccessException();
                }
                else
                {
                    Directory.CreateDirectory(folder);
                }
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Brak uprawnień\r\nUruchom jako administrator albo zmień lokalizację programu");
            }
            
            while (!String.IsNullOrEmpty(url = Console.ReadLine()))
            {
                if (url[0] == '\n')
                    break;
                Uri uriResult;
                if (Uri.TryCreate(url, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps))
                {
                    string subFolder = folder + uriResult.Host + @"\";
                    for (int i = 0; i < uriResult.Segments.Length; i++)
                    {
                        subFolder += uriResult.Segments[i] + @"\";
                    }
                    Directory.CreateDirectory(subFolder);

                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        if (excel.Workbook.Worksheets["Images"] == null)
                            excel.Workbook.Worksheets.Add("Images");
                        var worksheet = excel.Workbook.Worksheets["Images"];
                        List<string[]> rows = new List<string[]>()
                        {
                            new string[] {"Lp", "Uri", "File Name", "Size(kB)", "Alt" }
                        };

                        List<HtmlNode> nodes = new List<HtmlNode> { LoadHtmlDocument(uriResult) };
                        nodes = nodes.SelectMany(p => p.SelectNodes("//img")).ToList();
                        int i = 1;
                        Console.WriteLine($"Liczba obrazków na stronie: {nodes.Count}");
                        nodes.ForEach(t =>
                            {
                                string[] row = new string[5];
                                foreach (string ext in exts)
                                {
                                    if (t.Attributes["src"].Value.Contains(ext))
                                    {
                                        string imgUrl = t.Attributes["src"].Value.Substring(0, t.Attributes["src"].Value.LastIndexOf(ext) + ext.Length);
                                        Uri uriImg;
                                        if (!(Uri.TryCreate(imgUrl, UriKind.Absolute, out uriImg) && (uriImg.Scheme == Uri.UriSchemeHttp || uriImg.Scheme == Uri.UriSchemeHttps)))
                                        {
                                            if (imgUrl.StartsWith("//"))
                                                imgUrl = "https:" + imgUrl;
                                            else
                                                imgUrl = "https://" + imgUrl;
                                            uriImg = new Uri(imgUrl);
                                        }
                                        Console.WriteLine(uriImg.ToString());
                                        row[0] = i.ToString();
                                        row[1] = uriImg.ToString().Substring(0, uriImg.ToString().LastIndexOf("/"));
                                        row[2] = uriImg.ToString().Substring(uriImg.ToString().LastIndexOf("/") + 1);
                                        DownloadImage(subFolder, uriImg, new WebClient());
                                        rows.Add(row);
                                        i++;
                                    }
                                }
                            });

                        string headerRange = "A1:" + Char.ConvertFromUtf32(rows[0].Length + 64) + "1";
                        worksheet.Cells[headerRange].LoadFromArrays(rows);

                        excel.SaveAs(new FileInfo(subFolder + "images.xlsx"));
                    }
                }
                else
                    Console.WriteLine("Zły adres(tylko http/https)");
            }
        }

        private static HtmlNode LoadHtmlDocument(Uri uri)
        {
            var doc = new HtmlDocument();
            var wc = new WebClient();
            doc.LoadHtml(wc.DownloadString(uri));

            var documentNode = doc.DocumentNode;
            return documentNode;
        }

        private static void DownloadImage(string folderImagesPath, Uri url, WebClient webClient)
        {
            try
            {
                webClient.DownloadFile(url, Path.Combine(folderImagesPath, Path.GetFileName(url.ToString())));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static bool CheckFolderPermission(string folderPath)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(folderPath);
            try
            {
                DirectorySecurity dirAC = dirInfo.GetAccessControl(AccessControlSections.All);
                return true;
            }
            catch (PrivilegeNotHeldException)
            {
                return false;
            }
        }
    }
}