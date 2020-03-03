using HtmlAgilityPack;
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
            string url, folder = Directory.GetCurrentDirectory()+ @"\images\";
            List<string> exts = new List<string>()
            {
                ".jpg", ".png"
            };
            try
            {
                Directory.CreateDirectory(folder);
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Brak uprawnień\r\nUruchom jako administrator albo zmień lokalizację programu");
                return;
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
                            new string[] {"Lp", "Url", "File Name", "Size(kB)", "Alt" }
                        };

                        HtmlNode html = LoadHtmlDocument(uriResult);
                        List<HtmlNode> nodes = new List<HtmlNode>();
                        if (html != null)
                            nodes.Add(html);
                        List<HtmlNode> imgNodes = new List<HtmlNode>();

                        List<string> fileNames = new List<string>();
                        imgNodes.AddRange(nodes.SelectMany(p => p.SelectNodes(".//img") ?? new HtmlNodeCollection(null)));
                        imgNodes.AddRange(nodes.SelectMany(p => p.SelectNodes(".//picture//source")?? new HtmlNodeCollection(null)));
                        int i = 1, duplicates = 0;
                        Console.WriteLine($"Liczba obrazków na stronie: {imgNodes.Count}");
                        imgNodes.ForEach(t =>
                            {
                            string[] row = new string[5];
                            foreach (string ext in exts)
                            {
                                string imgUrl = "";
                                if (t.Attributes["src"] != null && t.Attributes["src"].Value.Contains(ext))
                                {
                                    imgUrl = t.Attributes["src"].Value.Substring(0, t.Attributes["src"].Value.LastIndexOf(ext) + ext.Length);
                                }
                                else if (t.Attributes["data-src-pc"] != null && t.Attributes["data-src-pc"].Value.Contains(ext))
                                {
                                    imgUrl = t.Attributes["data-src-pc"].Value.Substring(0, t.Attributes["data-src-pc"].Value.LastIndexOf(ext) + ext.Length);
                                }
                                else if (t.Attributes["srcset"] != null && t.Attributes["srcset"].Value.Contains(ext))
                                {
                                    imgUrl = t.Attributes["srcset"].Value.Substring(0, t.Attributes["srcset"].Value.LastIndexOf(ext) + ext.Length);
                                }
                                if (imgUrl != "")
                                {
                                    Uri uriImg;
                                    if (!(Uri.TryCreate(imgUrl, UriKind.Absolute, out uriImg) && (uriImg.Scheme == Uri.UriSchemeHttp || uriImg.Scheme == Uri.UriSchemeHttps)))
                                    {
                                        if (imgUrl.StartsWith("//"))
                                            imgUrl = "https:" + imgUrl;
                                        else if (imgUrl.StartsWith("/"))
                                            imgUrl = "https:/" + imgUrl;
                                        else
                                            imgUrl = "https://" + imgUrl;
                                        try
                                        {
                                            uriImg = new Uri(imgUrl);
                                        }
                                        catch (Exception e)
                                        {
                                            Console.WriteLine(e.Message);
                                            continue;
                                        }
                                    }

                                    row[0] = i.ToString();
                                    row[1] = uriImg.ToString();
                                    row[2] = uriImg.ToString().Substring(uriImg.ToString().LastIndexOf("/") + 1);
                                    row[4] = t.Attributes["alt"] != null ? t.Attributes["alt"].Value : "";

                                    if(fileNames.Contains(row[2]))
                                    {
                                        duplicates++;
                                        continue;
                                    }

                                    if (!DownloadImage(subFolder, uriImg, new WebClient()))
                                        continue;

                                    string filePath = subFolder + uriImg.ToString().Substring(uriImg.ToString().LastIndexOf("/"));
                                    if (File.Exists(filePath))
                                    {
                                        FileInfo fi = new FileInfo(filePath);
                                        row[3] = Math.Round(fi.Length / 1024f).ToString();
                                        Console.WriteLine($"Saving as {fi.FullName}");
                                        fileNames.Add(fi.Name);
                                    }
                                    rows.Add(row);
                                    i++;
                                }
                            }
                        });

                        string headerRange = "A1:" + Char.ConvertFromUtf32(rows[0].Length + 64) + "1";
                        worksheet.Cells[headerRange].LoadFromArrays(rows);
                        worksheet.Cells[$"A1:{'A'+ i}{i}"].AutoFitColumns();

                        try
                        {
                            excel.SaveAs(new FileInfo(subFolder + "images.xlsx"));
                            Console.WriteLine($"\r\nZnaleziono obrazków: {imgNodes.Count}");
                            Console.WriteLine($"Duplikatów: {duplicates}");
                            Console.WriteLine($"Pobrano: {i-1}\r\n\r\nKoniec\r\n\r\nWklej kolejny adres strony");
                        }
                        catch
                        {
                            Console.WriteLine("Nie można nadpisać pliku xlsx, prawdopodobnie jest używany przez inny proces");
                        }
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
            HtmlNode documentNode = null;
            try
            {
                doc.LoadHtml(wc.DownloadString(uri));
                documentNode = doc.DocumentNode;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return documentNode;
        }

        private static bool DownloadImage(string folderImagesPath, Uri url, WebClient webClient)
        {
            try
            {
                webClient.DownloadFile(url, Path.Combine(folderImagesPath, Path.GetFileName(url.ToString())));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }
    }
}

