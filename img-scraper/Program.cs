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

                        List<HtmlNode> nodes = new List<HtmlNode> { LoadHtmlDocument(uriResult) };
                        List<string> fileNames = new List<string>();
                        nodes = nodes.SelectMany(p => p.SelectNodes("//img")).ToList();
                        int i = 1;
                        Console.WriteLine($"Liczba obrazków na stronie: {nodes.Count}");
                        nodes.ForEach(t =>
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
                                    if (imgUrl != "")
                                    {
                                        Uri uriImg;
                                        if (!(Uri.TryCreate(imgUrl, UriKind.Absolute, out uriImg) && (uriImg.Scheme == Uri.UriSchemeHttp || uriImg.Scheme == Uri.UriSchemeHttps)))
                                        {
                                            if (imgUrl.StartsWith("//"))
                                                imgUrl = "https:" + imgUrl;
                                            else
                                                imgUrl = "https://" + imgUrl;
                                            uriImg = new Uri(imgUrl);
                                        }
                                        row[0] = i.ToString();
                                        row[1] = uriImg.ToString();
                                        row[2] = uriImg.ToString().Substring(uriImg.ToString().LastIndexOf("/") + 1);
                                        row[4] = t.Attributes["alt"].Value;
                                        DownloadImage(subFolder, uriImg, new WebClient());
                                        string filePath = subFolder + uriImg.ToString().Substring(uriImg.ToString().LastIndexOf("/"));
                                        if (File.Exists(filePath))
                                        {
                                            FileInfo fi = new FileInfo(filePath);
                                            row[3] = (fi.Length / 1024f).ToString("N2");
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
                            var q = fileNames.GroupBy(x => x)
                                            .Select(g => new { Value = g.Key, Count = g.Count() });
                            Console.WriteLine($"\r\nZnaleziono obrazków: {nodes.Count}");
                            Console.WriteLine($"Duplikatów: {q.Where(x=>x.Count >1).Count()}");
                            Console.WriteLine($"Pobrano: {i-2}\r\n\r\nKoniec\r\n\r\nWklej kolejny adres strony");
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
    }
}
