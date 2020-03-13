using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.AccessControl;
using OfficeOpenXml;
using System.Text;

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
            List<string> attributes = new List<string>()
            {
                "data-media-s4", "data-media-s3", "data-media-s2", "data-media-s1", "data-src-mobile", "srcset", "data-src-pc", "src"
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
                            new string[] {"Lp", "Url", "File name", "File extension", "Size(kB)", "HTML tag", "attribute", "Alt" }
                        };

                        HtmlNode html = LoadHtmlDocument(uriResult);

                        List<HtmlNode> nodes = new List<HtmlNode>();
                        if (html != null)
                            nodes.Add(html);
                        List<string> fileNames = new List<string>();

                        List<ImageInfo> images = GetImagesInfo(nodes, exts, attributes, subFolder);
                        int i = 1, duplicates = 0;
                        foreach(ImageInfo image in images)
                        {
                            string[] row = new string[8];

                            row[0] = i.ToString();
                            row[1] = image.Uri.ToString();
                            row[2] = image.FileInfo.Name.Substring(0, image.FileInfo.Name.LastIndexOf("."));
                            row[3] = image.FileInfo.Extension;
                            row[5] = image.Name;
                            row[6] = image.Attribute;
                            row[7] = image.Alt;
                            if(fileNames.Contains(image.FileInfo.Name))
                            {
                                duplicates++;
                            }
                            else
                            {
                                Console.WriteLine($"Saving as {image.FileInfo.FullName}");
                                if (!DownloadImage(image.FileInfo.Directory.ToString(), image.Uri, new WebClient()))
                                    continue;
                            }
                            if (File.Exists(image.FileInfo.FullName))
                            {
                                row[4] = Math.Round(image.FileInfo.Length / 1024f).ToString();
                                fileNames.Add(image.FileInfo.Name);
                            }
                            rows.Add(row);
                            i++;
                        };

                        string headerRange = "A1:" + Char.ConvertFromUtf32(rows[0].Length + 64) + "1";
                        worksheet.Cells[headerRange].LoadFromArrays(rows);
                        worksheet.Cells[$"A1:{'A'+ Char.ConvertFromUtf32(rows[0].Length + 64)}{i+1}"].AutoFitColumns();

                        int groupIndex = 0;
                        foreach (var group in images.GroupBy(x=>x.Parent))
                        {
                            if(groupIndex % 2 == 1)
                            {
                                int min = 0, max = 0;
                                min = group.Min(x => images.IndexOf(x))+2;
                                max = group.Max(x => images.IndexOf(x))+2;
                                var test = $"A{min}:{'A' + Char.ConvertFromUtf32(rows[0].Length + 64)}{max}";
                                worksheet.Cells[$"A{min}:{'A' + Char.ConvertFromUtf32(rows[0].Length + 64)}{max}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[$"A{min}:{'A' + Char.ConvertFromUtf32(rows[0].Length + 64)}{max}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            }
                            groupIndex += 1;
                        }

                        string fileName = "";
                        for (int j = 0; j < uriResult.Segments.Length; j++)
                        {
                            fileName += uriResult.Segments[j];
                        }
                        fileName = fileName.Replace("/", "-");
                        if (fileName.Length > 100)
                            fileName = fileName.Substring(0, 100);

                        try
                        {
                            excel.SaveAs(new FileInfo(Directory.GetCurrentDirectory() + "/Images-cc"+fileName+".xlsx"));
                            Console.WriteLine($"\r\n\r\nLiczba obrazków na stronie(bez <noscript></noscript>): {images.Count}");
                            Console.WriteLine($"Liczba tagów <picture>: {images.Where(x => x.Name == "source").GroupBy(x => x.Parent).Select(x => x.First()).Count()}");
                            Console.WriteLine($"Liczba tagów <img>: {images.Where(x => x.Name == "img" && x.Parent.Name != "picture").GroupBy(x => x.Parent).Select(x => x.First()).Count()}");

                            Console.WriteLine($"\r\nZapisano do .xlsx: { i - 1 }");
                            Console.WriteLine($"Duplikatów: {duplicates}\r\n");
                            Console.WriteLine($"Pobrano plików: {fileNames.GroupBy(x => x).Select(x => x.First()).Count()}\r\n\r\nKoniec\r\n\r\nWklej kolejny adres strony");
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

        private static List<ImageInfo> GetImagesInfo(List<HtmlNode> htmlNodes, List<string> fileExtensions, List<string> attributesToSearch, string destinationPath)
        {
            List<ImageInfo> images = new List<ImageInfo>();
            foreach(HtmlNode pictureNode in htmlNodes.SelectMany(p => p.SelectNodes(".//picture") ?? new HtmlNodeCollection(null)))
            { 
                HtmlNode imgNode = pictureNode.ChildNodes.Where(x => x.Name.Equals("img")).FirstOrDefault();
                HtmlNode sourceNode = pictureNode.ChildNodes.Where(x => x.Name.Equals("source")).FirstOrDefault();
                foreach (string attribute in attributesToSearch)
                {
                    if (imgNode != null)
                    {
                        string imgUrl = imgNode.GetAttributeValue(attribute, "");
                        if (!String.IsNullOrWhiteSpace(imgUrl))
                        {
                            foreach (string extension in fileExtensions)
                            {
                                if (!imgUrl.Contains(extension))
                                    continue;
                                imgUrl = imgUrl.Substring(0, imgUrl.LastIndexOf(extension) + extension.Length);
                                Uri uri;
                                if (!(Uri.TryCreate(imgUrl, UriKind.Absolute, out uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)))
                                {
                                    if (imgUrl.StartsWith("//"))
                                        imgUrl = "https:" + imgUrl;
                                    else if (imgUrl.StartsWith("/"))
                                        imgUrl = "https:/" + imgUrl;
                                    else
                                        imgUrl = "https://" + imgUrl;
                                    try
                                    {
                                        uri = new Uri(imgUrl);
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                        continue;
                                    }
                                }
                                images.Add(new ImageInfo(destinationPath + "/" + uri.ToString().Substring(uri.ToString().LastIndexOf("/") + 1), uri, attribute, imgNode.GetAttributeValue("alt", sourceNode.GetAttributeValue("alt", "")), pictureNode, "img"));
                            }
                        }
                    }
                    if (sourceNode != null)
                    {
                        string sourceUrl = sourceNode.GetAttributeValue(attribute, "");
                        if (!String.IsNullOrWhiteSpace(sourceUrl))
                        {
                            foreach (string extension in fileExtensions)
                            {
                                if (!sourceUrl.Contains(extension))
                                    continue;
                                sourceUrl = sourceUrl.Substring(0, sourceUrl.LastIndexOf(extension) + extension.Length);
                                Uri uri;
                                if (!(Uri.TryCreate(sourceUrl, UriKind.Absolute, out uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)))
                                {
                                    if (sourceUrl.StartsWith("//"))
                                        sourceUrl = "https:" + sourceUrl;
                                    else if (sourceUrl.StartsWith("/"))
                                        sourceUrl = "https:/" + sourceUrl;
                                    else
                                        sourceUrl = "https://" + sourceUrl;
                                    try
                                    {
                                        uri = new Uri(sourceUrl);
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                        continue;
                                    }
                                }
                                images.Add(new ImageInfo(destinationPath + "/" + uri.ToString().Substring(uri.ToString().LastIndexOf("/") + 1), uri, attribute, sourceNode.GetAttributeValue("alt", imgNode.GetAttributeValue("alt", "")), pictureNode, "source"));
                            }
                        }
                    }
                }
            }

            foreach (HtmlNode imageNode in htmlNodes.SelectMany(p => p.SelectNodes(".//img") ?? new HtmlNodeCollection(null)).Where(x=>!x.ParentNode.Name.Equals("noscript") && !x.ParentNode.Name.Equals("picture")))
            {
                foreach (string attribute in attributesToSearch)
                {
                    if (imageNode != null)
                    {
                        string imgUrl = imageNode.GetAttributeValue(attribute, "");
                        if (String.IsNullOrWhiteSpace(imgUrl))
                            continue;
                        foreach (string extension in fileExtensions)
                        {
                            if (!imgUrl.Contains(extension))
                                continue;
                            imgUrl = imgUrl.Substring(0, imgUrl.LastIndexOf(extension) + extension.Length);
                            Uri uri;
                            if (!(Uri.TryCreate(imgUrl, UriKind.Absolute, out uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)))
                            {
                                if (imgUrl.StartsWith("//"))
                                    imgUrl = "https:" + imgUrl;
                                else if (imgUrl.StartsWith("/"))
                                    imgUrl = "https:/" + imgUrl;
                                else
                                    imgUrl = "https://" + imgUrl;
                                try
                                {
                                    uri = new Uri(imgUrl);
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    continue;
                                }
                            }
                            images.Add(new ImageInfo(destinationPath + "/" + uri.ToString().Substring(uri.ToString().LastIndexOf("/") + 1), uri, attribute, imageNode.GetAttributeValue("alt",  ""), imageNode.ParentNode, "img"));
                        }
                    }
                }
            }
            return images;
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

    public struct ImageInfo
    {
        private FileInfo fileInfo;
        private Uri uri;
        private string attribute;
        private string alt;
        private HtmlNode parent;
        private string name;

        public ImageInfo(FileInfo fileInfo, Uri uri, string attribute, string alt, HtmlNode parent, string name)
        {
            this.parent = parent;
            this.attribute = attribute;
            this.alt = alt;
            this.fileInfo = fileInfo;
            this.uri = uri;
            this.name = name;
        }

        public ImageInfo(string fileName, Uri uri, string attribute, string alt, HtmlNode parent, string name) : this(new FileInfo(fileName), uri, attribute, alt, parent, name) { }

        public FileInfo FileInfo { get => fileInfo; }
        public Uri Uri { get => uri; }
        public string Attribute { get => attribute; }
        public string Alt { get => alt; set => alt = value; }
        public string Name { get => name;  }
        public HtmlNode Parent { get => parent; }
    }
}

