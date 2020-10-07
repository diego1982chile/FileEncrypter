using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Reg17Generator
{
    class WebSiteReader
    {

        System.Net.WebClient wc = new System.Net.WebClient();
        string extract;
        bool translate = true;

        public string extractSummary(string url)
        {            
            byte[] raw = wc.DownloadData(url);

            string webData = System.Text.Encoding.UTF8.GetString(raw);
            var doc = new HtmlDocument();
            doc.LoadHtml(webData);

            try
            {
                SearchSummary(doc.DocumentNode);
                if(this.extract == null)
                {
                    System.Console.WriteLine("Es null!! - url: " + url);
                }

                if (translate)
                {
                    this.extract = TranslateText(this.extract).Replace("\\ n", Environment.NewLine).Replace("\\ t","• ");                    
                }

                return this.extract;
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);

                if (translate)
                {
                    this.extract = TranslateText(this.extract).Replace("\\ n", Environment.NewLine).Replace("\\ t", "• ");
                }

                return this.extract;
            }
            
        }

        public string extractTitle(string url)
        {
            byte[] raw = wc.DownloadData(url);

            string webData = System.Text.Encoding.UTF8.GetString(raw);
            var doc = new HtmlDocument();
            doc.LoadHtml(webData);

            try
            {
                SearchTitle(doc.DocumentNode);
                if(translate)
                {
                    this.extract = TranslateText(this.extract);
                }                
                return this.extract;
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);

                if (translate)
                {
                    this.extract = TranslateText(this.extract);
                }

                return this.extract;
            }            
        }

        Boolean SearchSummary(HtmlNode node)
        {
            foreach (var child in node.ChildNodes)
            {
                SearchSummary(child);
            }

            if (node.InnerHtml.Contains("var microsoft = microsoft || {};\r\nmicrosoft.support = microsoft.support") ||
                node.InnerHtml.Contains("Latest security intelligence updates for Microsoft Defender Antivirus and other Microsoft antimalware - Microsoft Security Intelligence"))
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(node.InnerHtml);                
                bool flag = false;

                foreach (var line in node.InnerHtml.Split(new char[] { '\r', '\n' }))
                {
                    if (line.Contains("Summary") || line.Contains("Highlights"))
                    {
                        flag = true;
                        translate = true;
                    }

                    if (line.Contains("Mejoras y correcciones"))
                    {
                        flag = true;
                        translate = false;
                    }

                    if (line.Contains("Improvements and fixes"))
                    {
                        flag = true;
                        translate = true;
                    }

                    if (line.Contains("Security intelligence updates for Microsoft Defender Antivirus and other Microsoft antimalware"))
                    {
                        flag = true;
                        translate = true;
                    }                    

                    if ( (line.Contains("<p>") || line.Contains("Microsoft continually updates")) && flag)
                    {
                        this.extract = line.Replace("<p>", "").Replace("</p>", "\n");
                        this.extract = Regex.Replace(line, @"<[^>]*>", "").Replace("&nbsp;", " ").Replace("\\ N", "").Replace("\"", "").Replace("\\ n", Environment.NewLine);
                        throw new Exception("Summary Extracted!");
                    }
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        Boolean SearchTitle(HtmlNode node)
        {
            foreach (var child in node.ChildNodes)
            {
                SearchTitle(child);
            }

            if (node.InnerHtml.Contains("var microsoft = microsoft || {};\r\nmicrosoft.support = microsoft.support") ||
                node.InnerHtml.Contains("Latest security intelligence updates for Microsoft Defender Antivirus and other Microsoft antimalware - Microsoft Security Intelligence"))
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(node.InnerHtml);                                

                foreach (var line in node.InnerHtml.Split(new char[] { '\r', '\n' }))
                {
                    if (line.Contains("heading"))
                    {
                        this.extract = line.Split(':')[1].Replace("\"", "").Replace("\",", "");
                        throw new Exception("Title Extracted!");
                    }

                    if (line.Contains("Latest security intelligence updates for Microsoft Defender Antivirus and other Microsoft antimalware - Microsoft Security Intelligence"))
                    {
                        this.extract = line;
                        throw new Exception("Title Extracted!");
                    }
                }

                return true;
            }
            else
            {
                return false;
            }
        }
        public static string TranslateText(string input)
        {
            try
            {
                // Set the language from/to in the url (or pass it into this function)
                string url = String.Format("https://translate.googleapis.com/translate_a/single?client=gtx&sl={0}&tl={1}&dt=t&q={2}", "en", "es", Uri.EscapeUriString(input));
                HttpClient httpClient = new HttpClient();
                string result = httpClient.GetStringAsync(url).Result;

                // Get all json data
                var jsonData = new JavaScriptSerializer().Deserialize<List<dynamic>>(result);

                // Extract just the first array element (This is the only data we are interested in)
                var translationItems = jsonData[0];

                // Translation Data
                string translation = "";

                // Loop through the collection extracting the translated objects
                foreach (object item in translationItems)
                {
                    // Convert the item array to IEnumerable
                    IEnumerable translationLineObject = item as IEnumerable;

                    // Convert the IEnumerable translationLineObject to a IEnumerator
                    IEnumerator translationLineString = translationLineObject.GetEnumerator();

                    // Get first object in IEnumerator
                    translationLineString.MoveNext();

                    // Save its value (translated text)
                    translation += string.Format(" {0}", Convert.ToString(translationLineString.Current));
                }

                // Remove first blank character
                if (translation.Length > 1) { translation = translation.Substring(1); };

                // Return translation
                return translation;
            }
            catch(Exception e)
            {
                System.Console.WriteLine(e.InnerException.Message);
            }

            return "";
            
        }


    }
}
