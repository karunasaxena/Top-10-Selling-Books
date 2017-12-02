using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.util;
using System.Xml;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Pdf.Tables;
using System.Drawing;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web;
using System.Diagnostics;

namespace BookStore
{
    class Program
    {
        //server=localhost;database=test;Persist Security Info=True;
        static string MyConnString = "server=localhost;port=3307;database=mysql;uid=root;pwd=admin";
        static MySqlConnection con = new MySqlConnection(MyConnString);

        static void addPdfContent(Document document, string imageURL, string filename, string para, bool last)
        {
            try
            {
                if(document == null)
                {
                    Console.WriteLine("Error in PDF file creation");
                    return;
                }


                //Process Image
                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);

                //Resize image depend upon your need

                jpg.ScaleToFit(140f, 120f);

                //Give space before image

                jpg.SpacingBefore = 10f;

                //Give some space after the image

                jpg.SpacingAfter = 1f;

                jpg.Alignment = Element.ALIGN_LEFT;

                document.Add(jpg);
                Paragraph paragraph = new Paragraph(para);
                document.Add(paragraph);

                if(last)
                  document.Close();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                Console.WriteLine(filename + "PDF created.");
            }
        }
        static Document generate_pdf(string filename, string para, string date )
        {
            // step 1: creation of a document-object
            Document document = new Document(PageSize.A4, 10f, 10f, 10f, 0f);

            try
            {

                // step 2:
                // we create a writer that listens to the document
                // and directs a XML-stream to a file
				DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                string path = @"" + di + "\\BookStore";
                bool exists = System.IO.Directory.Exists(path);

                if (!exists)
                  System.IO.Directory.CreateDirectory((path));

                path += "\\" + filename;
                exists = System.IO.Directory.Exists(path);
                if (!exists)
                  System.IO.Directory.CreateDirectory((path));
                string file = path + "\\" + date + ".pdf";
                exists = System.IO.File.Exists(file);
                if (exists)
                    File.Delete(file);
                PdfWriter.GetInstance(document, new FileStream(file, FileMode.Append));
                document.Open();
                iTextSharp.text.Font header = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 25f, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.UNDERLINE, BaseColor.BLACK);
                
                document.Add(new Paragraph("Top 10 Best Selling Books", header));
                             
                
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.StackTrace);
                Console.Error.WriteLine(e.Message);

                if (e.InnerException != e)
                {
                    Console.Error.WriteLine(e.InnerException.Message);
                }
            }
            return document;
        }
        
        static void createTable()
        { 
            try
            {
                MySqlCommand cmd = con.CreateCommand();
                con.Open();
                cmd.CommandText = "SELECT * FROM information_schema.tables where table_name = 'bookstore'";
                cmd.CommandType = System.Data.CommandType.Text;

                MySqlDataReader reader = cmd.ExecuteReader();

                if (!reader.Read())
                {
                    reader.Close();
                    //now(), '" + bookName + "', '" + authName + "', '" + price + "', '" + reviews + "', '" + productDetail + "', '" + bookLink + "', '" + image + "             
                    cmd.CommandText = "CREATE TABLE bookstore (timestamp datetime, bookName text, author text, price decimal(20,3), price_type varchar(15), reviews text, productDetail text,  booklink text, image text)";
                    cmd.ExecuteNonQuery();
                }

                reader.Close();

            }
            catch(Exception e)
            {
                Console.WriteLine(e.StackTrace.ToString());
               
            }
            finally
            {
                if (con.State == System.Data.ConnectionState.Open)
                {
                    con.Close();
                }
            }

        }
        static private void dump_in_db(string bookName, string authName, float price, char pricetype, string reviews, string productDetail, string bookLink, string image)
        {
         
          try
          {
                con.Open();
                Trace.Listeners.Clear();
                Console.WriteLine(Path.GetTempPath() + " " + AppDomain.CurrentDomain.FriendlyName);
                TextWriterTraceListener twtl = new TextWriterTraceListener(Path.Combine(Path.GetTempPath(), AppDomain.CurrentDomain.FriendlyName));
                twtl.Name = "TextLogger";
                twtl.TraceOutputOptions = TraceOptions.ThreadId | TraceOptions.DateTime;

                ConsoleTraceListener ctl = new ConsoleTraceListener(false);
                ctl.TraceOutputOptions = TraceOptions.DateTime;

                Trace.Listeners.Add(twtl);
                Trace.Listeners.Add(ctl);
                Trace.AutoFlush = true;

               
            
                String query = "INSERT INTO bookstore values ( now(), '" + bookName + "', '" + authName + "', " + price + ", '" + pricetype + "','" + reviews.Trim() + "', '" + productDetail.Trim().Replace("\n"," ").Replace("\t"," ") + "', '" + bookLink + "', '" + image + "')";
                Trace.WriteLine(query);

               MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.ExecuteNonQuery();

          }
          catch (Exception e)
          {
              Console.WriteLine(e.ToString());
          }
            finally
            {
                if(con.State == System.Data.ConnectionState.Open)
                {
                    con.Close();
                }
            }
        }

        static private void createNode(XmlTextWriter writer, string[] fields)
        {
            try
            {
                writer.WriteStartElement("Row");

                for (int i = 0; i < fields.Length; i++)
                {
                    writer.WriteStartElement("Cell");
                    //writer.WriteAttributeString("ss:StyleID", "s53");
                    writer.WriteStartElement("Data");
                    //writer.WriteAttributeString("ss:Type", "String");
                    writer.WriteString(fields[i].Trim());

                    writer.WriteEndElement();

                    writer.WriteEndElement();

                }

                writer.WriteEndElement();
            }
            catch (Exception e) { Console.WriteLine(e.StackTrace); }

        }

        static private XmlTextWriter createWriter(string date, string dir)
        {
            DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            string path = @"" + di + "\\BookStore";
            bool exists = System.IO.Directory.Exists(path);

            if (!exists)
                System.IO.Directory.CreateDirectory((path));

            path += "\\" + dir;
            exists = System.IO.Directory.Exists(path);
            if (!exists)
                System.IO.Directory.CreateDirectory((path));
            string file = path + "\\" + date + ".xml";

            XmlTextWriter writer = new XmlTextWriter(file, System.Text.Encoding.UTF8);
            //.WriteStartDocument(true);
            writer.Formatting = Formatting.Indented;
            writer.Indentation = 2;
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><?mso-application progid = \"Excel.Sheet\"?>");
            createNode(writer, new string[] { "Book Name", "Author Name", "Price", "Book Link ", "Book Image Link", "Customer Reviews", "Product Detail" });
            return writer;
        }

        static void Main(string[] args)
        {
            System.Net.ServicePointManager.Expect100Continue = true;

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            HtmlWeb web = new HtmlWeb();

            //Create table in database.
            createTable();

            Console.WriteLine("****************************************************");
            Console.WriteLine("*****************TOP 10 BOOK DETAILS****************");
            Console.WriteLine("****************************************************");

            Console.WriteLine("\n\n\n\n\n\n");
            Console.WriteLine("---- Please Enter File Format(0 for XML/1 for PDF/2 for DB) -----");
            string fileFormat = Console.ReadLine();


            /**************** BLACK WELL *******************************************************************/
            HtmlDocument doc = web.Load("http://bookshop.blackwell.co.uk/bookshop/bestsellers");
            int i = 0;
            //IList<HtmlNode> bookList = doc.QuerySelectorAll(".s-access-detail-page");
            //IList<HtmlNode> bookList = doc.QuerySelectorAll("a");
            

            IList<HtmlNode> bookList = doc.QuerySelectorAll(".product-name");
            Console.WriteLine("booklist = " + bookList.Count);
            string bookLink = "";
            string productDetail = "";
            string bookName = "";
            string customerReviews = "";
            string authorName = "";
            string price = "";
            string bookImage = "";
            String[,] dataForPdf = new String[10,8];
            string para = "";
            bool isLast = false;
            Document document = null;

            //DateTime date = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, r, DateTime.Now.AddMinutes(5).Minute, 00); ;
            string date = DateTime.Now.Date.ToString("dd-MM-yyyy");
            XmlTextWriter writer = null;

            for (i = 0; i < 20; i++)
            {
                if (bookList[i] == null)
                {
                    i++;
                    continue;
                }
                bookLink = "http://bookshop.blackwell.co.uk" + bookList[i].GetAttributeValue("href", "");
                doc = web.Load(bookLink);
                bookName = doc.QuerySelector(".product__name").InnerText.Trim().Replace("\n","");

                
                authorName = doc.QuerySelector(".product__author").InnerText.Trim().Split(',')[0].Replace("(author)", "");
                HtmlNode review = doc.QuerySelector("#acrCustomerReviewText");
                price = HttpUtility.HtmlDecode(doc.QuerySelector(".product-price--current").InnerText);
                bookImage = "http://bookshop.blackwell.co.uk" + doc.QuerySelector(".picture-wrapper img").GetAttributeValue("src", "");
                if (review != null)
                  customerReviews = HttpUtility.HtmlDecode(review.InnerText);
                IList<HtmlNode> productDetailArr = doc.QuerySelectorAll(".u-separator-right td");
                productDetail = "";
                for (int j = 0; j < productDetailArr.Count; j++)
                {
                    productDetail += HttpUtility.HtmlDecode(productDetailArr[j].InnerText.Replace("\n",""));
                    if(j%2 != 0)
                      productDetail += "\n";
                }
                if (i == 0)
                {
                    /*if (writer != null)
                    {
                        writer.WriteEndElement();
                        writer.Close();

                    }
                    else*/
                        writer = createWriter(date, "BlackWell");
                }

                int count = 0;
                string[] bdata = new string[] {  bookName, authorName, price, bookLink, bookImage, customerReviews,productDetail};

                //create node in xml
                if(fileFormat == "0")
                  createNode(writer, bdata);

                para += "\n Book Name\t : " + bookName + "\n Author\t :" + authorName + "\n Price\t : " + price + "\n Link\t : " + bookLink + "\n Reviews\t : " + customerReviews + "\n Product Detail\t :" + productDetail + "\n\n";

                if (fileFormat == "1")
                {
                    if(i == 0)
                        document = generate_pdf("BlackWell", para, date);

                    if (i == 18)
                    {
                        isLast = true;
                    }
                                      
                    addPdfContent(document, bookImage, "BlackWell", para, isLast);
                    
                    
                }
                count++; //for blackwell data array
                i++; //To take the alternate eg 0, 2, 4
            }
            /******** BlackWell End ****************/


            /********* WaterStone *******************/
            para = "";
            document = null;
            isLast = false;
            doc = web.Load("https://www.waterstones.com/books/bestsellers");
            bookList = doc.QuerySelectorAll(".title-wrap a");
            for (i = 0; i < 10; i++)
            {
                if (bookList[i] == null)
                {
                    continue;
                }

                bookLink = "https://www.waterstones.com" + bookList[i].GetAttributeValue("href", "");
                doc = web.Load(bookLink);
                bookName = doc.QuerySelector(".book-title").InnerText;

                price = doc.QuerySelector(".price-rrp").InnerText;
                authorName = doc.QuerySelector(".text-gold span").InnerText.Trim();
                IList<HtmlNode> customerReviewsArr = doc.QuerySelectorAll(".p-medium");
                bookImage =  doc.QuerySelector("#scope_book_image").GetAttributeValue("src","");
                customerReviews = "";
                for (int j = 0; j < customerReviewsArr.Count; j++)
                {
                    customerReviews += doc.QuerySelectorAll(".intro")[j].InnerText + customerReviewsArr[j].InnerText;
                }
                customerReviews = HttpUtility.HtmlDecode(customerReviews);
                
                if(doc.QuerySelectorAll(".spec").Count > 1)
                  productDetail = HttpUtility.HtmlDecode(doc.QuerySelectorAll(".spec")[1].InnerText.Trim());

                if (i == 0)
                {
                    if (writer != null)
                    {
                        //writer.WriteEndElement();
                        writer.Close();

                    }
                    writer = createWriter(date, "WaterStones");
                }
                Console.WriteLine("Waterstone started " + (i+1));
                string[] bdata = new string[] { bookName, authorName, price, bookLink, bookImage, customerReviews, productDetail };

                //create node in xml
                if (fileFormat == "0")
                    createNode(writer, bdata);

                //create paragraph for pdf
                para += "\n " + bookImage + "\n Book Name\t : " + bookName + "\n Author\t :" + authorName + "\n Price\t : " + price + "\n Link\t : " + bookLink + "\n Reviews\t : " + customerReviews + "\n Product Detail\t :" + productDetail + "\n\n";

                if (fileFormat == "1")
                {
                    if (i == 0)
                       document = generate_pdf("WaterStones", para, date);

                    if (i == 9)
                    {
                        isLast = true;
                    }

                    addPdfContent(document, bookImage, "WaterStones", para, isLast);
                   
                }

                if(fileFormat == "2")
                {
                   dump_in_db(bookName, authorName,  float.Parse(price.Substring(1)), price[0], customerReviews, productDetail, bookLink, bookImage);
                }
            }
            /********** Water Stone Ends Here *************/

            /********** Booktopia Begins ******************/
            para = "";
            isLast = false;
            document = null;
            doc = web.Load(" https://www.booktopia.com.au/bestsellers/promo294.html");
            bookList = doc.QuerySelectorAll("#top-10 a");
            for (i = 5; i < 15; i++)
            {
                if (bookList[i] == null)
                {
                    continue;
                }

                bookLink = "https://www.booktopia.com.au/" + bookList[i].GetAttributeValue("href", "");
                doc = web.Load(bookLink);
                bookName = doc.QuerySelector("#product-title h1").InnerText;

                price = doc.QuerySelector(".sale-price").InnerText;
               authorName = doc.QuerySelector("#contributors").InnerText.Trim();
                IList<HtmlNode> customerReviewsArr = doc.QuerySelectorAll(".p-medium");
                bookImage = doc.QuerySelector("#product #image img").GetAttributeValue("src", "");
                customerReviews = "";
                for (int j = 0; j < customerReviewsArr.Count; j++)
                {
                    customerReviews += doc.QuerySelectorAll(".intro")[j].InnerText + customerReviewsArr[j].InnerText;
                }
                customerReviews = HttpUtility.HtmlDecode(customerReviews);
                productDetail = HttpUtility.HtmlDecode(doc.QuerySelector("#details p").InnerText).Trim();
                string pattern = "\\s+";
                string replacement = " ";
                Regex rgx = new Regex(pattern);
                string result = rgx.Replace(productDetail, replacement);

                if (i == 5)
                {
                    if (writer != null)
                    {
                        //writer.WriteEndElement();
                        writer.Close();

                    }
                    writer = createWriter(date, "BookTopia");
                }
                Console.WriteLine("Booktopia Row ");
                string[] bdata = new string[] { bookName, authorName, price, bookLink, bookImage, customerReviews, productDetail };

                //create node in xml
                if (fileFormat == "0")
                    createNode(writer, bdata);

                para += "\n " + bookImage + "\n Book Name\t : " + bookName + "\n Author\t :" + authorName + "\n Price\t : " + price + "\n Link\t : " + bookLink + "\n Reviews\t : " + customerReviews + "\n Product Detail\t :" + productDetail + "\n\n";

                if (fileFormat == "1")
                {
                    if (i == 5)
                        document = generate_pdf("BookTopia", para, date);

                    if (i == 14)
                    {
                        isLast = true;
                    }

                    addPdfContent(document, bookImage, "BookTopia", para, isLast);
                }
                if (fileFormat == "2")
                {
                    dump_in_db(bookName, authorName, float.Parse(price.Substring(1)), price[0], customerReviews, productDetail, bookLink, bookImage);
                }
            }
            //writer.WriteEndElement();
            writer.Close();
        }
    }
}
