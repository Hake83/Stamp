// pdfc program used as stamp or whatever user wants to call it
// B&B Engineering, written in visual studio 2017 community
// Ryan Hake, latest revision 12/12/18

using System;
using System.Linq;
using iTextSharp.text.pdf;
using System.IO;
using System.Xml;
using iTextSharp.text;




namespace pdfc
{
    class Program
    {
        static void Main()
        {
            // directory object for searching through files in the current directory
            string dir = Directory.GetCurrentDirectory();
            
            // gets array listing files starting with rfq and ending in .pdf
            string[] files = Directory.GetFiles(dir, "rfq*.pdf", SearchOption.TopDirectoryOnly);
            

            for (int n = 0; n < files.Count(); n++)
            {
                if (files[n] != null && files[n].Length >= 0)
                {
                    if (File.Exists(files[n]))
                    {
                        using (FileStream t = new FileStream(@"C:\User Programs\Pdf Template\rfq_template.xml", FileMode.Open, FileAccess.ReadWrite))

                        {
                            PdfReader reader = new PdfReader(files[n]);
                            XfaForm xfa = new XfaForm(reader);
                            XmlDocument doc = xfa.DomDocument;
                            // template is the xml document from Pdf Template folder doc is the xml document pulled from the pdf being altered
                            XmlDocument template = new XmlDocument();
                            template.Load(t);
                            
                            XmlNodeList partNum = doc.GetElementsByTagName("PartNumDrawNum");
                            XmlNodeList partNumTemplate = template.GetElementsByTagName("PartNumDrawNum");
                            //Exit the program if there are too many part number lines.
                            if (partNum.Count > 2)
                            {
                                for (int i = 0; i < partNum.Count-3; i++)
                                {
                                    // Create new body row section to expand part number table
                                    XmlNodeList sfDynTable = template.GetElementsByTagName("sfDynTable");
                                    XmlNode tbl = sfDynTable[0].FirstChild;
                                    XmlElement bodyRow = (XmlElement)template.CreateElement("BodyRow");
                                    bodyRow.AppendChild(template.CreateElement("PartNumDrawNum"));
                                    bodyRow.AppendChild(template.CreateElement("Qty"));
                                    bodyRow.AppendChild(template.CreateElement("UnitPrice"));
                                    bodyRow.AppendChild(template.CreateElement("DelDatetoBuyDock"));
                                    tbl.AppendChild(bodyRow);
                                    
                                }

                            }


                            for (int i = 0; i < partNum.Count-1; i++)
                            {
                                partNumTemplate[i].InnerText = partNum[i].InnerText;
                            }

                            Console.WriteLine("Enter delivery lead time:");
                            string userLeadTime = Console.ReadLine();

                            XmlNodeList leadTime = doc.GetElementsByTagName("DelLeadTime");
                            XmlNodeList leadTimeTemplate = template.GetElementsByTagName("DelLeadTime");
                            for (int i = 0; i < leadTime.Count; i++)
                            {
                                if (leadTime[i].InnerText == ("")) { }
                                else
                                {
                                    leadTimeTemplate[0].InnerText = userLeadTime+" WEEKS ARO";
                                }
                            }

                            XmlNodeList delDate = doc.GetElementsByTagName("DelDatetoBuyDock");
                            XmlNodeList delDateTemplate = template.GetElementsByTagName("DelDatetoBuyDock");

                            // Commented out for Andrew, activate this and remove next block for normal
                            /*for (int i = 0; i < partNum.Count; i++)
                            {
                                if (partNum[i].InnerText == ("")) { }
                                else
                                {
                                    delDateTemplate[i].InnerText = userLeadTime + " WEEKS ARO";
                                }
                            }*/

                            // Changed this to userLeadTime from delDate
                            for (int i = 0; i < delDate.Count; i++)
                            {
                                if (partNum[i].InnerText == ("")) { }
                                else
                                {
                                    delDateTemplate[0].InnerText = userLeadTime + " WEEKS ARO";
                                }
                            }


                            //Return Request for Quote No Later Than:
                            XmlNodeList returnDate = doc.GetElementsByTagName("returnDate");
                            XmlNodeList returnDateTemplate = template.GetElementsByTagName("returnDate");
                            for (int i = 0; i < returnDate.Count; i++)
                            {
                                if (returnDate[i].InnerText == ("")) { }
                                else
                                {
                                    returnDateTemplate[0].InnerText = returnDate[i].InnerText;
                                }
                            }
                            
                          
                            XmlNodeList qtyList = doc.GetElementsByTagName("Qty");
                            XmlNodeList qtyListTemplate = template.GetElementsByTagName("Qty");
                            
                            for (int i = 0; i < qtyList.Count-1; i++)
                             {
                                if (qtyList[i].InnerText == "") { }
                                else
                                {
                                    qtyListTemplate[i].InnerText = qtyList[i].InnerText;
                                }
                             }

                            XmlNodeList date = doc.GetElementsByTagName("date");
                            XmlNodeList dateTemplate = template.GetElementsByTagName("date");

                            for (int i = 0; i < date.Count; i++)
                            {
                                if (date[i].InnerText == ("")) { }
                                else
                                {
                                    dateTemplate[0].InnerText = date[i].InnerText;
                                }
                            }

                            XmlNodeList email = doc.GetElementsByTagName("email");
                            XmlNodeList emailTemplate = template.GetElementsByTagName("email");
                            for (int i = 0; i < email.Count; i++)
                            {
                                if (email[i].InnerText == ("")) { }
                                else
                                {
                                    emailTemplate[0].InnerText = email[i].InnerText;
                                }
                            }

                            XmlNodeList BuyerFillin = doc.GetElementsByTagName("Buyer_Fillin");
                            XmlNodeList BuyerFillinTemplate = template.GetElementsByTagName("Buyer_Fillin");
                            for (int i = 0; i < BuyerFillin.Count; i++)
                            {
                                if (BuyerFillin[i].InnerText == ("")) { }
                                else
                                {
                                    BuyerFillinTemplate[0].InnerText = BuyerFillin[i].InnerText;
                                }
                            }

                            XmlNodeList mat = doc.GetElementsByTagName("matSizeAndLength");
                            XmlNodeList matTemplate = template.GetElementsByTagName("matSizeAndLength");
                            for (int i = 0; i < mat.Count; i++)
                            {
                                if (mat[i].InnerText == ("")) { }
                                else
                                {
                                    matTemplate[0].InnerText = mat[i].InnerText;
                                }
                            }

                            XmlNodeList sincerely = doc.GetElementsByTagName("sincerely");
                            XmlNodeList sincerelyTemplate = template.GetElementsByTagName("sincerely");
                            for (int i = 0; i < sincerely.Count; i++)
                            {
                                if (sincerely[i].InnerText == ("")) { }
                                else
                                {
                                    sincerelyTemplate[0].InnerText = sincerely[i].InnerText;
                                }
                            }

                            XmlNodeList title = doc.GetElementsByTagName("title");
                            XmlNodeList titleTemplate = template.GetElementsByTagName("title");
                            for (int i = 0; i < title.Count; i++)
                            {
                                if (title[i].InnerText == ("")) { }
                                else
                                {
                                    titleTemplate[0].InnerText = title[i].InnerText;
                                }

                            }


                            XmlNodeList org = doc.GetElementsByTagName("org");
                            XmlNodeList orgTemplate = template.GetElementsByTagName("org");
                            for (int i = 0; i < org.Count; i++)
                            {
                                if (org[i].InnerText == ("")) { }
                                else
                                {
                                    orgTemplate[0].InnerText = org[i].InnerText;
                                }
                            }

                            XmlNodeList phone = doc.GetElementsByTagName("phone");
                            XmlNodeList phoneTemplate = template.GetElementsByTagName("phone");
                            for (int i = 0; i < phone.Count; i++)
                            {
                                if (phone[i].InnerText == ("")) { }

                                else
                                {
                                    phoneTemplate[i].InnerXml = phone[i].InnerXml;
                                }
                            }


                            /*Block for the make stamp
                            XmlNodeList comments = doc.GetElementsByTagName("addComments");
                            XmlNodeList commentsTemplate = template.GetElementsByTagName("addComments");
                            for (int i = 0; i < comments.Count; i++)
                            {
                                if (i==0 && comments[i].InnerText ==(""))
                                {
                                    commentsTemplate[i].InnerText = "Best manufacturing practices.";
                                }
                                else if (comments[i].InnerText == ("")) { } 
                                else if (comments[i].InnerText.Contains("Best manufacturing practices."))
                                {
                                    commentsTemplate[i].InnerText = comments[i].InnerText;
                                }
                                else
                                {
                                    commentsTemplate[i].InnerText = comments[i].InnerText + " Best manufacturing practices.";
                                }
                            }*/

                            // Block for design
                            XmlNodeList comments = doc.GetElementsByTagName("addComments");
                            XmlNodeList commentsTemplate = template.GetElementsByTagName("addComments");
                            for (int i = 0; i < comments.Count; i++)
                            {
                                //Console.WriteLine("number of comment lines = " + i);
                                if (comments[i].InnerText == ("")) { }
                                else
                                {
                                    commentsTemplate[i].InnerText = comments[i].InnerText;
                                }
                            }

                            XmlNodeList text1 = doc.GetElementsByTagName("TextField1");
                            XmlNodeList text1Template = template.GetElementsByTagName("TextField1");
                            for (int i = 0; i < text1.Count; i++)
                            {
                                if (text1[i].InnerText == ("")) { }
                                else
                                {
                                    text1Template[i].InnerText = text1[i].InnerText;
                                }
                            }

                            XmlNodeList text2 = doc.GetElementsByTagName("TextField2");
                            XmlNodeList text2Template = template.GetElementsByTagName("TextField2");
                            for (int i = 0; i < text2.Count; i++)
                            {
                                if (text2[i].InnerText == ("")) { }
                                else
                                {
                                    text2Template[i].InnerText = text2[i].InnerText;
                                }
                            }

                            XmlNodeList phoneNum = doc.GetElementsByTagName("PhoneNum");
                            XmlNodeList phoneNumTemplate = template.GetElementsByTagName("PhoneNum");
                            for (int i = 0; i < phoneNum.Count; i++)
                            {
                                if (phoneNum[i].InnerText == ("")) { }
                                else
                                {
                                    phoneNumTemplate[i].InnerXml = phoneNum[i].InnerXml;
                                }
                            }

                            
                            


                            XmlNodeList unitPrice = doc.GetElementsByTagName("UnitPrice");
                            XmlNodeList unitPriceTemplate = template.GetElementsByTagName("UnitPrice");
                            for (int i = 0; i < unitPrice.Count; i++)
                            {
                                if (unitPrice[i].InnerText == ("")) { }
                                else
                                {
                                    unitPriceTemplate[i].InnerText = unitPrice[i].InnerText;
                                }
                            }
                            Console.WriteLine("Enter Unit Price:");
                            string userPrice = Console.ReadLine();
                            unitPriceTemplate[0].InnerText = userPrice;
                            


                            XmlNodeList sellDateTemplate = template.GetElementsByTagName("selldate");
                            sellDateTemplate[0].InnerText = (DateTime.Today.ToString("M/d/yyyy"));

                            template.Save(dir + "/test.xml");
                            reader.Close();




                            using (FileStream existingPdf = new FileStream(files[n], FileMode.Open, FileAccess.ReadWrite))
                            using (FileStream xml = new FileStream(dir + "/test.xml", FileMode.Open, FileAccess.ReadWrite))
                            using (FileStream filledPdf = new FileStream(dir + "/newPdf.pdf", FileMode.Create, FileAccess.ReadWrite))
                            {
                                PdfReader.unethicalreading = true;
                                PdfReader reader1 = new PdfReader(existingPdf);
                                PdfStamper stamper = new PdfStamper(reader1, filledPdf, '\0', true);

                                stamper.AcroFields.Xfa.FillXfaForm(xml);
                                stamper.Close();
                                reader1.Close();
                            }

                            //Lines added to create small business classification paperwork
                            /*FileInfo check = new FileInfo(dir + "/J-3.1 NAICS 333514.pdf");
                            if (check.Exists)
                            { }
                            else
                            {
                                using (FileStream src = new FileStream(@"C:\User Programs\Pdf Template\small_business_template.pdf", FileMode.Open, FileAccess.Read))
                                using (FileStream newFile = new FileStream(dir + "/J-3.1 NAICS 333514.pdf", FileMode.Create, FileAccess.ReadWrite))
                                {
                                    PdfReader reader2 = new PdfReader(src);
                                    PdfStamper stamper2 = new PdfStamper(reader2, newFile, '\0', true);

                                    stamper2.RotateContents.Equals(false);
                                    PdfContentByte canvas = stamper2.GetOverContent(1);

                                    // Write part number & date
                                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(partNumTemplate[0].InnerText), 240, 685, 0);
                                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(DateTime.Today.ToString("M/dd/yyyy")), 180, 88, 0);

                                    stamper2.Close();
                                    reader2.Close();
                                }
                            }*/
                        }   

                        string fileName = (files[n]);
                        File.Delete(files[n]);
                        File.Move(dir + "/newPdf.pdf",fileName);
                        File.Delete(dir + "/test.xml");

                    }
                }
                else
                {
                    Console.WriteLine("Can't find a file to stamp");
                    Console.ReadLine();
                }
            }
              
                                  
        }

    }
}
