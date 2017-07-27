using AWDCustomRestService.DataContract;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace AWDCustomRestService.DataAccess
{
    public class CreatePDF
    {
        public byte[] GetBinaryData(WorkInstance oWork)
        {
            //Generate PDF and store in Temp location
            Byte[] lobjBuffer = null;
            string FileName = Convert.ToString(Guid.NewGuid());
            string FilePath = System.Configuration.ConfigurationManager.AppSettings["PDFRepository"] + "\\" + FileName + ".pdf";
            
            Document document = new Document();
            try
            {
                using (FileStream lobjFileStream = new FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfWriter.GetInstance(document, lobjFileStream);
                    document.Open();
                    document.Add(new Paragraph());
                    AWDDataAccess objAwdDataAccess = new AWDDataAccess();
                    TemplateDetails objTempDetails = objAwdDataAccess.getTemplateDetails(oWork.WorkTemplateID);
                    Font lobjArialH = FontFactory.GetFont("Arial", 10, Color.BLACK);
                    PdfPTable lobjMainTable = new PdfPTable(2);
                    lobjMainTable.WidthPercentage = 100;
                    PdfPCell lobjCell = new PdfPCell();
                    string imagePath = HttpContext.Current.Server.MapPath("~/Content/WNLLogo.jpg");
                    lobjMainTable.AddCell(new PdfPCell(iTextSharp.text.Image.GetInstance(imagePath))
                    {
                        BorderColor = Color.WHITE,
                    });

                    CreateUpperSubTable(oWork, objTempDetails, lobjArialH, lobjMainTable);


                    #region Header Information printing....
                    PdfPTable lobjPolicyInfoTable1;
                    PdfPTable lobjPolicyInfoTable2;
                    string strNewDiv;
                    //call to Header Creator Function
                    CreateHeaderRegion(oWork, out lobjPolicyInfoTable1, out lobjPolicyInfoTable2, out strNewDiv);
                    #endregion
                    
                    document.Add(lobjMainTable);
                    document.Add(lobjPolicyInfoTable1);
                    document.Add(lobjPolicyInfoTable2);

                    if (!String.IsNullOrWhiteSpace(strNewDiv))
                    {
                        PdfPTable lobjFraudTable = CreateFraudTable(strNewDiv);
                        document.Add(lobjFraudTable);
                    }

                    PdfPTable table = CreateFormFieldTable(oWork, objTempDetails);
                    
                    //add the form field table to the document
                    document.Add(table);
                    document.Close();
                }
            }
            catch (System.Exception ex)
            {
                Utilities.LogError(ex.StackTrace, "GetBinaryData");
                throw;
            }

            lobjBuffer = System.IO.File.ReadAllBytes(FilePath);
            return lobjBuffer;

        }

        private static PdfPTable CreateFraudTable(string strNewDiv)
        {
            PdfPTable lobjFraudTable = new PdfPTable(1);

            lobjFraudTable.HorizontalAlignment = 2;
            lobjFraudTable.WidthPercentage = 99;

            Font lobjArial = FontFactory.GetFont("Arial", 10, Color.RED);
            lobjFraudTable.AddCell(new PdfPCell(new Phrase(strNewDiv, lobjArial))
            {
                BorderColor = Color.RED,
                PaddingLeft = 5.0F,
                PaddingRight = 5.0F,
                PaddingBottom = 5.0F,

                BackgroundColor = new Color(255, 227, 231)
            });
            return lobjFraudTable;
        }

        private static PdfPTable CreateFormFieldTable(WorkInstance oWork, TemplateDetails objTempDetails)
        {
            Font lobjArialF = FontFactory.GetFont("Arial", 9, Color.BLACK);
            Font lobjArialG = FontFactory.GetFont("Arial", 11, Color.BLACK);
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 65;
            table.HorizontalAlignment = 0;
            table.SpacingBefore = 30f;

            table.AddCell(new PdfPCell(new Phrase("Form Field Values", lobjArialG))
            {
                Colspan = 2,
                PaddingBottom = 5.0F,
                BackgroundColor = Color.LIGHT_GRAY,
                BorderColor = Color.WHITE,
                HorizontalAlignment = 1
            });
            int LOBCount = 0;
            LOBCount = oWork.FieldValues.Count;
            int NonLOBCount = 0;
            if (oWork.NonLOBFields != null)
            {
                NonLOBCount = oWork.NonLOBFields.Count;
            }

            int TotalCount = LOBCount + NonLOBCount;
            int j = 0;
            for (int i = 1; i <= TotalCount; i++)
            {
                var LOBQuery = from LOB in oWork.FieldValues
                               where LOB.Sequence == i
                               select LOB;
                if (LOBQuery != null)
                {

                    foreach (FieldValue FV in LOBQuery)
                    {
                        String FieldDisplayName = (from T in objTempDetails.TemplateFields
                                                   where T.Dataname == FV.Name
                                                   select T.DisplayText).FirstOrDefault();

                        if (!String.IsNullOrWhiteSpace(FieldDisplayName))
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(FieldDisplayName, lobjArialF));
                            cell.BorderColor = Color.WHITE;
                            cell.HorizontalAlignment = 0;
                            cell.PaddingRight = 5.0F;
                            cell.PaddingBottom = 2.0F;

                            PdfPCell cell0 = new PdfPCell(new Phrase(FV.Value, lobjArialF));
                            cell0.BorderColor = Color.WHITE;
                            cell0.HorizontalAlignment = 0;
                            cell0.PaddingLeft = 5.0f;
                            cell.PaddingBottom = 2.0F;

                            if (j % 2 == 0)
                            {
                                cell.BackgroundColor = Color.WHITE;
                                cell0.BackgroundColor = Color.WHITE;
                            }

                            else
                            {
                                cell.BackgroundColor = Color.LIGHT_GRAY;
                                cell0.BackgroundColor = Color.LIGHT_GRAY;

                            }
                            j = j + 1;
                            table.AddCell(cell);
                            table.AddCell(cell0);
                        }
                    }
                }
                
                    var NonLOBQuery = from NonLOB in oWork.NonLOBFields
                                      where NonLOB.Sequence == i
                                      select NonLOB;

                    if (NonLOBQuery != null)
                    {

                        foreach (FieldValue FV in NonLOBQuery)
                        {

                            if (FV.Name != "chkUser" && FV.Name != "fldComment")
                            {
                                String FieldDisplayName = (from T in objTempDetails.TemplateFields
                                                           where T.Dataname == FV.Name
                                                           select T.DisplayText).FirstOrDefault();

                                PdfPCell cell = new PdfPCell(new Phrase(FieldDisplayName, lobjArialF));
                                cell.BorderColor = Color.WHITE;
                                cell.HorizontalAlignment = 0;
                                cell.PaddingRight = 5.0F;
                                cell.PaddingBottom = 2.0F;

                                PdfPCell cell0 = new PdfPCell(new Phrase(FV.Value, lobjArialF));
                                cell0.BorderColor = Color.WHITE;
                                cell0.HorizontalAlignment = 0;
                                cell0.PaddingLeft = 5.0f;
                                cell.PaddingBottom = 2.0F;

                                if (j % 2 == 0)
                                {
                                    cell.BackgroundColor = Color.WHITE;
                                    cell0.BackgroundColor = Color.WHITE;
                                }

                                else
                                {
                                    cell.BackgroundColor = Color.LIGHT_GRAY;
                                    cell0.BackgroundColor = Color.LIGHT_GRAY;

                                }
                                j = j + 1;

                                table.AddCell(cell);
                                table.AddCell(cell0);
                            }
                        }
                    }

            }
            return table;
        }

        private static void CreateHeaderRegion(WorkInstance oWork, out PdfPTable lobjPolicyInfoTable1, out PdfPTable lobjPolicyInfoTable2, out string strNewDiv)
        {
            lobjPolicyInfoTable1 = new PdfPTable(5);
            lobjPolicyInfoTable1.HorizontalAlignment = 0;
            lobjPolicyInfoTable1.SpacingBefore = 2.0F;
            lobjPolicyInfoTable1.SpacingAfter = 2.0F;

            lobjPolicyInfoTable2 = new PdfPTable(5);
            lobjPolicyInfoTable2.HorizontalAlignment = 0;
            lobjPolicyInfoTable2.WidthPercentage = 50;
            lobjPolicyInfoTable2.SpacingBefore = 2.0F;
            lobjPolicyInfoTable2.SpacingAfter = 10.0F;

            List<String> lstrTitle = new List<string>();
            List<String> lstrValue = new List<string>();
            strNewDiv = String.Empty;
            if (!String.IsNullOrWhiteSpace(oWork.HeaderString))
            {
                oWork.HeaderString = oWork.HeaderString.Substring(0, oWork.HeaderString.Length - 1);
                String[] StrValueArray = oWork.HeaderString.Split(',');

                int FraudCount = StrValueArray.Count();
                if (FraudCount == 15)
                {
                    strNewDiv = StrValueArray.ElementAt(FraudCount - 1);
                    StrValueArray = StrValueArray.Take(14).ToArray();
                }
                Int32 liCounter = 0;
                String lstrPolicyInformation = String.Empty;

                foreach (String s in StrValueArray)
                {
                    if (liCounter == 0)
                    {
                        lobjPolicyInfoTable1.AddCell(new PdfPCell(new Phrase(s, FontFactory.GetFont("Arial", 18, Color.GRAY)))
                        {
                            BorderColor = Color.WHITE,
                            PaddingLeft = 5.0F,
                            PaddingRight = 5.0F,
                            PaddingBottom = 5.0F,
                            NoWrap = true,
                            HorizontalAlignment = 0
                        });
                    }
                    if (liCounter == 2)
                    {
                        lobjPolicyInfoTable1.AddCell(new PdfPCell(new Phrase(s, FontFactory.GetFont("Arial", 18, Color.GRAY)))
                        {
                            BorderColor = Color.WHITE,
                            PaddingLeft = 5.0F,
                            PaddingRight = 5.0F,
                            PaddingBottom = 5.0F,
                            NoWrap = true,
                            HorizontalAlignment = 1
                        });
                    }
                    if (liCounter == 1)
                    {
                        lobjPolicyInfoTable1.AddCell(new PdfPCell(new Phrase(s, FontFactory.GetFont("Arial", 20, 1, new Color(0, 191, 255))))
                        {
                            BorderColor = Color.WHITE,
                            PaddingLeft = 5.0F,
                            PaddingRight = 5.0F,
                            PaddingBottom = 5.0F,
                            NoWrap = true,
                            HorizontalAlignment = 0
                        });
                    }
                    if (liCounter == 3)
                    {
                        lobjPolicyInfoTable1.AddCell(new PdfPCell(new Phrase(s, FontFactory.GetFont("Arial", 18, Color.DARK_GRAY)))
                        {
                            BorderColor = Color.WHITE,
                            PaddingLeft = 5.0F,
                            PaddingRight = 5.0F,
                            PaddingBottom = 5.0F,
                            Colspan = 2,
                            HorizontalAlignment = 0
                        });
                    }
                    if (liCounter >= 4)
                    {
                        if (liCounter % 2 == 0)
                            lstrValue.Add(s);
                        else
                            lstrTitle.Add(s);
                    }

                    liCounter++;
                }


                foreach (String s in lstrValue)
                {
                    Font lobjArial = FontFactory.GetFont("Arial", 12, 1, Color.BLACK);
                    lobjPolicyInfoTable2.AddCell(new PdfPCell(new Phrase(s, lobjArial))
                    {
                        UseVariableBorders = true,
                        BorderColorLeft = Color.WHITE,
                        BorderColorTop = Color.WHITE,
                        BorderColorBottom = Color.WHITE,
                        BorderColorRight = Color.GRAY,
                        PaddingLeft = 5.0F,
                        PaddingRight = 5.0F
                    });
                }
                foreach (String s in lstrTitle)
                {
                    Font lobjArial = FontFactory.GetFont("Arial", 7, Color.GRAY);
                    lobjPolicyInfoTable2.AddCell(new PdfPCell(new Phrase(s, lobjArial))
                    {
                        UseVariableBorders = true,
                        BorderColorLeft = Color.WHITE,
                        BorderColorTop = Color.WHITE,
                        BorderColorBottom = Color.WHITE,
                        BorderColorRight = Color.GRAY,
                        PaddingLeft = 5.0F,
                        PaddingRight = 5.0F
                    });
                }
            }
        }

        private static void CreateUpperSubTable(WorkInstance oWork, TemplateDetails objTempDetails, Font lobjArialH, PdfPTable lobjMainTable)
        {
            PdfPTable lobjInfoTable = new PdfPTable(2);
            lobjInfoTable.HorizontalAlignment = 2;
            lobjInfoTable.AddCell(new PdfPCell(new Phrase("Template Name:", lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });
            lobjInfoTable.AddCell(new PdfPCell(new Phrase(objTempDetails.TemplateFields[0].TemplateName, lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });

            lobjInfoTable.AddCell(new PdfPCell(new Phrase("User ID:", lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });

            lobjInfoTable.AddCell(new PdfPCell(new Phrase(oWork.UserID, lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });

            lobjInfoTable.AddCell(new PdfPCell(new Phrase("Created Date/Time:", lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });
            lobjInfoTable.AddCell(new PdfPCell(new Phrase(System.DateTime.Now.ToString(), lobjArialH))
            {
                BackgroundColor = new Color(220, 220, 220),
                BorderColor = new Color(220, 220, 220)
            });

            lobjMainTable.AddCell(new PdfPCell(lobjInfoTable) { BorderColor = Color.WHITE, PaddingTop = 20.0F, PaddingBottom = 30.0F, GrayFill = 2.0F });
        }
    }
}