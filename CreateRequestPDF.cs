using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Aig.Safg.Wnl.BusinessProcess.Utilities.DataServiceContract;
using System.Configuration;
using System.IO;
using System.Data;
using System.Xml;
using System.Globalization;


namespace AWDCustomRestService.DataAccess
{
    public class CreateRequestPDF
    {
        public static XmlDocument doc;

        public PDFStatusObject PerformAction(RequestConfigContract oConfig)
        {
            PDFStatusObject oReturn = new PDFStatusObject();
            RequestDataAccess objReqData = new RequestDataAccess();
            List<KeyData> oKeyData=null;
            RequestAgencyData oReqData = null;
            
            objReqData.GetAWDComments(ref oConfig);
            if (oConfig.ContextBA == Utilities.GetEnumValue("BusinessArea", "Agency"))
            {
                if (string.IsNullOrWhiteSpace(oConfig.CommonFields.ReferenceNum) || string.IsNullOrWhiteSpace(oConfig.CommonFields.CompanyCode) || string.IsNullOrWhiteSpace(oConfig.CommonFields.CompKey))
                {
                    string TaxId = (from iField in oConfig.workInstance.FieldValues
                                    where iField.Name == Utilities.GetEnumValue("RequestAgencyTrans", "TaxID")
                                    select iField.Value).FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(TaxId))
                    {
                        oReqData = Utilities.GetRequestAgencyFields(TaxId);
                        oConfig.CommonFields.ReferenceNum = oReqData.AgentNumber;
                        oConfig.CommonFields.CompKey = oReqData.CompKey;
                        oConfig.CommonFields.CompanyCode = oReqData.CompCode;
                        oConfig.CommonFields.AssociationType = "Agency";
                        oConfig.CommonFields.AssociateionTypeID = 2;
                    }
                }
            }

            #region GetInitialData
            System.Data.DataSet oDs = GetKeyInfo(oConfig);
            if (oDs != null)
            {
                if ((oDs.Tables.Count > 0) && (oDs.Tables[0].Rows.Count > 0))
                {
                    oKeyData = new List<KeyData>();
                    KeyData oField;
                    DataRow oRow = oDs.Tables[0].Rows[0];
                    for (int i = 0; i < oRow.ItemArray.Length; i++)
                    {
                        oField = new KeyData();
                        oField.FieldName = oDs.Tables[0].Columns[i].ColumnName;
                        oField.FieldValue = oRow[i].ToString();
                        oKeyData.Add(oField);
                    }
                }
            }

            string str_FontName = ConfigurationManager.AppSettings["FontName"];
            string str_TabReqData = "REQUEST DATA";
            string str_TabAttach = "ATTACHMENTS";
            string str_TabNotes = "NOTES";
            string str_TabHistory = "HISTORY";
            string str_TabKeyData = "KEY DATA";
            string str_TabComent = "COMMENTS";

            // Pre-defined Columns in the Datagrids.
            int MainTableColumn = 3;
            int CommonDataColumn = 2;
            int KeyDataColumn = 2;
            int RequestDataMaster = 3;
            int RequestDataChild1 = 2;
            int RequestDataChild2 = 2;
            int RequestAttColumn = 5;
            int HistoryColumn = 6;
            int CommentColumn = 3;

            //pre-defined Rows in DataRequest grid
            int n_NoRows = 6;

            //Data Count variable
            int n_ReqData = 0;
            int n_AttData = 0;
            int n_HistoryData = 0;
            int n_CommentData = 0;
            int NonGridCount = 0;

            //Dynamic Grid Flag
            bool b_firstflg = false;

            // FONT Style declartion
            iTextSharp.text.Font f_Header = iTextSharp.text.FontFactory.GetFont(str_FontName, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLACK);
            iTextSharp.text.Font f_SubCaption = iTextSharp.text.FontFactory.GetFont(str_FontName, 9, iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLUE);
            iTextSharp.text.Font f_TableHeader = iTextSharp.text.FontFactory.GetFont(str_FontName, 9, iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLACK);
            iTextSharp.text.Font f_Commondata = iTextSharp.text.FontFactory.GetFont(str_FontName, 8, iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLACK);
            iTextSharp.text.Font f_data = iTextSharp.text.FontFactory.GetFont(str_FontName, 7, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.BLACK);

            try
            {
                string str_RequestType = " REQUEST TYPE : " + oConfig.CommonFields.RequestTypeName;//oTemplate.CommonFields.RequestTypeName;
                string str_RequestItem = " REQUEST ITEM #  " + oConfig.CommonFields.RequestItemID.ToString();//oTemplate.CommonFields.RequestItemID.ToString();

                Document document = new Document(PageSize.A4);

                //Image/Logo for the page
                string imagePath = HttpContext.Current.Server.MapPath("~/Content/WNLLogo.jpg");
                //Jpeg img_WNLTitle = new Jpeg(new Uri(ConfigurationManager.AppSettings["IMGFileName"]));//new Jpeg(ConfigurationManager.AppSettings["IMGFileName"]);
                //img_WNLTitle.ScalePercent(70f);

                Phrase p_BlankLine = new Phrase();
                Phrase p_SubHeading = new Phrase();
                Phrase p_Data = new Phrase();
                Phrase[] p_GLITEM = new Phrase[1];

                //Assigning  Font and SubHeading Caption to Chunk object
                // Chunk c_RequestType = new Chunk(str_RequestType.ToUpper(), f_Header);
                Chunk c_SubHeading_ReqData = new Chunk(str_TabReqData, f_SubCaption);
                Chunk c_SubHeading_AttData = new Chunk(str_TabAttach, f_SubCaption);
                Chunk c_SubHeading_History = new Chunk(str_TabHistory, f_SubCaption);
                Chunk c_SubHeading_Comments = new Chunk(str_TabComent, f_SubCaption);

                Chunk c_SubHeading_Notes = new Chunk(str_TabNotes, f_SubCaption);
                Chunk[] c_GLITEM = new Chunk[1];

                //Header Table
                PdfPTable t_HeaderTable = new PdfPTable(1);
                t_HeaderTable.WidthPercentage = 50f;
                t_HeaderTable.HorizontalAlignment = Element.ALIGN_LEFT;
                t_HeaderTable.DefaultCell.Border = 0;
                t_HeaderTable.DefaultCell.BorderColor = iTextSharp.text.Color.WHITE;
                //New Line declration
                p_BlankLine.Add(Environment.NewLine);


                //PDF declaration
                string FileName = ConfigurationManager.AppSettings["PDFRepository"] + oConfig.CommonFields.PackageID.ToString() + ".pdf";  //oRequest.PackageID.ToString()
                try
                {
                    FileStream lobjFileStream = new FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None);
                    PdfWriter.GetInstance(document, lobjFileStream);
                    document.Open();
                }
                catch (Exception ex)
                {
                    Utilities.LogError(ex.Message, "CreateRequestPDF");
                    oReturn.IsSuccess = false;
                    oReturn.DisplayMessage = ex.Message;
                    return oReturn;
                    //throw;
                }
                #endregion
                #region Common And Key Data
                // RequestData Tab is formated to PDF  -- 1
                // Assiging Row and Column to RequestData PDF table

                //Outer Table
                PdfPTable t_MainTable = new PdfPTable(MainTableColumn);
                t_MainTable.WidthPercentage = 100f;
                t_MainTable.HorizontalAlignment = Element.ALIGN_LEFT;
                t_MainTable.DefaultCell.Border = 0;

                //CommonData Table
                PdfPTable t_CommData = new PdfPTable(CommonDataColumn);
                t_CommData.WidthPercentage = 100f;
                t_CommData.HorizontalAlignment = Element.ALIGN_LEFT;

                if (oKeyData == null)
                {
                    PdfPCell CommTbHeader1 = new PdfPCell(new Paragraph("COMMON DATA", f_TableHeader));
                    CommTbHeader1.GrayFill = 0.7f;
                    //Alignment
                    CommTbHeader1.HorizontalAlignment = Element.ALIGN_CENTER;
                    CommTbHeader1.VerticalAlignment = Element.ALIGN_MIDDLE;
                    t_CommData.AddCell(new PdfPCell(new Paragraph("COMMON DATA", f_TableHeader))
                    {
                        Colspan = 2,
                        GrayFill = 0.7f,
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        VerticalAlignment = Element.ALIGN_MIDDLE
                    });
                }

                //KeyData Table
                PdfPTable t_KeyData = new PdfPTable(KeyDataColumn);
                t_KeyData.WidthPercentage = 100f;
                t_KeyData.HorizontalAlignment = Element.ALIGN_LEFT;

                // Data Table creation for Common Data fields
                DataTable dt_Commtable = new DataTable();
                dt_Commtable.Columns.Add("ColName", typeof(string));
                dt_Commtable.Columns.Add("Value", typeof(string));


                if (oConfig.CommonFields != null)
                {

                    List<PickList> pickList = oConfig.PickLists;
                    Int64 pickVal = Convert.ToInt32(oConfig.CommonFields.ReasonCode);
                    PickList Query = pickList.Single(S => S.PicklistDetailsID.Equals(pickVal));
                    string reqVal = Query.PicklistDetailsDescription;
                    string Stat = String.Empty;
                    if (oConfig.CommonFields.StatusID == 1)
                    {
                        Stat = "InProgess";
                    }
                    else if (oConfig.CommonFields.StatusID == 2)
                    {
                        Stat = "Complete";
                    }
                    else if (oConfig.CommonFields.StatusID == 3)
                    {
                        Stat = "Cancelled";
                    }
                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));

                    dt_Commtable.Rows.Add("Association Type", oConfig.CommonFields.AssociationType); //oTemplate.CommonFields.AssociationType
                    dt_Commtable.Rows.Add(oConfig.CommonFields.AssociationType + " Number", oConfig.CommonFields.ReferenceNum);  //oTemplate.CommonFields.ReferenceNum)
                    dt_Commtable.Rows.Add("Company", oConfig.CommonFields.CompanyCode);  //oTemplate.CommonFields.CompanyCode
                    dt_Commtable.Rows.Add("Reason for Request", reqVal);  //oTemplate.CommonFields.Reason 
                    dt_Commtable.Rows.Add("Status", Stat);  //oTemplate.CommonFields.Status
                    dt_Commtable.AcceptChanges();
                }

                // Data Table creation for Key Data Fields
                DataTable dt_KeyData = new DataTable();
                dt_KeyData.Columns.Add("ColName1", typeof(string));
                dt_KeyData.Columns.Add("Value1", typeof(string));

                //Outer Table Header
                // Setting Width of the Column to Outer Table
                float[] OuterTablewidth = new float[] { 3f, 0.1f, 3f };
                t_MainTable.SetWidths(OuterTablewidth);
                t_MainTable.HorizontalAlignment = 0;

                // Setting Width of the Column for CommonData and KeyData Table
                float[] CommonTablewidth = new float[] { 1.5f, 2.5f };
                t_CommData.SetWidths(CommonTablewidth);
                t_KeyData.SetWidths(CommonTablewidth);
                t_CommData.HorizontalAlignment = 0;
                t_KeyData.HorizontalAlignment = 0;
                
                PdfPCell MainTbHeader1 = new PdfPCell(new Paragraph("COMMON DATA", f_TableHeader));
                PdfPCell MainTbHeader2 = new PdfPCell(new Paragraph("", f_TableHeader));
                PdfPCell MainTbHeader3;
                if (ValidTab(oConfig, str_TabKeyData))
                {
                    MainTbHeader3 = new PdfPCell(new Paragraph("KEY DATA", f_TableHeader));
                }
                else
                {
                    MainTbHeader3 = new PdfPCell(new Paragraph("", f_TableHeader));
                }

                //Header Color
                MainTbHeader1.GrayFill = 0.7f;
                MainTbHeader3.GrayFill = 0.7f;

                //Alignment
                MainTbHeader1.HorizontalAlignment = Element.ALIGN_CENTER;
                MainTbHeader1.VerticalAlignment = Element.ALIGN_MIDDLE;
                MainTbHeader3.HorizontalAlignment = Element.ALIGN_CENTER;
                MainTbHeader3.VerticalAlignment = Element.ALIGN_MIDDLE;

                //Adding OuterHeader to the Outer Table
                if (oKeyData != null)
                {
                    t_MainTable.AddCell(MainTbHeader1);
                    t_MainTable.AddCell("");
                    t_MainTable.AddCell(MainTbHeader3);
                }

                //CommonData Table Header
                PdfPCell TableHeader1 = new PdfPCell(new Paragraph("FIELD", f_TableHeader));
                PdfPCell TableHeader2 = new PdfPCell(new Paragraph("VALUE", f_TableHeader));
                TableHeader1.GrayFill = 0.7f;
                TableHeader2.GrayFill = 0.7f;
                TableHeader1.HorizontalAlignment = Element.ALIGN_CENTER;
                TableHeader1.VerticalAlignment = Element.ALIGN_MIDDLE;
                TableHeader2.HorizontalAlignment = Element.ALIGN_CENTER;
                TableHeader2.VerticalAlignment = Element.ALIGN_MIDDLE;

                //Adding CommonData Header to the CommonData Table
                t_CommData.AddCell(TableHeader1);
                t_CommData.AddCell(TableHeader2);

                if (dt_Commtable.Rows.Count > 0)
                {
                    int ColCount = 0;
                    int RowCount = 0;
                    PdfPCell C_Commondt_Field;
                    PdfPCell C_Commondt_Value;

                    if (ValidTab(oConfig, str_TabKeyData) && oKeyData != null)
                    {
                        for (int i = 0; i < oKeyData.Count; i++)
                        {
                            if (RowCount < dt_Commtable.Rows.Count)
                            {
                                C_Commondt_Field = new PdfPCell(new Paragraph(dt_Commtable.Rows[i][ColCount].ToString(), f_data));
                                ColCount++;
                                C_Commondt_Value = new PdfPCell(new Paragraph(dt_Commtable.Rows[i][ColCount].ToString(), f_data));
                                ColCount = 0;
                                RowCount++;
                            }
                            else
                            {
                                C_Commondt_Field = new PdfPCell(new Paragraph("", f_data));
                                C_Commondt_Value = new PdfPCell(new Paragraph("", f_data));
                            }

                            //Set the Cell Height
                            C_Commondt_Field.FixedHeight = 15f;
                            C_Commondt_Value.FixedHeight = 15f;

                            //Aligning the Cells
                            C_Commondt_Field.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_Commondt_Field.VerticalAlignment = Element.ALIGN_CENTER;
                            C_Commondt_Value.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_Commondt_Value.VerticalAlignment = Element.ALIGN_CENTER;

                            //ADD the cell to CommonData Table
                            t_CommData.AddCell(C_Commondt_Field);
                            t_CommData.AddCell(C_Commondt_Value);
                        }
                    }
                    else
                    {
                        // If Key Data is not loading
                        for (int i = 0; i < dt_Commtable.Rows.Count; i++)
                        {
                            C_Commondt_Field = new PdfPCell(new Paragraph(dt_Commtable.Rows[i][ColCount].ToString(), f_data));
                            ColCount++;
                            C_Commondt_Value = new PdfPCell(new Paragraph(dt_Commtable.Rows[i][ColCount].ToString(), f_data));
                            ColCount = 0;
                            RowCount++;

                            //Set the Cell Height
                            C_Commondt_Field.FixedHeight = 15f;
                            C_Commondt_Value.FixedHeight = 15f;

                            //Aligning the Cells
                            C_Commondt_Field.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_Commondt_Field.VerticalAlignment = Element.ALIGN_CENTER;
                            C_Commondt_Value.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_Commondt_Value.VerticalAlignment = Element.ALIGN_CENTER;

                            //ADD the cell to CommonData Table
                            t_CommData.AddCell(C_Commondt_Field);
                            t_CommData.AddCell(C_Commondt_Value);
                        }
                    }
                }
                t_KeyData.AddCell(TableHeader1);
                t_KeyData.AddCell(TableHeader2);

                PdfPCell c_CommonTable = new PdfPCell(t_CommData);
                PdfPCell c_EmptyCell = new PdfPCell(new Paragraph(""));
                PdfPCell c_KeyDataTable;
                if (oKeyData != null)
                {
                    if (oKeyData.Count > 0)
                    {

                        foreach (KeyData objKeyFld in oKeyData)
                        {

                            PdfPCell C_KeyData_Field = new PdfPCell(new Paragraph(objKeyFld.FieldName, f_data));
                            PdfPCell C_KeyData_Value = new PdfPCell(new Paragraph(objKeyFld.FieldValue, f_data));

                            //Set the Cell Height
                            C_KeyData_Field.FixedHeight = 15f;
                            C_KeyData_Value.FixedHeight = 15f;

                            //Aligning the Cells
                            C_KeyData_Field.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_KeyData_Field.VerticalAlignment = Element.ALIGN_CENTER;
                            C_KeyData_Value.HorizontalAlignment = Element.ALIGN_LEFT;
                            C_KeyData_Value.VerticalAlignment = Element.ALIGN_CENTER;

                            //ADD the cell to CommonData Table
                            t_KeyData.AddCell(C_KeyData_Field);
                            t_KeyData.AddCell(C_KeyData_Value);

                        }


                        //Adding CommonData Table in the MainTable

                        if (ValidTab(oConfig, str_TabKeyData))
                        {
                            c_KeyDataTable = new PdfPCell(t_KeyData);
                        }
                        else
                        {
                            c_KeyDataTable = new PdfPCell(new Paragraph(""));
                            c_KeyDataTable.BorderColor = iTextSharp.text.Color.WHITE;
                        }

                        //Setting Empty Cell border 
                        //c_EmptyCell.BorderColor = iTextSharp.text.Color.WHITE;
                        c_EmptyCell.UseVariableBorders = true;
                        c_EmptyCell.BorderColorTop = iTextSharp.text.Color.WHITE;
                        c_EmptyCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                        c_EmptyCell.BorderColorLeft = iTextSharp.text.Color.BLACK;
                        c_EmptyCell.BorderColorRight = iTextSharp.text.Color.WHITE;

                        //Assiging CommonTable and KeyData table to the Master table
                        t_MainTable.AddCell(c_CommonTable);
                        t_MainTable.AddCell(c_EmptyCell);
                        t_MainTable.AddCell(c_KeyDataTable);
                    }
                }
               
                #endregion
                #region Fields For Request Data Tab
                //********************************************************************          
                // RequestData Tab is formated to PDF  -- 2
                // Creating Outer Table
                PdfPTable t_RequestDataMaster = new PdfPTable(RequestDataMaster);
                t_RequestDataMaster.WidthPercentage = 100f;
                t_RequestDataMaster.HorizontalAlignment = Element.ALIGN_LEFT;
                t_RequestDataMaster.DefaultCell.Border = 0;

                // Setting Width of the Column width
                float[] RequestDataMaster_Tablewidth = new float[] { 3f, 0.1f, 3f };
                t_RequestDataMaster.SetWidths(RequestDataMaster_Tablewidth);
                t_RequestDataMaster.HorizontalAlignment = 0;

                // Creating Inner Table1
                PdfPTable t_RequestDataChild1 = new PdfPTable(RequestDataChild1);
                t_RequestDataChild1.WidthPercentage = 100f;
                t_RequestDataChild1.HorizontalAlignment = Element.ALIGN_LEFT;

                // Creating Inner Table2
                PdfPTable t_RequestDataChild2 = new PdfPTable(RequestDataChild2);
                t_RequestDataChild2.WidthPercentage = 100f;
                t_RequestDataChild2.HorizontalAlignment = Element.ALIGN_LEFT;

                // Setting Width of the Column width
                float[] RequestData1_Tablewidth = new float[] { 1.5f, 2.5f };
                t_RequestDataChild1.SetWidths(RequestData1_Tablewidth);
                t_RequestDataChild2.SetWidths(RequestData1_Tablewidth);
                t_RequestDataChild1.HorizontalAlignment = 0;
                t_RequestDataChild2.HorizontalAlignment = 0;

                PdfPTable t_RequestData_DynamicMaster = new PdfPTable(1);
                t_RequestData_DynamicMaster.WidthPercentage = 100f;
                t_RequestData_DynamicMaster.HorizontalAlignment = Element.ALIGN_LEFT;
                t_RequestData_DynamicMaster.DefaultCell.Border = 0;
                bool HasGrid = false;

                if (ValidTab(oConfig, str_TabReqData))
                {
                    if (oConfig.TemplateFields != null)
                    {

                        PdfPCell Header1 = new PdfPCell(new Paragraph("FIELD", f_TableHeader));
                        PdfPCell Header2 = new PdfPCell(new Paragraph("VALUE", f_TableHeader));
                        Header1.GrayFill = 0.7f;
                        Header2.GrayFill = 0.7f;
                        Header1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Header1.VerticalAlignment = Element.ALIGN_MIDDLE;
                        Header2.HorizontalAlignment = Element.ALIGN_CENTER;
                        Header2.VerticalAlignment = Element.ALIGN_MIDDLE;


                        t_RequestDataChild1.AddCell(Header1);
                        t_RequestDataChild1.AddCell(Header2);

                        int i_DataCount = 0;
                        //if (oTemplate.TemplateFields.Count() > n_NoRows)
                        {
                            t_RequestDataChild2.AddCell(Header1);
                            t_RequestDataChild2.AddCell(Header2);
                        }

                        int gridcnt = 0;
                        n_NoRows = oConfig.TemplateFields.Count();
                        string MigratedReqItem = string.Empty;
                        if (oConfig.workInstance != null)
                        {
                            MigratedReqItem = (from ifield in oConfig.workInstance.FieldValues
                                               where ifield.Name == Utilities.GetEnumValue("RequestData", "PackageID")
                                               select ifield.Value).FirstOrDefault();
                        }
                        foreach (TemplateField objTempFld in oConfig.TemplateFields.OrderBy(ofld => ofld.SortOrder))
                        {
                            if (!String.IsNullOrWhiteSpace(objTempFld.NewValue) || objTempFld.DataTypeID == 7)
                            {
                                #region "Getting Templates Data"
                                // Starts
                                if (objTempFld.DataTypeID == Convert.ToInt32(GetEnumValue("DataType", "DataGrid")))
                                {
                                    // Change DEC 4
                                    if (objTempFld.GridDataRows == null)
                                        continue;

                                    DatagridDataRow[] arr_dg_data = new DatagridDataRow[objTempFld.GridDataRows.Count()];
                                    arr_dg_data = objTempFld.GridDataRows.ToArray();


                                    PdfPTable t_RequestData_DynamicGrid = null;
                                    b_firstflg = false;


                                    Array.Resize(ref c_GLITEM, c_GLITEM.Length + 1);
                                    Array.Resize(ref p_GLITEM, c_GLITEM.Length + 1);

                                    //Adding Columns of the grid
                                    foreach (DatagridDataRow obj_DataRow in arr_dg_data)
                                    {
                                        int hVar = 0;


                                        GridColumnValue[] arr_dg_Value = obj_DataRow.GridColumnValues.ToArray();
                                        foreach (GridColumnValue obj_data_value in arr_dg_Value)
                                        {
                                            string Htext = objTempFld.GridColumns[hVar].HeaderText;

                                            if (!b_firstflg)
                                            {
                                                c_GLITEM[gridcnt] = new Chunk(objTempFld.DisplayText, f_SubCaption);
                                                p_GLITEM[gridcnt] = new Phrase();
                                                p_GLITEM[gridcnt].Add(c_GLITEM[gridcnt]);

                                                t_RequestData_DynamicMaster.AddCell(p_GLITEM[gridcnt]);
                                                ++gridcnt;

                                                t_RequestData_DynamicGrid = new PdfPTable(arr_dg_Value.Count());
                                                t_RequestData_DynamicGrid.WidthPercentage = 100f;
                                                t_RequestData_DynamicGrid.HorizontalAlignment = Element.ALIGN_LEFT;
                                                b_firstflg = true;
                                                HasGrid = true;
                                            }
                                            PdfPCell GridColumn = new PdfPCell(new Paragraph(Htext, f_TableHeader)); //4
                                            GridColumn.GrayFill = 0.7f;
                                            t_RequestData_DynamicGrid.AddCell(GridColumn);
                                            hVar++;

                                        }
                                        break;
                                    }

                                    foreach (DatagridDataRow obj_DataRow in arr_dg_data)
                                    {
                                        GridColumnValue[] arr_dg_Value = obj_DataRow.GridColumnValues.ToArray();
                                        GridColumnValue obj_data_value = null;
                                        int i = 0;
                                        foreach (DataGridProperties oColumn in objTempFld.GridColumns)
                                        {

                                            foreach (GridColumnValue item in obj_DataRow.GridColumnValues)
                                            {
                                                if (item.ColumnName == oColumn.ColumnMasterFieldID.ToString())
                                                {
                                                    obj_data_value = item;
                                                    break;
                                                }
                                            }

                                            if (objTempFld.DataGridID == oColumn.DataGridID)
                                            {
                                                if (oColumn.DataTypeID == Convert.ToInt32(GetEnumValue("DataType", "Dollar")))
                                                {
                                                    Double oDoub;
                                                    string strValue = obj_data_value.ColumnValue;
                                                    if (Double.TryParse(strValue, NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, new CultureInfo("en-US"), out oDoub))
                                                        strValue = String.Format("{0:C2}", Convert.ToDouble(oDoub));

                                                    PdfPCell GridValue = new PdfPCell(new Paragraph(strValue, f_data));
                                                    t_RequestData_DynamicGrid.AddCell(GridValue);
                                                }
                                                else if (oColumn.DataTypeID == Convert.ToInt32(GetEnumValue("DataType", "DropDownList")))
                                                {
                                                    PdfPCell GridValue;
                                                    string reqVal;
                                                    RequestConfigContract reqCon = oConfig;
                                                    List<PickList> pickList = reqCon.PickLists;

                                                    var x = obj_data_value.ColumnValue;
                                                    Int64 pickVal = Convert.ToInt64(obj_data_value.ColumnValue);
                                                    PickList Query = pickList.Single(S => S.PicklistDetailsID.Equals(pickVal));
                                                    reqVal = Query.PicklistDetailsDescription;
                                                    GridValue = new PdfPCell(new Paragraph(reqVal, f_data));

                                                    //if (string.IsNullOrWhiteSpace(MigratedReqItem))
                                                    //{
                                                    //    Int64 pickVal = Convert.ToInt64(obj_data_value.ColumnValue);
                                                    //    PickList Query = pickList.Single(S => S.PicklistDetailsID.Equals(pickVal));
                                                    //    reqVal = Query.PicklistDetailsDescription;
                                                    //    GridValue = new PdfPCell(new Paragraph(reqVal, f_data));
                                                    //}
                                                    //else
                                                    //{
                                                    //    reqVal = obj_data_value.ColumnValue;
                                                    //    GridValue = new PdfPCell(new Paragraph(reqVal, f_data));
                                                    //}
                                                    t_RequestData_DynamicGrid.AddCell(GridValue);
                                                }
                                                else
                                                {
                                                    PdfPCell GridValue = new PdfPCell(new Paragraph(obj_data_value.ColumnValue, f_data));
                                                    t_RequestData_DynamicGrid.AddCell(GridValue);
                                                }
                                            }
                                            i = i + 1;
                                        }
                                    }
                                    t_RequestData_DynamicMaster.AddCell(t_RequestData_DynamicGrid);
                                    n_ReqData++;
                                }

                                // Ends
                                else
                                {
                                    string strValue = objTempFld.NewValue;
                                    PdfPCell c1 = new PdfPCell(new Paragraph(objTempFld.DisplayText, f_data));
                                    PdfPCell c2 = null;

                                    //Check For FieldVisibility
                                    if (IsFieldVisible(oConfig, objTempFld))
                                    {

                                        // Change DEC 4
                                        if (objTempFld.DataTypeID == Convert.ToInt32(GetEnumValue("DataType", "Dollar")))
                                        {
                                            Double oDoub;
                                            if (Double.TryParse(strValue, NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, new CultureInfo("en-US"), out oDoub))
                                                strValue = String.Format("{0:C2}", Convert.ToDouble(oDoub));
                                        }
                                        else if (objTempFld.DataTypeID == Convert.ToInt32(GetEnumValue("DataType", "DropDownList")))
                                        {
                                            //if (strValue == "" || strValue == "0")
                                            //    strValue = "----Select----";
                                            RequestConfigContract reqCon = oConfig;
                                            List<PickList> pickList = reqCon.PickLists;
                                            if (string.IsNullOrWhiteSpace(MigratedReqItem))
                                            {
                                                Int64 pickVal = Convert.ToInt64(strValue);
                                                PickList Query = pickList.Single(S => S.PicklistDetailsID.Equals(pickVal));
                                                strValue = Query.PicklistDetailsDescription;
                                            }
                                        }
                                    }
                                    else
                                        strValue = "";

                                    c2 = new PdfPCell(new Paragraph(strValue, f_data));

                                    c1.FixedHeight = 15f;

                                    c1.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c1.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    c2.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c2.VerticalAlignment = Element.ALIGN_MIDDLE;

                                    if (i_DataCount < n_NoRows / 2)
                                    {
                                        t_RequestDataChild1.AddCell(c1);
                                        t_RequestDataChild1.AddCell(c2);
                                        i_DataCount++;
                                    }
                                    else
                                    {
                                        t_RequestDataChild2.AddCell(c1);
                                        t_RequestDataChild2.AddCell(c2);
                                    }
                                    n_ReqData++;
                                    NonGridCount++;
                                }
                                #endregion
                            }
                        }

                        //Adding CommonData Table in the MainTable
                        PdfPCell c_RequestData1 = new PdfPCell(t_RequestDataChild1);
                        //PdfPCell c_EmptyCell = new PdfPCell(new Paragraph(""));
                        PdfPCell c_RequestData2 = new PdfPCell(t_RequestDataChild2);

                        if (t_RequestDataChild2.Rows.Count == 1)
                        {
                            t_RequestDataChild2.Rows.RemoveAt(0);
                            c_RequestData2.Border = 0;
                        }

                        //Setting Empty Cell border 
                        c_EmptyCell = new PdfPCell(new Paragraph(""));
                        c_EmptyCell.UseVariableBorders = true;
                        c_EmptyCell.BorderColorTop = iTextSharp.text.Color.WHITE;
                        c_EmptyCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                        c_EmptyCell.BorderColorLeft = iTextSharp.text.Color.BLACK;
                        c_EmptyCell.BorderColorRight = iTextSharp.text.Color.WHITE;

                        if (i_DataCount < n_NoRows)
                        {
                            c_RequestData2.BorderColor = iTextSharp.text.Color.WHITE;
                        }
                        else
                        {
                            c_RequestData2.BorderColor = iTextSharp.text.Color.BLACK;
                        }

                        // Adding Child tables to the Master Table
                        t_RequestDataMaster.AddCell(c_RequestData1);
                        t_RequestDataMaster.AddCell(c_EmptyCell);
                        t_RequestDataMaster.AddCell(c_RequestData2);
                    }
                }
                #endregion
                #region Format Attachment
                // RequestAttachment Data is formatted to PDF -- 3
                // Assiging Row and Column to RequestAttachmentData PDF table

                PdfPTable t_RequestAttachment = new PdfPTable(RequestAttColumn);
                t_RequestAttachment.WidthPercentage = 100f;
                t_RequestAttachment.HorizontalAlignment = Element.ALIGN_LEFT;

                // To check the Valid Tab
                if (ValidTab(oConfig, str_TabAttach))
                {
                    if (oConfig.Attachments != null)
                    {
                        if (oConfig.Attachments.Count() > 0)
                        {

                            RequestAttachment[] arrAttachment = new RequestAttachment[oConfig.Attachments.Count()];
                            arrAttachment = oConfig.Attachments.ToArray();

                            PdfPCell Header1 = new PdfPCell(new Paragraph("MANDATORY", f_TableHeader));
                            PdfPCell Header2 = new PdfPCell(new Paragraph("DATE", f_TableHeader));
                            PdfPCell Header3 = new PdfPCell(new Paragraph("ATTACHMENT TYPE", f_TableHeader));
                            PdfPCell Header4 = new PdfPCell(new Paragraph("USER ID", f_TableHeader));
                            PdfPCell Header5 = new PdfPCell(new Paragraph("PROXY ID", f_TableHeader));

                            Header1.GrayFill = 0.7f;
                            Header2.GrayFill = 0.7f;
                            Header3.GrayFill = 0.7f;
                            Header4.GrayFill = 0.7f;
                            Header5.GrayFill = 0.7f;

                            Header1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header1.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header2.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header2.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header3.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header3.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header4.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header4.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header5.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header5.VerticalAlignment = Element.ALIGN_MIDDLE;

                            t_RequestAttachment.AddCell(Header1); //Mandatory
                            t_RequestAttachment.AddCell(Header2); //Date
                            t_RequestAttachment.AddCell(Header3); // Type
                            t_RequestAttachment.AddCell(Header4);
                            t_RequestAttachment.AddCell(Header5);

                            foreach (RequestAttachment objAttachment in arrAttachment)
                            {
                                // Printing Y/N in the Mandatory Field

                                PdfPCell c1 = new PdfPCell(new Paragraph(objAttachment.Mandatory, f_data));
                                PdfPCell c2 = new PdfPCell(new Paragraph(objAttachment.AttachmentDate.ToString(), f_data));
                                PdfPCell c3 = new PdfPCell(new Paragraph(objAttachment.AttachmentTypeName, f_data));
                                PdfPCell c4 = new PdfPCell(new Paragraph(objAttachment.UserID, f_data));
                                PdfPCell c5 = new PdfPCell(new Paragraph(objAttachment.ProxyUser, f_data));


                                c1.FixedHeight = 15f;

                                c1.HorizontalAlignment = Element.ALIGN_LEFT;
                                c1.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c2.HorizontalAlignment = Element.ALIGN_LEFT;
                                c2.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c3.HorizontalAlignment = Element.ALIGN_LEFT;
                                c3.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c4.HorizontalAlignment = Element.ALIGN_LEFT;
                                c4.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c5.HorizontalAlignment = Element.ALIGN_LEFT;
                                c5.VerticalAlignment = Element.ALIGN_MIDDLE;

                                t_RequestAttachment.AddCell(c1);
                                t_RequestAttachment.AddCell(c2);
                                t_RequestAttachment.AddCell(c3);
                                t_RequestAttachment.AddCell(c4);
                                t_RequestAttachment.AddCell(c5);
                                n_AttData++;

                            }
                        }
                    }
                }
                #endregion
                #region Format Comments

                PdfPTable t_RequestComments = new PdfPTable(CommentColumn);
                t_RequestComments.WidthPercentage = 100f;
                float[] RequestComments_Tablewidth = new float[] { 2f, 5f, 3f };
                t_RequestComments.SetWidths(RequestComments_Tablewidth);
                t_RequestComments.HorizontalAlignment = Element.ALIGN_LEFT;

                if (oConfig.Comments != null)
                {
                    if (oConfig.Comments.Count > 0)
                    {
                        Comment[] arrComment = new Comment[oConfig.Comments.Count];
                        arrComment = oConfig.Comments.ToArray();

                        PdfPCell Header1 = new PdfPCell(new Paragraph("DATE", f_TableHeader));
                        PdfPCell Header2 = new PdfPCell(new Paragraph("COMMENT", f_TableHeader));
                        PdfPCell Header3 = new PdfPCell(new Paragraph("AUTHOR", f_TableHeader));

                        Header1.GrayFill = 0.7f;
                        Header2.GrayFill = 0.7f;
                        Header3.GrayFill = 0.7f;

                        Header1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Header1.VerticalAlignment = Element.ALIGN_MIDDLE;
                        Header2.HorizontalAlignment = Element.ALIGN_CENTER;
                        Header2.VerticalAlignment = Element.ALIGN_MIDDLE;
                        Header3.HorizontalAlignment = Element.ALIGN_CENTER;
                        Header3.VerticalAlignment = Element.ALIGN_MIDDLE;

                        t_RequestComments.AddCell(Header1);
                        t_RequestComments.AddCell(Header2);
                        t_RequestComments.AddCell(Header3);

                        foreach (Comment objComm in arrComment)
                        {
                            PdfPCell c1 = new PdfPCell(new Paragraph(objComm.CommentDate, f_data));
                            PdfPCell c2 = new PdfPCell(new Paragraph(objComm.CommentText.ToString(), f_data));
                            PdfPCell c3 = new PdfPCell(new Paragraph(objComm.Author, f_data));

                            c1.FixedHeight = 15f;

                            c1.HorizontalAlignment = Element.ALIGN_LEFT;
                            c1.VerticalAlignment = Element.ALIGN_MIDDLE;
                            c2.HorizontalAlignment = Element.ALIGN_LEFT;
                            c2.VerticalAlignment = Element.ALIGN_MIDDLE;
                            c3.HorizontalAlignment = Element.ALIGN_LEFT;
                            c3.VerticalAlignment = Element.ALIGN_MIDDLE;

                            t_RequestComments.AddCell(c1);
                            t_RequestComments.AddCell(c2);
                            t_RequestComments.AddCell(c3);

                            n_CommentData++;
                        }
                    }
                }

                #endregion
                #region Format History

                // History is formatted to PDF -- 4
                // Assiging Row and Column to HistoryData PDF table
                PdfPTable t_History = new PdfPTable(HistoryColumn);
                t_History.WidthPercentage = 100f;
                t_History.HorizontalAlignment = Element.ALIGN_LEFT;

                //To check the Valid Tab
                if (ValidTab(oConfig, str_TabHistory))
                {
                    if (oConfig.History != null)
                    {
                        if (oConfig.History.Count() > 0)
                        {

                            RequestHistory[] arHistory = new RequestHistory[oConfig.History.Count()];
                            //Changed By kartheek to order the history detail by date
                            //arHistory = oTemplate.History.ToArray();
                            //arHistory = oConfig.History.OrderByDescending(His => His.CreatedDate).ToArray();
                            arHistory = oConfig.History.ToArray();
                            //End Change

                            PdfPCell Header1 = new PdfPCell(new Paragraph("DATE", f_TableHeader));
                            PdfPCell Header2 = new PdfPCell(new Paragraph("USER", f_TableHeader));
                            PdfPCell Header3 = new PdfPCell(new Paragraph("DESCRIPTION", f_TableHeader));
                            PdfPCell Header4 = new PdfPCell(new Paragraph("FROM", f_TableHeader));
                            PdfPCell Header5 = new PdfPCell(new Paragraph("TO", f_TableHeader));
                            PdfPCell Header6 = new PdfPCell(new Paragraph("PROXY USER", f_TableHeader));

                            Header1.GrayFill = 0.7f;
                            Header2.GrayFill = 0.7f;
                            Header3.GrayFill = 0.7f;
                            Header4.GrayFill = 0.7f;
                            Header5.GrayFill = 0.7f;
                            Header6.GrayFill = 0.7f;

                            Header1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header1.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header2.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header2.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header3.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header3.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header4.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header4.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header5.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header5.VerticalAlignment = Element.ALIGN_MIDDLE;
                            Header6.HorizontalAlignment = Element.ALIGN_CENTER;
                            Header6.VerticalAlignment = Element.ALIGN_MIDDLE;

                            t_History.AddCell(Header1);
                            t_History.AddCell(Header2);
                            t_History.AddCell(Header3);
                            t_History.AddCell(Header4);
                            t_History.AddCell(Header5);
                            t_History.AddCell(Header6);

                            foreach (RequestHistory objHistory in arHistory)
                            {

                                // PdfPCell c1 = new PdfPCell(new Paragraph(DateTime.Today.ToString("MM/dd/yyyy"), f_data));
                                PdfPCell c1 = new PdfPCell(new Paragraph(objHistory.CreatedDate.ToString(), f_data));

                                PdfPCell c2 = new PdfPCell(new Paragraph(objHistory.UserID.Trim(), f_data));
                                PdfPCell c3 = new PdfPCell(new Paragraph(objHistory.Description.Trim(), f_data));

                                c1.FixedHeight = 15f;

                                c1.HorizontalAlignment = Element.ALIGN_LEFT;
                                c1.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c2.HorizontalAlignment = Element.ALIGN_LEFT;
                                c2.VerticalAlignment = Element.ALIGN_MIDDLE;
                                c3.HorizontalAlignment = Element.ALIGN_LEFT;
                                c3.VerticalAlignment = Element.ALIGN_MIDDLE;

                                t_History.AddCell(c1);
                                t_History.AddCell(c2);
                                t_History.AddCell(c3);

                                if (objHistory.FromData != null)
                                {
                                    PdfPCell c4 = new PdfPCell(new Paragraph(objHistory.FromData.Trim(), f_data));

                                    c4.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c4.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c4);
                                }
                                else
                                {
                                    PdfPCell c4 = new PdfPCell(new Paragraph(" ", f_data));

                                    c4.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c4.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c4);
                                }

                                if (objHistory.ToData != null)
                                {
                                    PdfPCell c5 = new PdfPCell(new Paragraph(objHistory.ToData.Trim(), f_data));
                                    c5.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c5.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c5);
                                }
                                else
                                {
                                    PdfPCell c5 = new PdfPCell(new Paragraph(" ", f_data));
                                    c5.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c5.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c5);
                                }

                                if (objHistory.ProxyUser != null)
                                {
                                    PdfPCell c6 = new PdfPCell(new Paragraph(objHistory.ProxyUser, f_data));
                                    c6.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c6.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c6);

                                }
                                else
                                {
                                    PdfPCell c6 = new PdfPCell(new Paragraph(" ", f_data));
                                    c6.HorizontalAlignment = Element.ALIGN_LEFT;
                                    c6.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    t_History.AddCell(c6);
                                }


                                n_HistoryData++;
                            }
                        }
                    }
                }
                #endregion
                #region "Print the Sections To Document"

                //********************************************************************    
                // Common Data and Key Data rendered to PDF document.
                //p_BlankLine.Add(Environment.NewLine);

                //Image + Requesst Details
                PdfPTable t_TotalHeader = new PdfPTable(2);
                t_TotalHeader.WidthPercentage = 100f;
                t_TotalHeader.DefaultCell.Border = 0;
                //t_TotalHeader.DefaultCell.BorderColor = iTextSharp.text.Color.WHITE;

                PdfPTable t_logoDetails = new PdfPTable(1);
                //t_logoDetails.DefaultCell.Border = 1;
                t_logoDetails.DefaultCell.Border = 0;
                t_logoDetails.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
                t_logoDetails.DefaultCell.VerticalAlignment = Element.ALIGN_BASELINE;
                t_logoDetails.WidthPercentage = 50f;
                PdfPCell c_logo = new PdfPCell(iTextSharp.text.Image.GetInstance(imagePath));
                c_logo.Border = 0;
                c_logo.VerticalAlignment = Element.ALIGN_BASELINE;
                c_logo.BorderColor = iTextSharp.text.Color.WHITE;
                t_logoDetails.AddCell(c_logo);

                t_TotalHeader.AddCell(t_logoDetails);

                PdfPTable t_reqDetails = new PdfPTable(1);
                t_reqDetails.DefaultCell.Border = 0;
                t_reqDetails.DefaultCell.VerticalAlignment = Element.ALIGN_BOTTOM;
                t_reqDetails.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;

                PdfPCell c_reqDet = new PdfPCell(new Phrase(str_RequestType + Environment.NewLine + str_RequestItem, f_Header));
                c_reqDet.PaddingBottom = 20;
                c_reqDet.HorizontalAlignment = Element.ALIGN_LEFT;
                c_reqDet.VerticalAlignment = Element.ALIGN_BOTTOM;
                c_reqDet.BorderColor = iTextSharp.text.Color.WHITE;

                t_reqDetails.AddCell(c_reqDet);
                t_reqDetails.WidthPercentage = 50f;

                t_TotalHeader.AddCell(t_reqDetails);

                document.Add(t_TotalHeader);

                document.Add(p_BlankLine);

                //Common Data -Image and Request Data

                 FormTab oTabKeyData = (from oTabs in oConfig.FormTabs
                                          where oTabs.TabName=="Key Data"
                                          select oTabs).FirstOrDefault();

                 if (oTabKeyData != null && oKeyData != null)
                 {
                     PdfPTable tk_tmp = new PdfPTable(1);
                     PdfPCell ck_emptyCell = new PdfPCell(new Paragraph(""));
                     Chunk c_KeyData = new Chunk(" ", f_SubCaption);
                     ck_emptyCell.Border = 0;
                     tk_tmp.WidthPercentage = 100f;
                     tk_tmp.DefaultCell.Border = 0;
                     tk_tmp.AddCell(ck_emptyCell);
                     tk_tmp.AddCell(new Phrase(c_KeyData));
                     tk_tmp.AddCell(t_MainTable);

                     document.Add(tk_tmp);
                     tk_tmp = null;
                 }

                 if (oKeyData == null)
                 {
                     PdfPTable tk_tmp = new PdfPTable(1);
                     PdfPCell ck_emptyCell = new PdfPCell(new Paragraph(""));
                     Chunk c_KeyData = new Chunk(" ", f_SubCaption);
                     ck_emptyCell.Border = 0;
                     tk_tmp.WidthPercentage = 50f;
                     tk_tmp.DefaultCell.Border = 0;
                     tk_tmp.AddCell(ck_emptyCell);
                     tk_tmp.AddCell(new Phrase(c_KeyData));
                     tk_tmp.AddCell(t_CommData);
                     tk_tmp.HorizontalAlignment = Element.ALIGN_LEFT;
                     document.Add(tk_tmp);
                     tk_tmp = null;
                 }

                foreach (FormTab oTab in oConfig.FormTabs.OrderBy(oTab => oTab.SortOrder))
                {
                    switch (oTab.TabName)
                    {
                        //case "Key Data":
                        //    PdfPTable tk_tmp = new PdfPTable(1);
                        //    PdfPCell ck_emptyCell = new PdfPCell(new Paragraph(""));
                        //    Chunk c_KeyData = new Chunk(" ", f_SubCaption);
                        //    ck_emptyCell.Border = 0;
                        //    tk_tmp.WidthPercentage = 100f;
                        //    tk_tmp.DefaultCell.Border = 0;
                        //    tk_tmp.AddCell(ck_emptyCell);
                        //    tk_tmp.AddCell(new Phrase(c_KeyData));
                        //    tk_tmp.AddCell(t_MainTable);

                        //    document.Add(tk_tmp);
                        //    tk_tmp = null;
                        //    break;
                        case "Attachments":
                            #region " Attachemnt"
                            // New Line added to PDF document  -  Attachment Data
                            if (ValidTab(oConfig, str_TabAttach))
                            {
                                //document.Add(p_BlankLine);
                                if (n_AttData > 0)
                                {

                                    PdfPTable t_tmp = new PdfPTable(1);
                                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                    c_emptyCell.Border = 0;
                                    t_tmp.WidthPercentage = 100f;
                                    t_tmp.DefaultCell.Border = 0;
                                    t_tmp.AddCell(c_emptyCell);
                                    t_tmp.AddCell(new Phrase(c_SubHeading_AttData));
                                    t_tmp.AddCell(t_RequestAttachment);
                                    document.Add(t_tmp);
                                    t_tmp = null;
                                }
                                else
                                {
                                    PdfPTable t_tmp = new PdfPTable(1);
                                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                    str_TabAttach = str_TabAttach + " --- NO DATA";
                                    Chunk c_SubHeading_AttData1 = new Chunk(str_TabAttach, f_SubCaption);
                                    c_emptyCell.Border = 0;
                                    t_tmp.WidthPercentage = 100f;
                                    t_tmp.DefaultCell.Border = 0;
                                    t_tmp.AddCell(c_emptyCell);
                                    t_tmp.AddCell(new Phrase(c_SubHeading_AttData1));
                                    document.Add(t_tmp);
                                    t_tmp = null;


                                }
                                p_SubHeading.Clear();
                            }
                            #endregion
                            break;
                        case "History":
                            #region "History"

                            // New Line added to PDF document  -  History
                            if (ValidTab(oConfig, str_TabHistory))
                            {

                                //document.Add(p_BlankLine);
                                if (n_HistoryData > 0)
                                {
                                    PdfPTable t_tmp = new PdfPTable(1);
                                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                    c_emptyCell.Border = 0;
                                    t_tmp.WidthPercentage = 100f;
                                    t_tmp.DefaultCell.Border = 0;
                                    t_tmp.AddCell(c_emptyCell);
                                    t_tmp.AddCell(new Phrase(c_SubHeading_History));
                                    t_tmp.AddCell(t_History);

                                    document.Add(t_tmp);
                                    t_tmp = null;
                                    //p_SubHeading.Add(c_SubHeading_History);
                                    //document.Add(p_SubHeading);
                                    //document.Add(t_History);
                                }
                                else
                                {

                                    PdfPTable t_tmp = new PdfPTable(1);
                                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                    str_TabHistory = str_TabHistory + " --- NO DATA";
                                    Chunk c_SubHeading_History1 = new Chunk(str_TabHistory, f_SubCaption);
                                    c_emptyCell.Border = 0;
                                    t_tmp.WidthPercentage = 100f;
                                    t_tmp.DefaultCell.Border = 0;
                                    t_tmp.AddCell(c_emptyCell);
                                    t_tmp.AddCell(new Phrase(c_SubHeading_History1));
                                    document.Add(t_tmp);
                                    t_tmp = null;
                                }
                                p_SubHeading.Clear();
                            }

                            #endregion
                            break;
                        case "Request Data":
                            #region "Request Data"


                            // New Line added to PDF document  -  Request Data
                            if (ValidTab(oConfig, str_TabReqData))
                            {
                                if (n_ReqData > 0)
                                {
                                    if (NonGridCount > 0)
                                    {
                                        PdfPTable t_tmp = new PdfPTable(1);
                                        PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                        c_emptyCell.Border = 0;
                                        t_tmp.WidthPercentage = 100f;
                                        t_tmp.DefaultCell.Border = 0;
                                        t_tmp.AddCell(c_emptyCell);
                                        t_tmp.AddCell(new Phrase(c_SubHeading_ReqData));
                                        t_tmp.AddCell(t_RequestDataMaster);

                                        if (HasGrid)
                                        {
                                            t_tmp.AddCell(c_emptyCell);
                                            t_tmp.AddCell(t_RequestData_DynamicMaster);
                                        }

                                        document.Add(t_tmp);
                                        t_tmp = null;
                                    }
                                    else
                                    {
                                        if (HasGrid)
                                        {
                                            PdfPTable t_tmp = new PdfPTable(1);
                                            PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                            c_emptyCell.Border = 0;
                                            t_tmp.WidthPercentage = 100f;
                                            t_tmp.DefaultCell.Border = 0;
                                            t_tmp.AddCell(c_emptyCell);
                                            t_tmp.AddCell(new Phrase(c_SubHeading_ReqData));
                                            t_tmp.AddCell(t_RequestData_DynamicMaster);
                                            document.Add(t_tmp);
                                            t_tmp = null;
                                        }
                                    }

                                }
                                else
                                {
                                    PdfPTable t_tmp = new PdfPTable(1);
                                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                                    str_TabReqData = str_TabReqData + " --- NO DATA";
                                    Chunk c_SubHeading_ReqData1 = new Chunk(str_TabReqData, f_SubCaption);
                                    c_emptyCell.Border = 0;
                                    t_tmp.WidthPercentage = 100f;
                                    t_tmp.DefaultCell.Border = 0;
                                    t_tmp.AddCell(c_emptyCell);
                                    t_tmp.AddCell(new Phrase(c_SubHeading_ReqData));
                                    document.Add(t_tmp);
                                    t_tmp = null;
                                }
                                p_SubHeading.Clear();
                            }
                            #endregion
                            break;
                    }
                }
                #endregion
                #region Adding Comments
                if (n_CommentData > 0)
                {

                    PdfPTable t_tmp = new PdfPTable(1);
                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                    c_emptyCell.Border = 0;
                    t_tmp.WidthPercentage = 100f;
                    t_tmp.DefaultCell.Border = 0;
                    t_tmp.AddCell(c_emptyCell);
                    t_tmp.AddCell(new Phrase(c_SubHeading_Comments));
                    t_tmp.AddCell(t_RequestComments);
                    document.Add(t_tmp);
                    t_tmp = null;
                }
                else
                {
                    PdfPTable t_tmp = new PdfPTable(1);
                    PdfPCell c_emptyCell = new PdfPCell(new Paragraph(""));
                    str_TabComent = str_TabComent + " --- NO DATA";
                    Chunk c_SubHeading_Comm1 = new Chunk(str_TabComent, f_SubCaption);
                    c_emptyCell.Border = 0;
                    t_tmp.WidthPercentage = 100f;
                    t_tmp.DefaultCell.Border = 0;
                    t_tmp.AddCell(c_emptyCell);
                    t_tmp.AddCell(new Phrase(c_SubHeading_Comm1));
                    document.Add(t_tmp);
                    t_tmp = null;
                }
                p_SubHeading.Clear();
                #endregion
                // PDF closed
                document.Close();
                oReturn.DocumentBinary = File.ReadAllBytes(FileName);

                //File.Delete(FileName);
                oReturn.IsSuccess = true;
                oReturn.ResultActionCD = Convert.ToInt32(GetEnumValue("ResultActions", "DisplayPDF"));
                oReturn.ResultAction = GetEnumValue("ResultActions", "DisplayPDF").ToString();
                oReturn.DisplayMessage = FileName;

                // Exception      
            }
            catch (Exception ex)
            {
                oReturn.IsSuccess = false;
                oReturn.DisplayMessage = "Error in Creating PDF File.";
                Utilities.LogError(ex.Message.ToString() + "|" + ex.StackTrace.ToString(), "PDF Error");
            }
            return oReturn;
        }

        private System.Data.DataSet GetKeyInfo(RequestConfigContract oConfig)
        {
            Int32 CompKeyType = 0;
            string CompKey = string.Empty;

            RequestDataAccess objData = new RequestDataAccess();
            if (oConfig.workInstance.BusinessArea == Utilities.GetEnumValue("BusinessArea", "Policy").ToString())
            {
                CompKey = (from ifield in oConfig.workInstance.FieldValues
                           where ifield.Name == Utilities.GetEnumValue("RequestPolicyTrans", "CompKey")
                           select ifield.Value).FirstOrDefault();

                CompKeyType = 1;
            }
            else if (oConfig.workInstance.BusinessArea == Utilities.GetEnumValue("BusinessArea", "Agency").ToString())
            {
                CompKey = (from ifield in oConfig.workInstance.FieldValues
                           where ifield.Name == Utilities.GetEnumValue("RequestAgencyTrans", "CompKey")
                           select ifield.Value).FirstOrDefault();

                CompKeyType = 2;
            }

            System.Data.DataSet oDs = null;
            if (CompKeyType > 0)
            {
                oDs = objData.SetKeyInfo(CompKeyType, oConfig.CommonFields.ReferenceNum, CompKey);
            }

            return oDs;
        } 
        
        
    private bool ValidTab(RequestConfigContract objConfigContract, String TabName)
        {
            if ((objConfigContract.FormTabs != null) && (objConfigContract.FormTabs.Count() > 0))
            {
                FormTab[] arrTab = new FormTab[objConfigContract.FormTabs.Count()];
                arrTab = objConfigContract.FormTabs.ToArray();

                foreach (FormTab objTab in arrTab)
                {
                    if (objTab.TabName.ToUpper().Equals(TabName))
                    {
                        return true;
                    }

                }
                return false;
            }
            else return false;
        }

        public static string GetEnumValue(string itemName, string key)
        {
            //it will return string.Empty if key not found

            if (doc == null)
            {
                doc = new XmlDocument();
                string filePath = System.Configuration.ConfigurationManager.AppSettings["APP_XML"];
                doc.Load(HttpContext.Current.Server.MapPath(filePath));
            }

            string template = "/Entries/{0}/Item[@key='{1}']";
            string search = string.Format(template, itemName, key);
            XmlNode item;
            XmlElement root = doc.DocumentElement;
            item = root.SelectSingleNode(search);
            string value = (item != null ? item.Attributes["value"].Value : string.Empty);
            return value;
        }

        private Boolean IsFieldVisible(RequestConfigContract objConfigContract, TemplateField oField)
        {
            foreach (FieldAccessibility iAccess in objConfigContract.FieldsAccessibility)
            {
                if (oField.RequestTemplateFieldID == iAccess.RequestTemplateFieldID)
                {
                    if (iAccess.OverrideAccessCode == Convert.ToInt32(GetEnumValue("DataType", "DataGrid")))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

    }
}
