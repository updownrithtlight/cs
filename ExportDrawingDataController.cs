using design_cmf.Models.ExportDrawingData;
using design_cmf.Utils;
using Microsoft.AspNetCore.Mvc;
using Oracle.ManagedDataAccess.Client;
using System.Collections.Generic;
using System.Web.Http.Cors;
using System.Xml;
using design_cmf.Models.TcColorChart;
using design_cmf.Models.XtoneInfor;
using design_cmf.Models.InteriorColor;
using System;
using System.Linq;
using design_cmf.Controllers.TcColorChart;
using design_cmf.Controllers.InteriorColor;
using design_cmf.Controllers.XtoneInfor;
using System.IO;
using System.Text;

namespace design_cmf.Controllers.ExportDrawingData
{
    [EnableCors("*", "*", "*")]
    [Route("api/ExportDrawingData")]
    [Produces("application/json")]
    [Consumes("application/json")]
    [ApiController]
    public class ExportDrawingDataController : ControllerBase
    {
        private readonly ExportDrawingDataContext _context;

        public ExportDrawingDataController(ExportDrawingDataContext context)
        {
            _context = context;
        }

        private readonly string[] MNG_ROW_NM = new string[13]
        { "PART NO.", "PART NAME", "COMPLETION DATE",
            "REVISION NO.", "DWG ISSUE ACTION", "ISSUE DATE",
            "DRAWN", "DESIGNED", "CHECKER", "RESPONSIBLE",
            "SUPPLIER'S NAME", "SUPPLIER'S DESIGNERS", "SUPPLIER'S APPROVER" };

        private readonly string[] BASIC_ROW_NM = new string[18]
        { "MATERIAL", "THICKNESS(mm)", "MASS TYPE",
            "MASS(kg)", "REFERENCE STANDARD", "HNS",
            "NH", "Q", "Q COUNT", "NF",
            "C MARK", "COLOR", "ATTACH DATA","SPR","SUPPLIER DWG REQ.",
            "ATTENDANCE REQ. AT INSPECTION","EXPECTED ISSUE DATE","INTERCHANGEABILITY" };

        private readonly string[] COLOR_ROW_NM = new string[17]
        { "SECTION","SUB-SECTION","L1",
            "LVL","PART NO.","PARENT PART NO.",
            "PART NAME","PLANT","MODEL","DESTINATION",
            "GRADE","FEATURE","ITEM","REMARKS","COLOR",
            "L1_BASE PART NO","BASE PART NO"};

        private readonly string[] TC_NM = new string[5]
        { "PLANT","MODEL", "DESTINATION","GRADE","FOP" };

        private readonly string[] EXT_SUMMARY_NM = new string[12]
        { "SECTION","SUB-SECTION", "MAIN NO.","PART NAME",
            "PLANT","MODEL" ,"DESTINATION","GRADE",
            "FEATURE","ITEM","REMARKS","COLOR"};

        private readonly string[] INT_SUMMARY_NM = new string[12]
        { "GROUP","SECTION", "SUB-SECTION","MAIN NO.",
            "PART NAME","PLANT" ,"MODEL","DESTINATION",
            "GRADE","FEATURE","ITEM","REMARKS"};

        private readonly string[] GROUP_ROW_NM = new string[14]
{
            "SECTION","SUB-SECTION","L1","LVL","PART NO.",
            "PARENT PART NO.","PART NAME","PLANT","MODEL",
            "DESTINATION","GRADE","FEATURE","ITEM","REMARKS"};

        private const string ns_ss = "urn:schemas-microsoft-com:office:spreadsheet";

        private const string ns_x = "urn:schemas-microsoft-com:office:excel";

        private long bodyColorHeaderCount = 0;

        List<string> colorHeaderName = new List<string>();

        // GET: api/ExportDrawingData/Ext
        [HttpGet("Ext")]
        public ExportDrawingDataResponse ExportDrawingDataExtXml([FromQuery] ExportDrawingDataGetRequest request)
        {

            ExportDrawingDataResponse response = new ExportDrawingDataResponse();
            response.status = "200";
            response.error = "";

            LogUtil.WriteLog("★EXTのXML出力開始");

            try
            {

                if (!CheckFolder(request.output_path))
                {
                    LogUtil.WriteLog("EXTのXML出力エラー:" + "The file path does not exist");

                    response.status = "400";
                    response.error = "The file path does not exist";
                    return response;
                }

                string xmlPath = request.output_path + "\\" + request.file_nm;

                //XmlTextWriter writer = new XmlTextWriter(xmlPath, null);

                // Xml説明
                //writer.Formatting = Formatting.Indented;
                //writer.Indentation = 4;
                XmlDocument domDoc = new XmlDocument();
                XmlDeclaration nodeDeclar = domDoc.CreateXmlDeclaration("1.0", Encoding.UTF8.BodyName, "");
                domDoc.AppendChild(nodeDeclar);

                XmlProcessingInstruction spi = domDoc.CreateProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");
                domDoc.InsertBefore(spi, domDoc.DocumentElement);

                // rootノード取得
                XmlElement root = domDoc.CreateElement("Workbook");
                root.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
                root.SetAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel");
                root.SetAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40");

                // カラーチャート情報検索
                ColorChartGetRequest colorchartRequest = new ColorChartGetRequest();
                colorchartRequest.color_chart_id = request.color_chart_id;
                IDictionary<string, object> colorChartMap = GetColorChartItems(colorchartRequest);
                var colorchartResults = new List<ColorChartGetResponse>();
                colorchartResults = (List<ColorChartGetResponse>)colorChartMap["resultList"];

                // [DocumentProperties]ノート作成、追加
                CreateDocumentProperties(root, domDoc, colorchartResults);

                // OfficeDocumentSettings、ExcelWorkbook、Stylesの固定値作成、追加
                CreateCommonWorkSheet(root, domDoc);

                // WorkSheet[MANAGEMENT INFO]作成、追加
                if (colorchartResults.Count > 0)
                {
                    CreateMngWorkSheet(root, domDoc, colorchartResults[0].changed_part, colorchartResults[0].part_nm);
                }
                else
                {
                    CreateMngWorkSheet(root, domDoc, "", "");
                }

                // WorkSheet[BASIC SPEC]作成、追加
                CreateBasicWorkSheet(root, domDoc);

                // By bodycolor情報検索
                ByBodyColorInfo byBodyColorResults = GetColorByBodyColorItems(colorchartRequest);

                // Body Colorで入れ替えデータ格納リスト
                List<ColorListInfoGetResponse> bodyColorToHes = new List<ColorListInfoGetResponse>();

                // [BY BODYCOLOR]シートへの出力データ格納リスト
                List<ColorTabInfo> byBodyColorList = new List<ColorTabInfo>();

                // [COLOR]シートへの出力データ格納リスト
                List<ColorTabInfo> colorList = new List<ColorTabInfo>();

                // [BY BODYCOLOR]、[COLOR]シートデータ取得
                SetColorInfo(byBodyColorResults, bodyColorToHes, byBodyColorList, colorList);

                // [BY BODYCOLOR]作成
                XmlElement byBodyColorNode = CreateWorksheetByColor(root, domDoc, byBodyColorList);

                // [COLOR]作成、追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //CreateColorWorkSheet(root, domDoc, colorList);
                if (colorList.Count != 0)
                {
                    CreateColorWorkSheet(root, domDoc, colorList);
                }
                // MODIFY 2022/05/30 END

                // [BY BODYCOLOR]ノード追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //root.AppendChild(byBodyColorNode);
                if (byBodyColorList.Count != 0)
                {
                    root.AppendChild(byBodyColorNode);
                }
                // MODIFY 2022/05/30 END

                // [COLOR SUMMARY]作成、追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //CreateWorksheetColorSummary(root, domDoc, colorList);
                if (colorList.Count != 0)
                {
                    CreateWorksheetColorSummary(root, domDoc, colorList);
                }
                // MODIFY 2022/05/30 END

                // [BY BODYCOLOR SUMMARY]作成、追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //CreateWorksheetByBodyColorColorSummary(root, domDoc, byBodyColorList);
                if (byBodyColorList.Count != 0)
                {
                    CreateWorksheetByBodyColorColorSummary(root, domDoc, byBodyColorList);
                }
                // MODIFY 2022/05/30 END

                // [CODE LIST]作成、追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //CreateColorListWorkSheet(root, domDoc, request.color_chart_id, "Ext", bodyColorToHes);
                if (bodyColorToHes.Count != 0)
                {
                    CreateColorListWorkSheet(root, domDoc, request.color_chart_id, "Ext", bodyColorToHes);
                }
                // MODIFY 2022/05/30 END

                // アンマッチ情報検索
                // UnmatchInfo unmatchColorResults = GetUnmatchColorByBodyColorItems(request);

                // [UNMATCH LIST]シートへの出力データ格納リスト
                // List<UnmatchTabInfo> unmatchColorList = new List<UnmatchTabInfo>();

                // [UNMATCH]シートデータ作成
                // SetUnmatchInfo(unmatchColorResults, unmatchColorList);

                // if (unmatchColorList.Count != 0)
                // {
                //     // [UNMATCH LIST]作成
                //     XmlElement unmatchColorNode = CreateWorksheetunmatchColor(root, domDoc, unmatchColorList);
                //     root.AppendChild(unmatchColorNode);
                // }

                domDoc.AppendChild(root);
                domDoc.Save(xmlPath);
                //domDoc.WriteTo(writer);
                //writer.Flush();
                //writer.Close();

                // 余計な「amp;」除く
                ReplaceCharacter(xmlPath);

                LogUtil.WriteLog("★EXTのXML出力終了");

            }
            catch (Exception ex)
            {
                LogUtil.WriteLog("EXTのXML出力エラー:" + ex.Message);
                LogUtil.WriteLog("EXTのXML出力エラー:" + ex.StackTrace);
                response.status = "400";
                response.error = "Fail to output XML file.";
            }

            return response;

        }

        // GET: api/ExportDrawingData/Int
        [HttpGet("Int")]
        public ExportDrawingDataResponse ExportDrawingDataIntXml([FromQuery] ExportDrawingDataGetRequest request)
        {
            ExportDrawingDataResponse response = new ExportDrawingDataResponse();
            response.status = "200";
            response.error = "";

            LogUtil.WriteLog("★INTのXML出力開始");

            try
            {

                if (!CheckFolder(request.output_path))
                {

                    LogUtil.WriteLog("INTのXML出力エラー:" + "The file path does not exist");

                    response.status = "400";
                    response.error = "The file path does not exist";
                    return response;
                }

                string xmlPath = request.output_path + "\\" + request.file_nm;

                //XmlTextWriter writer = new XmlTextWriter(xmlPath, null);

                // Xml説明
                //writer.Formatting = Formatting.Indented;
                //writer.Indentation = 4;
                XmlDocument domDoc = new XmlDocument();
                XmlDeclaration nodeDeclar = domDoc.CreateXmlDeclaration("1.0", System.Text.Encoding.UTF8.BodyName, "");
                domDoc.AppendChild(nodeDeclar);

                XmlProcessingInstruction spi = domDoc.CreateProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");
                domDoc.InsertBefore(spi, domDoc.DocumentElement);

                // rootノード取得
                XmlElement root = domDoc.CreateElement("Workbook");
                root.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
                root.SetAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel");
                root.SetAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40");

                // カラーチャート情報検索
                ColorChartGetRequest colorchartRequest = new ColorChartGetRequest();
                colorchartRequest.color_chart_id = request.color_chart_id;

                IDictionary<string, object> colorChartMap = GetColorChartItems(colorchartRequest);
                var colorchartResults = new List<ColorChartGetResponse>();
                colorchartResults = (List<ColorChartGetResponse>)colorChartMap["resultList"];

                // [DocumentProperties]ノート作成、追加
                CreateDocumentProperties(root, domDoc, colorchartResults);

                // OfficeDocumentSettings、ExcelWorkbook、Stylesの固定値作成、追加
                CreateCommonWorkSheet(root, domDoc);

                // WorkSheet[MANAGEMENT INFO]作成、追加
                if (colorchartResults.Count > 0)
                {
                    CreateMngWorkSheet(root, domDoc, colorchartResults[0].changed_part, colorchartResults[0].part_nm);
                }
                else
                {
                    CreateMngWorkSheet(root, domDoc, "", "");
                }

                // WorkSheet[BASIC SPEC]作成、追加
                CreateBasicWorkSheet(root, domDoc);

                // [SUMMARY]ヘッダの集合に使う
                Dictionary<string, GroupIntColorTypeResponse> sumTableHeader = new Dictionary<string, GroupIntColorTypeResponse>();

                //重複した列を保存するためのID
                List<GroupIntColorTypeResponseSum> sameColumId = new List<GroupIntColorTypeResponseSum>();

                // [SUMMARY]データの集合に使う
                List<SummaryItemsResponse> sumTableData = new List<SummaryItemsResponse>();

                // [GROUP]作成、追加
                int createGroupSuccessFlag = CreatGroupWorkSheet(root, domDoc, request.color_chart_id, sumTableHeader, sumTableData, (long)request.color_chart_tc_id);
                if (createGroupSuccessFlag == 0)
                {
                    response.status = "9999";
                    response.error = "Groups are empty";
                    //writer.Close();
                    System.IO.File.Delete(xmlPath);
                    return response;
                }

                // [SUMMARY]ヘッダの集合処理
                CreateWorksheetIntSummaryHead(sumTableHeader, sameColumId);

                // [SUMMARY]表のデータ集合処理
                CreateWorksheetIntSummaryData(sumTableData);

                // [SUMMARY]作成、追加
                // MODIFY 2022/05/30 BEGIN COLORシートルール
                //CreateWorksheetIntSummary(root, domDoc, sumTableHeader, sumTableData, sameColumId);
                if (sumTableData.Count != 0)
                {
                    CreateWorksheetIntSummary(root, domDoc, sumTableHeader, sumTableData, sameColumId);
                }
                // MODIFY 2022/05/30 END

                // [CODE LIST]作成、追加
                CreateColorListWorkSheet(root, domDoc, request.color_chart_id, "Int", null);

                // CreatUnmatchWorkSheet(root, domDoc, request.color_chart_id, (long)request.color_chart_tc_id);

                domDoc.AppendChild(root);
                domDoc.Save(xmlPath);
                //domDoc.WriteTo(writer);
                //writer.Flush();
                //writer.Close();

                // 余計な「amp;」除く
                ReplaceCharacter(xmlPath);

                LogUtil.WriteLog("★INTのXML出力終了");
            }

            catch (Exception ex)
            {
                LogUtil.WriteLog("INTのXML出力エラー:" + ex.Message);
                LogUtil.WriteLog("INTのXML出力エラー:" + ex.StackTrace);

                response.status = "400";
                response.error = "Fail to output XML file.";
            }

            return response;

        }

        // GET: api/ExportDrawingData/TC
        [HttpGet("TC")]
        public ExportDrawingDataResponse ExportDrawingDataTcXml([FromQuery] ExportDrawingDataGetRequest request)
        {

            ExportDrawingDataResponse exportDrawingDataResponse = new ExportDrawingDataResponse();
            exportDrawingDataResponse.status = "200";
            exportDrawingDataResponse.error = "";

            LogUtil.WriteLog("★T/CのXML出力開始");

            try
            {

                if (!CheckFolder(request.output_path))
                {

                    LogUtil.WriteLog("T/CのXML出力エラー:" + "The file path does not exist");

                    exportDrawingDataResponse.status = "400";
                    exportDrawingDataResponse.error = "The file path does not exist";
                    return exportDrawingDataResponse;
                }
                // ADD 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
                string xmlPath = request.output_path + "\\" + request.file_nm;
                // ADD 2022/04/06 END

                //XmlTextWriter writer = new XmlTextWriter(request.output_path + "\\" + request.file_nm, null);

                // Xml説明
                //writer.Formatting = Formatting.Indented;
                //writer.Indentation = 4;
                XmlDocument domDoc = new XmlDocument();
                XmlDeclaration nodeDeclar = domDoc.CreateXmlDeclaration("1.0", System.Text.Encoding.UTF8.BodyName, "");
                domDoc.AppendChild(nodeDeclar);

                XmlProcessingInstruction spi = domDoc.CreateProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");
                domDoc.InsertBefore(spi, domDoc.DocumentElement);

                // rootノード取得
                XmlElement root = domDoc.CreateElement("Workbook");
                root.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
                root.SetAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel");
                root.SetAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet");
                root.SetAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40");

                // カラーチャート情報検索
                ColorChartGetRequest colorchartRequest = new ColorChartGetRequest();
                colorchartRequest.color_chart_id = request.color_chart_id;

                IDictionary<string, object> colorChartMap = GetColorChartItems(colorchartRequest);
                var colorchartResults = new List<ColorChartGetResponse>();
                colorchartResults = (List<ColorChartGetResponse>)colorChartMap["resultList"];

                // [DocumentProperties]ノート作成、追加
                CreateDocumentProperties(root, domDoc, colorchartResults);

                // OfficeDocumentSettings、ExcelWorkbook、Stylesの固定値作成、追加
                CreateCommonWorkSheet(root, domDoc);

                // WorkSheet[MANAGEMENT INFO]作成、追加
                if (colorchartResults.Count > 0)
                {
                    CreateMngWorkSheet(root, domDoc, colorchartResults[0].changed_part, colorchartResults[0].part_nm);
                }
                else
                {
                    CreateMngWorkSheet(root, domDoc, "", "");
                }

                // WorkSheet[BASIC SPEC]作成、追加
                CreateBasicWorkSheet(root, domDoc);

                // TC情報取得
                TcColorChartGetRequest tcColorChartGetRequest = new TcColorChartGetRequest();
                tcColorChartGetRequest.color_chart_base_id = request.color_chart_base_id;
                tcColorChartGetRequest.color_chart_base_exp_id = request.color_chart_base_exp_id;
                tcColorChartGetRequest.color_chart_tc_id = request.color_chart_tc_id;

                TcChartColorController tcChartColorController = new TcChartColorController(null);

                TcColorChartGetResponse response = tcChartColorController.GetTcColorChartRequestItems(tcColorChartGetRequest);

                // ADD 2022/04/27 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
                List<ColorChartMdgInfo> mdgList = new List<ColorChartMdgInfo>();
                Dictionary<string, int> pmdgf_map = new Dictionary<string, int>();
                foreach (var mdg_obj in response.color_chart_mdg_list)
                {
                    if (!pmdgf_map.ContainsKey(mdg_obj.production + mdg_obj.model_code + mdg_obj.destination_cd + mdg_obj.grade + mdg_obj.fop))
                    {
                        pmdgf_map.Add(mdg_obj.production + mdg_obj.model_code + mdg_obj.destination_cd + mdg_obj.grade + mdg_obj.fop, 1);
                        ColorChartMdgInfo mdg_info = new ColorChartMdgInfo();
                        mdg_info.mdg_id = mdg_obj.mdg_id;
                        mdg_info.family_code = mdg_obj.family_code;
                        mdg_info.model_code = mdg_obj.model_code;
                        mdg_info.destination_cd = mdg_obj.destination_cd;
                        mdg_info.grade = mdg_obj.grade;
                        mdg_info.mt_code = mdg_obj.mt_code;
                        mdg_info.production = mdg_obj.production;
                        mdg_info.color_chart_tc_id = mdg_obj.color_chart_tc_id;
                        mdg_info.color_chart_tc_mdgid = mdg_obj.color_chart_tc_mdgid;
                        mdg_info.fop = mdg_obj.fop;
                        mdg_info.copy_tc_mdgid = mdg_obj.copy_tc_mdgid;

                        mdgList.Add(mdg_info);
                    }
                }
                response.color_chart_mdg_list = mdgList;
                // ADD 2022/04/27 END

                // [TC]作成
                CreateWorksheetTc(root, domDoc, response);

                // [Int Color Type]作成
                CreateWorksheetIntColorTypeList(root, domDoc, request.color_chart_tc_id);

                // [BODY COLOR LIST]作成
                CreateWorksheetBodyColorList(root, domDoc, response);

                // [Xtone]作成
                CreateWorksheetXtoneColorList(root, domDoc, request.color_chart_tc_id);

                domDoc.AppendChild(root);
                domDoc.Save(xmlPath);
                //domDoc.WriteTo(writer);
                //writer.Flush();
                //writer.Close();

                // ADD 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
                // 余計な「amp;」除く
                ReplaceCharacter(xmlPath);
                // ADD 2022/04/06 END

                LogUtil.WriteLog("★T/CのXML出力終了");
            }
            catch (Exception ex)
            {
                LogUtil.WriteLog("T/CのXML出力エラー:" + ex.Message);
                LogUtil.WriteLog("T/CのXML出力エラー:" + ex.StackTrace);

                exportDrawingDataResponse.status = "400";
                exportDrawingDataResponse.error = "Fail to output XML file.";
            }

            return exportDrawingDataResponse;
        }

        // [DocumentProperties]ノート作成
        [NonAction]
        public void CreateDocumentProperties(XmlElement root, XmlDocument domDoc, List<ColorChartGetResponse> colorchartResults)
        {

            // [DocumentProperties]ノート作成、追加
            XmlElement docPro = domDoc.CreateElement("DocumentProperties");
            docPro.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:office");

            XmlNode authorNode = domDoc.CreateNode("element", "Author", "");
            if (colorchartResults.Count > 0)
            {
                authorNode.InnerText = colorchartResults[0].modified_by;
            }
            XmlNode lastAuthorNode = domDoc.CreateNode("element", "LastAuthor", "");
            if (colorchartResults.Count > 0)
            {
                lastAuthorNode.InnerText = colorchartResults[0].modified_by;
            }

            XmlNode createNode = domDoc.CreateNode("element", "Created", "");

            DateTime dt = DateTime.UtcNow;
            string format = "yyyy-MM-ddTHH:mm:ssZ";
            string utcDate = dt.ToString(format);
            createNode.InnerText = utcDate;

            XmlNode verNode = domDoc.CreateNode("element", "Version", "");
            verNode.InnerText = "16.00";

            XmlNode typeNode = domDoc.CreateNode("element", "Type", "");
            typeNode.InnerText = "VPMReference";

            XmlNode nameNode = domDoc.CreateNode("element", "Name", "");
            if (colorchartResults.Count > 0)
            {
                nameNode.InnerText = colorchartResults[0].color_dpm_name;
            }

            XmlNode revisionNode = domDoc.CreateNode("element", "Revision", "");
            if (colorchartResults.Count > 0)
            {
                revisionNode.InnerText = colorchartResults[0].color_dpm_revsion;
            }

            docPro.AppendChild(authorNode);
            docPro.AppendChild(lastAuthorNode);
            docPro.AppendChild(createNode);
            docPro.AppendChild(verNode);
            docPro.AppendChild(typeNode);
            docPro.AppendChild(nameNode);
            docPro.AppendChild(revisionNode);

            root.AppendChild(docPro);
        }

        // Worksheet[MANAGEMENT INFO]作成
        [NonAction]
        public void CreateMngWorkSheet(XmlElement root, XmlDocument domDoc, string drawingNo, string partNm)
        {

            // Worksheetノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "MANAGEMENT INFO");

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, "2");
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, "13");
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            XmlElement colNode = domDoc.CreateElement("Column");
            colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
            colNode.SetAttribute("header", "true");

            tableNode.AppendChild(colNode);

            // Columnノード
            XmlElement colNode2 = domDoc.CreateElement("Column");
            colNode2.SetAttribute("AutoFitWidth", ns_ss, "1");

            tableNode.AppendChild(colNode2);

            for (var i = 0; i < 13; i++)
            {

                // Rowノード
                XmlElement rowNode = domDoc.CreateElement("Row");

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");

                // Dataノード
                XmlElement dateNode = domDoc.CreateElement("Data");
                dateNode.SetAttribute("Type", ns_ss, "String");
                dateNode.InnerText = MNG_ROW_NM[i];

                cellNode.AppendChild(dateNode);


                // Cell(2行目)ノード
                XmlElement cell2Node = domDoc.CreateElement("Cell");

                // Data2(2行目)ノード
                XmlElement date2Node = domDoc.CreateElement("Data");
                date2Node.SetAttribute("Type", ns_ss, "String");
                date2Node.InnerText = "";
                if (i == 0)
                {
                    date2Node.InnerText = drawingNo;
                }

                if (i == 1)
                {
                    date2Node.InnerText = partNm;
                }

                cell2Node.AppendChild(date2Node);

                rowNode.AppendChild(cellNode);
                rowNode.AppendChild(cell2Node);
                tableNode.AppendChild(rowNode);
            }
            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "mng");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // Worksheet[BASIC SPEC]作成
        [NonAction]
        public void CreateBasicWorkSheet(XmlElement root, XmlDocument domDoc)
        {

            // Worksheetノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");
            tabNode.SetAttribute("Name", ns_ss, "BASIC SPEC");

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, "2");
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, "19");
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            XmlElement colNode = domDoc.CreateElement("Column");
            colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
            colNode.SetAttribute("header", "true");

            tableNode.AppendChild(colNode);

            // Columnノード
            XmlElement colNode2 = domDoc.CreateElement("Column");
            colNode2.SetAttribute("AutoFitWidth", ns_ss, "1");

            tableNode.AppendChild(colNode2);

            for (var i = 0; i < 18; i++)
            {

                // Rowノード
                XmlElement rowNode = domDoc.CreateElement("Row");

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");

                // Dataノード
                XmlElement dateNode = domDoc.CreateElement("Data");
                dateNode.SetAttribute("Type", ns_ss, "String");
                dateNode.InnerText = BASIC_ROW_NM[i];

                cellNode.AppendChild(dateNode);

                // Cell(2行目)ノード
                XmlElement cell2Node = domDoc.CreateElement("Cell");

                // Data2(2行目)ノード
                XmlElement date2Node = domDoc.CreateElement("Data");
                date2Node.SetAttribute("Type", ns_ss, "String");
                date2Node.InnerText = "";

                cell2Node.AppendChild(date2Node);

                rowNode.AppendChild(cellNode);
                rowNode.AppendChild(cell2Node);
                tableNode.AppendChild(rowNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "basic");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }


        // Worksheet[GROUP]作成
        [NonAction]
        public int CreatGroupWorkSheet(XmlElement root, XmlDocument domDoc, long color_chart_id, Dictionary<string, GroupIntColorTypeResponse> sumTableHeader, List<SummaryItemsResponse> sumTableData, long tc_id)
        {
            //GROUP 情報検索
            GroupChartGetRequest groupchartRequest = new GroupChartGetRequest();
            groupchartRequest.color_chart_id = color_chart_id;

            // Worksheet情報検索
            IDictionary<string, object> groupNameMap = GetWorksheetInt(groupchartRequest);
            var groupResults = new List<GroupPartGetResponse>();
            groupResults = (List<GroupPartGetResponse>)groupNameMap["resultList"];

            bool groupExitFlag = false;
            for (int i = 0; i < groupResults.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(groupResults[i].part_group))
                {
                    groupExitFlag = true;
                    break;
                }
            }
            // XML出力せず
            if (!groupExitFlag)
            {
                return 0;
            }


            for (var i = 0; i < groupResults.Count; i++)
            {

                // [Worksheet]ノード作成
                if (!string.IsNullOrWhiteSpace(groupResults[i].part_group))
                {
                    XmlElement tabNode = domDoc.CreateElement("Worksheet");
                    tabNode.SetAttribute("Name", ns_ss, "GROUP " + groupResults[i].part_group);
                    // ExpandedColumnCount:INT Color Type数情報検索          
                    IDictionary<string, object> int_color_Map = GetIntColorTypeItems(groupchartRequest, groupResults[i].part_group, tc_id);
                    var intcolorResults = new List<GroupIntColorTypeResponse>();
                    intcolorResults = (List<GroupIntColorTypeResponse>)int_color_Map["resultList"];

                    //INT Color Type数
                    long int_color_type_count = intcolorResults.Count;

                    //該当グループの部品情報検索
                    IDictionary<string, object> part_items_Map = GetGroupPartItems(groupchartRequest, groupResults[i].part_group);
                    var partItemsResults = new List<GroupPartItemsResponse>();
                    partItemsResults = (List<GroupPartItemsResponse>)part_items_Map["resultList"];
                    // ADD 2022/05/30 BEGIN COLORシートルール
                    if (partItemsResults.Count == 0)
                    {
                        continue;
                    }
                    // ADD 2022/05/30 END

                    //該当グループの部品数
                    long part_count = partItemsResults.Count;

                    //Tableノード
                    XmlElement tableNode = domDoc.CreateElement("Table");
                    tableNode.SetAttribute("ExpandedColumnCount", ns_ss, (int_color_type_count + 16).ToString());
                    tableNode.SetAttribute("ExpandedRowCount", ns_ss, (part_count + 1).ToString());
                    tableNode.SetAttribute("FullColumns", ns_x, "1");
                    tableNode.SetAttribute("FullRows", ns_x, "1");
                    tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

                    // Columnノード
                    for (var j = 0; j < int_color_type_count + 16; j++)
                    {
                        XmlElement colNode = domDoc.CreateElement("Column");
                        colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                        tableNode.AppendChild(colNode);
                    }

                    // Rowノード（ヘッダ）
                    XmlElement rowNode = domDoc.CreateElement("Row");
                    rowNode.SetAttribute("Height", ns_ss, "32.4");
                    rowNode.SetAttribute("header", "true");

                    for (var k = 0; k < 14; k++)
                    {
                        // Cellノード
                        XmlElement cellNode = domDoc.CreateElement("Cell");
                        cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                        cellNode.SetAttribute("Index", ns_ss, (k + 1).ToString());

                        // Dataノード
                        XmlElement dataNode = domDoc.CreateElement("Data");
                        dataNode.SetAttribute("Type", ns_ss, "String");
                        dataNode.InnerText = GROUP_ROW_NM[k];

                        cellNode.AppendChild(dataNode);
                        rowNode.AppendChild(cellNode);
                    }
                    tableNode.AppendChild(rowNode);

                    // Cell設定（ヘッダ）
                    for (var m = 14; m < int_color_type_count + 14; m++)
                    {
                        // Cellノード
                        XmlElement cellNode = domDoc.CreateElement("Cell");
                        cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");
                        cellNode.SetAttribute("Index", ns_ss, (m + 1).ToString());

                        // Dataノード
                        XmlElement dataNode = domDoc.CreateElement("Data");
                        dataNode.SetAttribute("Type", ns_ss, "String");
                        dataNode.InnerText = intcolorResults[m - 14].SYMBOL + "&#10;" + intcolorResults[m - 14].INT_COLOR_TYPE;

                        //sum 表の頭 データセット
                        sumTableHeader.Add("sheet_" + i + "_" + (m - 14), intcolorResults[m - 14]);

                        cellNode.AppendChild(dataNode);
                        rowNode.AppendChild(cellNode);
                    }
                    tableNode.AppendChild(rowNode);
                    //2つのヘッダの設定が最適である
                    for (var n = int_color_type_count + 14; n < int_color_type_count + 16; n++)
                    {
                        // Cellノード
                        XmlElement cellNode = domDoc.CreateElement("Cell");
                        cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                        cellNode.SetAttribute("Index", ns_ss, (n + 1).ToString());

                        // Dataノード
                        XmlElement dataNode = domDoc.CreateElement("Data");
                        dataNode.SetAttribute("Type", ns_ss, "String");

                        if (n == int_color_type_count + 14)
                        {
                            dataNode.InnerText = "L1_BASE PART NO";
                        }
                        else if (n == int_color_type_count + 15)
                        {
                            dataNode.InnerText = "BASE PART NO";
                        }

                        cellNode.AppendChild(dataNode);
                        rowNode.AppendChild(cellNode);
                    }
                    tableNode.AppendChild(rowNode);


                    //ROW部品設定
                    for (var j = 0; j < partItemsResults.Count; j++)
                    {
                        XmlElement rowNodeData = domDoc.CreateElement("Row");
                        // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                        List<string> highLightList = getHighLight(partItemsResults[j].CROSS_HIGHLIGHTS);
                        foreach (var highLightInfo in highLightList)
                        {
                            XmlElement HighlightData = domDoc.CreateElement("Highlight");
                            HighlightData.SetAttribute("id", highLightInfo);
                            rowNodeData.AppendChild(HighlightData);
                        }
                        //XmlElement HighlightData = domDoc.CreateElement("Highlight");
                        //HighlightData.SetAttribute("id", partItemsResults[j].CROSS_HIGHLIGHTS);
                        //rowNodeData.AppendChild(HighlightData);
                        // MODIFY 2022/05/27 END

                        // Cell1ノード(SECTION)
                        XmlElement cellNode1 = domDoc.CreateElement("Cell");
                        cellNode1.SetAttribute("Index", ns_ss, (1).ToString());
                        // Data1ノード
                        XmlElement dataNode1 = domDoc.CreateElement("Data");
                        dataNode1.SetAttribute("Type", ns_ss, "String");
                        dataNode1.InnerText = partItemsResults[j].SECTION;
                        cellNode1.AppendChild(dataNode1);

                        // Cell2ノード(SUB-SECTION)
                        XmlElement cellNode2 = domDoc.CreateElement("Cell");
                        cellNode2.SetAttribute("Index", ns_ss, (2).ToString());
                        // Data2ノード
                        XmlElement dataNode2 = domDoc.CreateElement("Data");
                        dataNode2.SetAttribute("Type", ns_ss, "String");

                        // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                        // dataNode2.InnerText = partItemsResults[j].SUB_SECTION;
                        // Subsection6,7桁目が空白の際は空白除去
                        dataNode2.InnerText = partItemsResults[j].SUB_SECTION.TrimEnd();
                        // MODIFY 2022/04/08 END
                        cellNode2.AppendChild(dataNode2);

                        // Cell3ノード(L1)
                        XmlElement cellNode3 = domDoc.CreateElement("Cell");
                        cellNode3.SetAttribute("Index", ns_ss, (3).ToString());
                        // Data3ノード
                        XmlElement dataNode3 = domDoc.CreateElement("Data");
                        dataNode3.SetAttribute("Type", ns_ss, "String");
                        if (partItemsResults[j].UNMATCH_FLG == 1 || partItemsResults[j].L1_UNMATCH_FLG == 1)
                        {
                            // L1部番がアンマッチの場合
                            dataNode3.InnerText = ChangeUnmatchNo(partItemsResults[j].L1_CHANGED_PART);
                        }
                        else
                        {
                            dataNode3.InnerText = ChangePartNo(partItemsResults[j].L1_CHANGED_PART);
                        }

                        cellNode3.AppendChild(dataNode3);

                        // Cell4ノード(LVL)
                        XmlElement cellNode4 = domDoc.CreateElement("Cell");
                        cellNode4.SetAttribute("Index", ns_ss, (4).ToString());
                        // Data4ノード
                        XmlElement dataNode4 = domDoc.CreateElement("Data");
                        dataNode4.SetAttribute("Type", ns_ss, "String");
                        if (string.IsNullOrEmpty(partItemsResults[j].LVL.ToString()))
                        {
                            dataNode4.InnerText = "";
                        }
                        else
                        {
                            dataNode4.InnerText = partItemsResults[j].LVL.ToString();
                        }
                        cellNode4.AppendChild(dataNode4);

                        // Cell5ノード(PART NO.)
                        XmlElement cellNode5 = domDoc.CreateElement("Cell");
                        cellNode5.SetAttribute("Index", ns_ss, (5).ToString());
                        // Data5ノード
                        XmlElement dataNode5 = domDoc.CreateElement("Data");
                        dataNode5.SetAttribute("Type", ns_ss, "String");
                        if (partItemsResults[j].UNMATCH_FLG == 1)
                        {
                            dataNode5.InnerText = ChangeUnmatchNo(partItemsResults[j].CHANGED_PART);
                        }
                        else
                        {
                            dataNode5.InnerText = ChangePartNo(partItemsResults[j].CHANGED_PART);
                        }

                        cellNode5.AppendChild(dataNode5);

                        // Cell6ノード(PARENT PART NO.)
                        XmlElement cellNode6 = domDoc.CreateElement("Cell");
                        cellNode6.SetAttribute("Index", ns_ss, (6).ToString());
                        // Data6ノード
                        XmlElement dataNode6 = domDoc.CreateElement("Data");
                        dataNode6.SetAttribute("Type", ns_ss, "String");
                        if (partItemsResults[j].UNMATCH_FLG == 1 || partItemsResults[j].PARENT_UNMATCH_FLG == 1)
                        {
                            dataNode6.InnerText = ChangeUnmatchNo(setUnderline(partItemsResults[j].PARENT_CHANGED_PART));
                        }
                        else
                        {
                            dataNode6.InnerText = ChangePartNo(setUnderline(partItemsResults[j].PARENT_CHANGED_PART));
                        }

                        cellNode6.AppendChild(dataNode6);

                        // Cell7ノード(PART NAME)
                        XmlElement cellNode7 = domDoc.CreateElement("Cell");
                        cellNode7.SetAttribute("Index", ns_ss, (7).ToString());
                        // Data7ノード
                        XmlElement dataNode7 = domDoc.CreateElement("Data");
                        dataNode7.SetAttribute("Type", ns_ss, "String");
                        dataNode7.InnerText = partItemsResults[j].PART_NM;
                        cellNode7.AppendChild(dataNode7);

                        // Cell8ノード(PLANT)
                        XmlElement cellNode8 = domDoc.CreateElement("Cell");
                        cellNode8.SetAttribute("Index", ns_ss, (8).ToString());
                        // Data8ノード
                        XmlElement dataNode8 = domDoc.CreateElement("Data");
                        dataNode8.SetAttribute("Type", ns_ss, "String");
                        dataNode8.InnerText = partItemsResults[j].PRODUCTIO;
                        cellNode8.AppendChild(dataNode8);

                        // Cell9ノード(MODEL)
                        XmlElement cellNode9 = domDoc.CreateElement("Cell");
                        cellNode9.SetAttribute("Index", ns_ss, (9).ToString());
                        // Data9ノード
                        XmlElement dataNode9 = domDoc.CreateElement("Data");
                        dataNode9.SetAttribute("Type", ns_ss, "String");
                        dataNode9.InnerText = partItemsResults[j].MODEL_CODE;
                        cellNode9.AppendChild(dataNode9);

                        // Cell10ノード(DESTINATION)
                        XmlElement cellNode10 = domDoc.CreateElement("Cell");
                        cellNode10.SetAttribute("Index", ns_ss, (10).ToString());
                        // Data10ノード
                        XmlElement dataNode10 = domDoc.CreateElement("Data");
                        dataNode10.SetAttribute("Type", ns_ss, "String");
                        dataNode10.InnerText = partItemsResults[j].DESTINATION;
                        cellNode10.AppendChild(dataNode10);

                        // Cell11ノード(GRADE)
                        XmlElement cellNode11 = domDoc.CreateElement("Cell");
                        cellNode11.SetAttribute("Index", ns_ss, (11).ToString());
                        // Data11ノード
                        XmlElement dataNode11 = domDoc.CreateElement("Data");
                        dataNode11.SetAttribute("Type", ns_ss, "String");
                        dataNode11.InnerText = partItemsResults[j].GRADE;
                        cellNode11.AppendChild(dataNode11);

                        // Cell12ノード(FEATURE)
                        XmlElement cellNode12 = domDoc.CreateElement("Cell");
                        cellNode12.SetAttribute("Index", ns_ss, (12).ToString());
                        // Data12ノード
                        XmlElement dataNode12 = domDoc.CreateElement("Data");
                        dataNode12.SetAttribute("Type", ns_ss, "String");
                        dataNode12.InnerText = partItemsResults[j].EQUIPMENT;
                        cellNode12.AppendChild(dataNode12);

                        // Cell13ノード(ITEM)
                        XmlElement cellNode13 = domDoc.CreateElement("Cell");
                        cellNode13.SetAttribute("Index", ns_ss, (13).ToString());
                        // Data13ノード
                        XmlElement dataNode13 = domDoc.CreateElement("Data");
                        dataNode13.SetAttribute("Type", ns_ss, "String");
                        dataNode13.InnerText = partItemsResults[j].ITEM;
                        cellNode13.AppendChild(dataNode13);

                        // Cell14ノード(REMARKS)
                        XmlElement cellNode14 = domDoc.CreateElement("Cell");
                        cellNode14.SetAttribute("Index", ns_ss, (14).ToString());
                        // Data14ノード
                        XmlElement dataNode14 = domDoc.CreateElement("Data");
                        dataNode14.SetAttribute("Type", ns_ss, "String");
                        dataNode14.InnerText = partItemsResults[j].REMARK;
                        cellNode14.AppendChild(dataNode14);

                        rowNodeData.AppendChild(cellNode1);
                        rowNodeData.AppendChild(cellNode2);
                        rowNodeData.AppendChild(cellNode3);
                        rowNodeData.AppendChild(cellNode4);
                        rowNodeData.AppendChild(cellNode5);
                        rowNodeData.AppendChild(cellNode6);
                        rowNodeData.AppendChild(cellNode7);
                        rowNodeData.AppendChild(cellNode8);
                        rowNodeData.AppendChild(cellNode9);
                        rowNodeData.AppendChild(cellNode10);
                        rowNodeData.AppendChild(cellNode11);
                        rowNodeData.AppendChild(cellNode12);
                        rowNodeData.AppendChild(cellNode13);
                        rowNodeData.AppendChild(cellNode14);

                        //SUMMARY データの記入
                        SummaryItemsResponse summaryItemsResponse = new SummaryItemsResponse();
                        summaryItemsResponse.part_group = groupResults[i].part_group;
                        summaryItemsResponse.SECTION = partItemsResults[j].SECTION;
                        summaryItemsResponse.SUB_SECTION = partItemsResults[j].SUB_SECTION;
                        summaryItemsResponse.CHANGED_PART = partItemsResults[j].CHANGED_PART;
                        summaryItemsResponse.PART_NM = partItemsResults[j].PART_NM;
                        summaryItemsResponse.PRODUCTIO = partItemsResults[j].PRODUCTIO;
                        summaryItemsResponse.MODEL_CODE = partItemsResults[j].MODEL_CODE;
                        summaryItemsResponse.DESTINATION = partItemsResults[j].DESTINATION;
                        summaryItemsResponse.GRADE = partItemsResults[j].GRADE;
                        summaryItemsResponse.EQUIPMENT = partItemsResults[j].EQUIPMENT;
                        summaryItemsResponse.ITEM = partItemsResults[j].ITEM;
                        summaryItemsResponse.REMARK = partItemsResults[j].REMARK;
                        summaryItemsResponse.CHKMEMO = partItemsResults[j].CHKMEMO;
                        List<string> highLight = new List<string>();
                        highLight.Add(partItemsResults[j].CROSS_HIGHLIGHTS);
                        summaryItemsResponse.CROSS_HIGHLIGHTS = highLight;
                        summaryItemsResponse.ColorHesPartItems = new List<ColorHesItemsResponse>();
                        // ADD 2022/05/12 BEGIN XML出力サマリ
                        summaryItemsResponse.RL_PAIR_NO = partItemsResults[j].RL_PAIR_NO;
                        summaryItemsResponse.RL_FLAG = partItemsResults[j].RL_FLAG;
                        summaryItemsResponse.SUMMARY_NO = partItemsResults[j].SUMMARY_NO;
                        summaryItemsResponse.SUMMARY_GROUP = partItemsResults[j].SUMMARY_GROUP;
                        summaryItemsResponse.SUMMARY_REMARK = partItemsResults[j].SUMMARY_REMARK;
                        // ADD 2022/05/12 END
                        summaryItemsResponse.CHILD_UNMATCH_FLAG = partItemsResults[j].CHILD_UNMATCH_FLAG;

                        // Data部
                        //INT_COLOR_ID(列) -> COLOR_HES
                        IDictionary<string, object> ICI_CH_Map = GetColorHesPartItems(partItemsResults[j].PART_APPLY_RELATION_ID.ToString());
                        var ICI_CH_Results = new List<ColorHesItemsResponse>();
                        ICI_CH_Results = (List<ColorHesItemsResponse>)ICI_CH_Map["resultList"];
                        for (var k = 15; k < int_color_type_count + 15; k++)
                        {
                            bool flag = false;
                            for (var m = 0; m < ICI_CH_Results.Count; m++)
                            {
                                //表のIDとデータのIDを同じセルに設定
                                if (intcolorResults[k - 15].INT_COLOR_ID == ICI_CH_Results[m].INT_COLOR_ID)
                                {

                                    XmlElement cellNode = domDoc.CreateElement("Cell");
                                    cellNode.SetAttribute("Index", ns_ss, (k).ToString());

                                    XmlElement dataNode = domDoc.CreateElement("Data");
                                    dataNode.SetAttribute("Type", ns_ss, "String");


                                    if (string.IsNullOrEmpty(ICI_CH_Results[m].COLOR_HES))
                                    {
                                        // 部品リレーションの子部品不一致フラグが1の場合、「-」に設定
                                        if (partItemsResults[j].CHILD_UNMATCH_FLAG == 1)
                                        {
                                            dataNode.InnerText = "-";
                                        }
                                        else
                                        {
                                            dataNode.InnerText = "";
                                        }
                                    }
                                    else
                                    {
                                        dataNode.InnerText = ICI_CH_Results[m].COLOR_HES;
                                    }

                                    cellNode.AppendChild(dataNode);
                                    rowNodeData.AppendChild(cellNode);
                                    flag = true;

                                    //SUMMARY データの記入
                                    summaryItemsResponse.ColorHesPartItems.Add(ICI_CH_Results[m]);

                                    break;
                                }
                            }
                            //表の現在の列には、本条のデータに対応するidが表示されていません
                            if (!flag)
                            {
                                XmlElement cellNode = domDoc.CreateElement("Cell");
                                cellNode.SetAttribute("Index", ns_ss, (k).ToString());

                                XmlElement dataNode = domDoc.CreateElement("Data");
                                dataNode.SetAttribute("Type", ns_ss, "String");
                                if (partItemsResults[j].CHILD_UNMATCH_FLAG == 1)
                                {
                                    dataNode.InnerText = "-";
                                }
                                else
                                {
                                    dataNode.InnerText = "";
                                }
                                cellNode.AppendChild(dataNode);
                                rowNodeData.AppendChild(cellNode);
                            }

                        }

                        //SUMMARY データの集合追加
                        sumTableData.Add(summaryItemsResponse);

                        // L1_BASE PART NO ノード
                        XmlElement cellNodeL1No = domDoc.CreateElement("Cell");
                        cellNodeL1No.SetAttribute("Index", ns_ss, (int_color_type_count + 15).ToString());
                        // L1_BASE PART NO ノード
                        XmlElement dataNodeL1No = domDoc.CreateElement("Data");
                        dataNodeL1No.SetAttribute("Type", ns_ss, "String");

                        if (partItemsResults[j].UNMATCH_FLG == 1 || partItemsResults[j].L1_UNMATCH_FLG == 1)
                        {
                            dataNodeL1No.InnerText = ChangeUnmatchNo(partItemsResults[j].L1_BASE_PART);
                        }
                        else
                        {
                            dataNodeL1No.InnerText = ChangePartNo(partItemsResults[j].L1_BASE_PART);
                        }

                        cellNodeL1No.AppendChild(dataNodeL1No);
                        rowNodeData.AppendChild(cellNodeL1No);
                        // cell n+16ノード(BASE PART NO)
                        XmlElement cellBasePartNo = domDoc.CreateElement("Cell");
                        cellBasePartNo.SetAttribute("Index", ns_ss, (int_color_type_count + 16).ToString());
                        // Data n+16ノード
                        XmlElement dataBasePartNo = domDoc.CreateElement("Data");
                        dataBasePartNo.SetAttribute("Type", ns_ss, "String");

                        if (partItemsResults[j].UNMATCH_FLG == 1)
                        {
                            dataBasePartNo.InnerText = ChangeUnmatchNo(partItemsResults[j].BASE_PART);
                        }
                        else
                        {
                            dataBasePartNo.InnerText = ChangePartNo(partItemsResults[j].BASE_PART);
                        }

                        cellBasePartNo.AppendChild(dataBasePartNo);
                        rowNodeData.AppendChild(cellBasePartNo);

                        tableNode.AppendChild(rowNodeData);

                    }

                    tabNode.AppendChild(tableNode);
                    // WorksheetOptionsノード
                    XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "color");
                    tabNode.AppendChild(worksheetOptions);

                    root.AppendChild(tabNode);
                }
                else
                {
                    continue;
                }
            }
            return 1;
        }

        // Worksheet[COLOR]作成
        [NonAction]
        public void CreateColorWorkSheet(XmlElement root, XmlDocument domDoc, List<ColorTabInfo> colorTabInfoList)
        {

            // [COLOR]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");
            tabNode.SetAttribute("Name", ns_ss, "COLOR");

            long colCount = 17;

            long rowCount = colorTabInfoList.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            for (var i = 0; i < colCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = COLOR_ROW_NM[i];

                cellNode.AppendChild(dataNode);
                rowNode.AppendChild(cellNode);

            }

            tableNode.AppendChild(rowNode);

            for (var i = 0; i < colorTabInfoList.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // Highlightノード
                // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                List<string> highLightList = getHighLight(colorTabInfoList[i].cross_highlights);
                foreach (var highLightInfo in highLightList)
                {
                    XmlElement highlightNode = domDoc.CreateElement("Highlight");
                    highlightNode.SetAttribute("id", highLightInfo);
                    rowDataNode.AppendChild(highlightNode);
                }
                //XmlElement highlightNode = domDoc.CreateElement("Highlight");
                //highlightNode.SetAttribute("id", colorTabInfoList[i].cross_highlights);
                //rowDataNode.AppendChild(highlightNode);
                // MODIFY 2022/05/27 END
                // Cellノード(SECTION)
                setCell(domDoc, colorTabInfoList[i].section, "1", rowDataNode);

                // Cellノード(SUB-SECTION)
                // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                // setCell(domDoc, colorTabInfoList[i].sub_section, "2", rowDataNode);
                setCell(domDoc, colorTabInfoList[i].sub_section.TrimEnd(), "2", rowDataNode);
                // MODIFY 2022/04/08 END

                // Cellノード(L1)
                if (colorTabInfoList[i].unmatch_flg == 1 || colorTabInfoList[i].l1_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(colorTabInfoList[i].l1_part_no), "3", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(colorTabInfoList[i].l1_part_no), "3", rowDataNode);
                }

                // Cellノード(LVL)
                setCell(domDoc, colorTabInfoList[i].lvl, "4", rowDataNode);

                // Cellノード(PART NO.)
                if (colorTabInfoList[i].unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(colorTabInfoList[i].changed_part), "5", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(colorTabInfoList[i].changed_part), "5", rowDataNode);
                }

                // Cellノード(PARENT PART NO.)
                if (colorTabInfoList[i].unmatch_flg == 1 || colorTabInfoList[i].parent_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(setUnderline(colorTabInfoList[i].parent_part_no)), "6", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(setUnderline(colorTabInfoList[i].parent_part_no)), "6", rowDataNode);
                }

                // Cellノード(PART NAME)
                setCell(domDoc, colorTabInfoList[i].part_name, "7", rowDataNode);

                // Cellノード(PLANT)
                setCell(domDoc, colorTabInfoList[i].plant, "8", rowDataNode);

                // Cellノード(MODEL)
                setCell(domDoc, colorTabInfoList[i].model, "9", rowDataNode);

                // Cellノード(DESTINATION)
                setCell(domDoc, colorTabInfoList[i].destination, "10", rowDataNode);

                // Cellノード(GRADE)
                setCell(domDoc, colorTabInfoList[i].grade, "11", rowDataNode);

                // Cellノード(FEATURE)
                setCell(domDoc, colorTabInfoList[i].feature, "12", rowDataNode);

                // Cellノード(ITEM)
                setCell(domDoc, colorTabInfoList[i].item, "13", rowDataNode);

                // Cellノード(REMARKS)
                setCell(domDoc, colorTabInfoList[i].remark, "14", rowDataNode);

                // Cellノード(COLOR)
                bool bodyColorFlg = true;

                bool isEmptyFlg = true;

                foreach (var colorItem in colorTabInfoList[i].color_item)
                {
                    if (!string.IsNullOrEmpty(colorItem))
                    {
                        isEmptyFlg = false;
                        break;
                    }
                }

                if (!isEmptyFlg)
                {
                    for (var j = 0; j < colorTabInfoList[i].color_header.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(colorTabInfoList[i].color_item[j]))
                        {
                            if (!colorTabInfoList[i].color_header[j].Equals(colorTabInfoList[i].color_item[j]))
                            {
                                bodyColorFlg = false;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    bodyColorFlg = false;
                }

                if (colorTabInfoList[i].child_unmatch_flag == 1)
                {
                    setCell(domDoc, "-", "15", rowDataNode);
                    // DELETE 2022/06/21 BEGIN XMLサマリー修正
                    //colorTabInfoList[i].body_color = "-";
                    // DELETE 2022/06/21 END
                }
                else
                {
                    if (bodyColorFlg)
                    {
                        setCell(domDoc, "BODY COLOR", "15", rowDataNode);
                        colorTabInfoList[i].body_color = "BODY COLOR";
                    }
                    else
                    {

                        bool itemFlg = true;

                        string hes_color = "";

                        string tmpHes = "";

                        // 色情報が複数件存在するの場合、一致するか比較する
                        if (colorTabInfoList[i].color_item.Count > 1)
                        {

                            // 一番目の空白以外のデータ取得
                            for (var j = 0; j < colorTabInfoList[i].color_item.Count; j++)
                            {

                                if (!string.IsNullOrWhiteSpace(colorTabInfoList[i].color_item[j]))
                                {
                                    tmpHes = colorTabInfoList[i].color_item[j];
                                    break;
                                }

                            }

                            // 色情報は1件以上存在
                            if (!string.IsNullOrWhiteSpace(tmpHes))
                            {

                                for (var j = 0; j < colorTabInfoList[i].color_item.Count; j++)
                                {

                                    if (!string.IsNullOrWhiteSpace(colorTabInfoList[i].color_item[j]))
                                    {
                                        // 色情報差異ある場合
                                        if (!colorTabInfoList[i].color_item[j].Equals(tmpHes))
                                        {
                                            itemFlg = false;
                                            break;
                                        }
                                    }

                                }
                            }

                            hes_color = tmpHes;
                        }
                        else if (colorTabInfoList[i].color_item.Count == 1)
                        {
                            hes_color = colorTabInfoList[i].color_item[0];
                        }

                        // 色情報差異ない場合
                        if (itemFlg)
                        {
                            // 該当色情報設定
                            if (!string.IsNullOrWhiteSpace(hes_color))
                            {
                                setCell(domDoc, hes_color, "15", rowDataNode);
                            }
                            else
                            {
                                continue;
                                //setCell(domDoc, "", "15", rowDataNode);
                            }

                            colorTabInfoList[i].body_color = hes_color;
                        }
                        else
                        {
                            // 色情報差異ある場合、「-」で設定
                            setCell(domDoc, "-", "15", rowDataNode);
                            colorTabInfoList[i].body_color = "-";
                        }
                    }

                }

                // Cellノード(L1_BASE PART NO)
                if (colorTabInfoList[i].unmatch_flg == 1 || colorTabInfoList[i].l1_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(colorTabInfoList[i].l1_base_part_no), "16", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(colorTabInfoList[i].l1_base_part_no), "16", rowDataNode);
                }

                // Cellノード(BASE PART NO)
                if (colorTabInfoList[i].unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(colorTabInfoList[i].base_part_no), "17", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(colorTabInfoList[i].base_part_no), "17", rowDataNode);
                }

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "color");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // Cell設定
        [NonAction]
        public void setCell(XmlDocument domDoc, string val, string idx, XmlElement rowDataNode)
        {
            XmlElement cellNode = domDoc.CreateElement("Cell");
            cellNode.SetAttribute("Index", ns_ss, idx);

            // Dataノード
            XmlElement dataNode = domDoc.CreateElement("Data");
            dataNode.SetAttribute("Type", ns_ss, "String");
            dataNode.InnerText = val;

            cellNode.AppendChild(dataNode);
            rowDataNode.AppendChild(cellNode);
        }

        // Cell設定
        [NonAction]
        public void setCell2(XmlDocument domDoc, string val, string idx, XmlElement rowDataNode)
        {
            XmlElement cellNode = domDoc.CreateElement("Cell");
            cellNode.SetAttribute("Index", ns_ss, idx);

            // Dataノード
            XmlElement dataNode = domDoc.CreateElement("Data");
            dataNode.SetAttribute("Type", ns_ss, "String");
            dataNode.InnerText = val;

            cellNode.AppendChild(dataNode);
            rowDataNode.AppendChild(cellNode);
        }

        // COLOR LISTシート作成
        [NonAction]
        public void CreateColorListWorkSheet(XmlElement root, XmlDocument domDoc, long color_chart_id, string opt, List<ColorListInfoGetResponse> bodyColorToHes)
        {

            // タブノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");
            tabNode.SetAttribute("Name", ns_ss, "COLOR LIST");

            // カラーリスト情報検索
            ColorChartGetRequest colorchartRequest = new ColorChartGetRequest();
            colorchartRequest.color_chart_id = color_chart_id;
            string[] compareArray = { "NH", "R", "YR", "Y", "GY", "G", "BG", "B", "PB", "RP" };

            IDictionary<string, object> colorListMap = null;

            var colorListResults = new List<ColorListInfoGetResponse>();

            if ("Ext".Equals(opt))
            {
                colorListMap = GetExtColorListItems(colorchartRequest);
                colorListResults = (List<ColorListInfoGetResponse>)colorListMap["resultList"];

                // Body Colorで入れ替えデータを追加
                foreach (var bodyColorItem in bodyColorToHes)
                {
                    colorListResults.Add(bodyColorItem);
                }
            }
            else
            {
                colorListMap = GetIntColorListItems(colorchartRequest);
                colorListResults = (List<ColorListInfoGetResponse>)colorListMap["resultList"];
            }

            // ADD 2022/05/30 BEGIN COLORシートルール
            if (colorListResults.Count == 0)
            {
                return;
            }
            // ADD 2022/05/30 END

            // 重複データ除去
            colorListResults = colorListResults.Where((x, y) => colorListResults.FindIndex(z => z.color_hes == x.color_hes) == y).ToList();

            // HESコード以外の形式の文字列含めるデータ除去
            List<ColorListInfoGetResponse> colorListTemp = new List<ColorListInfoGetResponse>();
            foreach (var colorHes in colorListResults)
            {
                if (CheckHesCode(colorHes.color_hes))
                {
                    colorListTemp.Add(colorHes);
                }
            }

            colorListResults = colorListTemp;

            // 自動整列
            colorListResults = colorListResults.OrderBy(e =>
            {
                var index = 0;
                index = Array.IndexOf(compareArray, e.color_hes.Split('-')[0]);
                if (index != -1) { return index; }
                else
                {
                    return int.MaxValue;
                }
            }).ThenBy(e => e.color_hes.Split('-')[1]).ToList();

            int rowCount = colorListResults.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, "2");
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            XmlElement colNode = domDoc.CreateElement("Column");
            colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
            tableNode.AppendChild(colNode);

            // Columnノード
            XmlElement col2Node = domDoc.CreateElement("Column");
            col2Node.SetAttribute("AutoFitWidth", ns_ss, "1");
            tableNode.AppendChild(col2Node);

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            for (var i = 0; i < 2; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                if (i == 0)
                {

                    dataNode.InnerText = "CODE";
                }
                else
                {
                    dataNode.InnerText = "COLOR NAME";
                }

                cellNode.AppendChild(dataNode);
                rowNode.AppendChild(cellNode);

            }

            tableNode.AppendChild(rowNode);



            for (var i = 0; i < colorListResults.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // Cellノード(CODE)
                XmlElement codeCellNode = domDoc.CreateElement("Cell");
                codeCellNode.SetAttribute("Index", ns_ss, "1");

                // Dataノード
                XmlElement codeDataNode = domDoc.CreateElement("Data");
                codeDataNode.SetAttribute("Type", ns_ss, "String");
                codeDataNode.InnerText = colorListResults[i].color_hes;
                codeCellNode.AppendChild(codeDataNode);
                rowDataNode.AppendChild(codeCellNode);


                // Cellノード(COLOR NAME)
                XmlElement codeNmCellNode = domDoc.CreateElement("Cell");
                codeNmCellNode.SetAttribute("Index", ns_ss, "2");

                // Dataノード
                XmlElement codeNmDataNode = domDoc.CreateElement("Data");
                codeNmDataNode.SetAttribute("Type", ns_ss, "String");
                codeNmDataNode.InnerText = colorListResults[i].color_name;
                codeNmCellNode.AppendChild(codeNmDataNode);
                rowDataNode.AppendChild(codeNmCellNode);

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "colorlist");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // 固定値作成
        [NonAction]
        public void CreateCommonWorkSheet(XmlElement root, XmlDocument domDoc)
        {

            // OfficeDocumentSettingsノード
            XmlElement officeNode = domDoc.CreateElement("OfficeDocumentSettings");
            officeNode.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:office");

            XmlNode allowPng = domDoc.CreateNode("element", "AllowPNG", "");
            officeNode.AppendChild(allowPng);

            root.AppendChild(officeNode);

            // ExcelWorkbookノード
            XmlElement excelWorkbookNode = domDoc.CreateElement("ExcelWorkbook");
            excelWorkbookNode.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:excel");

            XmlNode windowHeight = domDoc.CreateNode("element", "WindowHeight", "");
            windowHeight.InnerText = "10596";
            excelWorkbookNode.AppendChild(windowHeight);

            XmlNode windowWidth = domDoc.CreateNode("element", "WindowWidth", "");
            windowWidth.InnerText = "23040";
            excelWorkbookNode.AppendChild(windowWidth);

            XmlNode windowTopX = domDoc.CreateNode("element", "WindowTopX", "");
            windowTopX.InnerText = "32767";
            excelWorkbookNode.AppendChild(windowTopX);

            XmlNode windowTopY = domDoc.CreateNode("element", "WindowTopY", "");
            windowTopY.InnerText = "32767";
            excelWorkbookNode.AppendChild(windowTopY);

            XmlNode activeSheet = domDoc.CreateNode("element", "ActiveSheet", "");
            activeSheet.InnerText = "14";
            excelWorkbookNode.AppendChild(activeSheet);

            XmlNode firstVisibleSheet = domDoc.CreateNode("element", "FirstVisibleSheet", "");
            firstVisibleSheet.InnerText = "10";
            excelWorkbookNode.AppendChild(firstVisibleSheet);

            XmlNode protectStructure = domDoc.CreateNode("element", "ProtectStructure", "");
            protectStructure.InnerText = "False";
            excelWorkbookNode.AppendChild(protectStructure);

            XmlNode protectWindows = domDoc.CreateNode("element", "ProtectWindows", "");
            protectWindows.InnerText = "False";
            excelWorkbookNode.AppendChild(protectWindows);

            root.AppendChild(excelWorkbookNode);

            // Stylesノード
            XmlElement stylesNode = domDoc.CreateElement("Styles");

            // Style[Default]
            XmlElement styleDefault = domDoc.CreateElement("Style");

            styleDefault.SetAttribute("ID", "urn:schemas-microsoft-com:office:spreadsheet", "Default");
            styleDefault.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", "Normal");

            XmlElement styleDefaultAlignment = domDoc.CreateElement("Alignment");
            styleDefaultAlignment.SetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", "Center");
            styleDefault.AppendChild(styleDefaultAlignment);

            XmlElement styleDefaultBorders = domDoc.CreateElement("Borders");
            styleDefault.AppendChild(styleDefaultBorders);

            XmlElement styleDefaultFont = domDoc.CreateElement("Font");
            styleDefaultFont.SetAttribute("FontName", "urn:schemas-microsoft-com:office:spreadsheet", "游ゴシック Medium");
            styleDefaultFont.SetAttribute("CharSet", ns_x, "128");
            styleDefaultFont.SetAttribute("Family", ns_x, "Modern");
            styleDefaultFont.SetAttribute("Color", ns_ss, "#000000");
            styleDefault.AppendChild(styleDefaultFont);

            XmlElement styleDefaultInterior = domDoc.CreateElement("Interior");
            styleDefault.AppendChild(styleDefaultInterior);

            XmlElement styleDefaultNumberFormat = domDoc.CreateElement("NumberFormat");
            styleDefault.AppendChild(styleDefaultNumberFormat);

            XmlElement styleDefaultProtection = domDoc.CreateElement("Protection");
            styleDefault.AppendChild(styleDefaultProtection);

            stylesNode.AppendChild(styleDefault);

            // Style[sHead]
            XmlElement styleHead = domDoc.CreateElement("Style");
            styleHead.SetAttribute("ID", ns_ss, "sHead");

            XmlElement styleHeadAlignment = domDoc.CreateElement("Alignment");
            styleHeadAlignment.SetAttribute("Horizontal", ns_ss, "Center");
            styleHeadAlignment.SetAttribute("Vertical", ns_ss, "Center");
            styleHead.AppendChild(styleHeadAlignment);

            XmlElement styleHeadInterior = domDoc.CreateElement("Interior");
            styleHeadInterior.SetAttribute("Color", ns_ss, "#F2F2F2");
            styleHeadInterior.SetAttribute("Pattern", ns_ss, "Solid");
            styleHead.AppendChild(styleHeadInterior);

            stylesNode.AppendChild(styleHead);

            // Style[sDate]
            XmlElement styleDate = domDoc.CreateElement("Style");
            styleDate.SetAttribute("ID", ns_ss, "sDate");

            XmlElement styleDateNumberFormat = domDoc.CreateElement("NumberFormat");
            styleDateNumberFormat.SetAttribute("Format", ns_ss, "Medium Date");
            styleDate.AppendChild(styleDateNumberFormat);

            stylesNode.AppendChild(styleDate);

            // Style[sMultiHead]
            XmlElement styleMultiHead = domDoc.CreateElement("Style");
            styleMultiHead.SetAttribute("ID", ns_ss, "sMultiHead");

            XmlElement styleMultiHeadAlignment = domDoc.CreateElement("Alignment");
            styleMultiHeadAlignment.SetAttribute("Horizontal", ns_ss, "Center");
            styleMultiHeadAlignment.SetAttribute("Vertical", ns_ss, "Center");
            styleMultiHeadAlignment.SetAttribute("WrapText", ns_ss, "1");
            styleMultiHead.AppendChild(styleMultiHeadAlignment);

            XmlElement styleMultiHeadInterior = domDoc.CreateElement("Interior");
            styleMultiHeadInterior.SetAttribute("Color", ns_ss, "#F2F2F2");
            styleMultiHeadInterior.SetAttribute("Pattern", ns_ss, "Solid");
            styleMultiHead.AppendChild(styleMultiHeadInterior);

            stylesNode.AppendChild(styleMultiHead);

            // Style[sWrapText]
            XmlElement styleWrapText = domDoc.CreateElement("Style");
            styleWrapText.SetAttribute("ID", ns_ss, "sWrapText");

            XmlElement styleAlignment = domDoc.CreateElement("Alignment");
            styleAlignment.SetAttribute("WrapText", ns_ss, "1");
            styleWrapText.AppendChild(styleAlignment);

            stylesNode.AppendChild(styleWrapText);

            root.AppendChild(stylesNode);

        }

        // WorksheetOptOptions設定
        private XmlElement SetWorksheetOptOptions(XmlDocument domDoc, string worksheetType)
        {

            // WorksheetOptionsノード
            XmlElement worksheetOptions = domDoc.CreateElement("WorksheetOptions");
            worksheetOptions.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:excel");

            // PageSetup
            XmlElement pageSetup = domDoc.CreateElement("PageSetup");

            XmlElement header = domDoc.CreateElement("Header");
            header.SetAttribute("Margin", ns_x, "0.3");
            pageSetup.AppendChild(header);

            XmlElement footer = domDoc.CreateElement("Footer");
            footer.SetAttribute("Margin", ns_x, "0.3");
            pageSetup.AppendChild(footer);

            XmlElement pageMargins = domDoc.CreateElement("PageMargins");
            pageMargins.SetAttribute("Bottom", ns_x, "0.75");
            pageMargins.SetAttribute("Left", ns_x, "0.7");
            pageMargins.SetAttribute("Right", ns_x, "0.7");
            pageMargins.SetAttribute("Top", ns_x, "0.75");
            pageSetup.AppendChild(pageMargins);

            worksheetOptions.AppendChild(pageSetup);

            if (worksheetType == "mng")
            {
                // Panes
                XmlElement panes = domDoc.CreateElement("Panes");

                XmlElement pane = domDoc.CreateElement("Pane");

                XmlElement numberNode = domDoc.CreateElement("Number");
                numberNode.InnerText = "3";
                pane.AppendChild(numberNode);

                XmlElement activeRow = domDoc.CreateElement("ActiveRow");
                activeRow.InnerText = "7";
                pane.AppendChild(activeRow);

                panes.AppendChild(pane);

                worksheetOptions.AppendChild(panes);
            }

            if (worksheetType == "tc")
            {
                // Panes
                XmlElement selected = domDoc.CreateElement("Selected");

                worksheetOptions.AppendChild(selected);

                // Panes
                XmlElement panes = domDoc.CreateElement("Panes");

                XmlElement pane = domDoc.CreateElement("Pane");

                XmlElement numberNode = domDoc.CreateElement("Number");
                numberNode.InnerText = "3";
                pane.AppendChild(numberNode);

                XmlElement rangeSelection = domDoc.CreateElement("RangeSelection");
                rangeSelection.InnerText = "R1C1:R2C1";
                pane.AppendChild(rangeSelection);

                panes.AppendChild(pane);

                worksheetOptions.AppendChild(panes);
            }

            // ProtectObjects
            XmlElement protectObjects = domDoc.CreateElement("ProtectObjects");
            protectObjects.InnerText = "False";
            worksheetOptions.AppendChild(protectObjects);


            XmlElement protectScenarios = domDoc.CreateElement("ProtectScenarios");
            protectScenarios.InnerText = "False";
            worksheetOptions.AppendChild(protectScenarios);

            return worksheetOptions;
        }

        // [BY BODYCOLOR]
        private XmlElement CreateWorksheetByColor(XmlElement root, XmlDocument domDoc, List<ColorTabInfo> byBodyColorList)
        {

            // [COLOR]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "BY BODYCOLOR");

            // 外装色ヘッダの数
            long headerCount = bodyColorHeaderCount;

            // 列数
            long colCount = 16 + headerCount;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, (colCount).ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, (byBodyColorList.Count + 1).ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("Height", ns_ss, "40");
            rowNode.SetAttribute("header", "true");

            // ヘッダ作成（固定部）
            for (var i = 0; i < 14; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                if (i < 14)
                {
                    cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                    cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                    // Dataノード
                    XmlElement dataNode = domDoc.CreateElement("Data");
                    dataNode.SetAttribute("Type", ns_ss, "String");
                    dataNode.InnerText = COLOR_ROW_NM[i];
                    cellNode.AppendChild(dataNode);
                }

                rowNode.AppendChild(cellNode);

            }

            // ヘッダ作成（BodyColor部）
            for (var i = 0; i < headerCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");

                cellNode.SetAttribute("Index", ns_ss, (15 + i).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");

                dataNode.InnerText = colorHeaderName[i];

                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            // ヘッダ作成(固定部)
            for (var i = 15 + headerCount; i < colCount + 1; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, i.ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = COLOR_ROW_NM[i - headerCount];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            tableNode.AppendChild(rowNode);

            // データ部作成
            for (var i = 0; i < byBodyColorList.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // Highlightノード
                // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                List<string> highLightList = getHighLight(byBodyColorList[i].cross_highlights);
                foreach (var highLightInfo in highLightList)
                {
                    XmlElement highlightNode = domDoc.CreateElement("Highlight");
                    highlightNode.SetAttribute("id", highLightInfo);
                    rowDataNode.AppendChild(highlightNode);
                }
                //XmlElement highlightNode = domDoc.CreateElement("Highlight");
                //highlightNode.SetAttribute("id", byBodyColorList[i].cross_highlights);
                //rowDataNode.AppendChild(highlightNode);
                // MODIFY 2022/05/27 END

                // Cellノード(SECTION)
                setCell(domDoc, byBodyColorList[i].section, "1", rowDataNode);

                // Cellノード(SUB-SECTION)
                // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                // setCell(domDoc, byBodyColorList[i].sub_section, "2", rowDataNode);
                setCell(domDoc, byBodyColorList[i].sub_section.TrimEnd(), "2", rowDataNode);
                // MODIFY 2022/04/08 END

                // Cellノード(L1)
                if (byBodyColorList[i].unmatch_flg == 1 || byBodyColorList[i].l1_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].l1_part_no), "3", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(byBodyColorList[i].l1_part_no), "3", rowDataNode);
                }

                // Cellノード(LVL)
                setCell(domDoc, byBodyColorList[i].lvl, "4", rowDataNode);

                // Cellノード(PART NO.)
                if (byBodyColorList[i].unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].changed_part), "5", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(byBodyColorList[i].changed_part), "5", rowDataNode);
                }

                // Cellノード(PARENT PART NO.)
                if (byBodyColorList[i].unmatch_flg == 1 || byBodyColorList[i].parent_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(setUnderline(byBodyColorList[i].parent_part_no)), "6", rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(setUnderline(byBodyColorList[i].parent_part_no)), "6", rowDataNode);
                }

                // Cellノード(PART NAME)
                setCell(domDoc, byBodyColorList[i].part_name, "7", rowDataNode);

                // Cellノード(PLANT)
                setCell(domDoc, byBodyColorList[i].plant, "8", rowDataNode);

                // Cellノード(MODEL)
                setCell(domDoc, byBodyColorList[i].model, "9", rowDataNode);

                // Cellノード(DESTINATION)
                setCell(domDoc, byBodyColorList[i].destination, "10", rowDataNode);

                // Cellノード(GRADE)
                setCell(domDoc, byBodyColorList[i].grade, "11", rowDataNode);

                // Cellノード(FEATURE)
                setCell(domDoc, byBodyColorList[i].feature, "12", rowDataNode);

                // Cellノード(ITEM)
                setCell(domDoc, byBodyColorList[i].item, "13", rowDataNode);

                // Cellノード(REMARKS)
                setCell(domDoc, byBodyColorList[i].remark, "14", rowDataNode);

                // Body Color
                // Cellノード(COLOR)
                for (var j = 0; j < byBodyColorList[i].color_item.Count; j++)
                {

                    if (!string.IsNullOrWhiteSpace(byBodyColorList[i].color_item[j]))
                    {

                        setCell(domDoc, byBodyColorList[i].color_item[j], (j + 15).ToString(), rowDataNode);
                    }
                    else
                    {
                        // 部品リレーションの子部品不一致フラグが2の場合、色指示を「-」に設定
                        if (byBodyColorList[i].child_unmatch_flag == 2)
                        {
                            setCell(domDoc, "-", (j + 15).ToString(), rowDataNode);
                        }
                        else
                        {
                            continue;
                            //setCell(domDoc, "", (j + 15).ToString(), rowDataNode);
                        }
                    }

                }

                // Cellノード(L1_BASE PART NO)
                if (byBodyColorList[i].unmatch_flg == 1 || byBodyColorList[i].l1_unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].l1_base_part_no), (colCount - 1).ToString(), rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(byBodyColorList[i].l1_base_part_no), (colCount - 1).ToString(), rowDataNode);
                }

                // Cellノード(BASE PART NO)
                if (byBodyColorList[i].unmatch_flg == 1)
                {
                    setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].base_part_no), colCount.ToString(), rowDataNode);
                }
                else
                {
                    setCell(domDoc, ChangePartNo(byBodyColorList[i].base_part_no), colCount.ToString(), rowDataNode);
                }

                // L1部品等の-色指示になっている部品がBY BODYCOLOR側で出力されている
                tableNode.AppendChild(rowDataNode);

            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "by color");
            tabNode.AppendChild(worksheetOptions);

            return tabNode;
        }

        // [TC]Worksheet
        private void CreateWorksheetTc(XmlElement root, XmlDocument domDoc, TcColorChartGetResponse response)
        {

            // [COLOR]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "TC");

            int colCount = 5 + response.ext_list.Count;
            int rowCount = 2 + response.color_chart_mdg_list.Count;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            // MODIFY 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
            // rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("Height", ns_ss, "40");
            // MODIFY 2022/04/06 END
            rowNode.SetAttribute("header", "true");

            // Rowノード（2色目ヘッダ）
            XmlElement row2Node = domDoc.CreateElement("Row");
            row2Node.SetAttribute("AutoFitHeight", ns_ss, "1");
            row2Node.SetAttribute("header", "true");

            for (var i = 0; i < colCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                if (i < 5)
                {
                    cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                    cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());
                    cellNode.SetAttribute("MergeDown", ns_ss, "1");

                    // Dataノード
                    XmlElement dataNode = domDoc.CreateElement("Data");
                    dataNode.SetAttribute("Type", ns_ss, "String");
                    dataNode.InnerText = TC_NM[i];
                    cellNode.AppendChild(dataNode);
                }
                else
                {
                    cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");
                    cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                    // Dataノード
                    XmlElement dataNode = domDoc.CreateElement("Data");
                    dataNode.SetAttribute("Type", ns_ss, "String");
                    // MODIFY 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
                    // dataNode.InnerText = response.ext_list[i - 5].ext_color1_name;
                    dataNode.InnerText = getColor1Name(response.ext_list[i - 5].ext_color1_name);
                    // MODIFY 2022/04/06 END

                    cellNode.AppendChild(dataNode);
                    var mergeAcross = 0;
                    List<int> indexArr = new List<int>();
                    for (var j = i - 4; j < response.ext_list.Count; j++)
                    {
                        // MODIFY 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
                        // if (response.ext_list[j].ext_color1_name == dataNode.InnerText)
                        var tempExtColor1Name = getColor1Name(response.ext_list[j].ext_color1_name);
                        if (tempExtColor1Name == dataNode.InnerText)
                        // MODIFY 2022/04/06 END
                        {
                            if (indexArr.Count == 0)
                            {
                                indexArr.Add(i);
                            }
                            indexArr.Add(i + 1);
                            mergeAcross++;
                            i++;
                        }
                    }

                    if (mergeAcross != 0)
                    {
                        cellNode.SetAttribute("MergeAcross", ns_ss, mergeAcross.ToString());

                        for (var index = 0; index < indexArr.Count; index++)
                        {
                            // Cellノード
                            XmlElement cell2Node = domDoc.CreateElement("Cell");
                            cell2Node.SetAttribute("StyleID", ns_ss, "sMultiHead");
                            cell2Node.SetAttribute("Index", ns_ss, (indexArr[index] + 1).ToString());

                            // Dataノード
                            XmlElement data2Node = domDoc.CreateElement("Data");
                            data2Node.SetAttribute("Type", ns_ss, "String");
                            if (string.IsNullOrWhiteSpace(response.ext_list[indexArr[index] - 5].ext_color2_name))
                            {
                                data2Node.InnerText = "-";
                            }
                            else
                            {
                                data2Node.InnerText = response.ext_list[indexArr[index] - 5].ext_color2_name.Split(" ")[0];
                            }

                            cell2Node.AppendChild(data2Node);
                            row2Node.AppendChild(cell2Node);
                        }

                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(response.ext_list[i - 5].ext_color2_name))
                        {
                            cellNode.SetAttribute("MergeDown", ns_ss, "1");
                        }
                        else
                        {
                            // Cellノード
                            XmlElement cell2Node = domDoc.CreateElement("Cell");
                            cell2Node.SetAttribute("StyleID", ns_ss, "sMultiHead");
                            cell2Node.SetAttribute("Index", ns_ss, (i + 1).ToString());

                            // Dataノード
                            XmlElement data2Node = domDoc.CreateElement("Data");
                            data2Node.SetAttribute("Type", ns_ss, "String");
                            data2Node.InnerText = response.ext_list[i - 5].ext_color2_name;

                            cell2Node.AppendChild(data2Node);
                            row2Node.AppendChild(cell2Node);
                        }
                    }

                }

                rowNode.AppendChild(cellNode);

            }
            tableNode.AppendChild(rowNode);

            tableNode.AppendChild(row2Node);

            // Data部
            for (var i = 0; i < response.color_chart_mdg_list.Count; i++)
            {
                // Rowノード（ヘッダ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // PLANT
                setCell2(domDoc, response.color_chart_mdg_list[i].production, "1", rowDataNode);

                // MODEL
                setCell2(domDoc, response.color_chart_mdg_list[i].model_code, "2", rowDataNode);

                // DESTINATION
                setCell2(domDoc, response.color_chart_mdg_list[i].destination_cd, "3", rowDataNode);

                // GRADE
                setCell2(domDoc, response.color_chart_mdg_list[i].grade, "4", rowDataNode);

                // FOP
                if (!string.IsNullOrWhiteSpace(response.color_chart_mdg_list[i].fop))
                {
                    setCell2(domDoc, response.color_chart_mdg_list[i].fop, "5", rowDataNode);
                }
                else
                {
                    setCell2(domDoc, "-", "5", rowDataNode);
                }

                for (var j = 0; j < response.ext_list.Count; j++)
                {
                    XmlElement cellNode = domDoc.CreateElement("Cell");
                    cellNode.SetAttribute("Index", ns_ss, (j + 6).ToString());

                    // Dataノード
                    XmlElement dataNode = domDoc.CreateElement("Data");
                    dataNode.SetAttribute("Type", ns_ss, "String");
                    dataNode.InnerText = "";

                    for (var k = 0; k < response.int_list.Count; k++)
                    {
                        if (response.color_chart_mdg_list[i].color_chart_tc_mdgid == response.int_list[k].color_chart_tc_mdgid &&
                             response.ext_list[j].color_chart_tc_ext_color_id == response.int_list[k].color_chart_tc_ext_color_id)
                        {
                            if (string.IsNullOrWhiteSpace(dataNode.InnerText))
                            {
                                dataNode.InnerText = response.int_list[k].symbol;
                            }
                            else
                            {
                                dataNode.InnerText = dataNode.InnerText + "," + response.int_list[k].symbol;
                            }
                        }
                    }

                    cellNode.AppendChild(dataNode);
                    rowDataNode.AppendChild(cellNode);
                }

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "tc");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }


        // [INT COLOR TYPE LIST]Worksheet
        private void CreateWorksheetIntColorTypeList(XmlElement root, XmlDocument domDoc, long? color_chart_tc_id)
        {

            // [INT COLOR TYPE LIST]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "INT COLOR TYPE LIST");

            // Int Color  Type List取得
            InteriorColorGetRequest request = new InteriorColorGetRequest();
            request.color_chart_tc_id = color_chart_tc_id;

            InteriorColorController interiorColorController = new InteriorColorController(null);

            IDictionary<string, object> intListMap = interiorColorController.GetInteriorColorItems(request);
            var intListResults = new List<InteriorColorGetResponse>();
            intListResults = (List<InteriorColorGetResponse>)intListMap["resultList"];

            int rowCount = intListResults.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, "2");
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < 2; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            // Cellノード [SYMBOL]
            XmlElement cellNode = domDoc.CreateElement("Cell");

            cellNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNode.SetAttribute("Index", ns_ss, "1");

            // Dataノード [SYMBOL]
            XmlElement dataNode = domDoc.CreateElement("Data");
            dataNode.SetAttribute("Type", ns_ss, "String");
            dataNode.InnerText = "SYMBOL";
            cellNode.AppendChild(dataNode);

            rowNode.AppendChild(cellNode);

            // Cellノード [NAME]
            XmlElement cellNmNode = domDoc.CreateElement("Cell");

            cellNmNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNmNode.SetAttribute("Index", ns_ss, "2");

            // Dataノード [NAME]
            XmlElement dataNmNode = domDoc.CreateElement("Data");
            dataNmNode.SetAttribute("Type", ns_ss, "String");
            dataNmNode.InnerText = "NAME";
            cellNmNode.AppendChild(dataNmNode);

            rowNode.AppendChild(cellNmNode);

            tableNode.AppendChild(rowNode);

            // Data部
            for (var i = 0; i < intListResults.Count; i++)
            {
                // Rowノード（ヘッダ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // SYMBOL
                setCell2(domDoc, intListResults[i].symbol, "1", rowDataNode);

                // NAME
                setCell2(domDoc, intListResults[i].int_color_type, "2", rowDataNode);

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "int color type list");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // [X-tone color list]Worksheet
        private void CreateWorksheetXtoneColorList(XmlElement root, XmlDocument domDoc, long? color_chart_tc_id)
        {

            // [INT COLOR TYPE LIST]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "X-TONE COLOR LIST");

            // Xtone取得
            XtoneInforGetRequest request = new XtoneInforGetRequest();
            request.color_chart_tc_id = color_chart_tc_id;

            XtoneInforController xtoneInforController = new XtoneInforController(null);
            IDictionary<string, object> xtoneListMap = xtoneInforController.GetXtoneInforItems(request);
            var xtoneListResults = new List<Models.XtoneInfor.XtoneInforGetResponse>();
            xtoneListResults = (List<Models.XtoneInfor.XtoneInforGetResponse>)xtoneListMap["resultList"];

            // ADD 2022/05/30 BEGIN COLORシートルール
            if (xtoneListResults.Count == 0)
            {
                return;
            }
            // ADD 2022/05/30 END

            // 重複なデータ除去
            xtoneListResults = xtoneListResults.Where((x, y) => xtoneListResults.FindIndex(z => z.exterior_color_nm == x.exterior_color_nm) == y).ToList();

            int colCount = xtoneListResults.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, "2");
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < 2; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            // Cellノード [CODE]
            XmlElement cellNode = domDoc.CreateElement("Cell");

            cellNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNode.SetAttribute("Index", ns_ss, "1");

            // Dataノード [CODE]
            XmlElement dataNode = domDoc.CreateElement("Data");
            dataNode.SetAttribute("Type", ns_ss, "String");
            dataNode.InnerText = "CODE";
            cellNode.AppendChild(dataNode);

            rowNode.AppendChild(cellNode);

            // Cellノード [COLOR NAME]
            XmlElement cellNmNode = domDoc.CreateElement("Cell");

            cellNmNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNmNode.SetAttribute("Index", ns_ss, "2");

            // Dataノード [COLOR NAME]
            XmlElement dataNmNode = domDoc.CreateElement("Data");
            dataNmNode.SetAttribute("Type", ns_ss, "String");
            dataNmNode.InnerText = "COLOR NAME";
            cellNmNode.AppendChild(dataNmNode);

            rowNode.AppendChild(cellNmNode);

            tableNode.AppendChild(rowNode);

            // Data部
            for (var i = 0; i < xtoneListResults.Count; i++)
            {
                // Rowノード（ヘッダ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                var exterior_color_nm = xtoneListResults[i].exterior_color_nm.Split(' ');

                string codeNm = "";

                for (var j = 1; j < exterior_color_nm.Length; j++)
                {
                    if (j == 1)
                    {
                        codeNm = exterior_color_nm[j];
                    }
                    else
                    {
                        codeNm = codeNm + " " + exterior_color_nm[j];
                    }

                }

                // SYMBOL
                XmlElement cellCodeNode = domDoc.CreateElement("Cell");
                cellCodeNode.SetAttribute("Index", ns_ss, "1");

                // Dataノード
                XmlElement dataCodeNode = domDoc.CreateElement("Data");
                dataCodeNode.SetAttribute("Type", ns_ss, "String");
                dataCodeNode.InnerText = exterior_color_nm[0];
                cellCodeNode.AppendChild(dataCodeNode);

                // NAME
                XmlElement cellNameNode = domDoc.CreateElement("Cell");
                cellNameNode.SetAttribute("Index", ns_ss, "2");

                // Dataノード
                XmlElement dataNameNode = domDoc.CreateElement("Data");
                dataNameNode.SetAttribute("Type", ns_ss, "String");
                dataNameNode.InnerText = codeNm;
                cellNameNode.AppendChild(dataNameNode);


                rowDataNode.AppendChild(cellCodeNode);
                rowDataNode.AppendChild(cellNameNode);

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "int color type list");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // [BODY COLOR LIST]Worksheet
        private void CreateWorksheetBodyColorList(XmlElement root, XmlDocument domDoc, TcColorChartGetResponse response)
        {

            // [BODY COLOR LIST]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "BODY COLOR LIST");

            int colCount = 2;

            int rowCount = response.ext_list.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            // Cellノード [CODE]
            XmlElement cellNode = domDoc.CreateElement("Cell");

            cellNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNode.SetAttribute("Index", ns_ss, "1");

            // Dataノード [CODE]
            XmlElement dataNode = domDoc.CreateElement("Data");
            dataNode.SetAttribute("Type", ns_ss, "String");
            dataNode.InnerText = "CODE";
            cellNode.AppendChild(dataNode);

            rowNode.AppendChild(cellNode);

            // Cellノード [COLOR NAME]
            XmlElement cellNmNode = domDoc.CreateElement("Cell");

            cellNmNode.SetAttribute("StyleID", ns_ss, "sHead");
            cellNmNode.SetAttribute("Index", ns_ss, "2");

            // Dataノード [COLOR NAME]
            XmlElement dataNmNode = domDoc.CreateElement("Data");
            dataNmNode.SetAttribute("Type", ns_ss, "String");
            dataNmNode.InnerText = "COLOR NAME";
            cellNmNode.AppendChild(dataNmNode);

            rowNode.AppendChild(cellNmNode);

            tableNode.AppendChild(rowNode);

            string ext_color1_code = "";
            string ext_color1_name = "";

            // 重複なデータ除去
            List<ColorChartExtInfo> bodyColorList = new List<ColorChartExtInfo>();
            bodyColorList = response.ext_list;
            bodyColorList = bodyColorList.Where((x, y) => bodyColorList.FindIndex(z => z.ext_color1_name == x.ext_color1_name) == y).ToList();

            // Data部
            for (var i = 0; i < bodyColorList.Count; i++)
            {
                // Row(データ)
                XmlElement rowDataNode = domDoc.CreateElement("Row");
                rowDataNode.SetAttribute("AutoFitHeight", ns_ss, "1");

                if (!string.IsNullOrWhiteSpace(bodyColorList[i].ext_color1_name))
                {
                    var ext_color1_name_arr = bodyColorList[i].ext_color1_name.Split(" ");
                    ext_color1_code = ext_color1_name_arr[0];

                    for (var j = 1; j < ext_color1_name_arr.Length; j++)
                    {
                        if (j == 1)
                        {
                            ext_color1_name = ext_color1_name_arr[j];
                        }
                        else
                        {
                            ext_color1_name = ext_color1_name + " " + ext_color1_name_arr[j];
                        }

                    }
                }

                // Cellノード [CODE]
                XmlElement cellCodeNode = domDoc.CreateElement("Cell");
                cellCodeNode.SetAttribute("Index", ns_ss, "1");

                // Dataノード [CODE]
                XmlElement dataCodeNode = domDoc.CreateElement("Data");
                dataCodeNode.SetAttribute("Type", ns_ss, "String");
                dataCodeNode.InnerText = ext_color1_code;

                cellCodeNode.AppendChild(dataCodeNode);

                // Cellノード [CODE NAME]
                XmlElement cellCodeNameNode = domDoc.CreateElement("Cell");
                cellCodeNameNode.SetAttribute("Index", ns_ss, "2");

                // Dataノード [CODE NAME]
                XmlElement dataCodeNameNode = domDoc.CreateElement("Data");
                dataCodeNameNode.SetAttribute("Type", ns_ss, "String");

                dataCodeNameNode.InnerText = ext_color1_name;
                cellCodeNameNode.AppendChild(dataCodeNameNode);

                rowDataNode.AppendChild(cellCodeNode);
                rowDataNode.AppendChild(cellCodeNameNode);

                tableNode.AppendChild(rowDataNode);

            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "body color list");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // [COLOR SUMMARY]Worksheet
        private void CreateWorksheetColorSummary(XmlElement root, XmlDocument domDoc, List<ColorTabInfo> colorTabInfoList)
        {
            List<ColorSummaryInfo> colorSummaryList = new List<ColorSummaryInfo>();

            string mainNoColor = "";

            string mainNoSummary = "";

            ColorSummaryInfo colorSummaryInfo = new ColorSummaryInfo();

            long isExistFlg = 0;
            int idx = 0;

            // COLOR SUMMARY出力情報作成
            for (var i = 0; i < colorTabInfoList.Count; i++)
            {
                // 2行目以降の場合
                if (colorSummaryList.Count > 0)
                {
                    isExistFlg = 0;
                    idx = 0;

                    // 該当レコードは既にあるかどうか判断
                    for (var j = 0; j < colorSummaryList.Count; j++)
                    {
                        if (!string.IsNullOrWhiteSpace(colorTabInfoList[i].changed_part) && colorTabInfoList[i].changed_part.Length > 5)
                        {
                            mainNoColor = colorTabInfoList[i].changed_part.Substring(0, 5);
                        }
                        else
                        {
                            mainNoColor = NullConvertToString(colorTabInfoList[i].changed_part);
                        }

                        if (colorSummaryList[j].changed_part.Length > 5)
                        {
                            mainNoSummary = colorSummaryList[j].changed_part.Substring(0, 5);
                        }
                        else
                        {
                            mainNoSummary = colorSummaryList[j].changed_part;
                        }

                        if (NullConvertToString(colorTabInfoList[i].sub_section).Equals(colorSummaryList[j].sub_section)
                            && mainNoColor.Equals(mainNoSummary)
                            && NullConvertToString(colorTabInfoList[i].part_name).Equals(colorSummaryList[j].part_name))
                        {

                            isExistFlg = 1;
                            idx = j;

                            if (NullConvertToString(colorTabInfoList[i].body_color).Equals(NullConvertToString(colorSummaryList[j].body_color)))
                            {
                                isExistFlg = 2;
                                idx = j;
                                break;
                            }

                        }
                    }

                    // 同じデータが存在しない場合
                    if (isExistFlg != 2)
                    {
                        // 新しいデータをリストに追加
                        AddNewColorSummaryData(colorSummaryInfo, colorTabInfoList[i], colorSummaryList);
                    }
                    else
                    {
                        // 同じデータが存在する場合、Condition項目設定
                        if (!NullConvertToString(colorTabInfoList[i].plant).Equals(NullConvertToString(colorSummaryList[idx].plant)))
                        {
                            colorSummaryList[idx].plant = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].model).Equals(NullConvertToString(colorSummaryList[idx].model)))
                        {
                            colorSummaryList[idx].model = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].destination).Equals(NullConvertToString(colorSummaryList[idx].destination)))
                        {
                            colorSummaryList[idx].destination = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].grade).Equals(NullConvertToString(colorSummaryList[idx].grade)))
                        {
                            colorSummaryList[idx].grade = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].feature).Equals(NullConvertToString(colorSummaryList[idx].feature)))
                        {
                            colorSummaryList[idx].feature = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].item).Equals(NullConvertToString(colorSummaryList[idx].item)))
                        {
                            colorSummaryList[idx].item = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].remark).Equals(NullConvertToString(colorSummaryList[idx].remark)))
                        {
                            colorSummaryList[idx].remark = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].summary_remark).Equals(NullConvertToString(colorSummaryList[idx].summary_remark)))
                        {
                            colorSummaryList[idx].summary_remark = "*";
                        }

                        List<string> highlightsList = new List<string>();

                        if (colorSummaryList[idx].cross_highlights != null)
                        {

                            highlightsList = colorSummaryList[idx].cross_highlights;

                            if (!highlightsList.Contains(colorTabInfoList[i].cross_highlights))
                            {
                                highlightsList.Add(colorTabInfoList[i].cross_highlights);
                            }
                            highlightsList = highlightsList.Where((x, y) => highlightsList.FindIndex(z => z == x) == y).ToList();

                            colorSummaryList[idx].cross_highlights = highlightsList;

                            if (highlightsList.Count < colorSummaryList[idx].cross_highlights.Count)
                            {
                                for (var k = colorSummaryList[idx].cross_highlights.Count; k > highlightsList.Count; k--)
                                {

                                    colorSummaryList[idx].cross_highlights.Remove(colorSummaryList[idx].cross_highlights[k - 1]);

                                }
                            }
                        }
                    }
                }
                else
                {
                    // 1行目の場合、リストに追加
                    AddNewColorSummaryData(colorSummaryInfo, colorTabInfoList[i], colorSummaryList);
                }

            }
            // ADD 2022/05/12 BEGIN XML出力サマリ
            Dictionary<int, List<ColorSummaryInfo>> summaryMap = new Dictionary<int, List<ColorSummaryInfo>>();
            List<ColorSummaryInfo> colorSummaryListCache = new List<ColorSummaryInfo>();
            // ロジック①
            foreach (var colorSummary in colorSummaryList)
            {
                if (!summaryMap.ContainsKey(colorSummary.summary_no))
                {
                    if (colorSummary.summary_no != 0)
                    {
                        List<ColorSummaryInfo> summaryListCache = new List<ColorSummaryInfo>();
                        summaryListCache.Add(colorSummary);
                        summaryMap.Add(colorSummary.summary_no, summaryListCache);
                    }
                }
                else
                {
                    List<ColorSummaryInfo> summaryListCache = summaryMap[colorSummary.summary_no];
                    summaryMap.Remove(colorSummary.summary_no);
                    summaryListCache.Add(colorSummary);
                    summaryMap.Add(colorSummary.summary_no, summaryListCache);
                }
            }
            foreach (var key in summaryMap.Keys)
            {
                List<ColorSummaryInfo> summaryListCache = summaryMap[key];
                ColorSummaryInfo summaryInfo = summaryListCache[0];
                List<string> subSectionList = new List<string>();
                List<string> mainNoList = new List<string>();
                string subSection = "";
                string mainNo = "";
                // ロジック②
                foreach (var summaryCache in summaryListCache)
                {
                    // DELETE 2022/06/21 BEGIN XMLサマリー修正
                    /*if (summaryCache.child_unmatch_flag == 1)
                    {
                        summaryInfo.child_unmatch_flag = 1;
                    }*/
                    // DELETE 2022/06/21 END
                    if (subSectionList.IndexOf(summaryCache.sub_section) == -1)
                    {
                        subSectionList.Add(summaryCache.sub_section);
                        if (subSection != "")
                        {
                            subSection = subSection.TrimEnd() + "/" + summaryCache.sub_section.TrimEnd();
                        }
                        else
                        {
                            subSection = summaryCache.sub_section.TrimEnd();
                        }
                    }
                    if (mainNoList.IndexOf(summaryCache.main_no) == -1)
                    {
                        mainNoList.Add(summaryCache.main_no);
                        if (mainNo != "")
                        {
                            mainNo = mainNo + "/" + summaryCache.main_no;
                        }
                        else
                        {
                            mainNo = summaryCache.main_no;
                        }
                    }
                }
                summaryInfo.sub_section = subSection;
                summaryInfo.main_no = mainNo;
                // ロジック③とロジック④
                // MODIFY 2022/05/25 BEGIN XML出力サマリ
                //List<string> partNameList = new List<string>();
                //if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                //{
                //    summaryInfo.part_name = summaryListCache[0].summary_group;
                //}
                //else
                //{
                //    if (summaryListCache[0].rl_pair_no != 0)
                //    {
                //        foreach (var summaryCache in summaryListCache)
                //        {
                //            partNameList.Add(summaryCache.part_name);
                //        }
                //        summaryInfo.part_name = GetPartName(partNameList);
                //    }
                //}
                if (summaryListCache[0].rl_pair_no != 0)
                {
                    List<string> partNameListCache = new List<string>();
                    foreach (var summaryCache in summaryListCache)
                    {
                        if (partNameListCache.IndexOf(summaryCache.part_name) == -1)
                        {
                            partNameListCache.Add(summaryCache.part_name);
                        }
                    }
                    if (partNameListCache.Count == 2)
                    {
                        summaryInfo.part_name = GetPartName(partNameListCache);
                    }
                    else
                    {
                        if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                        {
                            summaryInfo.part_name = summaryListCache[0].summary_group;
                        }
                    }

                }
                else
                {
                    if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                    {
                        summaryInfo.part_name = summaryListCache[0].summary_group;
                    }
                }
                // MODIFY 2022/05/25 END
                // ロジック⑤
                summaryInfo.remark = summaryListCache[0].summary_remark;
                foreach (var summaryCache in summaryListCache)
                {
                    if (summaryInfo.remark != summaryCache.summary_remark)
                    {
                        summaryInfo.remark = "*";
                        break;
                    }
                }
                // ADD 2022/06/21 BEGIN XMLサマリー修正
                List<string> bodyColorList = new List<string>();
                bool childUnmatchFlag = false;
                foreach (var summaryCache in summaryListCache)
                {
                    if (bodyColorList.IndexOf(summaryCache.body_color) == -1 && summaryCache.body_color != null && summaryCache.body_color != "")
                    {
                        bodyColorList.Add(summaryCache.body_color);
                    }
                    if (summaryCache.child_unmatch_flag == 0)
                    {
                        childUnmatchFlag = true;
                    }
                }
                if (childUnmatchFlag)
                {
                    if (bodyColorList.Count > 1)
                    {
                        summaryInfo.body_color = "-";
                    }
                    else if (bodyColorList.Count == 0)
                    {
                        summaryInfo.body_color = "";
                    }
                    else
                    {
                        summaryInfo.body_color = bodyColorList[0];
                    }
                }
                else
                {
                    summaryInfo.body_color = "-";
                }
                // ADD 2022/06/21 END
                colorSummaryListCache.Add(summaryInfo);
            }
            colorSummaryList = colorSummaryListCache;
            // ADD 2022/05/12 END

            // [COLOR SUMMARY]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "COLOR SUMMARY");

            int colCount = 12;

            int rowCount = colorSummaryList.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("AutoFitHeight", ns_ss, "1");
            rowNode.SetAttribute("header", "true");

            for (var i = 0; i < colCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = EXT_SUMMARY_NM[i];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            tableNode.AppendChild(rowNode);

            // Data部
            for (var i = 0; i < colorSummaryList.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                if (colorSummaryList[i].cross_highlights != null)
                {
                    for (var j = 0; j < colorSummaryList[i].cross_highlights.Count; j++)
                    {
                        // Highlightノード
                        // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                        List<string> highLightList = getHighLight(colorSummaryList[i].cross_highlights[j]);
                        foreach (var highLightInfo in highLightList)
                        {
                            XmlElement highlightNode = domDoc.CreateElement("Highlight");
                            highlightNode.SetAttribute("id", highLightInfo);
                            rowDataNode.AppendChild(highlightNode);
                        }
                        //XmlElement highlightNode = domDoc.CreateElement("Highlight");
                        //highlightNode.SetAttribute("id", colorSummaryList[i].cross_highlights[j]);
                        //rowDataNode.AppendChild(highlightNode);
                        // MODIFY 2022/05/27 END
                    }
                }
                else
                {
                    // Highlightノード
                    XmlElement highlightNode = domDoc.CreateElement("Highlight");
                    highlightNode.SetAttribute("id", "");
                    rowDataNode.AppendChild(highlightNode);
                }


                // Cellノード(SECTION)
                setCell(domDoc, colorSummaryList[i].section, "1", rowDataNode);

                // Cellノード(SUB-SECTION)
                // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                // setCell(domDoc, colorSummaryList[i].sub_section, "2", rowDataNode);
                setCell(domDoc, colorSummaryList[i].sub_section.TrimEnd(), "2", rowDataNode);
                // MODIFY 2022/04/08 END

                // Cellノード(L1)
                setCell(domDoc, colorSummaryList[i].main_no, "3", rowDataNode);

                // Cellノード(PART NAME)
                setCell(domDoc, colorSummaryList[i].part_name, "4", rowDataNode);

                // Cellノード(PLANT)
                setCell(domDoc, colorSummaryList[i].plant, "5", rowDataNode);

                // Cellノード(MODEL)
                setCell(domDoc, colorSummaryList[i].model, "6", rowDataNode);

                // Cellノード(DESTINATION)
                setCell(domDoc, colorSummaryList[i].destination, "7", rowDataNode);

                // Cellノード(GRADE)
                setCell(domDoc, colorSummaryList[i].grade, "8", rowDataNode);

                // Cellノード(FEATURE)
                setCell(domDoc, colorSummaryList[i].feature, "9", rowDataNode);

                // Cellノード(ITEM)
                setCell(domDoc, colorSummaryList[i].item, "10", rowDataNode);

                // Cellノード(REMARKS)
                setCell(domDoc, colorSummaryList[i].remark, "11", rowDataNode);

                // Cellノード(COLOR)
                // MODIFY 2022/06/21 BEGIN XMLサマリー修正
                /*if (colorSummaryList[i].child_unmatch_flag == 1)
                {
                    setCell(domDoc, "-", "12", rowDataNode);
                }
                else
                {
                    setCell(domDoc, colorSummaryList[i].body_color, "12", rowDataNode);
                }*/
                setCell(domDoc, colorSummaryList[i].body_color, "12", rowDataNode);
                // MODIFY 2022/06/21 END

                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "Color Summary");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // COLOR SUMMARYデータ追加
        private void AddNewColorSummaryData(ColorSummaryInfo colorSummaryInfo, ColorTabInfo colorTabInfo, List<ColorSummaryInfo> colorSummaryList)
        {

            if (colorTabInfo.body_color == null || colorTabInfo.body_color.Equals(""))
                return;

            colorSummaryInfo = new ColorSummaryInfo();

            colorSummaryInfo.section = colorTabInfo.section;
            colorSummaryInfo.sub_section = colorTabInfo.sub_section;
            colorSummaryInfo.l1_part_no = colorTabInfo.l1_part_no;
            colorSummaryInfo.lvl = colorTabInfo.lvl;
            colorSummaryInfo.changed_part = colorTabInfo.changed_part;
            colorSummaryInfo.parent_part_no = colorTabInfo.parent_part_no;
            colorSummaryInfo.part_name = colorTabInfo.part_name;
            colorSummaryInfo.plant = colorTabInfo.plant;
            colorSummaryInfo.model = colorTabInfo.model;
            colorSummaryInfo.destination = colorTabInfo.destination;
            colorSummaryInfo.grade = colorTabInfo.grade;
            colorSummaryInfo.feature = colorTabInfo.feature;
            colorSummaryInfo.item = colorTabInfo.item;
            colorSummaryInfo.remark = colorTabInfo.remark;
            colorSummaryInfo.l1_base_part_no = colorTabInfo.l1_base_part_no;
            colorSummaryInfo.base_part_no = colorTabInfo.base_part_no;
            colorSummaryInfo.body_color = colorTabInfo.body_color;
            colorSummaryInfo.child_unmatch_flag = colorTabInfo.child_unmatch_flag;
            // ADD 2022/05/12 BEGIN XML出力サマリ
            colorSummaryInfo.summary_remark = colorTabInfo.summary_remark;
            colorSummaryInfo.rl_pair_no = colorTabInfo.rl_pair_no;
            colorSummaryInfo.rl_flag = colorTabInfo.rl_flag;
            colorSummaryInfo.summary_no = colorTabInfo.summary_no;
            colorSummaryInfo.summary_group = colorTabInfo.summary_group;
            // ADD 2022/05/12 END
            if (!string.IsNullOrWhiteSpace(colorTabInfo.changed_part) && colorTabInfo.changed_part.Length > 5)
            {
                colorSummaryInfo.main_no = colorTabInfo.changed_part.Substring(0, 5);
            }
            else
            {
                colorSummaryInfo.main_no = NullConvertToString(colorTabInfo.changed_part);
            }

            // ハイライト設定
            List<string> highlightsList = new List<string>();

            highlightsList.Add(colorTabInfo.cross_highlights);

            highlightsList = highlightsList.Where((x, y) => highlightsList.FindIndex(z => z == x) == y).ToList();

            colorSummaryInfo.cross_highlights = highlightsList;

            colorSummaryList.Add(colorSummaryInfo);
        }

        // [BY BODYCOLOR SUMMARY]Worksheet
        private void CreateWorksheetByBodyColorColorSummary(XmlElement root, XmlDocument domDoc, List<ColorTabInfo> colorTabInfoList)
        {

            List<ByBodyColorSummaryInfo> colorSummaryList = new List<ByBodyColorSummaryInfo>();

            string mainNoColor = "";

            string mainNoSummary = "";

            ByBodyColorSummaryInfo byBodyColorSummaryInfo = null;

            bool colorItemFlg = true;

            long isExistFlg = 0;
            int idx = 0;

            // BY BODYCOLOR SUMMARY出力情報作成
            for (var i = 0; i < colorTabInfoList.Count; i++)
            {

                isExistFlg = 0;
                idx = 0;

                if (colorSummaryList.Count > 0)
                {

                    for (var j = 0; j < colorSummaryList.Count; j++)
                    {
                        colorItemFlg = true;

                        if (!string.IsNullOrWhiteSpace(colorTabInfoList[i].changed_part) && colorTabInfoList[i].changed_part.Length > 5)
                        {
                            mainNoColor = colorTabInfoList[i].changed_part.Substring(0, 5);
                        }
                        else
                        {
                            mainNoColor = NullConvertToString(colorTabInfoList[i].changed_part);
                        }

                        if (colorSummaryList[j].changed_part.Length > 5)
                        {
                            mainNoSummary = colorSummaryList[j].changed_part.Substring(0, 5);
                        }
                        else
                        {
                            mainNoSummary = colorSummaryList[j].changed_part;
                        }

                        if (NullConvertToString(colorTabInfoList[i].sub_section).Equals(colorSummaryList[j].sub_section)
                            && mainNoColor.Equals(mainNoSummary)
                            && NullConvertToString(colorTabInfoList[i].part_name).Equals(colorSummaryList[j].part_name))
                        {

                            for (var k = 0; k < colorTabInfoList[i].color_item.Count; k++)
                            {

                                if (!colorTabInfoList[i].color_item[k].Equals(colorSummaryList[j].color_item[k]))
                                {
                                    colorItemFlg = false;
                                }

                            }

                            // 色差異がある場合
                            if (colorItemFlg)
                            {
                                isExistFlg = 2;
                                idx = j;
                                break;
                            }

                        }

                    }

                    // 同じデータが存在しない場合
                    if (isExistFlg != 2)
                    {
                        AddNewByBodyColorSummaryData(byBodyColorSummaryInfo, colorTabInfoList[i], colorSummaryList);
                    }
                    else
                    {
                        // 同じデータが存在する場合、Condition項目設定
                        if (!NullConvertToString(colorTabInfoList[i].plant).Equals(NullConvertToString(colorSummaryList[idx].plant)))
                        {
                            colorSummaryList[idx].plant = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].model).Equals(NullConvertToString(colorSummaryList[idx].model)))
                        {
                            colorSummaryList[idx].model = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].destination).Equals(NullConvertToString(colorSummaryList[idx].destination)))
                        {
                            colorSummaryList[idx].destination = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].grade).Equals(NullConvertToString(colorSummaryList[idx].grade)))
                        {
                            colorSummaryList[idx].grade = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].feature).Equals(NullConvertToString(colorSummaryList[idx].feature)))
                        {
                            colorSummaryList[idx].feature = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].item).Equals(NullConvertToString(colorSummaryList[idx].item)))
                        {
                            colorSummaryList[idx].item = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].remark).Equals(NullConvertToString(colorSummaryList[idx].remark)))
                        {
                            colorSummaryList[idx].remark = "*";
                        }

                        if (!NullConvertToString(colorTabInfoList[i].summary_remark).Equals(NullConvertToString(colorSummaryList[idx].summary_remark)))
                        {
                            colorSummaryList[idx].summary_remark = "*";
                        }

                        List<string> highlightsList = new List<string>();

                        highlightsList = colorSummaryList[idx].cross_highlights;

                        if (highlightsList != null)
                        {

                            if (!highlightsList.Contains(colorTabInfoList[i].cross_highlights))
                            {
                                highlightsList.Add(colorTabInfoList[i].cross_highlights);
                            }

                            highlightsList = highlightsList.Where((x, y) => highlightsList.FindIndex(z => z == x) == y).ToList();

                            colorSummaryList[idx].cross_highlights = highlightsList;

                            if (highlightsList.Count < colorSummaryList[idx].cross_highlights.Count)
                            {
                                for (var k = colorSummaryList[idx].cross_highlights.Count; k > highlightsList.Count; k--)
                                {

                                    colorSummaryList[idx].cross_highlights.Remove(colorSummaryList[idx].cross_highlights[k - 1]);

                                }
                            }
                        }
                    }

                }
                else
                {
                    AddNewByBodyColorSummaryData(byBodyColorSummaryInfo, colorTabInfoList[i], colorSummaryList);
                }
            }

            // ADD 2022/05/12 BEGIN XML出力サマリ
            Dictionary<int, List<ByBodyColorSummaryInfo>> summaryMap = new Dictionary<int, List<ByBodyColorSummaryInfo>>();
            List<ByBodyColorSummaryInfo> colorSummaryListCache = new List<ByBodyColorSummaryInfo>();
            // ロジック①
            foreach (var colorSummary in colorSummaryList)
            {
                if (!summaryMap.ContainsKey(colorSummary.summary_no))
                {
                    if (colorSummary.summary_no != 0)
                    {
                        List<ByBodyColorSummaryInfo> summaryListCache = new List<ByBodyColorSummaryInfo>();
                        summaryListCache.Add(colorSummary);
                        summaryMap.Add(colorSummary.summary_no, summaryListCache);
                    }
                }
                else
                {
                    List<ByBodyColorSummaryInfo> summaryListCache = summaryMap[colorSummary.summary_no];
                    summaryMap.Remove(colorSummary.summary_no);
                    summaryListCache.Add(colorSummary);
                    summaryMap.Add(colorSummary.summary_no, summaryListCache);
                }
            }
            foreach (var key in summaryMap.Keys)
            {
                List<ByBodyColorSummaryInfo> summaryListCache = summaryMap[key];
                ByBodyColorSummaryInfo summaryInfo = summaryListCache[0];
                List<string> subSectionList = new List<string>();
                List<string> mainNoList = new List<string>();
                string subSection = "";
                string mainNo = "";
                // ロジック②
                foreach (var summaryCache in summaryListCache)
                {
                    if (subSectionList.IndexOf(summaryCache.sub_section) == -1)
                    {
                        subSectionList.Add(summaryCache.sub_section);
                        if (subSection != "")
                        {
                            subSection = subSection.TrimEnd() + "/" + summaryCache.sub_section.TrimEnd();
                        }
                        else
                        {
                            subSection = summaryCache.sub_section.TrimEnd();
                        }
                    }
                    if (mainNoList.IndexOf(summaryCache.main_no) == -1)
                    {
                        mainNoList.Add(summaryCache.main_no);
                        if (mainNo != "")
                        {
                            mainNo = mainNo + "/" + summaryCache.main_no;
                        }
                        else
                        {
                            mainNo = summaryCache.main_no;
                        }
                    }
                }
                summaryInfo.sub_section = subSection;
                summaryInfo.main_no = mainNo;
                // ロジック③とロジック④
                // MODIFY 2022/05/25 BEGIN XML出力サマリ
                //List<string> partNameList = new List<string>();
                //if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                //{
                //    summaryInfo.part_name = summaryListCache[0].summary_group;
                //}
                //else
                //{
                //    if (summaryListCache[0].rl_pair_no != 0)
                //    {
                //        foreach (var summaryCache in summaryListCache)
                //        {
                //            partNameList.Add(summaryCache.part_name);
                //        }
                //        summaryInfo.part_name = GetPartName(partNameList);
                //    }
                //}
                if (summaryListCache[0].rl_pair_no != 0)
                {
                    List<string> partNameListCache = new List<string>();
                    foreach (var summaryCache in summaryListCache)
                    {
                        if (partNameListCache.IndexOf(summaryCache.part_name) == -1)
                        {
                            partNameListCache.Add(summaryCache.part_name);
                        }
                    }
                    if (partNameListCache.Count == 2)
                    {
                        summaryInfo.part_name = GetPartName(partNameListCache);
                    }
                    else
                    {
                        if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                        {
                            summaryInfo.part_name = summaryListCache[0].summary_group;
                        }
                    }

                }
                else
                {
                    if (summaryListCache[0].summary_group != null && summaryListCache[0].summary_group != "")
                    {
                        summaryInfo.part_name = summaryListCache[0].summary_group;
                    }
                }
                // MODIFY 2022/05/25 END
                // ロジック⑤
                summaryInfo.remark = summaryListCache[0].summary_remark;
                foreach (var summaryCache in summaryListCache)
                {
                    if (summaryInfo.remark != summaryCache.summary_remark)
                    {
                        summaryInfo.remark = "*";
                        break;
                    }
                }
                colorSummaryListCache.Add(summaryInfo);
            }
            colorSummaryList = colorSummaryListCache;
            // ADD 2022/05/12 END

            // [BY BODYCOLOR SUMMARY]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "BY BODYCOLOR SUMMARY");

            long colCount = 11;

            colCount = colCount + bodyColorHeaderCount;

            int rowCount = colorSummaryList.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, colCount.ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("Height", ns_ss, "40");
            rowNode.SetAttribute("header", "true");

            for (var i = 0; i < 11; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = EXT_SUMMARY_NM[i];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            for (var i = 0; i < bodyColorHeaderCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 12).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = colorHeaderName[i];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);
            }

            tableNode.AppendChild(rowNode);

            // Data部
            for (var i = 0; i < colorSummaryList.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                if (colorSummaryList[i].cross_highlights != null)
                {
                    // Highlightノード
                    for (var j = 0; j < colorSummaryList[i].cross_highlights.Count; j++)
                    {
                        // Highlightノード
                        // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                        List<string> highLightList = getHighLight(colorSummaryList[i].cross_highlights[j]);
                        foreach (var highLightInfo in highLightList)
                        {
                            XmlElement highlightNode = domDoc.CreateElement("Highlight");
                            highlightNode.SetAttribute("id", highLightInfo);
                            rowDataNode.AppendChild(highlightNode);
                        }
                        //XmlElement highlightNode = domDoc.CreateElement("Highlight");
                        //highlightNode.SetAttribute("id", colorSummaryList[i].cross_highlights[j]);
                        //rowDataNode.AppendChild(highlightNode);
                        // MODIFY 2022/05/27 END
                    }
                }
                else
                {
                    // Highlightノード
                    XmlElement highlightNode = domDoc.CreateElement("Highlight");
                    highlightNode.SetAttribute("id", "");
                    rowDataNode.AppendChild(highlightNode);
                }

                // Cellノード(SECTION)
                setCell(domDoc, colorSummaryList[i].section, "1", rowDataNode);

                // Cellノード(SUB-SECTION)
                // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                // setCell(domDoc, colorSummaryList[i].sub_section, "2", rowDataNode);
                setCell(domDoc, colorSummaryList[i].sub_section.TrimEnd(), "2", rowDataNode);
                // MODIFY 2022/04/08 END

                // Cellノード(L1)
                setCell(domDoc, colorSummaryList[i].main_no, "3", rowDataNode);

                // Cellノード(PART NAME)
                setCell(domDoc, colorSummaryList[i].part_name, "4", rowDataNode);

                // Cellノード(PLANT)
                setCell(domDoc, colorSummaryList[i].plant, "5", rowDataNode);

                // Cellノード(MODEL)
                setCell(domDoc, colorSummaryList[i].model, "6", rowDataNode);

                // Cellノード(DESTINATION)
                setCell(domDoc, colorSummaryList[i].destination, "7", rowDataNode);

                // Cellノード(GRADE)
                setCell(domDoc, colorSummaryList[i].grade, "8", rowDataNode);

                // Cellノード(FEATURE)
                setCell(domDoc, colorSummaryList[i].feature, "9", rowDataNode);

                // Cellノード(ITEM)
                setCell(domDoc, colorSummaryList[i].item, "10", rowDataNode);

                // Cellノード(REMARKS)
                setCell(domDoc, colorSummaryList[i].remark, "11", rowDataNode);

                // Cellノード(COLOR)
                for (var j = 0; j < colorSummaryList[i].color_item.Count; j++)
                {

                    if (!string.IsNullOrWhiteSpace(colorSummaryList[i].color_item[j]))
                    {
                        setCell(domDoc, colorSummaryList[i].color_item[j], (j + 12).ToString(), rowDataNode);
                    }
                    else
                    {
                        if (colorSummaryList[i].child_unmatch_flag == 2)
                        {
                            setCell(domDoc, "-", (j + 12).ToString(), rowDataNode);
                        }
                        else
                        {
                            setCell(domDoc, "", (j + 12).ToString(), rowDataNode);
                        }

                    }

                }

                // L1部品等の-色指示になっている部品がBY BODYCOLOR側で出力されている
                tableNode.AppendChild(rowDataNode);
            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "by bodyColor Summary");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // BY BODYCOLOR SUMMARYデータ追加
        private void AddNewByBodyColorSummaryData(ByBodyColorSummaryInfo byBodyColorSummaryInfo, ColorTabInfo colorTabInfo, List<ByBodyColorSummaryInfo> colorSummaryList)
        {

            byBodyColorSummaryInfo = new ByBodyColorSummaryInfo();

            List<string> colorItemList = new List<string>();

            byBodyColorSummaryInfo.section = colorTabInfo.section;
            byBodyColorSummaryInfo.sub_section = colorTabInfo.sub_section;
            byBodyColorSummaryInfo.l1_part_no = colorTabInfo.l1_part_no;
            byBodyColorSummaryInfo.lvl = colorTabInfo.lvl;
            byBodyColorSummaryInfo.changed_part = colorTabInfo.changed_part;
            byBodyColorSummaryInfo.parent_part_no = colorTabInfo.parent_part_no;
            byBodyColorSummaryInfo.part_name = colorTabInfo.part_name;
            byBodyColorSummaryInfo.plant = colorTabInfo.plant;
            byBodyColorSummaryInfo.model = colorTabInfo.model;
            byBodyColorSummaryInfo.destination = colorTabInfo.destination;
            byBodyColorSummaryInfo.grade = colorTabInfo.grade;
            byBodyColorSummaryInfo.feature = colorTabInfo.feature;
            byBodyColorSummaryInfo.item = colorTabInfo.item;
            byBodyColorSummaryInfo.remark = colorTabInfo.remark;
            byBodyColorSummaryInfo.l1_base_part_no = colorTabInfo.l1_base_part_no;
            byBodyColorSummaryInfo.base_part_no = colorTabInfo.base_part_no;
            byBodyColorSummaryInfo.child_unmatch_flag = colorTabInfo.child_unmatch_flag;
            // ADD 2022/05/12 BEGIN XML出力サマリ
            byBodyColorSummaryInfo.summary_remark = colorTabInfo.summary_remark;
            byBodyColorSummaryInfo.rl_pair_no = colorTabInfo.rl_pair_no;
            byBodyColorSummaryInfo.rl_flag = colorTabInfo.rl_flag;
            byBodyColorSummaryInfo.summary_no = colorTabInfo.summary_no;
            byBodyColorSummaryInfo.summary_group = colorTabInfo.summary_group;
            // ADD 2022/05/12 END
            if (!string.IsNullOrWhiteSpace(colorTabInfo.changed_part) && colorTabInfo.changed_part.Length > 5)
            {
                byBodyColorSummaryInfo.main_no = colorTabInfo.changed_part.Substring(0, 5);
            }
            else
            {
                byBodyColorSummaryInfo.main_no = NullConvertToString(colorTabInfo.changed_part);
            }

            for (var j = 0; j < colorTabInfo.color_item.Count; j++)
            {
                colorItemList.Add(colorTabInfo.color_item[j]);


            }

            // カラー明細情報設定
            byBodyColorSummaryInfo.color_item = colorItemList;

            // カラー明細情報設定
            byBodyColorSummaryInfo.color_header = colorTabInfo.color_header_name;

            // ハイライト設定
            List<string> highlightsList = new List<string>();

            highlightsList.Add(colorTabInfo.cross_highlights);

            highlightsList = highlightsList.Where((x, y) => highlightsList.FindIndex(z => z == x) == y).ToList();

            byBodyColorSummaryInfo.cross_highlights = highlightsList;

            colorSummaryList.Add(byBodyColorSummaryInfo);

        }
        // [INT SUMMARY]Worksheet
        private void CreateWorksheetIntSummary(XmlElement root, XmlDocument domDoc, Dictionary<string, GroupIntColorTypeResponse> summaryTableHeader, List<SummaryItemsResponse> sumTableData, List<GroupIntColorTypeResponseSum> sameColumId)
        {

            // ADD 2022/05/12 BEGIN XML出力サマリ
            Dictionary<int, List<SummaryItemsResponse>> summaryMap = new Dictionary<int, List<SummaryItemsResponse>>();
            List<SummaryItemsResponse> sumTableDataCache = new List<SummaryItemsResponse>();
            // ロジック①
            foreach (var sumTableInfo in sumTableData)
            {
                if (!string.IsNullOrWhiteSpace(sumTableInfo.CHANGED_PART) && sumTableInfo.CHANGED_PART.Length > 5)
                {
                    if (string.IsNullOrEmpty(sumTableInfo.CHANGED_PART.Substring(0, 5)))
                    {
                        sumTableInfo.CHANGED_PART = "";
                    }
                    else
                    {
                        sumTableInfo.CHANGED_PART = sumTableInfo.CHANGED_PART.Substring(0, 5);
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(sumTableInfo.CHANGED_PART))
                    {
                        sumTableInfo.CHANGED_PART = "";
                    }
                    else
                    {
                        sumTableInfo.CHANGED_PART = sumTableInfo.CHANGED_PART;
                    }
                }
                if (!summaryMap.ContainsKey(sumTableInfo.SUMMARY_NO))
                {
                    if (sumTableInfo.SUMMARY_NO != 0)
                    {
                        List<SummaryItemsResponse> summaryListCache = new List<SummaryItemsResponse>();
                        summaryListCache.Add(sumTableInfo);
                        summaryMap.Add(sumTableInfo.SUMMARY_NO, summaryListCache);
                    }
                }
                else
                {
                    List<SummaryItemsResponse> summaryListCache = summaryMap[sumTableInfo.SUMMARY_NO];
                    summaryMap.Remove(sumTableInfo.SUMMARY_NO);
                    summaryListCache.Add(sumTableInfo);
                    summaryMap.Add(sumTableInfo.SUMMARY_NO, summaryListCache);
                }
            }
            foreach (var key in summaryMap.Keys)
            {
                List<SummaryItemsResponse> summaryListCache = summaryMap[key];
                SummaryItemsResponse summaryInfo = summaryListCache[0];
                List<string> subSectionList = new List<string>();
                List<string> mainNoList = new List<string>();
                string subSection = "";
                string mainNo = "";
                // ロジック②
                foreach (var summaryCache in summaryListCache)
                {
                    if (subSectionList.IndexOf(summaryCache.SUB_SECTION) == -1)
                    {
                        subSectionList.Add(summaryCache.SUB_SECTION);
                        if (subSection != "")
                        {
                            subSection = subSection + "/" + summaryCache.SUB_SECTION.TrimEnd();
                        }
                        else
                        {
                            subSection = summaryCache.SUB_SECTION.TrimEnd();
                        }
                    }
                    if (mainNoList.IndexOf(summaryCache.CHANGED_PART) == -1)
                    {
                        mainNoList.Add(summaryCache.CHANGED_PART);
                        if (mainNo != "")
                        {
                            mainNo = mainNo + "/" + summaryCache.CHANGED_PART;
                        }
                        else
                        {
                            mainNo = summaryCache.CHANGED_PART;
                        }
                    }
                }
                summaryInfo.SUB_SECTION = subSection;
                summaryInfo.CHANGED_PART = mainNo;
                // ロジック③とロジック④
                // MODIFY 2022/05/25 BEGIN XML出力サマリ
                //List<string> partNameList = new List<string>();
                //if (summaryListCache[0].SUMMARY_GROUP != null && summaryListCache[0].SUMMARY_GROUP != "")
                //{
                //    summaryInfo.PART_NM = summaryListCache[0].SUMMARY_GROUP;
                //}
                //else
                //{
                //    if (summaryListCache[0].RL_PAIR_NO != 0 && summaryListCache.Count == 2)
                //    {
                //        foreach (var summaryCache in summaryListCache)
                //        {
                //            partNameList.Add(summaryCache.PART_NM);
                //        }
                //        summaryInfo.PART_NM = GetPartName(partNameList);
                //    }
                //}
                if (summaryListCache[0].RL_PAIR_NO != 0)
                {
                    List<string> partNameListCache = new List<string>();
                    foreach (var summaryCache in summaryListCache)
                    {
                        if (partNameListCache.IndexOf(summaryCache.PART_NM) == -1)
                        {
                            partNameListCache.Add(summaryCache.PART_NM);
                        }
                    }
                    if (partNameListCache.Count == 2)
                    {
                        summaryInfo.PART_NM = GetPartName(partNameListCache);
                    }
                    else
                    {
                        if (summaryListCache[0].SUMMARY_GROUP != null && summaryListCache[0].SUMMARY_GROUP != "")
                        {
                            summaryInfo.PART_NM = summaryListCache[0].SUMMARY_GROUP;
                        }
                    }

                }
                else
                {
                    if (summaryListCache[0].SUMMARY_GROUP != null && summaryListCache[0].SUMMARY_GROUP != "")
                    {
                        summaryInfo.PART_NM = summaryListCache[0].SUMMARY_GROUP;
                    }
                }
                // MODIFY 2022/05/25 END
                // ロジック⑤
                summaryInfo.REMARK = summaryListCache[0].SUMMARY_REMARK;
                foreach (var summaryCache in summaryListCache)
                {
                    if (summaryInfo.REMARK != summaryCache.SUMMARY_REMARK)
                    {
                        summaryInfo.REMARK = "*";
                        break;
                    }
                }
                // ADD 2022/06/21 BEGIN XMLサマリー修正
                List<ColorHesItemsResponse> colorHesPartItemCache = new List<ColorHesItemsResponse>();
                List<long> intColorIdList = new List<long>();
                foreach (var summaryCache in summaryListCache)
                {
                    for (int i = 0; i < summaryCache.ColorHesPartItems.Count; i++)
                    {
                        if (summaryCache.ColorHesPartItems[i] != null
                            && summaryCache.ColorHesPartItems[i].INT_COLOR_ID != null
                            && intColorIdList.IndexOf((long)summaryCache.ColorHesPartItems[i].INT_COLOR_ID) == -1)
                        {
                            intColorIdList.Add((long)summaryCache.ColorHesPartItems[i].INT_COLOR_ID);
                            colorHesPartItemCache.Add(summaryCache.ColorHesPartItems[i]);
                        }
                    }
                }
                summaryInfo.ColorHesPartItems = colorHesPartItemCache;
                // ADD 2022/06/21 END
                sumTableDataCache.Add(summaryInfo);
            }
            sumTableData = sumTableDataCache;
            // ADD 2022/05/12 END

            // [SUMMARY]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "SUMMARY");

            int colCount = 12;

            int rowCount = sumTableData.Count + 1;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, (colCount + summaryTableHeader.Count).ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, rowCount.ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount + summaryTableHeader.Count; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("Height", ns_ss, "32.4");
            rowNode.SetAttribute("header", "true");

            for (var i = 0; i < colCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = INT_SUMMARY_NM[i];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }
            tableNode.AppendChild(rowNode);

            // Cell設定（ヘッダ）
            var item1Count = 12;
            foreach (var item in summaryTableHeader)
            {
                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");
                cellNode.SetAttribute("Index", ns_ss, (item1Count + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = item.Value.SYMBOL + "&#10;" + item.Value.INT_COLOR_TYPE;
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);
                item1Count++;
                if (item1Count == 12 + summaryTableHeader.Count)
                {
                    break;
                }
            }
            tableNode.AppendChild(rowNode);

            // Data部
            for (var i = 0; i < sumTableData.Count; i++)
            {
                XmlElement rowNodeData = domDoc.CreateElement("Row");

                // HighLight設定
                foreach (var highLight in sumTableData[i].CROSS_HIGHLIGHTS)
                {
                    // MODIFY 2022/05/27 BEGIN クロスハイライト複数対応
                    List<string> highLightList = getHighLight(highLight);
                    foreach (var highLightInfo in highLightList)
                    {
                        XmlElement HighlightData = domDoc.CreateElement("Highlight");
                        HighlightData.SetAttribute("id", highLightInfo);
                        rowNodeData.AppendChild(HighlightData);
                    }
                    //XmlElement HighlightData = domDoc.CreateElement("Highlight");
                    //HighlightData.SetAttribute("id", highLight);
                    //rowNodeData.AppendChild(HighlightData);
                    // MODIFY 2022/05/27 END

                }

                // Cell1ノード
                XmlElement cellNode1 = domDoc.CreateElement("Cell");
                cellNode1.SetAttribute("Index", ns_ss, (1).ToString());
                // Data1ノード
                XmlElement dataNode1 = domDoc.CreateElement("Data");
                dataNode1.SetAttribute("Type", ns_ss, "String");
                dataNode1.InnerText = sumTableData[i].part_group;
                cellNode1.AppendChild(dataNode1);
                rowNodeData.AppendChild(cellNode1);

                // Cell2ノード
                XmlElement cellNode2 = domDoc.CreateElement("Cell");
                cellNode2.SetAttribute("Index", ns_ss, (2).ToString());
                // Data2ノード
                XmlElement dataNode2 = domDoc.CreateElement("Data");
                dataNode2.SetAttribute("Type", ns_ss, "String");
                dataNode2.InnerText = sumTableData[i].SECTION;

                cellNode2.AppendChild(dataNode2);
                rowNodeData.AppendChild(cellNode2);

                // Cell3ノード
                XmlElement cellNode3 = domDoc.CreateElement("Cell");
                cellNode3.SetAttribute("Index", ns_ss, (3).ToString());
                // Data3ノード
                XmlElement dataNode3 = domDoc.CreateElement("Data");
                dataNode3.SetAttribute("Type", ns_ss, "String");

                // MODIFY 2022/04/08 BEGIN XMLデータ出力改善(Subsection空白除去)
                // dataNode3.InnerText = sumTableData[i].SUB_SECTION;
                // Subsection6,7桁目が空白の際は空白除去
                dataNode3.InnerText = sumTableData[i].SUB_SECTION.TrimEnd();
                // MODIFY 2022/04/08 END
                cellNode3.AppendChild(dataNode3);
                rowNodeData.AppendChild(cellNode3);

                // Cell4ノード
                XmlElement cellNode4 = domDoc.CreateElement("Cell");
                cellNode4.SetAttribute("Index", ns_ss, (4).ToString());
                // Data4ノード
                XmlElement dataNode4 = domDoc.CreateElement("Data");
                dataNode4.SetAttribute("Type", ns_ss, "String");
                // MODIFY 2022/05/12 BEGIN XML出力サマリ
                //if (!string.IsNullOrWhiteSpace(sumTableData[i].CHANGED_PART) && sumTableData[i].CHANGED_PART.Length > 5)
                //{
                //    if (string.IsNullOrEmpty(sumTableData[i].CHANGED_PART.Substring(0, 5)))
                //    {
                //        dataNode4.InnerText = "";
                //    }
                //    else
                //    {
                //        dataNode4.InnerText = sumTableData[i].CHANGED_PART.Substring(0, 5);
                //    }
                //}
                //else
                //{
                //    if (string.IsNullOrEmpty(sumTableData[i].CHANGED_PART))
                //    {
                //        dataNode4.InnerText = "";
                //    }
                //    else
                //    {
                //        dataNode4.InnerText = sumTableData[i].CHANGED_PART;
                //    }
                //}
                dataNode4.InnerText = sumTableData[i].CHANGED_PART;
                // MODIFY 2022/05/12 END
                cellNode4.AppendChild(dataNode4);
                rowNodeData.AppendChild(cellNode4);

                // Cell5ノード
                XmlElement cellNode5 = domDoc.CreateElement("Cell");
                cellNode5.SetAttribute("Index", ns_ss, (5).ToString());
                // Data5ノード
                XmlElement dataNode5 = domDoc.CreateElement("Data");
                dataNode5.SetAttribute("Type", ns_ss, "String");
                dataNode5.InnerText = sumTableData[i].PART_NM;
                cellNode5.AppendChild(dataNode5);
                rowNodeData.AppendChild(cellNode5);

                // Cell6ノード
                XmlElement cellNode6 = domDoc.CreateElement("Cell");
                cellNode6.SetAttribute("Index", ns_ss, (6).ToString());
                // Data6ノード
                XmlElement dataNode6 = domDoc.CreateElement("Data");
                dataNode6.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].PRODUCTIO))
                {
                    dataNode6.InnerText = "";
                }
                else
                {
                    dataNode6.InnerText = sumTableData[i].PRODUCTIO;
                }
                cellNode6.AppendChild(dataNode6);
                rowNodeData.AppendChild(cellNode6);

                // Cell7ノード
                XmlElement cellNode7 = domDoc.CreateElement("Cell");
                cellNode7.SetAttribute("Index", ns_ss, (7).ToString());
                // Data7ノード
                XmlElement dataNode7 = domDoc.CreateElement("Data");
                dataNode7.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].MODEL_CODE))
                {
                    dataNode7.InnerText = "";
                }
                else
                {
                    dataNode7.InnerText = sumTableData[i].MODEL_CODE;
                }
                cellNode7.AppendChild(dataNode7);
                rowNodeData.AppendChild(cellNode7);

                // Cell8ノード
                XmlElement cellNode8 = domDoc.CreateElement("Cell");
                cellNode8.SetAttribute("Index", ns_ss, (8).ToString());
                // Data8ノード
                XmlElement dataNode8 = domDoc.CreateElement("Data");
                dataNode8.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].DESTINATION))
                {
                    dataNode8.InnerText = "";
                }
                else
                {
                    dataNode8.InnerText = sumTableData[i].DESTINATION;
                }
                cellNode8.AppendChild(dataNode8);
                rowNodeData.AppendChild(cellNode8);

                // Cell9ノード
                XmlElement cellNode9 = domDoc.CreateElement("Cell");
                cellNode9.SetAttribute("Index", ns_ss, (9).ToString());
                // Data9ノード
                XmlElement dataNode9 = domDoc.CreateElement("Data");
                dataNode9.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].GRADE))
                {
                    dataNode9.InnerText = "";
                }
                else
                {
                    dataNode9.InnerText = sumTableData[i].GRADE;
                }
                cellNode9.AppendChild(dataNode9);
                rowNodeData.AppendChild(cellNode9);

                // Cell10ノード
                XmlElement cellNode10 = domDoc.CreateElement("Cell");
                cellNode10.SetAttribute("Index", ns_ss, (10).ToString());
                // Data10ノード
                XmlElement dataNode10 = domDoc.CreateElement("Data");
                dataNode10.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].EQUIPMENT))
                {
                    dataNode10.InnerText = "";
                }
                else
                {
                    dataNode10.InnerText = sumTableData[i].EQUIPMENT;
                }
                cellNode10.AppendChild(dataNode10);
                rowNodeData.AppendChild(cellNode10);

                // Cell11ノード
                XmlElement cellNode11 = domDoc.CreateElement("Cell");
                cellNode11.SetAttribute("Index", ns_ss, (11).ToString());
                // Data11ノード
                XmlElement dataNode11 = domDoc.CreateElement("Data");
                dataNode11.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].ITEM))
                {
                    dataNode11.InnerText = "";
                }
                else
                {
                    dataNode11.InnerText = sumTableData[i].ITEM;
                }
                cellNode11.AppendChild(dataNode11);
                rowNodeData.AppendChild(cellNode11);

                // Cell12ノード
                XmlElement cellNode12 = domDoc.CreateElement("Cell");
                cellNode12.SetAttribute("Index", ns_ss, (12).ToString());
                // Data12ノード
                XmlElement dataNode12 = domDoc.CreateElement("Data");
                dataNode12.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(sumTableData[i].REMARK))
                {
                    dataNode12.InnerText = "";
                }
                else
                {
                    dataNode12.InnerText = sumTableData[i].REMARK;
                }
                cellNode12.AppendChild(dataNode12);
                rowNodeData.AppendChild(cellNode12);

                int count = 13;
                foreach (var item in summaryTableHeader)
                {
                    bool flag = false;
                    for (var k = 0; k < sumTableData[i].ColorHesPartItems.Count; k++)
                    {
                        //重複列IDを調べて値を設定
                        bool deleteColorIdFlag = false;
                        for (var m = 0; m < sameColumId.Count; m++)
                        {
                            if (item.Value.INT_COLOR_ID == sameColumId[m].INT_COLOR_ID)
                            {
                                for (var n = 0; n < sameColumId[m].INT_COLOR_ID_list.Count; n++)
                                {
                                    if (sumTableData[i].ColorHesPartItems[k].INT_COLOR_ID == sameColumId[m].INT_COLOR_ID_list[n])
                                    {
                                        deleteColorIdFlag = true;
                                    }
                                }
                            }
                        }

                        if (sumTableData[i].ColorHesPartItems[k].INT_COLOR_ID == item.Value.INT_COLOR_ID || deleteColorIdFlag)
                        {
                            XmlElement cellNode = domDoc.CreateElement("Cell");
                            cellNode.SetAttribute("Index", ns_ss, (count).ToString());

                            XmlElement dataNode = domDoc.CreateElement("Data");
                            dataNode.SetAttribute("Type", ns_ss, "String");
                            if (string.IsNullOrEmpty(sumTableData[i].ColorHesPartItems[k].COLOR_HES))
                            {
                                if (sumTableData[i].CHILD_UNMATCH_FLAG == 1)
                                {
                                    dataNode.InnerText = "-";
                                }
                                else
                                {
                                    dataNode.InnerText = "";
                                }

                            }
                            else
                            {
                                dataNode.InnerText = sumTableData[i].ColorHesPartItems[k].COLOR_HES;
                            }
                            cellNode.AppendChild(dataNode);
                            rowNodeData.AppendChild(cellNode);
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        XmlElement cellNode = domDoc.CreateElement("Cell");
                        cellNode.SetAttribute("Index", ns_ss, (count).ToString());

                        XmlElement dataNode = domDoc.CreateElement("Data");
                        dataNode.SetAttribute("Type", ns_ss, "String");
                        if (sumTableData[i].CHILD_UNMATCH_FLAG == 1)
                        {
                            dataNode.InnerText = "-";
                        }
                        else
                        {
                            dataNode.InnerText = "";
                        }
                        cellNode.AppendChild(dataNode);
                        rowNodeData.AppendChild(cellNode);
                    }
                    count++;
                }

                tableNode.AppendChild(rowNodeData);

            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "Int Summary");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        // [SUMMARY]ヘッダの集合処理
        private void CreateWorksheetIntSummaryHead(Dictionary<string, GroupIntColorTypeResponse> sumTableHeader, List<GroupIntColorTypeResponseSum> sameColumId)
        {
            Dictionary<string, GroupIntColorTypeResponse> SUMMARYTableHeader = new Dictionary<string, GroupIntColorTypeResponse>();
            SUMMARYTableHeader = sumTableHeader;
            foreach (var item in sumTableHeader)
            {
                GroupIntColorTypeResponseSum groupIntColorTypeResponseSum = new GroupIntColorTypeResponseSum();
                groupIntColorTypeResponseSum.INT_COLOR_ID_list = new List<long?>();
                //IDによる重複値の除去
                foreach (var itemNew in SUMMARYTableHeader)
                {
                    if (!item.Key.Equals(itemNew.Key))
                    {
                        if ((itemNew.Value.INT_COLOR_TYPE.Equals(item.Value.INT_COLOR_TYPE)) && (itemNew.Value.SYMBOL.Equals(item.Value.SYMBOL)))
                        {
                            groupIntColorTypeResponseSum.INT_COLOR_ID = item.Value.INT_COLOR_ID;
                            groupIntColorTypeResponseSum.INT_COLOR_ID_list.Add(itemNew.Value.INT_COLOR_ID);
                            SUMMARYTableHeader.Remove(itemNew.Key);
                            sumTableHeader.Remove(itemNew.Key);
                        }
                    }
                }
                if (groupIntColorTypeResponseSum.INT_COLOR_ID != null)
                {
                    sameColumId.Add(groupIntColorTypeResponseSum);
                }

            }
        }

        // [SUMMARY]表のデータ集合処理
        private void CreateWorksheetIntSummaryData(List<SummaryItemsResponse> sumTableData)
        {
            List<SummaryItemsResponse> sumTableDataCopy = sumTableData;
            for (int i = 0; i < sumTableData.Count; i++)
            {
                List<string> samCHKMEMO = new List<string>();
                for (int j = 0; j < sumTableDataCopy.Count; j++)
                {
                    if (j != i)
                    {
                        if ((!string.IsNullOrWhiteSpace(sumTableData[i].CHANGED_PART) && sumTableData[i].CHANGED_PART.Length > 5)
                            && (!string.IsNullOrWhiteSpace(sumTableDataCopy[j].CHANGED_PART) && sumTableDataCopy[j].CHANGED_PART.Length > 5))
                        {
                            // 固定値の比較
                            if ((sumTableData[i].part_group.Equals(sumTableDataCopy[j].part_group))
                            && (sumTableData[i].SUB_SECTION.Equals(sumTableDataCopy[j].SUB_SECTION))
                            && ((sumTableData[i].CHANGED_PART.Substring(0, 5)).Equals((sumTableDataCopy[j].CHANGED_PART.Substring(0, 5))))
                            && (sumTableData[i].PART_NM.Equals(sumTableDataCopy[j].PART_NM))
                            )
                            {
                                // ダイナミックな列の値の比較
                                List<ColorHesItemsResponse> List1 = sumTableData[i].ColorHesPartItems;
                                List<ColorHesItemsResponse> List2 = sumTableDataCopy[j].ColorHesPartItems;
                                int sameDataCount = 0;
                                if (List1.Count == List2.Count)
                                {
                                    for (var k = 0; k < List1.Count; k++)
                                    {
                                        if ((List1[k].INT_COLOR_ID == List2[k].INT_COLOR_ID) && (List1[k].COLOR_HES.Equals(List2[k].COLOR_HES)))
                                        {
                                            sameDataCount++;
                                        }
                                    }

                                    if (sameDataCount == List1.Count)
                                    {
                                        if (!(NullConvertToString(sumTableData[i].PRODUCTIO).Equals(NullConvertToString(sumTableDataCopy[j].PRODUCTIO))))
                                        {
                                            sumTableData[i].PRODUCTIO = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].MODEL_CODE).Equals(NullConvertToString(sumTableDataCopy[j].MODEL_CODE)))
                                        {
                                            sumTableData[i].MODEL_CODE = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].DESTINATION).Equals(NullConvertToString(sumTableDataCopy[j].DESTINATION)))
                                        {
                                            sumTableData[i].DESTINATION = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].GRADE).Equals(NullConvertToString(sumTableDataCopy[j].GRADE)))
                                        {
                                            sumTableData[i].GRADE = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].EQUIPMENT).Equals(NullConvertToString(sumTableDataCopy[j].EQUIPMENT)))
                                        {
                                            sumTableData[i].EQUIPMENT = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].ITEM).Equals(NullConvertToString(sumTableDataCopy[j].ITEM)))
                                        {
                                            sumTableData[i].ITEM = "*";
                                        }

                                        if (!NullConvertToString(sumTableData[i].REMARK).Equals(NullConvertToString(sumTableDataCopy[j].REMARK)))
                                        {
                                            sumTableData[i].REMARK = "*";
                                        }

                                        // HighLight設定
                                        List<string> highLight = new List<string>();
                                        highLight = sumTableData[i].CROSS_HIGHLIGHTS;

                                        foreach (var hight in sumTableDataCopy[j].CROSS_HIGHLIGHTS)
                                        {
                                            highLight.Add(hight);
                                        }

                                        // 重複データ除去
                                        highLight = highLight.Where((x, y) => highLight.FindIndex(z => z == x) == y).ToList();
                                        sumTableData[i].CROSS_HIGHLIGHTS = highLight;

                                        sumTableDataCopy.Remove(sumTableDataCopy[j]);
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 固定値の比較
                            if ((NullConvertToString(sumTableData[i].part_group).Equals(sumTableDataCopy[j].part_group))
                            && (NullConvertToString(sumTableData[i].SECTION).Equals(sumTableDataCopy[j].SECTION))
                            && (NullConvertToString(sumTableData[i].SUB_SECTION).Equals(sumTableDataCopy[j].SUB_SECTION))
                            && (NullConvertToString(sumTableData[i].CHANGED_PART).Equals(sumTableDataCopy[j].CHANGED_PART))
                            && (NullConvertToString(sumTableData[i].PART_NM).Equals(sumTableDataCopy[j].PART_NM))
                            )
                            {
                                // ダイナミックな列の値の比較
                                List<ColorHesItemsResponse> List1 = sumTableData[i].ColorHesPartItems;
                                List<ColorHesItemsResponse> List2 = sumTableDataCopy[j].ColorHesPartItems;
                                int sameDataCount = 0;
                                if (List1.Count == List2.Count)
                                {
                                    for (var k = 0; k < List1.Count; k++)
                                    {
                                        if ((List1[k].INT_COLOR_ID == List2[k].INT_COLOR_ID) && (List1[k].COLOR_HES.Equals(List2[k].COLOR_HES)))
                                        {
                                            sameDataCount++;
                                        }
                                    }

                                    if (sameDataCount == List1.Count)
                                    {
                                        if (!sumTableData[i].PRODUCTIO.Equals(sumTableDataCopy[j].PRODUCTIO))
                                        {
                                            sumTableData[i].PRODUCTIO = "*";
                                        }

                                        if (!sumTableData[i].MODEL_CODE.Equals(sumTableDataCopy[j].MODEL_CODE))
                                        {
                                            sumTableData[i].MODEL_CODE = "*";
                                        }

                                        if (!sumTableData[i].DESTINATION.Equals(sumTableDataCopy[j].DESTINATION))
                                        {
                                            sumTableData[i].DESTINATION = "*";
                                        }

                                        if (!sumTableData[i].GRADE.Equals(sumTableDataCopy[j].GRADE))
                                        {
                                            sumTableData[i].GRADE = "*";
                                        }

                                        if (!sumTableData[i].EQUIPMENT.Equals(sumTableDataCopy[j].EQUIPMENT))
                                        {
                                            sumTableData[i].EQUIPMENT = "*";
                                        }

                                        if (!sumTableData[i].ITEM.Equals(sumTableDataCopy[j].ITEM))
                                        {
                                            sumTableData[i].ITEM = "*";
                                        }

                                        if (!sumTableData[i].REMARK.Equals(sumTableDataCopy[j].REMARK))
                                        {
                                            sumTableData[i].REMARK = "*";
                                        }

                                        // HighLight設定
                                        List<string> highLight = new List<string>();
                                        highLight = sumTableData[i].CROSS_HIGHLIGHTS;

                                        foreach (var hight in sumTableDataCopy[j].CROSS_HIGHLIGHTS)
                                        {
                                            highLight.Add(hight);
                                        }

                                        // 重複データ除去
                                        highLight = highLight.Where((x, y) => highLight.FindIndex(z => z == x) == y).ToList();
                                        sumTableData[i].CROSS_HIGHLIGHTS = highLight;

                                        sumTableDataCopy.Remove(sumTableDataCopy[j]);
                                    }
                                }

                            }
                        }
                    }
                }
            }
        }

        // カラーチャート情報取得
        [NonAction]
        public IDictionary<string, object> GetColorChartItems(ColorChartGetRequest request)
        {

            string sql = "SELECT                                 " +
                        "    MODIFIED_BY                         " +
                        "    , COLOR_DPM_NAME                    " +
                        "    , COLOR_DPM_REVSION                 " +
                        "    , CHANGED_PART                      " +
                        "    , PART_NM                           " +
                        "FROM                                    " +
                        "    T_COLOR_CHART                       " +
                        "WHERE                                   " +
                        "    COLOR_CHART_ID = '" + request.color_chart_id + "'";

            DbUtil<ColorChartGetResponse> dbUtil = new DbUtil<ColorChartGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        // ColorList情報取得(EXT)
        [NonAction]
        public IDictionary<string, object> GetExtColorListItems(ColorChartGetRequest request)
        {

            // Color情報検索
            string sql = "SELECT                                                               " +
                         "  TSPCIL.COLOR_HES                                                   " +
                         "  , TSPCIL.COLOR_NAME                                                " +
                         "FROM                                                                 " +
                         "  T_COLOR_CHART_PART_RLT TCCPR                                       " +
                         " INNER JOIN T_EXT_COMMU_PLACE_DIST TECPD                             " +
                         "    ON TECPD.PART_APPLY_RELATION_ID = TCCPR.PART_APPLY_RELATION_ID   " +
                         // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                         "    AND TECPD.UNAPPLIED_FLAG = '0 '                                  " +
                         // ADD 2022/05/17 END
                         " INNER JOIN T_STYLING_PART_COLOR_INST_LIST TSPCIL                    " +
                         "    ON TSPCIL.COLOR_INST_ID = TECPD.COLOR_INST_ID                    " +
                         "    AND TSPCIL.CMF_STYLING_PART_ID = TECPD.CMF_STYLING_PART_ID       " +
                         "WHERE                                                                " +
                         "    TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'";

            DbUtil<ColorListInfoGetResponse> dbUtil = new DbUtil<ColorListInfoGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        // ColorList情報取得(INT)
        [NonAction]
        public IDictionary<string, object> GetIntColorListItems(ColorChartGetRequest request)
        {

            // Color情報検索
            string sql = "SELECT                                                               " +
                         "  TSPCIL.COLOR_HES                                                   " +
                         "  , TSPCIL.COLOR_NAME                                                " +
                         "FROM                                                                 " +
                         "  T_COLOR_CHART_PART_RLT TCCPR                                       " +
                         "  INNER JOIN T_INT_COMMU_PLACE_DIST TICPD                            " +
                         "    ON TICPD.PART_APPLY_RELATION_ID = TCCPR.PART_APPLY_RELATION_ID   " +
                         // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                         "    AND TICPD.UNAPPLIED_FLAG = '0'                                   " +
                         // ADD 2022/05/17 END
                         "  INNER JOIN T_STYLING_PART_COLOR_INST_LIST TSPCIL                   " +
                         "    ON TSPCIL.COLOR_INST_ID = TICPD.COLOR_INST_ID                    " +
                         "WHERE                                                                " +
                         "    TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'";

            DbUtil<ColorListInfoGetResponse> dbUtil = new DbUtil<ColorListInfoGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        // BodyColor情報取得
        [NonAction]
        public ByBodyColorInfo GetColorByBodyColorItems(ColorChartGetRequest request)
        {

            // 部品、部品リレーション情報検索
            string sql = "SELECT DISTINCT                                       " +
                        "   TCCPR.PART_APPLY_RELATION_ID                        " +
                        "   , TCCPR.PRODUCTIO                                   " +
                        "   , TCCPR.MODEL_CODE                                  " +
                        "   , TCCPR.DESTINATION                                 " +
                        "   , TCCPR.GRADE                                       " +
                        "   , TCCPR.EQUIPMENT                                   " +
                        "   , TCCPR.ITEM                                        " +
                        "   , TCCPR.REMARK                                      " +
                        "   , TCCPR.CROSS_HIGHLIGHTS                            " +
                        // ADD 2022/05/12 BEGIN BEGIN XML出力サマリ
                        "   , TCCPR.SUMMARY_REMARK                              " +
                        "   , TCCPR.RL_PAIR_NO                                  " +
                        "   , TCCPR.RL_FLAG                                     " +
                        "   , TCCPR.SUMMARY_NO                                  " +
                        "   , TCCPR.SUMMARY_GROUP                               " +
                        // ADD 2022/05/12 END
                        "   , TCCPR.UNMATCH_FLG                                 " +
                        "   , TP.BASE_PART                                      " +
                        "   , TP.CHANGED_PART                                   " +
                        "   , TP.PART_NM                                        " +
                        "   , TP.LVL                                            " +
                        "   , TP.SECTION                                        " +
                        "   , TP.SUB_SECTION                                    " +
                        "   , TP.L1_PART_ID                                     " +
                        "   , TP.PART_ID                                         " +
                        "   , TP2.CHANGED_PART AS L1_PART_NO                    " +
                        "   , TP2.BASE_PART AS L1_BASE_PART                     " +
                        "   , TP3.CHANGED_PART AS PARENT_PART_NO                " +
                        "   , TCCPR2.UNMATCH_FLG AS L1_UNMATCH_FLG              " +
                        "   , TCCPR3.UNMATCH_FLG AS PARENT_UNMATCH_FLG          " +
                        "   , TCCPR.CHILD_UNMATCH_FLAG                          " +
                        " FROM                                                  " +
                        "   T_COLOR_CHART_PART_RLT TCCPR                        " +
                        "   INNER JOIN T_PART TP                                " +
                        "     ON TP.PART_ID = TCCPR.PART_ID                     " +
                        "   LEFT JOIN T_PART TP2                                " +
                        "     ON TP2.PART_ID = TP.L1_PART_ID                    " +
                        "   LEFT JOIN T_PART TP3                                " +
                        "     ON TP3.PART_ID = TP.PARENT_PART_ID                " +
                        "   LEFT JOIN T_COLOR_CHART_PART_RLT TCCPR2             " +
                        "     ON TP2.PART_ID = TCCPR2.PART_ID                   " +
                        "   LEFT JOIN T_COLOR_CHART_PART_RLT TCCPR3             " +
                        "     ON TP3.PART_ID = TCCPR3.PART_ID                   " +
                        " WHERE                                                 " +
                        "   TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'" +
                        "   AND TP2.CHANGED_PART IS NOT NULL                    " +
                        "   AND (                                               " +
                        "     TP.LVL = 1                                        " +
                        "     OR(TP.LVL > 1 AND TP3.CHANGED_PART IS NOT NULL)   " +
                        "   )                                                   " +
                        //" ORDER BY L1_PART_NO, TP.LVL";
                        //2022.01.26 modify 
                        " ORDER BY                                              " +
                        "   TP.SECTION                                          " +
                        "   , TP.SUB_SECTION                                    " +
                        "   , TP.L1_PART_ID                                     " +
                        "   , TP.PART_ID                                        " +
                        "   , TCCPR.PART_APPLY_RELATION_ID                      ";


            DbUtil<ByBodyColorInfoGetResponse> dbUtil = new DbUtil<ByBodyColorInfoGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            ByBodyColorInfo byBodyColor = new ByBodyColorInfo();

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");

                // 部品、部品リレーション情報設定
                List<ByBodyColorInfoGetResponse> sel_objs = (List<ByBodyColorInfoGetResponse>)map["resultList"];
                byBodyColor.partInfo = sel_objs;
            }
            else
            {
                Response.StatusCode = 400;
            }

            // Color情報検索
            string color_sql = " SELECT                                                           " +
                               "  TTECI.EXT_COLOR1_NAME                                           " +
                               " , TTECI.EXT_COLOR2_NAME                                          " +
                               " , TEPD.EXT_COLOR_ID                                              " +
                               " , TEPD.CMF_STYLING_PART_ID                                       " +
                               " , TEPD.COLOR_INST_ID                                             " +
                               " , TEPD.PART_APPLY_RELATION_ID                                    " +
                               " FROM                                                             " +
                               "   T_COLOR_CHART_PART_RLT TCCPR                                   " +
                               " INNER JOIN T_EXT_COMMU_PLACE_DIST TEPD                           " +
                               "   ON TEPD.PART_APPLY_RELATION_ID = TCCPR.PART_APPLY_RELATION_ID  " +
                               // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                               "   AND TEPD.UNAPPLIED_FLAG = '0'                                  " +
                               // ADD 2022/05/17 END
                               " INNER JOIN T_TC_EXT_COLOR_INFO TTECI                             " +
                               "   ON TTECI.EXT_COLOR_ID = TEPD.EXT_COLOR_ID                      " +
                               " WHERE                                                            " +
                               "    TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'";

            DbUtil<BodyColor> color_dbUtil = new DbUtil<BodyColor>();

            IDictionary<string, object> color_map = color_dbUtil.SelectQuery(color_sql, conn, tr);

            if (color_map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                color_map.Remove("result");

                // COLOR情報設定
                List<BodyColor> sel_objs = (List<BodyColor>)color_map["resultList"];
                byBodyColor.bodyColor = sel_objs;

                List<ByBodyColorHes> byBodyColorHesList = new List<ByBodyColorHes>();

                foreach (var sel_obj in sel_objs)
                {

                    string hes_sql = " SELECT DISTINCT                                               " +
                                     "   TSPCIL.COLOR_HES                                            " +
                                     "   , TEPD.EXT_COLOR_ID                                         " +
                                     "   , TSPCIL.COLOR_INST_ID                                      " +
                                     "   , TSPCIL.CMF_STYLING_PART_ID                                " +
                                     " FROM                                                          " +
                                     "   T_STYLING_PART_COLOR_INST_LIST TSPCIL                       " +
                                     "   INNER JOIN T_EXT_COMMU_PLACE_DIST TEPD                      " +
                                     "     ON TEPD.CMF_STYLING_PART_ID = TSPCIL.CMF_STYLING_PART_ID  " +
                                     // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                                     "     AND TEPD.UNAPPLIED_FLAG = '0'                             " +
                                     // ADD 2022/05/17 END
                                     " WHERE                                                         " +
                                     "   TSPCIL.COLOR_INST_ID = '" + sel_obj.color_inst_id + "'" +
                                     "   AND TSPCIL.CMF_STYLING_PART_ID = '" + sel_obj.cmf_styling_part_id + "'";

                    DbUtil<ByBodyColorHes> hes_dbUtil = new DbUtil<ByBodyColorHes>();

                    IDictionary<string, object> hes_map = hes_dbUtil.SelectQuery(hes_sql, conn, tr);

                    if (hes_map["result"].Equals("OK"))
                    {
                        Response.StatusCode = 200;
                        hes_map.Remove("result");

                        List<ByBodyColorHes> hes_objs = (List<ByBodyColorHes>)hes_map["resultList"];

                        foreach (var hes_obj in hes_objs)
                        {
                            ByBodyColorHes hesInfo = new ByBodyColorHes();
                            hesInfo.ext_color_id = hes_obj.ext_color_id;
                            hesInfo.color_hes = hes_obj.color_hes;
                            hesInfo.color_inst_id = hes_obj.color_inst_id;
                            hesInfo.cmf_styling_part_id = hes_obj.cmf_styling_part_id;

                            byBodyColorHesList.Add(hesInfo);
                        }
                    }
                    else
                    {
                        Response.StatusCode = 400;
                    }

                }

                byBodyColor.colorHes = byBodyColorHesList;

            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return byBodyColor;
        }

        // Worksheet情報取得
        [NonAction]
        public IDictionary<string, object> GetWorksheetInt(GroupChartGetRequest request)
        {

            // Group sheet情報検索
            string sql = "SELECT                                               " +
                         "  PART_GROUP                                         " +
                         "FROM                                                 " +
                         "  T_COLOR_CHART_PART_RLT                             " +
                         "WHERE                                                " +
                         "    COLOR_CHART_ID = '" + request.color_chart_id + "'" +
                         "GROUP BY PART_GROUP                                  " +
                         "ORDER BY PART_GROUP ASC";

            DbUtil<GroupPartGetResponse> dbUtil = new DbUtil<GroupPartGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        //INT Color Type情報分記載
        [NonAction]
        public IDictionary<string, object> GetIntColorTypeItems(GroupChartGetRequest request, string group, long tc_id)
        {
            string sql = "SELECT                                               " +
                         "  TICI.INT_COLOR_TYPE                                " +
                         " , TICI.SYMBOL                                       " +
                         " , TICI.INT_COLOR_ID                                 " +
                         "FROM                                                 " +
                         "  T_TC_INT_COLOR_INFO TICI                           " +
                         "  INNER JOIN T_COLOR_CHART_TC_INT_COLOR CCTIC        " +
                         "    ON TICI.INT_COLOR_ID = CCTIC.INT_COLOR_ID        " +
                         "    AND CCTIC.COLOR_CHART_TC_ID = '" + tc_id + "'    " +
                         "WHERE                                                " +
                         "  TICI.INT_COLOR_ID IN (                             " +
                         "      SELECT                                         " +
                         "          distinct INT_COLOR_ID                      " +
                         "      FROM                                           " +
                         "          T_INT_COMMU_PLACE_DIST                     " +
                         "      WHERE                                          " +
                         "          PART_APPLY_RELATION_ID                     " +
                         "          IN(                                        " +
                         "              SELECT                                 " +
                         "                  PART_APPLY_RELATION_ID             " +
                         "              FROM                                   " +
                         "                  T_COLOR_CHART_PART_RLT             " +
                         "              WHERE                                  " +
                         "                  COLOR_CHART_ID = '" + request.color_chart_id + "')" +
                         // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                         "                  AND T_INT_COMMU_PLACE_DIST.UNAPPLIED_FLAG = '0'   " +
                         // ADD 2022/05/17 END
                         "      GROUP BY INT_COLOR_ID                          " +
                         ")                                                    " +
                         "ORDER BY CCTIC.DISPLAY_NO, TICI.SYMBOL";

            DbUtil<GroupIntColorTypeResponse> dbUtil = new DbUtil<GroupIntColorTypeResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        //GetGroupPartItems
        [NonAction]
        public IDictionary<string, object> GetGroupPartItems(GroupChartGetRequest request, string group)
        {
            string sql = "SELECT                                                    " +
                        "   TCCPR.PART_APPLY_RELATION_ID                            " +
                        "   , TCCPR.PRODUCTIO                                       " +
                        "   , TCCPR.MODEL_CODE                                      " +
                        "   , TCCPR.DESTINATION                                     " +
                        "   , TCCPR.GRADE                                           " +
                        "   , TCCPR.EQUIPMENT                                       " +
                        "   , TCCPR.ITEM                                            " +
                        "   , TCCPR.REMARK                                          " +
                        "   , TCCPR.CROSS_HIGHLIGHTS                                " +
                        // ADD 2022/05/12 END
                        "   , TCCPR.SUMMARY_REMARK                                  " +
                        "   , TCCPR.RL_PAIR_NO                                      " +
                        "   , TCCPR.RL_FLAG                                         " +
                        "   , TCCPR.SUMMARY_NO                                      " +
                        "   , TCCPR.SUMMARY_GROUP                                   " +
                        // ADD 2022/05/12 END
                        "   , TCCPR.UNMATCH_FLG                                     " +
                        "   , TCCPR.CHKMEMO                                         " +
                        "   , TP.BASE_PART                                          " +
                        "   , TP.CHANGED_PART                                       " +
                        "   , TP.PART_NM                                            " +
                        "   , TP.LVL                                                " +
                        "   , TP.SECTION                                            " +
                        "   , TP.SUB_SECTION                                        " +
                        "   , TP2.CHANGED_PART AS L1_CHANGED_PART                   " +
                        "   , TP2.BASE_PART AS L1_BASE_PART                         " +
                        "   , TP3.CHANGED_PART AS PARENT_CHANGED_PART               " +
                        "   , TP2.UNMATCH_FLG AS L1_UNMATCH_FLG                     " +
                        "   , TP3.UNMATCH_FLG AS PARENT_UNMATCH_FLG                 " +
                        "   , TCCPR.CHILD_UNMATCH_FLAG                              " +
                        " FROM                                                      " +
                        "   T_COLOR_CHART_PART_RLT TCCPR                            " +
                        "   INNER JOIN T_PART TP                                    " +
                        "     ON TP.PART_ID = TCCPR.PART_ID                         " +
                        "   LEFT JOIN (                                             " +
                        "     SELECT DISTINCT                                       " +
                        "       T_P.PART_ID                                         " +
                        "       , T_P.CHANGED_PART                                  " +
                        "       , T_P.BASE_PART                                     " +
                        "       , T_CCPR.UNMATCH_FLG                                " +
                        "     FROM                                                  " +
                        "       T_PART T_P                                          " +
                        "       INNER JOIN T_COLOR_CHART_PART_RLT T_CCPR            " +
                        "         ON T_P.PART_ID = T_CCPR.PART_ID                   " +
                        "   ) TP2                                                   " +
                        "     ON TP.L1_PART_ID = TP2.PART_ID                        " +
                        "   LEFT JOIN (                                             " +
                        "     SELECT DISTINCT                                       " +
                        "       T_P.PART_ID                                         " +
                        "       , T_P.CHANGED_PART                                  " +
                        "       , T_P.BASE_PART                                     " +
                        "       , T_CCPR.UNMATCH_FLG                                " +
                        "     FROM                                                  " +
                        "       T_PART T_P                                          " +
                        "       INNER JOIN T_COLOR_CHART_PART_RLT T_CCPR            " +
                        "         ON T_P.PART_ID = T_CCPR.PART_ID                   " +
                        "   ) TP3                                                   " +
                        "     ON TP.PARENT_PART_ID = TP3.PART_ID                    " +
                        " WHERE                                                     " +
                        "   TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "' " +
                        "   AND TCCPR.PART_GROUP = '" + group + "'                  " +
                        "   AND TP2.CHANGED_PART IS NOT NULL                        " +
                        "   AND (                                                   " +
                        "     TP.LVL = 1                                            " +
                        "     OR (TP.LVL > 1 AND TP3.CHANGED_PART IS NOT NULL)      " +
                        "   )                                                       " +
                        //"ORDER BY TP2.CHANGED_PART, TP.LVL";
                        //2022.01.26 modify 
                        " ORDER BY                                                  " +
                        "   TP.SECTION                                              " +
                        "   , TP.SUB_SECTION                                        " +
                        "   , TP.L1_PART_ID                                         " +
                        "   , TP.PART_ID                                            " +
                        "   , TCCPR.PART_APPLY_RELATION_ID                          ";

            DbUtil<GroupPartItemsResponse> dbUtil = new DbUtil<GroupPartItemsResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        //INT_COLOR_ID(列) -> COLOR_HES
        [NonAction]
        public IDictionary<string, object> GetColorHesPartItems(string part_apply_relation_id)
        {

            string sql = "SELECT                                                            " +
                         "  TSPCIL.INT_COLOR_ID                                             " +
                         "  , TSPCIL.COLOR_HES                                              " +
                         "FROM                                                              " +
                         "  T_STYLING_PART_COLOR_INST_LIST TSPCIL                           " +
                         "WHERE                                                             " +
                         "  EXISTS (                                                        " +
                         "    SELECT                                                        " +
                         "      1                                                           " +
                         "    FROM                                                          " +
                         "      T_INT_COMMU_PLACE_DIST TICPD                                " +
                         "    WHERE                                                         " +
                         "      TSPCIL.COLOR_INST_ID = TICPD.COLOR_INST_ID                  " +
                         "      AND TSPCIL.CMF_STYLING_PART_ID = TICPD.CMF_STYLING_PART_ID  " +
                         "      AND TSPCIL.INT_COLOR_ID = TICPD.INT_COLOR_ID                " +
                         // ADD 2022/05/17 BEGIN UNAPPLIEDフラグを追加
                         "      AND TICPD.UNAPPLIED_FLAG = '0'                              " +
                         // ADD 2022/05/17 END
                         "      AND TICPD.PART_APPLY_RELATION_ID = " + part_apply_relation_id +
                         "  ) ";

            DbUtil<ColorHesItemsResponse> dbUtil = new DbUtil<ColorHesItemsResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        [NonAction]
        public bool CheckFolder(string folderPath)
        {
            // フォルダ存在チェック
            DirectoryInfo dInfo = new DirectoryInfo(folderPath);
            if (dInfo.Exists)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        // 文字除く
        private void ReplaceCharacter(string xmpPath)
        {
            // XMLファイル読込
            StreamReader sr = new StreamReader(xmpPath);

            string xmlStr = "";

            while (!sr.EndOfStream)
            {
                string str = sr.ReadLine();
                xmlStr += str + "\n";
            }

            // [amp;除く]
            xmlStr = xmlStr.Replace("&amp;#10;", "&#10;");

            sr.Close();

            // XMLファイル書込
            System.IO.File.WriteAllText(xmpPath, xmlStr, Encoding.UTF8);
        }

        // NullからStringに変換
        private string NullConvertToString(string val)
        {

            if (val == null)
            {
                return "";
            }

            return val;
        }

        // 空欄を「-」に設定
        private string setUnderline(string val)
        {

            if (string.IsNullOrWhiteSpace(val))
            {
                return "-";
            }
            return val;
        }

        // PART NOを変換
        private string ChangePartNo(string partNo)
        {

            string newPartNo = "";

            // 文字列の最後のアンダーバーを除去し、それ以外のアンダーバーは空白に置換
            if (!string.IsNullOrWhiteSpace(partNo))
            {
                newPartNo = partNo.Replace("_", " ").TrimEnd();
            }

            return newPartNo;
        }

        // HESコード形式チェック
        private bool CheckHesCode(string hesCode)
        {
            // HESコード正式形式
            string[] ecopasArray = { "NH", "R", "YR", "Y", "GY", "G", "BG", "B", "PB", "RP" };

            if (!string.IsNullOrEmpty(hesCode))
            {
                string[] colorHesCode = hesCode.Split('-');
                if (colorHesCode.Length > 1)
                {
                    if (ecopasArray.Contains(colorHesCode[0]))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        // HES名前取得
        private string GetHesName(string colorName)
        {
            var colorNameArr = colorName.Split(' ');

            string nameTmp = "";

            for (var i = 1; i < colorNameArr.Length; i++)
            {
                if (i == 1)
                {
                    nameTmp = colorNameArr[i];
                }
                else
                {
                    nameTmp = nameTmp + " " + colorNameArr[i];
                }

            }

            return nameTmp;
        }

        // COLORシート、BY BODY COLORシート分ける処理
        private void SetColorInfo(ByBodyColorInfo colorResults, List<ColorListInfoGetResponse> bodyColorToHes,
            List<ColorTabInfo> byBodyColorList, List<ColorTabInfo> colorList)
        {
            // 外装色ヘッダ数
            long bodyColorCount = colorResults.bodyColor.Count;

            // 外装色ヘッダ重複除去
            List<BodyColorHeader> headerList = new List<BodyColorHeader>();

            // EXT COLOR ID臨時格納リスト
            List<long?> list = new List<long?>();

            // 外装色ヘッダ格納リスト（重複除く前）
            List<BodyColorHeader> headerBefor = new List<BodyColorHeader>();

            // 外装色ヘッダ重複のデータ除去
            for (var i = 0; i < bodyColorCount; i++)
            {
                if (!list.Contains(colorResults.bodyColor[i].ext_color_id))
                {
                    BodyColorHeader header = new BodyColorHeader();
                    header.col_no = i;
                    header.ext_color_id = colorResults.bodyColor[i].ext_color_id;

                    headerBefor.Add(header);
                    list.Add(colorResults.bodyColor[i].ext_color_id);
                }
            }

            for (var i = 0; i < headerBefor.Count; i++)
            {
                BodyColorHeader bodyColorHeader = new BodyColorHeader();
                bodyColorHeader.ext_color_id = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color_id;
                bodyColorHeader.col_no = i + 15;
                bodyColorHeader.ext_color1_name = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color1_name;
                bodyColorHeader.ext_color2_name = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color2_name;

                headerList.Add(bodyColorHeader);

            }

            // 外装色ヘッダ格納リスト（HESコード）
            List<string> colorHeaderList = new List<string>();

            // 外装色ヘッダ格納リスト（HESコード、カラー名）
            List<string> colorHeaderNmList = new List<string>();

            // ヘッダ作成（BodyColor部）
            for (var i = 0; i < headerList.Count; i++)
            {

                string extColor1Code = "";
                string extColor1Name = "";

                string extColor2Code = "";

                if (!string.IsNullOrWhiteSpace(headerList[i].ext_color1_name))
                {

                    var ext_color1_name = headerList[i].ext_color1_name.Split(' ');

                    extColor1Code = ext_color1_name[0];

                    for (var j = 1; j < ext_color1_name.Length; j++)
                    {
                        if (j == 1)
                        {
                            extColor1Name = ext_color1_name[j];
                        }
                        else
                        {
                            extColor1Name = extColor1Name + " " + ext_color1_name[j];
                        }

                    }

                }
                else
                {
                    extColor1Code = "-";
                }

                if (!string.IsNullOrWhiteSpace(headerList[i].ext_color2_name))
                {
                    var ext_color2_name = headerList[i].ext_color2_name.Split(' ');

                    extColor2Code = ext_color2_name[0];
                }

                string extColorTmp = headerList[i].ext_color1_name;
                if (!string.IsNullOrWhiteSpace(extColor2Code))
                {
                    extColorTmp = extColor1Code + "&#10;" + extColor1Name + "&#10;" + "(" + extColor2Code + ")";
                }
                else
                {
                    extColorTmp = extColor1Code + "&#10;" + extColor1Name;

                }

                colorHeaderList.Add(extColor1Code);

                colorHeaderNmList.Add(extColorTmp);

            }

            // ヘッダの数格納
            bodyColorHeaderCount = colorHeaderList.Count;

            // 外装色ヘッダ名リスト格納
            colorHeaderName = colorHeaderNmList;

            // データ部作成
            for (var i = 0; i < colorResults.partInfo.Count; i++)
            {
                ColorTabInfo colorTabInfo = new ColorTabInfo();

                colorTabInfo.color_header = colorHeaderList;

                colorTabInfo.color_header_name = colorHeaderNmList;

                // Cellノード(SECTION)
                colorTabInfo.section = colorResults.partInfo[i].section;

                // Cellノード(SUB-SECTION)
                colorTabInfo.sub_section = colorResults.partInfo[i].sub_section;

                // Cellノード(L1)
                colorTabInfo.l1_part_no = colorResults.partInfo[i].l1_part_no;

                // Cellノード(LVL)
                colorTabInfo.lvl = colorResults.partInfo[i].lvl;

                // Cellノード(PART NO.)
                colorTabInfo.changed_part = colorResults.partInfo[i].changed_part;

                // Cellノード(PARENT PART NO.)
                colorTabInfo.parent_part_no = colorResults.partInfo[i].parent_part_no;

                // Cellノード(PART NAME)
                colorTabInfo.part_name = colorResults.partInfo[i].part_nm;

                // Cellノード(PLANT)
                colorTabInfo.plant = colorResults.partInfo[i].productio;

                // Cellノード(MODEL)
                colorTabInfo.model = colorResults.partInfo[i].model_code;

                // Cellノード(DESTINATION)
                colorTabInfo.destination = colorResults.partInfo[i].destination;

                // Cellノード(GRADE)
                colorTabInfo.grade = colorResults.partInfo[i].grade;

                // Cellノード(FEATURE)
                colorTabInfo.feature = colorResults.partInfo[i].equipment;

                // Cellノード(ITEM)
                colorTabInfo.item = colorResults.partInfo[i].item;

                // Cellノード(REMARKS)
                colorTabInfo.remark = colorResults.partInfo[i].remark;

                // HIGHLIGHT
                colorTabInfo.cross_highlights = colorResults.partInfo[i].cross_highlights;

                // ADD 2022/05/12 BEGIN BEGIN XML出力サマリ
                // Cellノード(SUMMARY_REMARK)
                colorTabInfo.summary_remark = colorResults.partInfo[i].summary_remark;

                // Cellノード(RL_PAIR_NO)
                colorTabInfo.rl_pair_no = colorResults.partInfo[i].rl_pair_no;

                // Cellノード(RL_FLAG)
                colorTabInfo.rl_flag = colorResults.partInfo[i].rl_flag;

                // Cellノード(SUMMARY_NO)
                colorTabInfo.summary_no = colorResults.partInfo[i].summary_no;

                // Cellノード(SUMMARY_GROUP)
                colorTabInfo.summary_group = colorResults.partInfo[i].summary_group;
                // ADD 2022/05/12 END

                colorTabInfo.unmatch_flg = colorResults.partInfo[i].unmatch_flg;
                colorTabInfo.l1_unmatch_flg = colorResults.partInfo[i].l1_unmatch_flg;
                colorTabInfo.parent_unmatch_flg = colorResults.partInfo[i].parent_unmatch_flg;

                // 色指示格納リスト
                List<string> colorTabHesList = new List<string>();

                string color_hes = "";

                // Body Color
                for (var colIdx = 0; colIdx < headerList.Count; colIdx++)
                {
                    color_hes = "";

                    for (int j = 0; j < colorResults.bodyColor.Count; j++)
                    {

                        if (colorResults.bodyColor[j].ext_color_id == headerList[colIdx].ext_color_id)
                        {

                            if (colorResults.partInfo[i].part_apply_relation_id == colorResults.bodyColor[j].part_apply_relation_id)
                            {
                                for (int k = 0; k < colorResults.colorHes.Count; k++)
                                {
                                    if (colorResults.colorHes[k].cmf_styling_part_id == colorResults.bodyColor[j].cmf_styling_part_id &&
                                        colorResults.colorHes[k].color_inst_id == colorResults.bodyColor[j].color_inst_id)
                                    {
                                        // Cellノード(BODY COLOR)
                                        color_hes = colorResults.colorHes[k].color_hes;
                                        break;
                                    }

                                }
                            }
                        }

                    }

                    if (!string.IsNullOrWhiteSpace(color_hes))
                    {
                        // 色指示が「Body Color」の場合、ヘッダの1色目に設定
                        if (color_hes.Equals("Body Color") || color_hes.Equals("BODY COLOR"))
                        {
                            color_hes = headerList[colIdx].ext_color1_name.Split(" ")[0];

                            // 入れ替え後の色指示格納
                            ColorListInfoGetResponse newHes = new ColorListInfoGetResponse();
                            newHes.color_hes = color_hes;
                            newHes.color_name = GetHesName(headerList[colIdx].ext_color1_name);
                            bodyColorToHes.Add(newHes);

                            colorTabHesList.Add(color_hes + "「Body Color」");

                        }
                        // 色指示がHESコード形式ではない場合、空欄に設定
                        else if (!CheckHesCode(color_hes))
                        {
                            colorTabHesList.Add("");
                        }
                        else
                        {
                            colorTabHesList.Add(color_hes);
                        }
                    }
                    else
                    {
                        colorTabHesList.Add("");
                    }

                }

                // 子部品不一致フラグ
                colorTabInfo.child_unmatch_flag = colorResults.partInfo[i].child_unmatch_flag;

                colorTabInfo.color_item = colorTabHesList;

                // Cellノード(L1_BASE PART NO)
                colorTabInfo.l1_base_part_no = colorResults.partInfo[i].l1_base_part;

                // Cellノード(BASE PART NO)
                colorTabInfo.base_part_no = colorResults.partInfo[i].base_part;

                // COLORシート、BY BODY COLORシートに分ける 
                if (CheckColor(colorTabInfo))
                {
                    colorList.Add(colorTabInfo);
                }
                else
                {
                    byBodyColorList.Add(colorTabInfo);
                }
            }
        }

        // COLORチェック処理
        private bool CheckColor(ColorTabInfo colorTabInfo)
        {

            // MODIFY 2022/05/30 BEGIN COLORシートルール
            // Cellノード(COLOR)
            //bool bodyColorFlg = true;

            //bool isEmptyFlg = true;

            // 1行に全ての色指示が未設定であるチェック
            //foreach (var colorItem in colorTabInfo.color_item)
            //{
            //    if (!string.IsNullOrEmpty(colorItem))
            //    {
            //        isEmptyFlg = false;
            //        break;
            //    }
            //}

            //if (!isEmptyFlg)
            //{
            // 色指示は該当列の外装色ヘッダと同じであるかチェック
            //    for (var j = 0; j < colorTabInfo.color_header.Count; j++)
            //    {
            //        if (!string.IsNullOrEmpty(colorTabInfo.color_item[j]))
            //        {
            //            if (!colorTabInfo.color_header[j].Equals(colorTabInfo.color_item[j]))
            //            {
            //                bodyColorFlg = false;
            //                break;
            //            }
            //        }
            //    }
            //}
            //else
            //{
            // 該当行の色指示が全て未設定の場合、ByBodyColor対象とする
            //    return false;

            //}

            //if (bodyColorFlg)
            //{
            // COLOR対象とする
            //    return true;
            //}
            //else
            //{

            //    bool itemFlg = true;

            //    string tmpHes = "";

            // 色情報が複数件存在するの場合、一致するか比較する
            //    if (colorTabInfo.color_item.Count > 1)
            //    {

            // 一番目の空白以外のデータ取得
            //        for (var j = 0; j < colorTabInfo.color_item.Count; j++)
            //        {

            //            if (!string.IsNullOrWhiteSpace(colorTabInfo.color_item[j]))
            //            {
            //                tmpHes = colorTabInfo.color_item[j];
            //                break;
            //            }

            //        }

            // 色情報は1件以上存在
            //        if (!string.IsNullOrWhiteSpace(tmpHes))
            //        {

            //            for (var j = 0; j < colorTabInfo.color_item.Count; j++)
            //            {

            //                if (!string.IsNullOrWhiteSpace(colorTabInfo.color_item[j]))
            //                {
            // 色情報差異ある場合
            //                    if (!colorTabInfo.color_item[j].Equals(tmpHes))
            //                    {
            //                        itemFlg = false;
            //                        break;
            //                    }
            //                }

            //            }
            //        }
            //    }

            // 色情報差異ある場合
            //    if (!itemFlg)
            //    {
            // ByBodyColor対象とする
            //        return false;
            //    }

            //    return true;
            //}


            bool bodyColorFLg = true;
            for (var i = 0; i < colorTabInfo.color_item.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(colorTabInfo.color_item[i]) && !colorTabInfo.color_item[i].Contains("「Body Color」"))
                {
                    bodyColorFLg = false;
                }
                colorTabInfo.color_item[i] = colorTabInfo.color_item[i].Replace("「Body Color」", "");
            }

            if (bodyColorFLg)
            {
                for (var i = 0; i < colorTabInfo.color_item.Count; i++)
                {
                    if (!string.IsNullOrWhiteSpace(colorTabInfo.color_item[i]))
                    {
                        colorTabInfo.color_item[i] = "BODY COLOR";
                    }
                }
                return true;
            }
            else if (colorTabInfo.child_unmatch_flag == 0 || colorTabInfo.child_unmatch_flag == null)
            {
                string colorInfo = "";
                foreach (var colorItemInfo in colorTabInfo.color_item)
                {
                    if (colorInfo == "")
                    {
                        if (colorItemInfo != null || colorItemInfo != "")
                        {
                            colorInfo = colorItemInfo;
                        }
                    }
                    else
                    {
                        if (colorItemInfo != null && colorItemInfo != "" && !colorInfo.Equals(colorItemInfo))
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            else if (colorTabInfo.child_unmatch_flag == 1)
            {
                return true;
            }
            else if (colorTabInfo.child_unmatch_flag == 2)
            {
                return false;
            }
            else
            {
                return true;
            }
            // MODIFY 2022/05/30 END
        }
        // ADD 2022/04/06 BEGIN XMLデータ出力改善(T/Cのフォーマット変更)
        // Body Color列ヘッダを一列表現から二段表記にする
        private string getColor1Name(string beforeExtColorName)
        {
            string extColor1Code = "";
            string extColor1Name = "";

            if (!string.IsNullOrWhiteSpace(beforeExtColorName))
            {

                var ext_color1_name = beforeExtColorName.Split(' ');

                extColor1Code = ext_color1_name[0];

                for (var j = 1; j < ext_color1_name.Length; j++)
                {
                    if (j == 1)
                    {
                        extColor1Name = ext_color1_name[j];
                    }
                    else
                    {
                        extColor1Name = extColor1Name + " " + ext_color1_name[j];
                    }
                }
                return extColor1Code + "&#10;" + extColor1Name;
            }
            else
            {
                return beforeExtColorName;
            }
        }
        // ADD 2022/04/06 END 

        // ADD 2022/05/12 BEGIN XML出力サマリ
        private string GetPartName(List<string> partNameList)
        {
            int index = 0;
            int length = partNameList[0].Length > partNameList[1].Length ? partNameList[1].Length : partNameList[0].Length;
            for (int i = 0; i < length; i++)
            {
                if (partNameList[0].Substring(i, 1) != partNameList[1].Substring(i, 1))
                {
                    index = i;
                }
            }
            return partNameList[0].Substring(0, index) + "R/L" + partNameList[0].Substring(index + 1, length - index - 1);
        }
        // ADD 2022/05/12 END

        // ADD 2022/05/27 BEGIN クロスハイライト複数対応
        private List<string> getHighLight(string strHighLight)
        {
            List<string> highLightList = new List<string>();
            if (strHighLight != null)
            {
                if (strHighLight.Contains(":"))
                {
                    string[] text = strHighLight.Split(":");
                    string[] textList = text[1].Split(",");
                    for (int i = 0; i < textList.Length; i++)
                    {
                        highLightList.Add(text[0].Insert(strHighLight.IndexOf("/"), textList[i]));
                    }
                }
                else
                {
                    highLightList.Add(strHighLight);
                }
            }
            else
            {
                highLightList.Add("");
            }
            return highLightList;
        }
        // ADD 2022/05/27 END

        // UnmatchColor情報取得
        [NonAction]
        public UnmatchInfo GetUnmatchColorByBodyColorItems(ExportDrawingDataGetRequest request)
        {

            // 部品、部品リレーション情報検索
            string sql = " SELECT                                " +
                         "  TCCPR.PART_APPLY_RELATION_ID         " +
                         "  , TCCPR.PRODUCTIO                    " +
                         "  , TCCPR.MODEL_CODE                   " +
                         "  , TCCPR.DESTINATION                  " +
                         "  , TCCPR.GRADE                        " +
                         "  , TCCPR.EQUIPMENT                    " +
                         "  , TCCPR.ITEM                         " +
                         "  , TCCPR.REMARK                       " +
                         "  , TP.BASE_PART                       " +
                         "  , TP.CHANGED_PART                    " +
                         "  , TP.PART_NM                         " +
                         "  , TP.LVL                             " +
                         "  , TP.SECTION                         " +
                         "  , TP.SUB_SECTION                     " +
                         "  , TP2.BASE_PART AS L1_BASE_PART_NO   " +
                         "  , TP2.CHANGED_PART AS L1_PART_NO     " +
                         "  , TP3.CHANGED_PART AS PARENT_PART_NO " +
                         "  , TCCPR.CHILD_UNMATCH_FLAG           " +
                         "  , TCCPR.UNMATCH_FLG                  " +
                         " FROM                                  " +
                         "  T_COLOR_CHART_PART_RLT TCCPR         " +
                         "  INNER JOIN T_PART TP                 " +
                         "    ON TP.PART_ID = TCCPR.PART_ID      " +
                         "  LEFT JOIN T_PART TP2                 " +
                         "    ON TP2.PART_ID = TP.L1_PART_ID     " +
                         "  LEFT JOIN T_PART TP3                 " +
                         "    ON TP3.PART_ID = TP.PARENT_PART_ID " +
                         " WHERE                                 " +
                         "    TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'" +
                         "  AND                                                  " +
                         "    TCCPR.UNMATCH_FLG = 1                              " +
                         "  AND                                                  " +
                         "    TP2.CHANGED_PART IS NOT NULL                       " +
                         "  AND(                                                 " +
                         "    TP.LVL = 1                                         " +
                         "    OR(TP.LVL > 1 AND TP3.CHANGED_PART IS NOT NULL)    " +
                         "  )                                                    " +
                         " ORDER BY TP.SECTION, TP.SUB_SECTION, TP.L1_PART_ID, TP.PART_ID, TCCPR.PART_APPLY_RELATION_ID";

            DbUtil<UnmatchInfoGetResponse> dbUtil = new DbUtil<UnmatchInfoGetResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            UnmatchInfo unmatchColor = new UnmatchInfo();

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");

                // 部品、部品リレーション情報設定
                List<UnmatchInfoGetResponse> sel_objs = (List<UnmatchInfoGetResponse>)map["resultList"];
                unmatchColor.partInfo = sel_objs;

            }
            else
            {
                Response.StatusCode = 400;
            }

            // Color情報検索
            string color_sql = " SELECT                                                           " +
                               "  TTECI.EXT_COLOR1_NAME                                           " +
                               " , TTECI.EXT_COLOR2_NAME                                          " +
                               " , TEPD.EXT_COLOR_ID                                              " +
                               " , TEPD.CMF_STYLING_PART_ID                                       " +
                               " , TEPD.COLOR_INST_ID                                             " +
                               " , TEPD.PART_APPLY_RELATION_ID                                    " +
                               " FROM                                                             " +
                               "   T_COLOR_CHART_PART_RLT TCCPR                                   " +
                               " INNER JOIN T_EXT_COMMU_PLACE_DIST TEPD                           " +
                               "   ON TEPD.PART_APPLY_RELATION_ID = TCCPR.PART_APPLY_RELATION_ID  " +
                               "   AND TEPD.UNAPPLIED_FLAG = '0'                                  " +
                               " INNER JOIN T_TC_EXT_COLOR_INFO TTECI                             " +
                               "   ON TTECI.EXT_COLOR_ID = TEPD.EXT_COLOR_ID                      " +
                               " INNER JOIN T_COLOR_CHART_TC_EXT_COLOR TCCTEC                     " +
                               "   ON TTECI.EXT_COLOR_ID = TCCTEC.EXT_COLOR_ID                    " +
                               "   AND TCCTEC.COLOR_CHART_TC_ID = '" + request.color_chart_tc_id + "'" +
                               " WHERE                                                            " +
                               "       TCCPR.UNMATCH_FLG = 1                                      " +
                               "   AND TCCPR.COLOR_CHART_ID = '" + request.color_chart_id + "'    " +
                               " ORDER BY TCCTEC.DISPLAY_NO";


            DbUtil<BodyColor> color_dbUtil = new DbUtil<BodyColor>();

            IDictionary<string, object> color_map = color_dbUtil.SelectQuery(color_sql, conn, tr);

            if (color_map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                color_map.Remove("result");

                // COLOR情報設定
                List<BodyColor> sel_objs = (List<BodyColor>)color_map["resultList"];
                unmatchColor.bodyColor = sel_objs;

                List<ByBodyColorHes> unmatchColorHesList = new List<ByBodyColorHes>();

                foreach (var sel_obj in sel_objs)
                {

                    string hes_sql = " SELECT DISTINCT                                               " +
                                     "   TSPCIL.COLOR_HES                                            " +
                                     "   , TEPD.EXT_COLOR_ID                                         " +
                                     "   , TSPCIL.COLOR_INST_ID                                      " +
                                     "   , TSPCIL.CMF_STYLING_PART_ID                                " +
                                     " FROM                                                          " +
                                     "   T_STYLING_PART_COLOR_INST_LIST TSPCIL                       " +
                                     "   INNER JOIN T_EXT_COMMU_PLACE_DIST TEPD                      " +
                                     "     ON TEPD.CMF_STYLING_PART_ID = TSPCIL.CMF_STYLING_PART_ID  " +
                                     "     AND TEPD.UNAPPLIED_FLAG = '0'                             " +
                                     " WHERE                                                         " +
                                     "   TSPCIL.COLOR_INST_ID = '" + sel_obj.color_inst_id + "'" +
                                     "   AND TSPCIL.CMF_STYLING_PART_ID = '" + sel_obj.cmf_styling_part_id + "'";

                    DbUtil<ByBodyColorHes> hes_dbUtil = new DbUtil<ByBodyColorHes>();

                    IDictionary<string, object> hes_map = hes_dbUtil.SelectQuery(hes_sql, conn, tr);

                    if (hes_map["result"].Equals("OK"))
                    {
                        Response.StatusCode = 200;
                        hes_map.Remove("result");

                        List<ByBodyColorHes> hes_objs = (List<ByBodyColorHes>)hes_map["resultList"];

                        foreach (var hes_obj in hes_objs)
                        {
                            ByBodyColorHes hesInfo = new ByBodyColorHes();
                            hesInfo.ext_color_id = hes_obj.ext_color_id;
                            hesInfo.color_hes = hes_obj.color_hes;
                            hesInfo.color_inst_id = hes_obj.color_inst_id;
                            hesInfo.cmf_styling_part_id = hes_obj.cmf_styling_part_id;

                            unmatchColorHesList.Add(hesInfo);
                        }
                    }
                    else
                    {
                        Response.StatusCode = 400;
                    }

                }

                unmatchColor.colorHes = unmatchColorHesList;
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return unmatchColor;
        }

        // [UNMATCH]
        private XmlElement CreateWorksheetunmatchColor(XmlElement root, XmlDocument domDoc, List<UnmatchTabInfo> byBodyColorList)
        {

            // [COLOR]ノード作成
            XmlElement tabNode = domDoc.CreateElement("Worksheet");

            tabNode.SetAttribute("Name", ns_ss, "UNMATCH LIST");

            // 外装色ヘッダの数
            long headerCount = bodyColorHeaderCount;

            // 列数
            long colCount = 16 + headerCount;

            // Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, (colCount).ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, (byBodyColorList.Count + 1).ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var i = 0; i < colCount; i++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("Height", ns_ss, "40");
            rowNode.SetAttribute("header", "true");

            // ヘッダ作成（固定部）
            for (var i = 0; i < 14; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                if (i < 14)
                {
                    cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                    cellNode.SetAttribute("Index", ns_ss, (i + 1).ToString());

                    // Dataノード
                    XmlElement dataNode = domDoc.CreateElement("Data");
                    dataNode.SetAttribute("Type", ns_ss, "String");
                    dataNode.InnerText = COLOR_ROW_NM[i];
                    cellNode.AppendChild(dataNode);
                }

                rowNode.AppendChild(cellNode);

            }

            // ヘッダ作成（BodyColor部）
            for (var i = 0; i < headerCount; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");

                cellNode.SetAttribute("Index", ns_ss, (15 + i).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");

                dataNode.InnerText = colorHeaderName[i];

                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            // ヘッダ作成(固定部)
            for (var i = 15 + headerCount; i < colCount + 1; i++)
            {

                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");

                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, i.ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = COLOR_ROW_NM[i - headerCount];
                cellNode.AppendChild(dataNode);

                rowNode.AppendChild(cellNode);

            }

            tableNode.AppendChild(rowNode);

            // データ部作成
            for (var i = 0; i < byBodyColorList.Count; i++)
            {

                // Rowノード（データ）
                XmlElement rowDataNode = domDoc.CreateElement("Row");

                // Highlightノード
                XmlElement highlightNode = domDoc.CreateElement("Highlight");
                highlightNode.SetAttribute("id", "");
                rowDataNode.AppendChild(highlightNode);

                // Cellノード(SECTION)
                setCell(domDoc, byBodyColorList[i].section, "1", rowDataNode);

                // Cellノード(SUB-SECTION)
                setCell(domDoc, byBodyColorList[i].sub_section.TrimEnd(), "2", rowDataNode);

                // Cellノード(L1)
                setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].l1_part_no), "3", rowDataNode);

                // Cellノード(LVL)
                setCell(domDoc, byBodyColorList[i].lvl, "4", rowDataNode);

                // Cellノード(PART NO.)
                setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].changed_part), "5", rowDataNode);

                // Cellノード(PARENT PART NO.)
                setCell(domDoc, ChangeUnmatchNo(setUnderline(byBodyColorList[i].parent_part_no)), "6", rowDataNode);

                // Cellノード(PART NAME)
                setCell(domDoc, byBodyColorList[i].part_name, "7", rowDataNode);

                // Cellノード(PLANT)
                setCell(domDoc, byBodyColorList[i].plant, "8", rowDataNode);

                // Cellノード(MODEL)
                setCell(domDoc, byBodyColorList[i].model, "9", rowDataNode);

                // Cellノード(DESTINATION)
                setCell(domDoc, byBodyColorList[i].destination, "10", rowDataNode);

                // Cellノード(GRADE)
                setCell(domDoc, byBodyColorList[i].grade, "11", rowDataNode);

                // Cellノード(FEATURE)
                setCell(domDoc, byBodyColorList[i].feature, "12", rowDataNode);

                // Cellノード(ITEM)
                setCell(domDoc, byBodyColorList[i].item, "13", rowDataNode);

                // Cellノード(REMARKS)
                setCell(domDoc, byBodyColorList[i].remark, "14", rowDataNode);

                // Body Color
                // Cellノード(COLOR)
                for (var j = 0; j < byBodyColorList[i].color_item.Count; j++)
                {

                    if (!string.IsNullOrWhiteSpace(byBodyColorList[i].color_item[j]))
                    {

                        setCell(domDoc, byBodyColorList[i].color_item[j], (j + 15).ToString(), rowDataNode);
                    }
                    else
                    {
                        // 部品リレーションの子部品不一致フラグが2の場合、色指示を「-」に設定
                        if (byBodyColorList[i].child_unmatch_flag == 2)
                        {
                            setCell(domDoc, "-", (j + 15).ToString(), rowDataNode);
                        }
                        else
                        {
                            setCell(domDoc, "", (j + 15).ToString(), rowDataNode);
                        }
                    }

                }

                // Cellノード(L1_BASE PART NO)
                setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].l1_base_part_no), (colCount - 1).ToString(), rowDataNode);

                // Cellノード(BASE PART NO)
                setCell(domDoc, ChangeUnmatchNo(byBodyColorList[i].base_part_no), colCount.ToString(), rowDataNode);

                // L1部品等の-色指示になっている部品がBY BODYCOLOR側で出力されている
                tableNode.AppendChild(rowDataNode);

            }

            tabNode.AppendChild(tableNode);

            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "unmatch");
            tabNode.AppendChild(worksheetOptions);

            return tabNode;
        }

        private string ChangeUnmatchNo(string partNo)
        {
            string newPartNo = partNo;

            if (!string.IsNullOrWhiteSpace(partNo) && partNo.Length > 5)
            {
                newPartNo = partNo.Substring(0, 5);
            }

            return newPartNo;

        }

        private void SetUnmatchInfo(UnmatchInfo colorResults, List<UnmatchTabInfo> byBodyColorList)
        {
            // 外装色ヘッダ数
            long bodyColorCount = colorResults.bodyColor.Count;

            // 外装色ヘッダ重複除去
            List<BodyColorHeader> headerList = new List<BodyColorHeader>();

            // EXT COLOR ID臨時格納リスト
            List<long?> list = new List<long?>();

            // 外装色ヘッダ格納リスト（重複除く前）
            List<BodyColorHeader> headerBefor = new List<BodyColorHeader>();

            // 外装色ヘッダ重複のデータ除去
            for (var i = 0; i < bodyColorCount; i++)
            {
                if (!list.Contains(colorResults.bodyColor[i].ext_color_id))
                {
                    BodyColorHeader header = new BodyColorHeader();
                    header.col_no = i;
                    header.ext_color_id = colorResults.bodyColor[i].ext_color_id;

                    headerBefor.Add(header);
                    list.Add(colorResults.bodyColor[i].ext_color_id);
                }
            }

            for (var i = 0; i < headerBefor.Count; i++)
            {
                BodyColorHeader bodyColorHeader = new BodyColorHeader();
                bodyColorHeader.ext_color_id = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color_id;
                bodyColorHeader.col_no = i + 15;
                bodyColorHeader.ext_color1_name = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color1_name;
                bodyColorHeader.ext_color2_name = colorResults.bodyColor[(int)headerBefor[i].col_no].ext_color2_name;

                headerList.Add(bodyColorHeader);

            }

            // 外装色ヘッダ格納リスト（HESコード）
            List<string> colorHeaderList = new List<string>();

            // 外装色ヘッダ格納リスト（HESコード、カラー名）
            List<string> colorHeaderNmList = new List<string>();

            // ヘッダ作成（BodyColor部）
            for (var i = 0; i < headerList.Count; i++)
            {

                string extColor1Code = "";
                string extColor1Name = "";

                string extColor2Code = "";

                if (!string.IsNullOrWhiteSpace(headerList[i].ext_color1_name))
                {

                    var ext_color1_name = headerList[i].ext_color1_name.Split(' ');

                    extColor1Code = ext_color1_name[0];

                    for (var j = 1; j < ext_color1_name.Length; j++)
                    {
                        if (j == 1)
                        {
                            extColor1Name = ext_color1_name[j];
                        }
                        else
                        {
                            extColor1Name = extColor1Name + " " + ext_color1_name[j];
                        }
                    }

                }
                else
                {
                    extColor1Code = "-";
                }

                if (!string.IsNullOrWhiteSpace(headerList[i].ext_color2_name))
                {
                    var ext_color2_name = headerList[i].ext_color2_name.Split(' ');

                    extColor2Code = ext_color2_name[0];
                }

                string extColorTmp = headerList[i].ext_color1_name;
                if (!string.IsNullOrWhiteSpace(extColor2Code))
                {
                    extColorTmp = extColor1Code + "&#10;" + extColor1Name + "&#10;" + "(" + extColor2Code + ")";
                }
                else
                {
                    extColorTmp = extColor1Code + "&#10;" + extColor1Name;

                }

                colorHeaderList.Add(extColor1Code);

                colorHeaderNmList.Add(extColorTmp);

            }

            // ヘッダの数格納
            bodyColorHeaderCount = colorHeaderList.Count;

            // 外装色ヘッダ名リスト格納
            colorHeaderName = colorHeaderNmList;

            // データ部作成
            for (var i = 0; i < colorResults.partInfo.Count; i++)
            {
                UnmatchTabInfo colorTabInfo = new UnmatchTabInfo();

                colorTabInfo.color_header = colorHeaderList;

                colorTabInfo.color_header_name = colorHeaderNmList;

                // Cellノード(SECTION)
                colorTabInfo.section = colorResults.partInfo[i].section;

                // Cellノード(SUB-SECTION)
                colorTabInfo.sub_section = colorResults.partInfo[i].sub_section;

                // Cellノード(L1)
                colorTabInfo.l1_part_no = colorResults.partInfo[i].l1_part_no;

                // Cellノード(LVL)
                colorTabInfo.lvl = colorResults.partInfo[i].lvl;

                // Cellノード(PART NO.)
                colorTabInfo.changed_part = colorResults.partInfo[i].changed_part;

                // Cellノード(PARENT PART NO.)
                colorTabInfo.parent_part_no = colorResults.partInfo[i].parent_part_no;

                // Cellノード(PART NAME)
                colorTabInfo.part_name = colorResults.partInfo[i].part_nm;

                // Cellノード(PLANT)
                colorTabInfo.plant = colorResults.partInfo[i].productio;

                // Cellノード(MODEL)
                colorTabInfo.model = colorResults.partInfo[i].model_code;

                // Cellノード(DESTINATION)
                colorTabInfo.destination = colorResults.partInfo[i].destination;

                // Cellノード(GRADE)
                colorTabInfo.grade = colorResults.partInfo[i].grade;

                // Cellノード(FEATURE)
                colorTabInfo.feature = colorResults.partInfo[i].equipment;

                // Cellノード(ITEM)
                colorTabInfo.item = colorResults.partInfo[i].item;

                // Cellノード(REMARKS)
                colorTabInfo.remark = colorResults.partInfo[i].remark;

                // 色指示格納リスト
                List<string> colorTabHesList = new List<string>();

                string color_hes = "";

                // Body Color
                for (var colIdx = 0; colIdx < headerList.Count; colIdx++)
                {
                    color_hes = "";

                    for (int j = 0; j < colorResults.bodyColor.Count; j++)
                    {

                        if (colorResults.bodyColor[j].ext_color_id == headerList[colIdx].ext_color_id)
                        {

                            if (colorResults.partInfo[i].part_apply_relation_id == colorResults.bodyColor[j].part_apply_relation_id)
                            {
                                for (int k = 0; k < colorResults.colorHes.Count; k++)
                                {
                                    if (colorResults.colorHes[k].cmf_styling_part_id == colorResults.bodyColor[j].cmf_styling_part_id &&
                                        colorResults.colorHes[k].color_inst_id == colorResults.bodyColor[j].color_inst_id)
                                    {
                                        // Cellノード(BODY COLOR)
                                        color_hes = colorResults.colorHes[k].color_hes;
                                        break;
                                    }

                                }
                            }
                        }

                    }

                    if (!string.IsNullOrWhiteSpace(color_hes))
                    {
                        // 色指示が「Body Color」の場合、ヘッダの1色目に設定
                        if (color_hes.Equals("Body Color") || color_hes.Equals("BODY COLOR"))
                        {
                            color_hes = headerList[colIdx].ext_color1_name.Split(" ")[0];
                            colorTabHesList.Add(color_hes + "「Body Color」");

                        }
                        // 色指示がHESコード形式ではない場合、空欄に設定
                        else if (!CheckHesCode(color_hes))
                        {
                            colorTabHesList.Add("");
                        }
                        else
                        {
                            colorTabHesList.Add(color_hes);
                        }
                    }
                    else
                    {
                        colorTabHesList.Add("");
                    }

                }

                // 子部品不一致フラグ
                colorTabInfo.child_unmatch_flag = colorResults.partInfo[i].child_unmatch_flag;

                colorTabInfo.color_item = colorTabHesList;

                // Cellノード(L1_BASE PART NO)
                colorTabInfo.l1_base_part_no = colorResults.partInfo[i].l1_base_part_no;

                // Cellノード(BASE PART NO)
                colorTabInfo.base_part_no = colorResults.partInfo[i].base_part;

                for (var j = 0; j < colorTabInfo.color_item.Count; j++)
                {
                    colorTabInfo.color_item[j] = colorTabInfo.color_item[j].Replace("「Body Color」", "");
                }

                byBodyColorList.Add(colorTabInfo);
            }
        }

        // Worksheet[UNMATCH LIST]作成
        [NonAction]
        public void CreatUnmatchWorkSheet(XmlElement root, XmlDocument domDoc, long color_chart_id, long tc_id)
        {
            XmlElement tabNode = domDoc.CreateElement("Worksheet");
            tabNode.SetAttribute("Name", ns_ss, "UNMATCH LIST");
            // ExpandedColumnCount:INT Color Type数情報検索          
            IDictionary<string, object> int_color_Map = GetUnmatchIntColorTypeItems(color_chart_id, tc_id);
            var intcolorResults = new List<UnmatchGroupIntColorTypeResponse>();
            intcolorResults = (List<UnmatchGroupIntColorTypeResponse>)int_color_Map["resultList"];

            //INT Color Type数
            long int_color_type_count = intcolorResults.Count;

            //該当グループの部品情報検索
            IDictionary<string, object> part_items_Map = GetUnmatchGroupPartItems(color_chart_id);
            var partItemsResults = new List<UnmatchGroupPartItemsResponse>();
            partItemsResults = (List<UnmatchGroupPartItemsResponse>)part_items_Map["resultList"];
            if (partItemsResults.Count == 0)
            {
                return;
            }

            //該当グループの部品数
            long part_count = partItemsResults.Count;

            //Tableノード
            XmlElement tableNode = domDoc.CreateElement("Table");
            tableNode.SetAttribute("ExpandedColumnCount", ns_ss, (int_color_type_count + 16).ToString());
            tableNode.SetAttribute("ExpandedRowCount", ns_ss, (part_count + 1).ToString());
            tableNode.SetAttribute("FullColumns", ns_x, "1");
            tableNode.SetAttribute("FullRows", ns_x, "1");
            tableNode.SetAttribute("DefaultRowHeight", ns_ss, "16.2");

            // Columnノード
            for (var j = 0; j < int_color_type_count + 16; j++)
            {
                XmlElement colNode = domDoc.CreateElement("Column");
                colNode.SetAttribute("AutoFitWidth", ns_ss, "1");
                tableNode.AppendChild(colNode);
            }

            // Rowノード（ヘッダ）
            XmlElement rowNode = domDoc.CreateElement("Row");
            rowNode.SetAttribute("Height", ns_ss, "32.4");
            rowNode.SetAttribute("header", "true");

            for (var k = 0; k < 14; k++)
            {
                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (k + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = GROUP_ROW_NM[k];

                cellNode.AppendChild(dataNode);
                rowNode.AppendChild(cellNode);
            }
            tableNode.AppendChild(rowNode);

            // Cell設定（ヘッダ）
            for (var m = 14; m < int_color_type_count + 14; m++)
            {
                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sMultiHead");
                cellNode.SetAttribute("Index", ns_ss, (m + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");
                dataNode.InnerText = intcolorResults[m - 14].SYMBOL + "&#10;" + intcolorResults[m - 14].INT_COLOR_TYPE;

                cellNode.AppendChild(dataNode);
                rowNode.AppendChild(cellNode);
            }
            tableNode.AppendChild(rowNode);
            //2つのヘッダの設定が最適である
            for (var n = int_color_type_count + 14; n < int_color_type_count + 16; n++)
            {
                // Cellノード
                XmlElement cellNode = domDoc.CreateElement("Cell");
                cellNode.SetAttribute("StyleID", ns_ss, "sHead");
                cellNode.SetAttribute("Index", ns_ss, (n + 1).ToString());

                // Dataノード
                XmlElement dataNode = domDoc.CreateElement("Data");
                dataNode.SetAttribute("Type", ns_ss, "String");

                if (n == int_color_type_count + 14)
                {
                    dataNode.InnerText = "L1_BASE PART NO";
                }
                else if (n == int_color_type_count + 15)
                {
                    dataNode.InnerText = "BASE PART NO";
                }

                cellNode.AppendChild(dataNode);
                rowNode.AppendChild(cellNode);
            }
            tableNode.AppendChild(rowNode);

            //ROW部品設定
            for (var j = 0; j < partItemsResults.Count; j++)
            {
                XmlElement rowNodeData = domDoc.CreateElement("Row");
                // Highlightノード
                XmlElement highlightNode = domDoc.CreateElement("Highlight");
                highlightNode.SetAttribute("id", "");
                rowNodeData.AppendChild(highlightNode);

                // Cell1ノード
                XmlElement cellNode1 = domDoc.CreateElement("Cell");
                cellNode1.SetAttribute("Index", ns_ss, (1).ToString());
                // Data1ノード
                XmlElement dataNode1 = domDoc.CreateElement("Data");
                dataNode1.SetAttribute("Type", ns_ss, "String");
                dataNode1.InnerText = partItemsResults[j].SECTION;
                cellNode1.AppendChild(dataNode1);

                // Cell2ノード
                XmlElement cellNode2 = domDoc.CreateElement("Cell");
                cellNode2.SetAttribute("Index", ns_ss, (2).ToString());
                // Data2ノード
                XmlElement dataNode2 = domDoc.CreateElement("Data");
                dataNode2.SetAttribute("Type", ns_ss, "String");

                // Subsection6,7桁目が空白の際は空白除去
                dataNode2.InnerText = partItemsResults[j].SUB_SECTION.TrimEnd();
                cellNode2.AppendChild(dataNode2);

                // L1
                // Cell3ノード
                XmlElement cellNode3 = domDoc.CreateElement("Cell");
                cellNode3.SetAttribute("Index", ns_ss, (3).ToString());
                // Data3ノード
                XmlElement dataNode3 = domDoc.CreateElement("Data");
                dataNode3.SetAttribute("Type", ns_ss, "String");
                // DataL1 情報検索
                if (!string.IsNullOrEmpty(partItemsResults[j].L1_PART_NO))
                {
                    dataNode3.InnerText = ChangeUnmatchNo(partItemsResults[j].L1_PART_NO);
                }
                else
                {
                    dataNode3.InnerText = "";
                }
                cellNode3.AppendChild(dataNode3);

                // Cell4ノード
                XmlElement cellNode4 = domDoc.CreateElement("Cell");
                cellNode4.SetAttribute("Index", ns_ss, (4).ToString());
                // Data4ノード
                XmlElement dataNode4 = domDoc.CreateElement("Data");
                dataNode4.SetAttribute("Type", ns_ss, "String");
                if (string.IsNullOrEmpty(partItemsResults[j].LVL.ToString()))
                {
                    dataNode4.InnerText = "";
                }
                else
                {
                    dataNode4.InnerText = partItemsResults[j].LVL.ToString();
                }
                cellNode4.AppendChild(dataNode4);

                // Cell5ノード
                XmlElement cellNode5 = domDoc.CreateElement("Cell");
                cellNode5.SetAttribute("Index", ns_ss, (5).ToString());
                // Data5ノード
                XmlElement dataNode5 = domDoc.CreateElement("Data");
                dataNode5.SetAttribute("Type", ns_ss, "String");
                dataNode5.InnerText = ChangeUnmatchNo(partItemsResults[j].CHANGED_PART);
                cellNode5.AppendChild(dataNode5);

                // Cell6ノード
                XmlElement cellNode6 = domDoc.CreateElement("Cell");
                cellNode6.SetAttribute("Index", ns_ss, (6).ToString());
                // Data6ノード
                XmlElement dataNode6 = domDoc.CreateElement("Data");
                dataNode6.SetAttribute("Type", ns_ss, "String");
                // Data6情報検索
                if (!string.IsNullOrEmpty(partItemsResults[j].PARENT_PART_NO))
                {
                    dataNode6.InnerText = ChangeUnmatchNo(partItemsResults[j].PARENT_PART_NO);
                }
                else
                {
                    dataNode6.InnerText = "-";
                }

                cellNode6.AppendChild(dataNode6);

                // Cell7ノード
                XmlElement cellNode7 = domDoc.CreateElement("Cell");
                cellNode7.SetAttribute("Index", ns_ss, (7).ToString());
                // Data7ノード
                XmlElement dataNode7 = domDoc.CreateElement("Data");
                dataNode7.SetAttribute("Type", ns_ss, "String");
                dataNode7.InnerText = partItemsResults[j].PART_NM;
                cellNode7.AppendChild(dataNode7);

                // Cell8ノード
                XmlElement cellNode8 = domDoc.CreateElement("Cell");
                cellNode8.SetAttribute("Index", ns_ss, (8).ToString());
                // Data8ノード
                XmlElement dataNode8 = domDoc.CreateElement("Data");
                dataNode8.SetAttribute("Type", ns_ss, "String");
                dataNode8.InnerText = partItemsResults[j].PRODUCTIO;
                cellNode8.AppendChild(dataNode8);

                // Cell9ノード
                XmlElement cellNode9 = domDoc.CreateElement("Cell");
                cellNode9.SetAttribute("Index", ns_ss, (9).ToString());
                // Data9ノード
                XmlElement dataNode9 = domDoc.CreateElement("Data");
                dataNode9.SetAttribute("Type", ns_ss, "String");
                dataNode9.InnerText = partItemsResults[j].MODEL_CODE;
                cellNode9.AppendChild(dataNode9);

                // Cell10ノード
                XmlElement cellNode10 = domDoc.CreateElement("Cell");
                cellNode10.SetAttribute("Index", ns_ss, (10).ToString());
                // Data10ノード
                XmlElement dataNode10 = domDoc.CreateElement("Data");
                dataNode10.SetAttribute("Type", ns_ss, "String");
                dataNode10.InnerText = partItemsResults[j].DESTINATION;
                cellNode10.AppendChild(dataNode10);

                // Cell11ノード
                XmlElement cellNode11 = domDoc.CreateElement("Cell");
                cellNode11.SetAttribute("Index", ns_ss, (11).ToString());
                // Data11ノード
                XmlElement dataNode11 = domDoc.CreateElement("Data");
                dataNode11.SetAttribute("Type", ns_ss, "String");
                dataNode11.InnerText = partItemsResults[j].GRADE;
                cellNode11.AppendChild(dataNode11);

                // Cell12ノード
                XmlElement cellNode12 = domDoc.CreateElement("Cell");
                cellNode12.SetAttribute("Index", ns_ss, (12).ToString());
                // Data12ノード
                XmlElement dataNode12 = domDoc.CreateElement("Data");
                dataNode12.SetAttribute("Type", ns_ss, "String");
                dataNode12.InnerText = partItemsResults[j].EQUIPMENT;
                cellNode12.AppendChild(dataNode12);

                // Cell13ノード
                XmlElement cellNode13 = domDoc.CreateElement("Cell");
                cellNode13.SetAttribute("Index", ns_ss, (13).ToString());
                // Data13ノード
                XmlElement dataNode13 = domDoc.CreateElement("Data");
                dataNode13.SetAttribute("Type", ns_ss, "String");
                dataNode13.InnerText = partItemsResults[j].ITEM;
                cellNode13.AppendChild(dataNode13);

                // Cell14ノード
                XmlElement cellNode14 = domDoc.CreateElement("Cell");
                cellNode14.SetAttribute("Index", ns_ss, (14).ToString());
                // Data14ノード
                XmlElement dataNode14 = domDoc.CreateElement("Data");
                dataNode14.SetAttribute("Type", ns_ss, "String");
                dataNode14.InnerText = partItemsResults[j].REMARK;
                cellNode14.AppendChild(dataNode14);

                rowNodeData.AppendChild(cellNode1);
                rowNodeData.AppendChild(cellNode2);
                rowNodeData.AppendChild(cellNode3);
                rowNodeData.AppendChild(cellNode4);
                rowNodeData.AppendChild(cellNode5);
                rowNodeData.AppendChild(cellNode6);
                rowNodeData.AppendChild(cellNode7);
                rowNodeData.AppendChild(cellNode8);
                rowNodeData.AppendChild(cellNode9);
                rowNodeData.AppendChild(cellNode10);
                rowNodeData.AppendChild(cellNode11);
                rowNodeData.AppendChild(cellNode12);
                rowNodeData.AppendChild(cellNode13);
                rowNodeData.AppendChild(cellNode14);

                // Data部
                //INT_COLOR_ID(列) -> COLOR_HES
                IDictionary<string, object> ICI_CH_Map = GetUnmatchColorHesPartItems(partItemsResults[j].PART_APPLY_RELATION_ID.ToString());
                var ICI_CH_Results = new List<ColorHesItemsResponse>();
                ICI_CH_Results = (List<ColorHesItemsResponse>)ICI_CH_Map["resultList"];
                for (var k = 15; k < int_color_type_count + 15; k++)
                {
                    bool flag = false;
                    for (var m = 0; m < ICI_CH_Results.Count; m++)
                    {
                        //表のIDとデータのIDを同じセルに設定
                        if (intcolorResults[k - 15].INT_COLOR_ID == ICI_CH_Results[m].INT_COLOR_ID)
                        {

                            XmlElement cellNode = domDoc.CreateElement("Cell");
                            cellNode.SetAttribute("Index", ns_ss, (k).ToString());

                            XmlElement dataNode = domDoc.CreateElement("Data");
                            dataNode.SetAttribute("Type", ns_ss, "String");


                            if (string.IsNullOrEmpty(ICI_CH_Results[m].COLOR_HES))
                            {
                                // 部品リレーションの子部品不一致フラグが1の場合、「-」に設定
                                if (partItemsResults[j].CHILD_UNMATCH_FLAG == 1)
                                {
                                    dataNode.InnerText = "-";
                                }
                                else
                                {
                                    dataNode.InnerText = "";
                                }
                            }
                            else
                            {
                                dataNode.InnerText = ICI_CH_Results[m].COLOR_HES;
                            }

                            cellNode.AppendChild(dataNode);
                            rowNodeData.AppendChild(cellNode);
                            flag = true;
                            break;
                        }
                    }
                    //表の現在の列には、本条のデータに対応するidが表示されていません
                    if (!flag)
                    {
                        XmlElement cellNode = domDoc.CreateElement("Cell");
                        cellNode.SetAttribute("Index", ns_ss, (k).ToString());

                        XmlElement dataNode = domDoc.CreateElement("Data");
                        dataNode.SetAttribute("Type", ns_ss, "String");
                        if (partItemsResults[j].CHILD_UNMATCH_FLAG == 1)
                        {
                            dataNode.InnerText = "-";
                        }
                        else
                        {
                            dataNode.InnerText = "";
                        }
                        cellNode.AppendChild(dataNode);
                        rowNodeData.AppendChild(cellNode);
                    }

                }

                // L1_BASE PART NO ノード
                XmlElement cellNodeL1No = domDoc.CreateElement("Cell");
                cellNodeL1No.SetAttribute("Index", ns_ss, (int_color_type_count + 15).ToString());
                // L1_BASE PART NO ノード
                XmlElement dataNodeL1No = domDoc.CreateElement("Data");
                dataNodeL1No.SetAttribute("Type", ns_ss, "String");
                if (!string.IsNullOrEmpty(partItemsResults[j].L1_BASE_PART_NO))
                {
                    dataNodeL1No.InnerText = ChangeUnmatchNo(partItemsResults[j].L1_BASE_PART_NO);
                }
                else
                {
                    dataNodeL1No.InnerText = "";
                }

                cellNodeL1No.AppendChild(dataNodeL1No);
                rowNodeData.AppendChild(cellNodeL1No);
                // BasePartNoノード
                XmlElement cellBasePartNo = domDoc.CreateElement("Cell");
                cellBasePartNo.SetAttribute("Index", ns_ss, (int_color_type_count + 16).ToString());
                // BasePartNoノード
                XmlElement dataBasePartNo = domDoc.CreateElement("Data");
                dataBasePartNo.SetAttribute("Type", ns_ss, "String");
                dataBasePartNo.InnerText = ChangeUnmatchNo(partItemsResults[j].BASE_PART);
                cellBasePartNo.AppendChild(dataBasePartNo);
                rowNodeData.AppendChild(cellBasePartNo);

                tableNode.AppendChild(rowNodeData);
            }

            tabNode.AppendChild(tableNode);
            // WorksheetOptionsノード
            XmlElement worksheetOptions = SetWorksheetOptOptions(domDoc, "unmatch");
            tabNode.AppendChild(worksheetOptions);

            root.AppendChild(tabNode);
        }

        //INT Color Type情報取得
        [NonAction]
        public IDictionary<string, object> GetUnmatchIntColorTypeItems(long color_chart_id, long tc_id)
        {
            string sql = "SELECT DISTINCT                                      " +
                         "  TTICI.INT_COLOR_TYPE                               " +
                         " , TTICI.SYMBOL                                      " +
                         " , TTICI.INT_COLOR_ID                                " +
                         " , TCCTIC.DISPLAY_NO                                 " +
                         "FROM                                                 " +
                         "  T_TC_INT_COLOR_INFO TTICI                          " +
                         "  INNER JOIN T_COLOR_CHART_TC_INT_COLOR TCCTIC       " +
                         "    ON TTICI.INT_COLOR_ID = TCCTIC.INT_COLOR_ID      " +
                         "    AND TCCTIC.COLOR_CHART_TC_ID = '" + tc_id + "'   " +
                         "  INNER JOIN T_INT_COMMU_PLACE_DIST TICPD            " +
                         "    ON TTICI.INT_COLOR_ID =TICPD.INT_COLOR_ID        " +
                         "    AND TICPD.UNAPPLIED_FLAG = '0'                   " +
                         "  INNER JOIN T_COLOR_CHART_PART_RLT TCCPR            " +
                         "    ON TCCPR.PART_APPLY_RELATION_ID = TICPD.PART_APPLY_RELATION_ID" +
                         "    AND TCCPR.COLOR_CHART_ID = '" + color_chart_id + "'" +
                         "    AND TCCPR.UNMATCH_FLG = '1'                      " +
                         "ORDER BY TCCTIC.DISPLAY_NO, TTICI.SYMBOL";

            DbUtil<UnmatchGroupIntColorTypeResponse> dbUtil = new DbUtil<UnmatchGroupIntColorTypeResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        //GetGroupPartItems
        [NonAction]
        public IDictionary<string, object> GetUnmatchGroupPartItems(long color_chart_id)
        {
            string sql = "SELECT                              " +
                         "TP.SECTION                          " +
                         " , TP.SUB_SECTION                   " +
                         " , TP.LVL                           " +
                         " , TP.CHANGED_PART                  " +
                         " , TP.PARENT_PART_ID                " +
                         " , TP.PART_NM                       " +
                         " , TP.L1_PART_ID                    " +
                         " , TP.BASE_PART                     " +
                         " , TCCPR.PART_APPLY_RELATION_ID     " +
                         " , TCCPR.PART_ID                    " +
                         " , TCCPR.MODEL_CODE                 " +
                         " , TCCPR.DESTINATION                " +
                         " , TCCPR.GRADE                      " +
                         " , TCCPR.ITEM                       " +
                         " , TCCPR.REMARK                     " +
                         " , TCCPR.PRODUCTIO                  " +
                         " , TCCPR.EQUIPMENT                  " +
                         " , TCCPR.CHILD_UNMATCH_FLAG         " +
                         " , TP2.BASE_PART AS L1_BASE_PART_NO " +
                         " , TP2.CHANGED_PART AS L1_PART_NO   " +
                         " , TP3.CHANGED_PART AS PARENT_PART_NO " +
                         "FROM                                                 " +
                         "   T_COLOR_CHART_PART_RLT TCCPR                      " +
                         "INNER JOIN                                           " +
                         "   T_PART TP                                         " +
                         "ON                                                   " +
                         "   TCCPR.PART_ID = TP.PART_ID                        " +
                         "LEFT JOIN                                            " +
                         "  T_PART TP2                                         " +
                         "ON                                                   " +
                         "  TP2.PART_ID = TP.L1_PART_ID                        " +
                         "LEFT JOIN                                            " +
                         "  T_PART TP3                                         " +
                         "ON                                                   " +
                         "  TP3.PART_ID = TP.PARENT_PART_ID                    " +
                         "WHERE                                                " +
                         "    TCCPR.COLOR_CHART_ID = '" + color_chart_id + "'  " +
                         "AND TCCPR.UNMATCH_FLG = '1'                          " +
                         "AND TP2.CHANGED_PART IS NOT NULL                     " +
                         "AND(                                                 " +
                         "  TP.LVL = 1                                         " +
                         "  OR(TP.LVL > 1 AND TP3.CHANGED_PART IS NOT NULL)    " +
                         "  )                                                  " +
                         "ORDER BY TP.SECTION, TP.SUB_SECTION, TP.L1_PART_ID, TP.PART_ID, TCCPR.PART_APPLY_RELATION_ID";

            DbUtil<UnmatchGroupPartItemsResponse> dbUtil = new DbUtil<UnmatchGroupPartItemsResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

        //INT_COLOR_ID(列) -> COLOR_HES
        [NonAction]
        public IDictionary<string, object> GetUnmatchColorHesPartItems(string part_apply_relation_id)
        {

            string sql = "SELECT                                                            " +
                         "  TSPCIL.INT_COLOR_ID                                             " +
                         "  , TSPCIL.COLOR_HES                                              " +
                         "FROM                                                              " +
                         "  T_STYLING_PART_COLOR_INST_LIST TSPCIL                           " +
                         "INNER JOIN T_INT_COMMU_PLACE_DIST TICPD       " +
                         "    ON TSPCIL.CMF_STYLING_PART_ID = TICPD.CMF_STYLING_PART_ID     " +
                         "   AND TSPCIL.COLOR_INST_ID = TICPD.COLOR_INST_ID     " +
                         "   AND TSPCIL.INT_COLOR_ID = TICPD.INT_COLOR_ID     " +
                         "   AND TICPD.UNAPPLIED_FLAG = '0'     " +
                         "WHERE                                                             " +
                         "    TICPD.PART_APPLY_RELATION_ID = '" + part_apply_relation_id + "'";

            DbUtil<ColorHesItemsResponse> dbUtil = new DbUtil<ColorHesItemsResponse>();

            OracleConnection conn = dbUtil.ConnectionOpen();

            OracleTransaction tr = dbUtil.Transaction(conn);

            IDictionary<string, object> map = dbUtil.SelectQuery(sql, conn, tr);

            if (map["result"].Equals("OK"))
            {
                Response.StatusCode = 200;
                map.Remove("result");
            }
            else
            {
                Response.StatusCode = 400;
            }

            dbUtil.ConnectionClose(conn, tr);

            return map;
        }

    }
}
