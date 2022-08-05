using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;
using TFlex.Drawing;
using TFlex.Command;

namespace CREATEBYEXCEL_PLG
{
    public class CommandManager
    {
        public int Level, exportSet;
        public string BasePath;
        public string pathSTPRU, pathDXFRU, pathPDFRU;
        public bool exportDXF, exportSTEP, exportPDF, saveToDOCs, saveToEXCEL;
        public bool regenerated = false;
        DateTime t1, t2;
        private readonly System.Diagnostics.Stopwatch uptime = new System.Diagnostics.Stopwatch();
        int i;

        #region GET
        private Page GetPageDXF(Document doc, string namePage)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.Name == namePage)
                {
                    return page;
                }
            }
            return null;
        }

        private Page GetPageDXF2(Document doc, PageType pageType)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    return page;
                }
            }
            return null;
        }

        private List<Page> GetPagesPDF(Document doc, PageType pageType)
        {
            List<Page> pagesList = new List<Page>();
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    pagesList.Add(page);
                }
            }
            if (pagesList.Count != 0) return pagesList;
            else return null;
        }

        private List<ProductStructure> GetProductStructures(Document doc)
        {
            List<ProductStructure> elementsList = new List<ProductStructure>();
            foreach (ProductStructure element in doc.GetProductStructures())
            {
                if (element.DisplayName != null)
                {
                    elementsList.Add(element);
                }
            }
            if (elementsList.Count != 0) return elementsList;
            else return null;
        }
        #endregion GET

        #region READEXCEL
        public static void ReadExcel()
        {

        }

        public struct EXL
        {
            public Excel.Application Application;
            public Excel.Workbooks Workbooks;
            public Excel.Workbook Workbook;
            public Excel.Sheets Sheets;
            public Excel.Worksheet Sheet;
            public Excel.Range Cell;
        }

        public static void LoadExcel(EXL Excel_file, string xlFileName)
        {
            Excel_file.Application = new Excel.Application();
            Excel_file.Workbooks = Excel_file.Application.Workbooks;
            Excel_file.Workbook = Excel_file.Workbooks.Open(xlFileName);
            Excel_file.Sheets = Excel_file.Workbook.Worksheets;
            Excel_file.Sheet = Excel_file.Sheets.Item[1];
            Excel_file.Cell = Excel_file.Sheet.Cells[1, 1];
        }

        static private string GetLetter(int nn)
        {
            string p1;

            int n2 = nn / 26;
            if (n2 > 0)
            {
                p1 = ((char)((int)('A') + n2 - 1)).ToString() + ((char)((int)('A') + nn - n2 * 26)).ToString();
            }
            else
            {
                p1 = ((char)((int)('A') + nn)).ToString();
            }

            return p1;
        }
        #endregion READEXCEL

        #region OK
        public void OK(Document doc, ATTRIBUTES_COM parameter)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            EXL EX_DATA = new EXL();
            object[,] dataArr = null;
            List<string> variablesList = null;

            // Устанавливаем параметры числового формата
            NumberFormatInfo NFI;
            NFI = new CultureInfo("en-US", false).NumberFormat;
            // Устанавливаем делитель в числах
            NFI.NumberDecimalSeparator = ".";

            OpenFileDialog inputFile = new OpenFileDialog();

            inputFile.Filter = "Файлы Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
            inputFile.FilterIndex = 1;
            inputFile.RestoreDirectory = true;

            if (inputFile.ShowDialog() != DialogResult.OK)
                return;

            string xlFileName = inputFile.FileName;

            Excel.Range Rng;
            EX_DATA.Application = new Excel.Application();
            EX_DATA.Workbooks = EX_DATA.Application.Workbooks;
            EX_DATA.Workbook = EX_DATA.Workbooks.Open(xlFileName);
            EX_DATA.Sheets = EX_DATA.Workbook.Worksheets;
            EX_DATA.Sheet = EX_DATA.Sheets.Item[1];
            EX_DATA.Cell = EX_DATA.Sheet.Cells[1, 1];

            int iLastRow = EX_DATA.Cell[EX_DATA.Sheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
            int iLastCol = EX_DATA.Cell[1, EX_DATA.Sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            Rng = (Excel.Range)EX_DATA.Sheet.Range["A1", EX_DATA.Sheet.Cells[iLastRow, iLastCol]];

            dataArr = (object[,])Rng.Value2;

            string[] arrCol = new string[iLastCol];

            if (parameter.parameterDXF == 1) exportDXF = true;
            if (parameter.parameterSTEP == 1) exportSTEP = true;
            if (parameter.parameterPDF == 1) exportPDF = true;
            if (parameter.parameterDOCs == 1) saveToDOCs = true;
            if (parameter.parameterEXCEL == 1) saveToEXCEL = true;
            if (!exportDXF && !exportSTEP && !exportPDF && !saveToDOCs && !saveToEXCEL) exportSet = -1;
            else exportSet = 1;

            //if (parameter.parameterDOCs == 1) exportSet = 1;

            if (exportSet != -1)
            {
                //string logname = doc.FileName;
                //logname = logname.Replace(".grb", ".log");
                /*string logname = "C:\\Users\\litvinov.ls\\Documents\\_Горбатенко\\log.txt";
                using (StreamWriter sw = new StreamWriter(logname))*/
                if (doc == null)
                    return;

                try
                {
                    doc.BeginChanges("");
                    for (int i = 2; i <= iLastRow; i++)
                    {
                        for (int j = 1; j <= iLastCol; j++)
                        {
                            arrCol[j - 1] = dataArr[i, j].ToString();

                            if (doc.FindVariable($"{dataArr[1, j]}") != null)
                            {
                                doc.FindVariable($"{dataArr[1, j]}").RealValue = Convert.ToDouble(arrCol[j - 1]);
                            }
                            else
                            {
                                doc.FindVariable($"${dataArr[1, j]}").TextValue = arrCol[j - 1];
                            }
                        }

                        string oboz = "-";
                        string naim = "-";
                        Variable voboz;
                        Variable vnaim;
                        if (doc.FindVariable("$Обозначение") != null)
                        {
                            voboz = doc.FindVariable("$Обозначение");
                        }
                        else voboz = null;
                        if (doc.FindVariable("$Наименование") != null)
                        {
                            vnaim = doc.FindVariable("$Наименование");
                        }
                        else vnaim = null;

                        if (voboz != null)
                        {
                            oboz = voboz.TextValue;
                        }
                        else oboz = "";
                        if (vnaim != null)
                        {
                            naim = vnaim.TextValue;
                        }
                        else naim = "";

                        FileInfo parFile = new FileInfo(doc.FileName);
                        DirectoryInfo parDir = new DirectoryInfo(parFile.DirectoryName);
                        BasePath = parDir.FullName;
                        string subpathSTPRU = @"STP";
                        string subpathDXFRU = @"DXF";
                        string subpathPDFRU = @"PDF";
                        pathSTPRU = pathDXFRU = pathPDFRU = parFile.DirectoryName;
                        DirectoryInfo dirInfo = new DirectoryInfo(pathSTPRU);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        dirInfo.CreateSubdirectory($"{dataArr[i, 1]}");
                        pathSTPRU = pathDXFRU = pathPDFRU = $"{pathSTPRU}\\{dataArr[i, 1]}";

                        string file_name = doc.FileName;
                        Level = 0;
                        TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.AutoRefresh;

                        RegenerateOptions regenerationOptions = new RegenerateOptions();
                        regenerationOptions.Full = true;
                        regenerationOptions.UpdateAllLinks = true;
                        regenerationOptions.UpdateProductStructures = true;
                        regenerationOptions.UpdateBillOfMaterials = true;
                        doc.Regenerate(regenerationOptions);
                        regenerated = true;
                        //Выгрузка в DOCs
                        if (parameter.parameterDOCs == 1) doc.SaveInNomenclature(true, true);
                        //

                        //
                        GetFragmentData(doc, file_name, doc.FilePath, oboz, naim, Level);
                    }
                    doc.EndChanges();
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message, "Ошибка", System.Windows.Forms.MessageBoxButtons.OK);

                    TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";
                }
            }

            EX_DATA.Application.DisplayAlerts = false;
            EX_DATA.Workbooks.Close();
            EX_DATA.Application.Quit();
            EX_DATA.Application.DisplayAlerts = true;

            Marshal.ReleaseComObject(EX_DATA.Cell);
            Marshal.ReleaseComObject(EX_DATA.Sheet);
            Marshal.ReleaseComObject(EX_DATA.Sheets);
            Marshal.ReleaseComObject(EX_DATA.Workbook);
            Marshal.ReleaseComObject(EX_DATA.Workbooks);
            Marshal.ReleaseComObject(EX_DATA.Application);
            Marshal.ReleaseComObject(Rng);

            TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";

            doc.Selection.DeselectAll();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено\nRunTime: " + elapsedTime, "Создание по Excel");
            else System.Windows.Forms.MessageBox.Show("Не выбраны параметры экспорта", "Создание по Excel");
            return;
        }
        #endregion OK


        private void GetFragmentData(Document doc, string name, string path, string oboz, string naim, int lev)
        {
            RegenerateOptions regenerateOptions = new RegenerateOptions();
            //regenerateOptions.Full = true;
            regenerateOptions.UpdateAllLinks = true;
            regenerateOptions.UpdateProductStructures = true;
            regenerateOptions.UpdateBillOfMaterials = true;
            doc.Regenerate(regenerateOptions);

            string offset = "";
            string confName;

            for (int nn = 0; nn < lev; nn++)
                offset += "   ";

            confName = oboz;

            Variable voboz;
            Variable vnaim;

            if (doc.FindVariable("$Обозначение") != null)
            {
                voboz = doc.FindVariable("$Обозначение");
            }
            else voboz = null;

            if (doc.FindVariable("$Наименование") != null)
            {
                vnaim = doc.FindVariable("$Наименование");
            }
            else vnaim = null;

            if (voboz != null)
            {
                oboz = voboz.TextValue;
            }
            else oboz = "";

            if (vnaim != null)
            {
                naim = vnaim.TextValue;
            }
            else naim = "";

            #region EXPORT
            if (naim != "" || oboz != "")
            {
                if (exportSTEP)
                {
                    doc.Regenerate(regenerateOptions);
                    regenerated = true;
                    ExportToStep exportSTPRU = new ExportToStep(doc);
                    string fileNameSTPRU = ($"{pathSTPRU}\\{oboz}_{naim}.stp");
                    if (!File.Exists(fileNameSTPRU))
                    {
                        exportSTPRU.Export(fileNameSTPRU);
                    }
                }
                if (exportDXF)
                {
                    if (!regenerated) doc.Regenerate(regenerateOptions);
                    ExportToDXF exportDXFRU = new ExportToDXF(doc);
                    Page pgRUDXF = GetPageDXF(doc, "Развертка");
                    Page pgRUDXF2 = GetPageDXF(doc, "Unfolding");
                    if (pgRUDXF != null || pgRUDXF2 != null)
                    {
                        List<Page> pgDXFRU = new List<Page>();
                        pgDXFRU.Add(pgRUDXF);
                        pgDXFRU.Add(pgRUDXF2);
                        exportDXFRU.ExportPages = pgDXFRU;
                        string fileNameDXFRU = ($"{pathDXFRU}\\{oboz}_{naim}.dxf");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            exportDXFRU.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {
                    if (!regenerated) doc.Regenerate(regenerateOptions);
                    ExportToPDF exportPDFnormalRU = new ExportToPDF(doc);
                    List<Page> pgPDFnormalRU = GetPagesPDF(doc, PageType.Normal);
                    if (pgPDFnormalRU != null)
                    {
                        exportPDFnormalRU.ExportPages = pgPDFnormalRU;
                        exportPDFnormalRU.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{oboz}_{naim}.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            exportPDFnormalRU.Export(fileNamePDFRU);
                        }
                    }

                    ExportToPDF exportPDFBOMRU = new ExportToPDF(doc);
                    List<Page> pgPDFBOMRU = GetPagesPDF(doc, PageType.BillOfMaterials);
                    if (pgPDFBOMRU != null)
                    {
                        exportPDFBOMRU.ExportPages = pgPDFBOMRU;
                        exportPDFBOMRU.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{oboz}_{naim}_СП.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            exportPDFBOMRU.Export(fileNamePDFRU);
                        }
                    }
                    regenerated = false;
                }
                if (saveToEXCEL)
                {
                    foreach (ProductStructure product in doc.GetProductStructures())
                    {
                        if (product.Name.Contains("АСУП"))
                        {
                            doc.BeginChanges("");
                            product.Regenerate(true);
                            product.UpdateStructure();
                            ProductStructureExcelExportOptions options = new ProductStructureExcelExportOptions();
                            options.FilePath = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_{product.Name}.xlsx");
                            options.Silent = true;
                            TFlex.Model.Data.ProductStructure.GroupingRules item = new TFlex.Model.Data.ProductStructure.GroupingRules();
                            item.Name = "АСУП";
                            options.GroupingUID = item.ID;
                            product.ExportToExcel(options);
                            doc.EndChanges();
                        }
                    }
                }
            }
            #endregion EXPORT

            int n_fr = doc.GetFragments3D().Count;
            foreach (Fragment3D frag in doc.GetFragments3D())
            {
                if (frag.Suppression.Suppress) continue;
                if (!frag.VisibleInScene) continue;
                if (frag.Layer.Hidden) continue;
                if (frag.FileName.Contains("Болт") || frag.FileName.Contains("Винт") || frag.FileName.Contains("Заклепка") || frag.FileName.Contains("Кольцо") ||
                    frag.FileName.Contains("Ось") || frag.FileName.Contains("Гайка") || frag.FileName.Contains("Шайба") || frag.FileName.Contains("Уплотнитель") ||
                    frag.FileName.Contains("Подшипник") || frag.FileName.Contains("Шпилька") || frag.FileName.Contains("Шплинт") || frag.FileName.Contains("Шпонка") ||
                    frag.FileName.Contains("Штифт") || frag.FileName.Contains("Штуцер") || frag.FileName.Contains("Шуруп") || frag.FileName.Contains("Этикетка") ||
                    frag.FileName == "Фильтр.grb") continue;
                {
                    Document docFR = null;
                    string obozF = "-";
                    string naimF = "-";
                    bool err = false;
                    string str_err = "";
                    string FRname = frag.FullFilePath;
                    if (lev > 0)
                    {
                        FRname = frag.FilePath;
                        FRname = TFlex.Application.FindPathName(FRname);
                    }

                    if (File.Exists(FRname))
                    {
                        Fragment.OpenPartOptions options = new Fragment.OpenPartOptions();
                        options.DontShowDocument = true;
                        options.QuietMode = true;
                        options.SubstituteGeometry = true;
                        options.SubstituteVariables = true;
                        options.SubstituteStatus = true;
                        docFR = frag.OpenPart(options);

                        if (docFR != null)
                        {
                            if (vnaim != null)
                            {
                                naimF = vnaim.TextValue;
                            }
                            else
                            {
                                naimF = "Переменная $Наименование не найдена";
                            };
                            if (voboz != null)
                            {
                                obozF = voboz.TextValue;
                            }
                            else
                            {
                                obozF = "Переменная $Обозначение не найдена";
                            };
                        }
                        else
                        {
                            err = true;
                            str_err = "Ошибка открытия";
                        }
                    }
                    else
                    {
                        err = true;
                        str_err = "Файл не найден";
                    }

                    if (err == false)
                    {
                        GetFragmentData(docFR, FRname, frag.FullFilePath, obozF, naimF, lev + 1);
                        docFR.Close();
                    }
                }
            }
            return;
        }
    }

    public class ATTRIBUTES_COM
    {
        public Int16 attribute;
        public int parameterDXF // Экспорт в STP
        {
            get { return (attribute & 0x0001); }
            set
            {
                attribute = (Int16)(attribute & 0xFFFE);
                attribute = (Int16)(attribute | (Int16)value);
            }
        }

        public int parameterSTEP // Экспорт в DXF
        {
            get { return ((attribute & 0x0002) >> 1); }
            set
            {
                attribute = (Int16)(attribute & 0xFFFD);
                attribute = (Int16)(attribute | (Int16)(value << 1));
            }
        }
        public int parameterPDF // Экспорт в PDF
        {
            get { return ((attribute & 0x0004) >> 2); }
            set
            {
                attribute = (Int16)(attribute & 0xFFFB);
                attribute = (Int16)(attribute | (Int16)(value << 2));
            }
        }

        public int parameterDOCs // Экспорт в DOCs
        {
            get { return ((attribute & 0x0008) >> 3); }
            set
            {
                attribute = (Int16)(attribute & 0xFFF7);
                attribute = (Int16)(attribute | (Int16)(value << 3));
            }
        }

        public int parameterEXCEL // Экспорт в DOCs
        {
            get { return ((attribute & 0x0010) >> 4); }
            set
            {
                attribute = (Int16)(attribute & 0xFFEF);
                attribute = (Int16)(attribute | (Int16)(value << 4));
            }
        }


        public void Set(int dxfs, int steps, int pdfs)
        {
            parameterDXF = dxfs;
            parameterSTEP = steps;
            parameterPDF = pdfs;
        }
    }
}