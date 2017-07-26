using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Diagnostics;

namespace MWO_XMLReader
{
    class ExcelOps
    {
        public static Excel.Application ExApp = null;

        public ExcelOps()
        {
            ExApp = new Excel.Application();
        }

        public static bool isFileOpenable (string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists)
            {
                ExApp = new Excel.Application();
                Excel.Worksheet ExWS = new Excel.Worksheet();
                Excel.Workbook ExWB = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    ExWB = null;
                    //FileOps.RemoveAttributes(fileName);
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    if (fileInfo.Extension.ToUpper() == ".XLS")
                    {
                        ExWB.SaveAs(fileName, Excel.XlFileFormat.xlExcel8,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                    }
                    else
                    {
                        ExWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,//.xlOpenXMLStrictWorkbook,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                    }

                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    Marshal.FinalReleaseComObject(ExApp);
                }
                catch (Exception ex)
                {
                    try
                    {
                        if (ExWB != null)
                        {
                            ExWB.Close();
                            Marshal.FinalReleaseComObject(ExWB);
                        }
                        if (ExWS != null)
                        {
                            Marshal.FinalReleaseComObject(ExWS);
                        }
                        if (ExApp != null)
                        {
                            Marshal.FinalReleaseComObject(ExApp);
                        }
                        Logger.PrintLC("Error: " + ex.Message, 2, "ERR:");
                        return false;
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            return true;
        }



        /// <summary>
        /// Converts an Open Office (.xlsx) file to the proprietary MS old format (.xls).
        /// -A.Amato, 2016 11
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public static bool ConvertXLSXtoXLS(string fileName = null)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;

            //if (ExApp == null)
            //    return false;

            ExApp = new Excel.Application();
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists && (fileInfo.Extension.ToUpper() == ".XLSX"))
            {
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();
                    ExApp.DisplayAlerts = false;
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    fileName = Path.ChangeExtension(fileName, ".xls"); //.Replace(".xlsx", ".xls");
                    FileInfo FileToSaveInfo = new FileInfo(fileName);
                    if (FileToSaveInfo.Exists)
                    {
                        FileToSaveInfo.Delete();
                    }
                    ExWB.SaveAs(fileName, Excel.XlFileFormat.xlExcel8,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                    Marshal.FinalReleaseComObject(ExApp);
                    Logger.PrintLC("Successfully converted " + fileInfo.FullName + " to " + fileName, 2, "INFO:");
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Error converting XLSX to XLS: " + exp.Message, 2, "ERR:");
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Converts a proprietary MS old format (.xls) to the Open Office (.xlsx).
        /// -A.Amato, 2016 11
        /// </summary>
        /// <param name="fileName">Path and file name to convert.</param>
        /// <returns>True if successfull, False otherwise.</returns>
        public static bool ConvertXLStoXLSX(string fileName = null)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists && (fileInfo.Extension.ToUpper() == ".XLS"))
            {
                Excel.Workbook ExWB; // = new Excel.Workbook();
                try
                {
                    Excel.Worksheet ExWS = new Excel.Worksheet();
                    ExApp = new Excel.Application();
                    ExApp.DisplayAlerts = false;
                    ExWB = ExApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    fileName = Path.ChangeExtension(fileName, ".xlsx");
                    FileInfo FileToSaveInfo = new FileInfo(fileName);
                    if (FileToSaveInfo.Exists)
                    {
                        FileToSaveInfo.Delete();
                    }
                    ExWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,//.xlOpenXMLStrictWorkbook,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                    ExWB.Close();
                    ExApp.DisplayAlerts = true;
                    Logger.PrintLC("File " + fileInfo.Name + " converted successfully to XLSX", 3, "INFO:");
                    Marshal.FinalReleaseComObject(ExWB);
                    Marshal.FinalReleaseComObject(ExWS);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("File " + fileInfo.Name + " could not be converted to XLSX. Error: " + exp.Message, 3);
                    return false;
                }
            }
            else
                return false;
            return true;
        }

        
        public static void WriteQuirks(List<MechStats> list)
        {
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp = new Excel.Application();
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(new FileInfo(ConfigFile.TEMPLATE.FullName));

                ExApp.DisplayAlerts = true;
            }
            catch
            {
                try
                {
                    ExApp.DisplayAlerts = false;
                    Logger.PrintLC(ConfigFile.TEMPLATE.Name + " is opened by another application. Will skip its creation.", 2, "ERR:");
                }
                catch { }
                return;
            }
            WB = p.Workbook;
            ws = WB.Worksheets;
            ExcelWorksheet worksheet = null;
            foreach (ExcelWorksheet works in ws)
            {
                if(works.Name.Trim().ToUpper() == ConfigFile.QUIRK_PAGE.ToUpper())
                {
                    list = Worker.SortMechs(list);

                }
            }

        }
        /*
        public static void WriteErwinOutcome(FileAll12T file)
        {
            string extension = ".xlsx";
            string suffix = "_esito";
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            
            string mystring = string.Empty;
            mystring = Path.GetFullPath(file.FullName);
            mystring = mystring.Replace(Path.GetFileName(file.FullName), "");
            string fileToOpen = Path.Combine(mystring, ConfigFile.ALL12_DEST_FOLD_NAME, ConfigFile.TIMESTAMPFOLDER, Path.GetFileNameWithoutExtension(file.FullName) + suffix + extension);
            if(FileOps.CopyFile(ConfigFile.ESITI_TEMPLATE, fileToOpen))
            {
                Logger.PrintLC("Copied Outcome (Esiti) file " + ConfigFile.ESITI_TEMPLATE + " to " + Path.GetDirectoryName(fileToOpen) + 
                    " with name " + new FileInfo(fileToOpen).Name + ".", 3, ConfigFile.INFO);
                if(new FileInfo(fileToOpen).Exists)
                {
                    Logger.PrintLC("\n# ELABORATION OF OUTCOME (Esiti) FILE FOR " + file.SimpleName + " - START #", 2, ConfigFile.INFO);
                    try
                    {
                        if (!FileOps.isFileOpenable(fileToOpen))
                        {
                            Logger.PrintLC("File " + fileToOpen + " could not be opened.", 2, ConfigFile.ERROR);
                            Logger.PrintLC("# ELABORATION OF OUTCOME (Esiti) FILE FOR " + file.SimpleName + " - END #", 2, ConfigFile.INFO);
                            return;
                        }
                        FileInfo fileDaAprire = new FileInfo(fileToOpen);
                        try
                        {
                            ExApp = new Excel.Application();
                            ExApp.DisplayAlerts = false;
                            p = new ExcelPackage(new FileInfo(fileToOpen));
                            ExApp.DisplayAlerts = true;
                        }
                        catch
                        {
                            try
                            {
                                ExApp.DisplayAlerts = false;
                                Logger.PrintLC(fileDaAprire.Name + " is opened by another application. Will skip its creation.", 2, ConfigFile.ERROR);
                            }
                            catch { }
                            return;
                        }
                        WB = p.Workbook;
                        ws = WB.Worksheets;
                        
                        foreach (var worksheet in ws)
                        {
                            // Tables
                            if (worksheet.Name == "Esito Tabelle")
                            {
                                List<All12TableT> tables = file.ErwinTables;
                                if (tables.Count > 0)
                                {
                                    int row = 2;
                                    foreach(var table in tables)
                                    {
                                        try
                                        {
                                            worksheet.Cells[row, 1].Value = table.SSAAll12.ToString();
                                            worksheet.Cells[row, 2].Value = table.NomeTabellaAll12;// file.All12Data.Where(x=>x.B_TabellaBFD == "").FirstOrDefault();
                                            worksheet.Cells[row, 3].Value = table.DescrizioneTabellaDCP; // !string.IsNullOrWhiteSpace(table.DescrizioneTabellaDCP) ? table.DescrizioneTabellaDCP : file.All12Data.Where(x => x.B_TabellaBFD == table.NomeTabellaAll12).Count() > 0 ? file.All12Data.Where(x => x.B_TabellaBFD == table.NomeTabellaAll12).FirstOrDefault().DG_DescrizioneLungaDatafile : string.Empty;
                                            worksheet.Cells[row, 4].Value = table.TipologiaInformazioneDCP;
                                            worksheet.Cells[row, 5].Value = table.AreaDCP;
                                            worksheet.Cells[row, 6].Value = table.History.ERRMessages.Count > 0 ? "KO" : table.History.WARMessages.Count > 0 ? "WARN" : "OK";
                                            worksheet.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(worksheet.Cells[row, 6].Text.ToUpper() == "OK" ? Color.Green : worksheet.Cells[row, 6].Text.ToUpper() == "WARN" ? Color.Yellow : Color.Red);

                                            string messages = string.Empty;
                                            foreach (var mes in table.History.ERRMessages)
                                            {
                                                messages += " ERR: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            foreach (var mes in table.History.WARMessages)
                                            {
                                                messages += " WARN: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            worksheet.Cells[row, 7].Value = messages.Trim();
                                        }
                                        catch(Exception exp)
                                        {
                                            Logger.PrintLC("Exception while writing data for Table " + table.NomeTabellaAll12 + " with message: " + exp.Message);
                                        }
                                        row++;
                                    }
                                }
                            }
                            // Attributes
                            if (worksheet.Name == "Esito Colonne")
                            {
                                List<All12AttributeT> attributes = file.ErwinAttributes;
                                List<string> tableNames = attributes.Select(x => x.Table.NomeTabellaAll12).Distinct().ToList();
                                int row = 2;
                                foreach (string table in tableNames)
                                {
                                    List<All12AttributeT> attributesInTable = attributes.Where(x => x.Table.NomeTabellaAll12 == table).ToList();
                                    foreach (var attribute in attributesInTable)
                                    {
                                        try
                                        {
                                            worksheet.Cells[row, 1].Value = attribute.Table.NomeTabellaAll12;
                                            worksheet.Cells[row, 2].Value = string.IsNullOrWhiteSpace(attribute.B_Campo) ? attribute.B_Campo_Err : attribute.B_Campo;
                                            worksheet.Cells[row, 3].Value = attribute.B_Formato;
                                            //worksheet.Cells[row, 4].Value = attribute.B_KeyPosition != null ? attribute.B_KeyPosition.ToString() : string.Empty;
                                            worksheet.Cells[row, 4].Value = string.IsNullOrWhiteSpace(attribute.B_Key_Excel) ? string.Empty : attribute.B_Key_Excel;
                                            worksheet.Cells[row, 5].Value = attribute.B_Note;
                                            // uncomment for uniformity
                                            // worksheet.Cells[row, 6].Value = attribute.B_NULL == 0 ? "TRUE" : attribute.B_NULL == 1 ? "FALSE" : attribute.NULL_Excel;
                                            worksheet.Cells[row, 6].Value = string.IsNullOrWhiteSpace(attribute.NULL_Excel) ? string.Empty : attribute.NULL_Excel;
                                            worksheet.Cells[row, 7].Value = attribute.CommentDCP; //attribute.G_DescrizioneCampo;
                                            worksheet.Cells[row, 8].Value = attribute.DatoSensibileDCP;
                                            worksheet.Cells[row, 9].Value = attribute.History.ERRMessages.Count > 0 ? "KO" : attribute.History.WARMessages.Count > 0 ? "WARN" : "OK";
                                            worksheet.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(worksheet.Cells[row, 9].Text.ToUpper() == "OK" ? Color.Green : worksheet.Cells[row, 9].Text.ToUpper() == "WARN" ? Color.Yellow : Color.Red);

                                            string messages = string.Empty;
                                            foreach (var mes in attribute.History.ERRMessages)
                                            {
                                                messages += " ERR: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            foreach (var mes in attribute.History.WARMessages)
                                            {
                                                messages += " WARN: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            worksheet.Cells[row, 10].Value = messages.Trim();
                                        }
                                        catch (Exception exp)
                                        {
                                            Logger.PrintLC("Exception while writing data for Field " + attribute.B_Campo + " of Table " + attribute.Table.NomeTabellaAll12 + " with message: " + exp.Message);
                                        }
                                        row++;
                                    }
                                }
                            }
                            // Relations
                            if (worksheet.Name == "Esito Relazioni")
                            {
                                // get History from erwin relation struct
                                file.GlobalRelations.GlobalRelazioni.ForEach(x => 
                                {
                                    List<All12RelationT> relAllList = file.ErwinRelations.Where(y => y.IdentificativoRelazione == x.ID).ToList();
                                    List<RelationT> xRel = x.Relazioni;
                                    foreach(var elem in relAllList)
                                    {
                                        try
                                        {
                                            if (xRel != null)
                                            {
                                                if (xRel.Where(z => z.CampoPadre == elem.All12CampoPadre && z.CampoFiglio == elem.All12CampoFiglio && z.TabellaPadre == elem.All12TabellaPadre && z.TabellaFiglia == elem.All12TabellaFiglia).Count() > 0)
                                                {
                                                    elem.History = xRel.Where(z => z.CampoPadre == elem.All12CampoPadre && z.CampoFiglio == elem.All12CampoFiglio && z.TabellaPadre == elem.All12TabellaPadre && z.TabellaFiglia == elem.All12TabellaFiglia).FirstOrDefault().HistoryH;
                                                }
                                            }
                                        }
                                        catch(Exception exp)
                                        {
                                            Logger.PrintC("Exception while converting old Historical messages to new format for relation " + elem.IdentificativoRelazione + ". Message: " + exp.Message);
                                        }
                                    }
                                });

                                List<All12RelationT> relations = file.ErwinRelations;
                                List<string> relationNames = relations.Select(x =>  x.IdentificativoRelazione).Distinct().ToList();
                                int row = 2;
                                foreach (string relationID in relationNames)
                                {
                                    List<All12RelationT> relationsInID = relations.Where(x => x.IdentificativoRelazione == relationID).ToList();
                                    foreach (var relation in relationsInID)
                                    {
                                        try
                                        {
                                            worksheet.Cells[row, 1].Value = relation.IdentificativoRelazione;
                                            worksheet.Cells[row, 2].Value = relation.All12TabellaPadre;
                                            worksheet.Cells[row, 3].Value = relation.All12TabellaFiglia;
                                            worksheet.Cells[row, 4].Value = relation.CardinalitaStr;
                                            worksheet.Cells[row, 5].Value = relation.All12CampoPadre;
                                            worksheet.Cells[row, 6].Value = relation.All12CampoFiglio;
                                            worksheet.Cells[row, 7].Value = relation.IdentificativaStr;
                                            worksheet.Cells[row, 8].Value = relation.Eccezioni;
                                            worksheet.Cells[row, 9].Value = relation.TipoRelazioneStr;
                                            worksheet.Cells[row, 10].Value = relation.Note;
                                            worksheet.Cells[row, 11].Value = relation.dcpID.ToString();

                                            worksheet.Cells[row, 12].Value = relation.History.ERRMessages.Count > 0 ? "KO" : relation.History.WARMessages.Count > 0 ? "WARN" : "OK";
                                            worksheet.Cells[row, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(worksheet.Cells[row, 12].Text.ToUpper() == "OK" ? Color.Green : worksheet.Cells[row, 12].Text.ToUpper() == "WARN" ? Color.Yellow : Color.Red);

                                            string messages = string.Empty;
                                            foreach (var mes in relation.History.ERRMessages)
                                            {
                                                messages += " ERR: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            foreach (var mes in relation.History.WARMessages)
                                            {
                                                messages += " WARN: " + mes + (mes.ToArray().Last() != '.' ? "." : string.Empty);
                                            }
                                            worksheet.Cells[row, 13].Value = messages.Trim();
                                        }
                                        catch (Exception exp)
                                        {
                                            Logger.PrintLC("Exception while writing data for Relation ID " + relation.IdentificativoRelazione + " with Parent Field " + relation.All12CampoPadre+ 
                                                " and Child Field " +relation.All12CampoFiglio + ", with message: " + exp.Message);
                                        }
                                        row++;
                                    }
                                }

                            }
                        }
                    }
                    catch (Exception exp)
                    {
                        Logger.PrintLC("Exception (message: " + exp.Message + ") while elaborating Outcome (Esiti) file " + fileToOpen + ". Skipping its full creation (something may have been written).", 2, ConfigFile.ERROR);
                    }
                    p.Save();
                    Logger.PrintLC("# ELABORATION OF OUTCOME (Esiti) FILE FOR " + file.SimpleName + " - END #", 2);
                }
            }
            else
            {
                Logger.PrintLC("Could not copy " + ConfigFile.ESITI_TEMPLATE + " to " + Path.GetDirectoryName(fileToOpen) + ". Not going to create an Outcome (Esiti) file.", 2, ConfigFile.ERROR);
            }
        }
        

        /// <summary>
        /// Checks formal validity of a DCP file
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool ValidateDCP(string file, FileAll12T fileAll12)
        {
            int genericError = 0;
            string testoLog = string.Empty;
            string TxtControlloNonPassato = string.Empty;
            bool sheetFoundTabelle = false;
            bool sheetFoundAttributi = false;
            bool sheetFoundRelazioni = false;
            bool columnsFoundTabelle = false;
            bool columnsFoundAttributi = false;
            bool columnsFoundRelazioni = false;
            int columns = 0;
            int[] check_sheet = new int[3] { 0, 0, 0 };
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            FileInfo fileDaAprire = new FileInfo(file);
            bool isXLS = false;
            Logger.PrintLC("Starting validation of file " + fileDaAprire.FullName + ".", 3, ConfigFile.INFO);
            //SEZIONE TEST D'APERTURA
            try
            {
                //test se il file è un temporaneo
                char[] opened = fileDaAprire.Name.ToCharArray();
                if (opened[0] == '~')
                {
                    Logger.PrintLC(fileDaAprire.Name + " is a temporary file. Will not elaborate.", 2, ConfigFile.ERROR);
                    genericError = 1;
                    goto ERROR;
                }
                //test se il file apribile
                if (!FileOps.isFileOpenable(file))
                {
                    Logger.PrintLC("Cannot open file " + fileDaAprire.Name + ". Will not elaborate.", 2, ConfigFile.ERROR);
                    genericError = 2;
                    goto ERROR;
                }
                //string extension = fileDaAprire.Extension.ToUpper();
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                    {
                        if (!ConvertXLStoXLSX(file))
                        {
                            genericError = 3;
                            goto ERROR;
                        }
                    }
                    isXLS = true;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
            }
            catch
            {
                genericError = 4;
                goto ERROR;
            }

            try
            {
                ExApp = new Excel.ApplicationClass();
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
            }
            catch
            {
                try
                {
                    ExApp.DisplayAlerts = false;
                    Logger.PrintLC(fileDaAprire.Name + " is already open. Close it and try again.", 2, ConfigFile.ERROR);
                }
                catch { }
                return false;
            }
            WB = p.Workbook;
            ws = WB.Worksheets;

            foreach (var worksheet in ws)
            {
                // SEZIONE TABELLE
                if (worksheet.Name == ConfigFile.TABELLE)
                {
                    columns = 0;
                    check_sheet[0] += 1;
                    sheetFoundTabelle = true;
                    columnsFoundTabelle = false;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_TABELLE; 
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_TABELLE; 
                            columnsPosition++)
                    {   
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._TABELLE.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._TABELLE[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] not found in column position [" + columnsPosition + "] of Sheet " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                                value = "[Field without value, position: " + columnsPosition + "]";
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] is not a valid column in Sheet " + worksheet.Name;
                            Logger.PrintLC(fileDaAprire.Name + ": [" + value + "] is not a valid column in Sheet " + worksheet.Name + ". File cannot be elaborated.", 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_TABELLE)
                        columnsFoundTabelle = true;
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tIncorrect number of columns in Sheet " + worksheet.Name;
                    }
                }

                // SEZIONE ATTRIBUTI
                if (worksheet.Name == ConfigFile.ATTRIBUTI)
                {
                    check_sheet[1] += 1;
                    columns = 0;
                    columnsFoundAttributi = false;
                    sheetFoundAttributi = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_ATTRIBUTI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._ATTRIBUTI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._ATTRIBUTI[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] not found in column position [" + columnsPosition + "] of Sheet " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                            {
                                value = "[Field without value, position: " + columnsPosition + "]";
                            }
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] is not a valid column in Sheet " + worksheet.Name;
                            Logger.PrintLC(fileDaAprire.Name + ": [" + value + "] is not a valid column in Sheet " + worksheet.Name + ". File cannot be elaborated.", 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_ATTRIBUTI)
                    {
                        columnsFoundAttributi = true;
                    }
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tIncorrect number of columns in Sheet " + worksheet.Name;
                    }
                }

                // SEZIONE RELAZIONI
                if (worksheet.Name == ConfigFile.RELAZIONI)
                {
                    check_sheet[2] += 1;
                    columns = 0;
                    columnsFoundRelazioni = false;
                    sheetFoundRelazioni = true;
                    for (int columnsPosition = ConfigFile.HEADER_COLONNA_MIN_RELAZIONI;
                            columnsPosition <= ConfigFile.HEADER_COLONNA_MAX_RELAZIONI;
                            columnsPosition++)
                    {
                        string value = worksheet.Cells[ConfigFile.HEADER_RIGA, columnsPosition].Text;
                        if (ConfigFile._RELAZIONI.ContainsKey(value))
                        {
                            columns += 1;
                            if (ConfigFile._RELAZIONI[value] != columnsPosition)
                            {
                                TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] not found in column position [" + columnsPosition + "] of Sheet " + worksheet.Name;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(value.Trim()))
                            {
                                value = "[Field without value, position: " + columnsPosition + "]";
                            }
                            TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\t[" + value + "] is not a valid column in Sheet " + worksheet.Name;
                            testoLog = fileDaAprire.Name + ": [" + value + "] is not a valid column in Sheet " + worksheet.Name + ". File cannot be elaborated.";
                            Logger.PrintLC(testoLog, 2, ConfigFile.ERROR);
                        }
                    }
                    if (columns == ConfigFile.HEADER_MAX_COLONNE_RELAZIONI)
                    {
                        columnsFoundRelazioni = true;
                    }
                    else
                    {
                        TxtControlloNonPassato = TxtControlloNonPassato + Environment.NewLine + "\t\tIncorrect number of columns in Sheet " + worksheet.Name;
                    }
                }
            }

            ERROR:
            try
            {
                WB.Dispose();
                p.Dispose();
            }
            catch
            {
            }

            //MngProcesses.KillAllOf(MngProcesses.ProcList("EXCEL"));
            string fileError = fileAll12.FileKO; //Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(file) + "_KO.txt");
            string fileCorrect = fileAll12.FileOK; //Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(file) + "_OK.txt");
            string fileStampa = String.Empty;

            if (genericError != 0 || check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1 ||
                sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                if (File.Exists(fileError))
                {
                    //FileOps.RemoveAttributes(fileError);
                    //File.Delete(fileError);
                }
                if (File.Exists(fileCorrect))
                {
                    string fullFile = File.ReadAllText(fileCorrect) + Environment.NewLine;
                    FileOps.RemoveAttributes(fileCorrect);
                    File.Delete(fileCorrect);
                    Logger.PrintF(fileError, fullFile);
                }
                fileStampa = fileError;
            }
            else
            {
                if (File.Exists(fileError))
                {
                    fileStampa = fileError;
                }
                if (File.Exists(fileCorrect))
                {
                    fileStampa = fileCorrect;
                }
            }


            //if (File.Exists(fileError))
            //{
            //    FileOps.RemoveAttributes(fileError);
            //    File.Delete(fileError);
            //}
            //if (File.Exists(fileCorrect))
            //{
            //    FileOps.RemoveAttributes(fileCorrect);
            //    File.Delete(fileCorrect);
            //}
            //string fileStampa = String.Empty;

            Logger.PrintF(fileStampa, Environment.NewLine + "## er_driveup_lake ## - DCP Section -" + Environment.NewLine + "Check formal formatting: " + Environment.NewLine, false);
            if (genericError != 0)
            {
                switch (genericError)
                {
                    case 1:
                        Logger.PrintF(fileStampa, "File is temporary. Cannot be elaborated.", false);
                        break;
                    case 2:
                        Logger.PrintF(fileStampa, "File is unopenable. It's probably badly formatted (es: links to external tables). Cannot be elaborated.", false);
                        break;
                    case 3:
                        Logger.PrintF(fileStampa, "Was unable to convert file to '.XLSX' format. Cannot be elaborated.", false);
                        break;
                    case 4:
                        Logger.PrintF(fileStampa, "Unexpected error while opening file. Cannot be elaborated.", false);
                        break;
                }
                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
                return false;
            }

            //if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1 ||
            //    sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
            //    columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            //{
            //    fileStampa = fileError;
            //}
            //else
            //{
            //    fileStampa = fileCorrect;
            //}

            //Logger.PrintF(fileStampa, "er_driveup_lake – Caricamento Excel su ERwin", true);

            if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1)
            {
                Logger.PrintLC(fileDaAprire.Name + ": cannot be elaborated: a Sheet is not present or one of the Clumns does not conform.", 2, ConfigFile.ERROR);
                Logger.PrintF(fileStampa, fileDaAprire.Name + ": cannot be elaborated: one or more Sheets are not present :", false);
                if (check_sheet[0] != 1)
                {
                    Logger.PrintLC("\t\t'Foglio Censimento Tabelle' is not present.");
                    Logger.PrintF(fileStampa, "'Foglio Censimento Tabelle' is not present", false);
                }
                if (check_sheet[1] != 1)
                {
                    Logger.PrintLC("\t\t'Foglio Censimento Attributi' is not present.");
                    Logger.PrintF(fileStampa, "'Foglio Censimento Attributi' is not present.", false);
                }
                if (check_sheet[2] != 1)
                {
                    Logger.PrintLC("\t\t'Foglio Relazioni-ModelloDatiLegacy' is not present.");
                    Logger.PrintF(fileStampa, "'Foglio Relazioni-ModelloDatiLegacy' is not present.", false);
                }

                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
            }
            if (sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                Logger.PrintLC(fileDaAprire.Name + ": file could not be processed: Columns or Sheets are not in the expected format.", 2, ConfigFile.ERROR);
                Logger.PrintF(fileStampa, "Columns or Sheets are not in the expected format:", false);
                string[] val = TxtControlloNonPassato.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                if (val.Count() > 0)
                {
                    foreach (string valC in val)
                    {
                        if (!string.IsNullOrWhiteSpace(valC))
                            Logger.PrintF(fileStampa, valC, false);
                    }
                }
                else
                {
                    Logger.PrintF(fileStampa, TxtControlloNonPassato, false);
                }
            }

            if (check_sheet[0] != 1 || check_sheet[1] != 1 || check_sheet[2] != 1 ||
                sheetFoundTabelle != true || sheetFoundAttributi != true || sheetFoundRelazioni != true ||
                columnsFoundTabelle != true || columnsFoundAttributi != true || columnsFoundRelazioni != true)
            {
                if (isXLS == true)
                {
                    if (File.Exists(fileDaAprire.FullName))
                    {
                        File.Delete(fileDaAprire.FullName);
                    }
                }
                Logger.PrintLC("File NOT formatted correctly.", 4, ConfigFile.INFO);
                return false;
            }

            Logger.PrintLC("File formatted correctly.", 4, ConfigFile.INFO);
            Logger.PrintF(fileStampa, "File formatted correctly.", false);
            return true;
        }

        /// <summary>
        /// Searches for at least 1 Attribute in the Attribute Sheet for each Table 'name'
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="nome"></param>
        /// <returns></returns>
        /*
        public static bool TestAttributesExist(ExcelWorksheets ws, string nome)
        {
            bool attrExist = false;
            foreach (var sheetAtt in ws)
            {
                if (sheetAtt.Name == ConfigFile.ATTRIBUTI)
                {
                    bool FilesEndAtt = false;
                    for (int RowPosAtt = ConfigFile.HEADER_RIGA + 1;
                            FilesEndAtt != true;
                            RowPosAtt++)
                    {
                        string nomeAtt = sheetAtt.Cells[RowPosAtt, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text.Trim();
                        if (!string.IsNullOrWhiteSpace(nomeAtt))
                        {
                            if (nome == nomeAtt)
                            {
                                attrExist = true;
                                break;
                            }
                        }
                        else
                        {

                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossimeAtt = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(sheetAtt.Cells[RowPosAtt + i, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text))
                                prossimeAtt++;
                        }
                        if (prossimeAtt == 10)
                            FilesEndAtt = true;
                        //******************************************
                    }
                    break;
                }
            }
            return attrExist;
        }
        */
        /*
        /// <summary>
        /// Cicles all All12 files to validate them, setting their state
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        public static List<FileAll12T> FullValidateAll12(List<FileAll12T> fileList)
        {
            if (fileList.Where(x => x.Correct == true).ToList().Count > 0)
            {
                Logger.PrintLC("## 'Allegato 12' FILES VALIDATION SECTION - START ##", 2);
                fileList.ForEach(x =>
                {
                    x.Correct = ValidateAllegato12(new FileInfo(x.FullName), x.FullNameOriginal, x);
                });
                Logger.PrintLC("## 'Allegato 12' FILES VALIDATION SECTION - END ##", 2);
            }
            return fileList;
        }

        /// <summary>
        /// Cicles all DCP files to validate them, setting their state
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        public static List<FileAll12T> FullValidateDCP(List<FileAll12T> fileList)
        {
            if (fileList.Where(x => x.Correct == true).ToList().Count > 0)
            {
                Logger.PrintLC("## 'DCP' FILES VALIDATION SECTION - START ##", 2);
                fileList.ForEach(x =>
                {
                    //if (x.Correct)
                        x.DCP.Correct = ValidateDCP(x.DCP.FullName, x);
                });
                Logger.PrintLC("## 'DCP' FILES VALIDATION SECTION - END ##", 2);
            }
            return fileList;
        }

        /// <summary>
        /// Tests formal validity (formatting) of 'Allegato 12''s file
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool ValidateAllegato12(FileInfo file, string fileNameOriginal, FileAll12T fileAll12)
        {
            Logger.PrintLC("Starting validation of file " + fileNameOriginal + ".", 3, ConfigFile.INFO);
            string fileToOpen = string.Empty;
            // #########################
            // Convert if Excel is XLS
            if (file.Extension.ToUpper() == ".XLS")
            {
                if (ConvertXLStoXLSX(file.FullName))
                    fileToOpen = Path.ChangeExtension(file.FullName, ".xlsx");
                else
                {
                    Logger.PrintLC("File has an old format (tipically .XLS). The attempt to convert it FAILED. Skipping it.", 3, ConfigFile.ERROR);
                    return false;
                }
            }
            else
                fileToOpen = file.FullName;
            // #########################

            if (!isFileOpenable(fileToOpen))
            {
                Logger.PrintLC("File is not openable.", 3, ConfigFile.INFO);
                return false;
            }
            try
            {
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;

                // **********************
                try
                {
                    ExApp = new Excel.ApplicationClass();
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(new FileInfo(fileToOpen));
                    ExApp.DisplayAlerts = true;
                }
                catch
                {
                    Logger.PrintLC(file.Name + " already opened by another application. Close it and try again.", 3, ConfigFile.ERROR);
                    return false;
                }
                WB = p.Workbook;
                ws = WB.Worksheets;
                // ********************

                bool FilesEnd = false;
                int EmptyRow = 0;
                string errorHistory = string.Empty;
                string PrintValidationFile = string.Empty;
                // Searches our Sheet
                foreach (var worksheet in ws)
                {
                    if (worksheet.Name == ConfigFile.ALL12_SHEET_NAME)
                    {
                        // Cycles through all colums, validating each
                        for (int ColPos = ConfigFile.ALL12_FIRST_COLUMN;
                                ColPos <= ConfigFile.ALL12_LAST_COLUMN;
                                ColPos++)
                        {
                            string expectedValue = ConfigFile._ALLEGATO12.Where(x => x.Key == ColPos).Select(y => y.Value).FirstOrDefault().Trim();
                            string effectiveValue = worksheet.Cells[ConfigFile.ALL12_HEADER_ROW, ColPos].Text.ToString().Trim();
                            // Checks for correctness of Column's value vs. expected
                            if (!string.Equals(expectedValue, effectiveValue))
                            {
                                string error = "Was expecting column [" + expectedValue + "], found [" + effectiveValue + "] on position " + ColPos + ".";
                                Logger.PrintLC(error, 4, ConfigFile.ERROR);
                                if (!string.IsNullOrWhiteSpace(errorHistory))
                                    errorHistory += Environment.NewLine;
                                else
                                    errorHistory = "Formatting error in file " + file.Name + "." + Environment.NewLine;
                                errorHistory += error;
                            }
                        }
                        // We're getting path + name of Validation File
                        // Tuple Items:
                        //  Item1 = file OK
                        //  Item2 = file KO
                        Tuple<string, string> validationFileTuple = GetResponseFile(new FileInfo(fileToOpen));
                        fileAll12.FileOK = validationFileTuple.Item1;
                        fileAll12.FileKO = validationFileTuple.Item2;
                        string FileHeader = "## er_driveup_lake ## - Allegato 12 Section -" + Environment.NewLine + "Check formal formatting: " + Environment.NewLine;

                        bool correct = false;
                        if (string.IsNullOrWhiteSpace(errorHistory))
                        {
                            correct = true;
                            PrintValidationFile = validationFileTuple.Item1;
                            errorHistory = "File OK";
                        }
                        else
                        {
                            PrintValidationFile = validationFileTuple.Item2;
                            errorHistory += Environment.NewLine + "File KO";
                        }

                        FileHeader += Environment.NewLine + errorHistory;
                        Logger.PrintF(PrintValidationFile, FileHeader);

                        if (correct)
                        {
                            Logger.PrintLC("File formatted correctly.", 3, ConfigFile.INFO);
                            return true;
                        }
                        else
                        {
                            Logger.PrintLC("File NOT formatted correctly.", 3, ConfigFile.INFO);
                            return false;
                        }
                    }
                }
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Exception while validating file " + file.Name + ". Error n°" + exp.HResult + ", Message: " + exp.Message, 2, ConfigFile.ERROR);
                return false;
            }
            Logger.PrintLC("File NOT formatted correctly.", 4, ConfigFile.INFO);
            return false;
        }


        /// <summary>
        /// Loads all data from a single Excel row
        /// into a 'All12DataT' record
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="worksheet"></param>
        /// <param name="RowPos"></param>
        /// <returns></returns>
        private static All12DataT All12_LoadDataFromExcel(All12DataT dataRow, ExcelWorksheet worksheet, int RowPos)
        {
            dataRow.R_Row = RowPos;
            dataRow.DG_Host = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["HOST / DIP"]].Text.ToString().Trim();
            dataRow.DG_Clone = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CLONE/ MB"]].Text.ToString().Trim();
            dataRow.DG_DB = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["DB / FILE"]].Text.ToString().Trim();
            dataRow.DG_SSA = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["SSA"]].Text.ToString().Trim();
            dataRow.LG_TabellaMasterOrigine = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["TABELLA MASTER DI ORIGINE"]].Text.ToString().Trim();
            dataRow.LG_CampoOrigine = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CAMPO DI ORIGINE"]].Text.ToString().Trim();

            dataRow.LG_Oggetto = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["OGGETTO"]].Text.ToString().Trim();
            dataRow.LG_Campo = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CAMPO1"]].Text.ToString().Trim();
            // Missing fields:
            // FORMATO1
            // KEY1
            dataRow.DG_NomeHost = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NOME HOST (*)"]].Text.ToString().Trim();
            dataRow.DG_PathDatafile = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["PATH DATAFILE (*)"]].Text.ToString().Trim();
            dataRow.DG_NomeFileProdotto = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NOME FILE PRODOTTO (*)"]].Text.ToString().Trim();
            dataRow.DG_Campo = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CAMPO2"]].Text.ToString().Trim();
            dataRow.DG_FormatoODBC = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["FORMATO ODBC"]].Text.ToString().Trim();
            dataRow.DG_Lunghezza = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["LUNGHEZZA"]].Text.ToString().Trim();
            dataRow.DG_Decimali = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["DECIMALI"]].Text.ToString().Trim();
            dataRow.DG_StrutturaData = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["STRUTTURA DATA"]].Text.ToString().Trim();
            dataRow.DG_Key = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["KEY2"]].Text.ToString().Trim();
            dataRow.DG_Unique = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["UNIQUE"]].Text.ToString().Trim();
            dataRow.DG_Null = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NULL1"]].Text.ToString().Trim();
            dataRow.DG_Posizione = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["POSIZIONE"]].Text.ToString().Trim();
            dataRow.DG_Offset = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["OFFSET (FACOLTATIVO)"]].Text.ToString().Trim();
            dataRow.DG_DescrizioneCampo = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["DESCRIZIONE CAMPO"]].Text.ToString().Trim();
            dataRow.DG_DescrizioneBreveDatafile = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["DESCRIZIONE BREVE DATAFILE"]].Text.ToString().Trim();
            dataRow.DG_DescrizioneLungaDatafile = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["DESCRIZIONE LUNGA DATAFILE"]].Text.ToString().Trim();
            dataRow.DG_CodiceDatafile = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CODICE DATAFILE"]].Text.ToString().Trim();
            dataRow.DG_TabellaDominio = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["TABELLA DI DOMINIO"]].Text.ToString().Trim();
            dataRow.DG_Note = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NOTE"]].Text.ToString().Trim();

            dataRow.B_TabellaBFD = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["Tabella BFD"]].Text.ToString().Trim();
            dataRow.B_Campo = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["CAMPO3"]].Text.ToString().Trim();
            dataRow.B_Formato = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["FORMATO2"]].Text.ToString().Trim();
            dataRow.B_Key = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["KEY3"]].Text.ToString().Trim();
            dataRow.B_Note = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NOTE2"]].Text.ToString().Trim();
            dataRow.B_TabellaDominioBFD = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["Tabella di dominio BFD"]].Text.ToString().Trim();
            dataRow.B_Null = worksheet.Cells[RowPos, ConfigFile._ALLEGATO12_ref["NULL2"]].Text.ToString().Trim();

            return dataRow;
        }


        /// <summary>
        /// Formal checks for All12 fields from Excel
        /// </summary>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        private static Tuple<All12DataT, bool> All12_CheckValidity(All12DataT dataRow)
        {
            bool error = false;
            if (string.IsNullOrWhiteSpace(dataRow.B_Campo) &&
                string.IsNullOrWhiteSpace(dataRow.B_Formato) &&
                string.IsNullOrWhiteSpace(dataRow.B_Key) &&
                string.IsNullOrWhiteSpace(dataRow.B_Note) &&
                string.IsNullOrWhiteSpace(dataRow.B_Null) &&
                string.IsNullOrWhiteSpace(dataRow.B_TabellaBFD) &&
                string.IsNullOrWhiteSpace(dataRow.B_TabellaDominioBFD) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Campo) &&
                string.IsNullOrWhiteSpace(dataRow.LG_CampoOrigine) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Clone) &&
                string.IsNullOrWhiteSpace(dataRow.DG_CodiceDatafile) &&
                string.IsNullOrWhiteSpace(dataRow.DG_DB) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Decimali) &&
                string.IsNullOrWhiteSpace(dataRow.DG_DescrizioneBreveDatafile) &&
                string.IsNullOrWhiteSpace(dataRow.DG_DescrizioneCampo) &&
                string.IsNullOrWhiteSpace(dataRow.DG_DescrizioneLungaDatafile) &&
                string.IsNullOrWhiteSpace(dataRow.DG_FormatoODBC) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Host) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Key) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Lunghezza) &&
                string.IsNullOrWhiteSpace(dataRow.DG_NomeFileProdotto) &&
                string.IsNullOrWhiteSpace(dataRow.DG_NomeHost) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Null) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Offset) &&
                string.IsNullOrWhiteSpace(dataRow.DG_PathDatafile) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Posizione) &&
                string.IsNullOrWhiteSpace(dataRow.DG_SSA) &&
                string.IsNullOrWhiteSpace(dataRow.DG_StrutturaData) &&
                string.IsNullOrWhiteSpace(dataRow.DG_TabellaDominio) &&
                string.IsNullOrWhiteSpace(dataRow.LG_TabellaMasterOrigine) &&
                string.IsNullOrWhiteSpace(dataRow.DG_Unique) &&
                string.IsNullOrWhiteSpace(dataRow.LG_Campo) &&
                string.IsNullOrWhiteSpace(dataRow.LG_Oggetto))
            {
                // EMPTY ROW case
                error = false;
                dataRow.History.EmptyRow = true;
                dataRow.History.ERRMessages.Clear();
            }
            else
            {
                if (string.IsNullOrWhiteSpace(dataRow.DG_SSA))
                {
                    error = true;
                    dataRow.History.ERRMessages.Add("Mandatory field 'SSA (blue)' is empty.");
                }
                if (string.IsNullOrWhiteSpace(dataRow.B_TabellaBFD))
                {
                    error = true;
                    dataRow.History.ERRMessages.Add("Mandatory field 'Tabella BFD (blue)' is empty.");
                }
                if (string.IsNullOrWhiteSpace(dataRow.B_Campo))
                {
                    error = true;
                    dataRow.History.ERRMessages.Add("Mandatory field 'Campo (blue)' is empty.");
                }

                if (string.IsNullOrWhiteSpace(dataRow.B_Formato))
                {
                    error = true;
                    dataRow.History.ERRMessages.Add("Mandatory field 'Formato (blue)' is empty.");
                }
            }
            return new Tuple<All12DataT,bool> (dataRow, error);
        }


        /// <summary>
        /// Writes on the Excel all OK/KO conditions, 
        /// with the relative Message(s)
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="dataList"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /*
        private static bool All12_WriteErrorsOnExcel(ExcelWorksheet worksheet, List<All12DataT> dataList, string fileName)
        {
            Logger.PrintLC("Writing historical values to Excel file " + fileName, 4, ConfigFile.INFO);
            try
            {
                // For each record in the list we print the 
                // historical values in the correspective row 
                dataList.ForEach(record =>
                {
                    if (record.History.Validated == true)
                    {
                        // We collect all historical messages
                        string history = string.Empty;
                        if (record.History.ERRMessages.Count > 0)
                        {
                            record.History.ERRMessages.ForEach(message =>
                            {
                                history += message + " ";
                            });
                        }
                        else
                        {
                            history = string.Empty;
                        }

                        // Setting the columns formats
                        // OK/KO Column
                        worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.Font.Bold = true;
                        worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Column(ConfigFile.ALL12_COLUMN_OKKO).Width = 10;
                        // Historic Column
                        worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_HISTORY].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        worksheet.Column(ConfigFile.ALL12_COLUMN_HISTORY).Width = 100;
                        // We check if it was evaluated as an 'empty row'
                        if (record.History.EmptyRow)
                        {
                            worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 255));
                            worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Value = "Vuota";
                            // Historic Column
                            worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_HISTORY].Value = history;
                        }
                        else
                        {
                            // We check if the record was OK/KO
                            if (record.History.Valid)
                            {
                                // OK/KO Column
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Value = "OK";
                                // Historic Column
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_HISTORY].Value = history;
                            }
                            else
                            {
                                // We format all the individual messages into one string
                                // We print the historical messages in the excel row
                                // OK/KO Column
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_OKKO].Value = "OK";
                                // Historic Column
                                worksheet.Cells[(int)record.R_Row, ConfigFile.ALL12_COLUMN_HISTORY].Value = history;
                            }
                        }
                    }
                });
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Exception while writing historical data to excel " + fileName + ". Message: " + exp.Message,
                    4,
                    ConfigFile.ERROR);
            }
            return true;
        }
        */
        /*
        /// <summary>
        /// Reads all 'Allegato 12' file records and loads 
        /// its data in the 'FileAll12' structure
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static FileAll12T GetAllDataFromAll12(FileAll12T file)
        {
            Logger.PrintLC("# Loading all data (All12) from "+ file.FullName +".", 3, ConfigFile.INFO);
            // Check if XLS
            string fileToOpen = string.Empty;
            if (file.isXLS)
                fileToOpen = Path.ChangeExtension(file.FullName, ".xlsx");
            else
                fileToOpen = file.FullName;
            // ############
            if (!isFileOpenable(fileToOpen))
            {
                Logger.PrintLC("File " + fileToOpen + " is not openable. Skipping it.", 2, ConfigFile.ERROR);
                return null;
            }
            string fileName = Path.GetFileName(file.FullName);
            try
            {
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;

                //we try to open an Excel object
                try
                {
                    ExApp = new Excel.ApplicationClass();
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(new FileInfo(fileToOpen));
                    ExApp.DisplayAlerts = true;
                }
                catch
                {
                    ExApp.DisplayAlerts = false;
                    Logger.PrintLC(fileToOpen + " is opened by another application. Close it and try again.", 4, ConfigFile.ERROR);
                    return null;
                }
                WB = p.Workbook;
                ws = WB.Worksheets;
                // ********************

                bool FilesEnd = false;
                string errorHistory = string.Empty;
                string PrintValidationFile = string.Empty;
                List<All12DataT> dataList = new List<All12DataT>();
                // Searches our Sheet
                foreach (var worksheet in ws)
                {
                    if (worksheet.Name == ConfigFile.ALL12_SHEET_NAME)
                    {
                        // Cycles through all colums, validating each
                        for (int RowPos = ConfigFile.ALL12_HEADER_ROW + 2;
                                                    FilesEnd != true;
                                                    RowPos++)
                        {
                            All12DataT dataRow = new All12DataT();
                            dataRow.FileName = fileName;
                            dataRow.History = new HistoryT();
                            bool error = false;
                            // SECTION: data load
                            dataRow = All12_LoadDataFromExcel(dataRow, worksheet, RowPos);
                            // SECTION: formal validity check
                            // Item1: dataRow record
                            // Item2: error bool
                            Tuple<All12DataT, bool> checkValidity = All12_CheckValidity(dataRow);
                            dataRow = checkValidity.Item1;
                            error = checkValidity.Item2;
                            // ##############################
                            if(!dataRow.History.EmptyRow)
                                dataList.Add(dataRow);

                            //******************************************
                            // Verifies the next 10 rows to find end of table
                            int prossime = 0;
                            for (int i = 1; i < 11; i++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ALLEGATO12_ref["Tabella BFD"]].Text) && 
                                    string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ALLEGATO12_ref["CAMPO3"]].Text))
                                    prossime++;
                            }
                            if (prossime == 10)
                                FilesEnd = true;
                            //******************************************
                        }
                        // ## When all data on sheet has been read ->
                        // we set the historical values for each record
                        dataList = Funct.All12_SetHistoryValues(dataList);
                        // we set possible alternative values relative to DCP Table names
                        dataList = Funct.All12_ParseTableName(dataList);
                        // we assign our full data list to the parent structure
                        file.All12Data = dataList;
                        // we write all historical data in the original Excel (-- DEPRECATED --)
                        //All12_WriteErrorsOnExcel(worksheet, dataList, fileName);
                    }
                }

                // We try to save the file (important)
                try
                {
                    // DEPRECATED
                    //Logger.PrintLC("Saving validated file " + fileName, 3, ConfigFile.INFO);
                    //p.Save();
                }
                catch
                {
                    Logger.PrintLC("Exception while saving Excel " + fileName, 3, ConfigFile.ERROR);
                }

                // We try to dispose all Excel objects
                try
                {
                    ws.Dispose();
                    WB = null;
                    p = null;
                    ExApp = null;
                }
                catch
                {
                    Logger.PrintLC("Exception while disposing Excel " + fileName, 3, ConfigFile.WARNING);
                }
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Exception while loading 'All12' data from Excel " + file.FullName + ". Message: " + exp.Message, 3, ConfigFile.ERROR);
            }
            Logger.PrintLC("All data loaded.", 3, ConfigFile.INFO);
            return file;
        }


        /// <summary>
        /// Reads from an Excel and builds a structure (list)
        /// that contains all data
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static FileAll12T GetAllDataFromDCP(FileAll12T file)
        {
            Logger.PrintLC("# Loading all data (DCP) from " + file.DCP.FullName + ".", 3, ConfigFile.INFO);
            string fileToOpen = string.Empty;
            if (!isFileOpenable(file.DCP.FullName))
                return null;
            else
            {
                if (new FileInfo(file.DCP.FullName).Extension.ToUpper() == ".XLS")
                    fileToOpen = Path.ChangeExtension(file.DCP.FullName, ".xlsx");
                else
                    fileToOpen = file.DCP.FullName;
            }

            string ActualState = string.Empty;

            try
            {
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;

                // **********************
                try
                {
                    ExApp = new Excel.ApplicationClass();
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(new FileInfo(fileToOpen));
                    //ExApp.DisplayAlerts = true;
                }
                catch
                {
                    try
                    {
                        ExApp.DisplayAlerts = false;
                        Logger.PrintLC(fileToOpen + " is opened by another application. Close it and try again.", 3, ConfigFile.ERROR);
                    }
                    catch { }
                    return null;
                }
                WB = p.Workbook;
                ws = WB.Worksheets;
                // ********************

                bool FilesEnd = false;
                int EmptyRow = 0;
                string errorHistory = string.Empty;
                string PrintValidationFile = string.Empty;
                List<EntityT> dataEntity = new List<EntityT>();
                List<AttributeT> dataAttribute = new List<AttributeT>();
                List<RelationT> dataRelation = new List<RelationT>();
                // Searches our Sheet
                foreach (var worksheet in ws)
                {
                    // LOAD ALL TABLES
                    if (worksheet.Name == ConfigFile.TABELLE)
                    {
                        Logger.PrintLC("Loading Tables data...", 4, ConfigFile.INFO);
                        ActualState = "Tables";
                        // Cycles through all colums, validating each
                        for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                                                    FilesEnd != true;
                                                    RowPos++)
                        {
                            EntityT dataRow = new EntityT();
                            dataRow.HistoryH = new HistoryT();
                            //string expectedValue = ConfigFile._TABELLE.Where(x => x.Key == ColPos).Select(y => y.Value).First().Trim();
                            dataRow.Row = RowPos;
                            try
                            {
                                dataRow.SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text.ToString().Trim();
                                dataRow.HostName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text.ToString().Trim();
                                dataRow.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.ToString().Trim();
                                dataRow.Schema = worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text.ToString().Trim();
                                dataRow.TableName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text.ToString().Trim();
                                dataRow.TableDescr = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text.ToString().Trim();
                                dataRow.InfoType = worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text.ToString().Trim();
                                dataRow.TableLimit = worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text.ToString().Trim();
                                dataRow.TableGranularity = worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text.ToString().Trim();
                                dataRow.FlagBFD = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text.ToString().Trim();
                                dataRow.State = worksheet.Cells[RowPos, ConfigFile._TABELLE["State"]].Text.ToString().Trim().ToUpper() == "OK" ?
                                    true :
                                    false;
                            }
                            catch (Exception exp)
                            {
                                Logger.PrintLC("Exception while loading " + ActualState + " data from file " + fileToOpen + ". Message: " + exp.Message, 4, ConfigFile.ERROR);
                            }
                            dataEntity.Add(dataRow);
                            //******************************************
                            // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                            int prossime = 0;
                            for (int i = 1; i < 11; i++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Nome Tabella"]].Text) && string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Flag BFD"]].Text))
                                    prossime++;
                            }
                            if (prossime == 10)
                                FilesEnd = true;
                            //******************************************
                        }
                        FilesEnd = false;
                        file.DCP.EntityData = dataEntity;
                    }

                    // LOAD ALL ATTRIBUTES
                    if (worksheet.Name == ConfigFile.ATTRIBUTI)
                    {
                        Logger.PrintLC("Loading Attributes data...", 4, ConfigFile.INFO);
                        ActualState = "Attributes";
                        // Cycles through all colums, validating each
                        for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                                                    FilesEnd != true;
                                                    RowPos++)
                        {
                            AttributeT dataRow = new AttributeT();
                            //string expectedValue = ConfigFile._TABELLE.Where(x => x.Key == ColPos).Select(y => y.Value).First().Trim();
                            int tempInt = 0;
                            dataRow.Row= RowPos;
                            try
                            {
                                dataRow.SSA = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["SSA"]].Text.ToString().Trim();
                                dataRow.Area = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Area"]].Text.ToString().Trim();
                                dataRow.NomeTabellaLegacy = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text.ToString().Trim();
                                dataRow.NomeCampoLegacy = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text.ToString().Trim();
                                dataRow.DefinizioneCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text.ToString().Trim();
                                dataRow.TipologiaTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati"]].Text.ToString().Trim();
                                dataRow.DataType = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Datatype"]].Text.ToString().Trim();
                                dataRow.Lunghezza = int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text.ToUpper().Trim(), out tempInt) ? (int?)tempInt : null;
                                dataRow.Decimali = int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text.ToUpper().Trim(), out tempInt) ? (int?)tempInt : null;
                                dataRow.ChiaveStr = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave"]].Text.ToString().Trim();
                                dataRow.Unique = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Unique"]].Text.ToString().Trim();
                                dataRow.ChiaveLogica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave Logica"]].Text.ToString().Trim();
                                dataRow.MandatoryFlagStr = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Mandatory Flag"]].Text.ToString().Trim();
                                dataRow.Dominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dominio"]].Text.ToString().Trim();
                                dataRow.ProvenienzaDominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Provenienza dominio "]].Text.ToString().Trim();
                                dataRow.Note = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Note"]].Text.ToString().Trim();
                                dataRow.Storica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Storica"]].Text.ToString().Trim();
                                dataRow.DatoSensibile = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dato Sensibile"]].Text.ToString().Trim();
                                dataRow.State = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["State"]].Text.ToString().Trim().ToUpper() == "OK" ?
                                    true :
                                    false;
                            }
                            catch(Exception exp)
                            {
                                Logger.PrintLC("Exception while loading " + ActualState + " data from file " + fileToOpen + ". Message: " + exp.Message, 4, ConfigFile.ERROR);
                            }
                            dataAttribute.Add(dataRow);
                            //******************************************
                            // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                            int prossime = 0;
                            for (int i = 1; i < 11; i++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text) && string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text))
                                    prossime++;
                            }
                            if (prossime == 10)
                                FilesEnd = true;
                            //******************************************
                        }
                        FilesEnd = false;
                        file.DCP.AttributeData = dataAttribute;
                    }

                    // LOAD ALL RELATIONS
                    if (worksheet.Name == ConfigFile.RELAZIONI)
                    {
                        Logger.PrintLC("Loading Relations data...", 4, ConfigFile.INFO);
                        ActualState = "Relations";
                        // Cycles through all colums, validating each
                        for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                                                    FilesEnd != true;
                                                    RowPos++)
                        {
                            RelationT dataRow = new RelationT();
                            //string expectedValue = ConfigFile._TABELLE.Where(x => x.Key == ColPos).Select(y => y.Value).First().Trim();
                            int tempInt = 0;
                            dataRow.Row = RowPos;
                            try
                            {
                                dataRow.IdentificativoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativo relazione"]].Text.ToString().Trim();
                                dataRow.TabellaPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Padre"]].Text.ToString().Trim();
                                dataRow.TabellaFiglia = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Figlia"]].Text.ToString().Trim();
                                dataRow.CardinalitaStr = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Cardinalità"]].Text.ToString().Trim();
                                dataRow.CampoPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Padre"]].Text.ToString().Trim();
                                dataRow.CampoFiglio = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Figlio"]].Text.ToString().Trim();
                                dataRow.IdentificativaStr = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativa"]].Text.ToString().Trim();
                                dataRow.Eccezioni = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text.ToString().Trim();
                                dataRow.TipoRelazioneStr = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tipo Relazione"]].Text.ToString().Trim();
                                dataRow.Note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text.ToString().Trim();
                                dataRow.State = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["State"]].Text.ToString().Trim().ToUpper() == "OK" ?
                                    true :
                                    false;
                            }
                            catch (Exception exp)
                            {
                                Logger.PrintLC("Exception while loading " + ActualState + " data from file " + fileToOpen + ". Message: " + exp.Message, 3, ConfigFile.ERROR);
                            }

                            dataRelation.Add(dataRow);
                            //******************************************
                            // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                            int prossime = 0;
                            for (int i = 1; i < 11; i++)
                            {
                                if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._RELAZIONI["Tabella Padre"]].Text) && string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._RELAZIONI["Tabella Figlia"]].Text))
                                    prossime++;
                            }
                            if (prossime == 10)
                                FilesEnd = true;
                            //******************************************
                        }
                        FilesEnd = false;
                        file.DCP.RelationData = dataRelation;
                    }
                    FilesEnd = false;
                }
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Exception while loading " + ActualState + " data from file " + fileToOpen + ". Message: " + exp.Message, 3, ConfigFile.INFO);
                return file;
            }
            Logger.PrintLC("All data loaded.", 3, ConfigFile.INFO);
            return file;
        }

        /// <summary>
        /// Creates the path and file name for both the OK and KO text file,
        /// to allegate to a validated file
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <returns></returns>
        public static Tuple<string, string> GetResponseFile(FileInfo fileDaAprire)
        {
            string fileError = Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(fileDaAprire.Name) + "_KO.txt");
            string fileCorrect = Path.Combine(fileDaAprire.DirectoryName, Path.GetFileNameWithoutExtension(fileDaAprire.Name) + "_OK.txt");
            if (File.Exists(fileError))
            {
                FileOps.RemoveAttributes(fileError);
                File.Delete(fileError);
            }
            if (File.Exists(fileCorrect))
            {
                FileOps.RemoveAttributes(fileCorrect);
                File.Delete(fileCorrect);
            }
            return Tuple.Create(fileCorrect, fileError);
        }
        */

        /// <summary>
        /// Reads and processes Table data from excel's 'TABELLE' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /*
        public static List<EntityT> ReadXFileEntity(FileInfo fileDaAprire, string db, string sheet = ConfigFile.TABELLE)
        {
            string file = fileDaAprire.FullName;
            List<EntityT> listaFile = new List<EntityT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Tables. File " + fileDaAprire.Name + " doesn't exist.", 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                {
                    return listaFile = null;
                }
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Reading Tables. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            
            bool FilesEnd = false;
            int EmptyRow = 0;
            //int columns = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string error = string.Empty;
                        string nome = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text;
                        if (listaFile.Exists(x => x.TableName == nome))
                        {
                            incorrect = true;
                            error += "Una tabella con lo stesso NOME TABELLA è già presente. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        // #################################
                        // TEST ESISTENZA ATTRIBUTI PER LA TABELLA
                        bool attrExist = TestAttributesExist(ws, nome);     // 'attributes exist for table' flag
                        if (attrExist == false)
                        {
                            incorrect = true;
                            error += "La Tabella non possiede Attributi; non verrà inserita. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        // #################################
                        string SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text;
                        if (string.IsNullOrWhiteSpace(nome))
                        {
                            incorrect = true;
                            error += "Valore di SSA mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        
                        string Descr_Tab = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text;
                        if (string.IsNullOrWhiteSpace(nome))
                        {
                            incorrect = true;
                            error += "Valore di Descrizione Tabella mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        string flag = worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text;
                        if (string.IsNullOrWhiteSpace(nome))
                        {
                            incorrect = true;
                            error += "Valore di NOME TABELLA mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        if (!(Funct.ParseFlag(flag, "YES") || Funct.ParseFlag(flag, "NO")))
                        {
                            incorrect = true;
                            error += "Valore di FLAG BFD non conforme. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        // CODE 66
                        if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim()))
                        {
                            string databaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim(); //ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim();
                            if (!Funct.ValidateDatabaseName(databaseName))
                            {
                                incorrect = true;
                                error += "Valore di NOME DATABASE non conforme. ";
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                                worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            }
                        }
                        else
                        {
                            incorrect = true;
                            error += "Valore di NOME DATABASE mancante. ";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        }

                        if (incorrect == false)
                        { 
                            EmptyRow = 0;
                            EntityT ValRiga = new EntityT(row: RowPos, db: db, tName: nome);
                            ValRiga.TableName = nome;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text.Trim()))
                            //    ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._TABELLE["SSA"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text.Trim()))
                                ValRiga.HostName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome host"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text.Trim()))
                                ValRiga.Schema = worksheet.Cells[RowPos, ConfigFile._TABELLE["Schema"]].Text.Trim();
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text.Trim()))
                            //    ValRiga.TableDescr = worksheet.Cells[RowPos, ConfigFile._TABELLE["Descrizione Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text.Trim()))
                                ValRiga.InfoType = worksheet.Cells[RowPos, ConfigFile._TABELLE["Tipologia Informazione"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text.Trim()))
                                ValRiga.TableLimit = worksheet.Cells[RowPos, ConfigFile._TABELLE["Perimetro Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text.Trim()))
                                ValRiga.TableGranularity = worksheet.Cells[RowPos, ConfigFile._TABELLE["Granularità Tabella"]].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Flag BFD"]].Text.Trim()))
                            {
                                if (Funct.ParseFlag(flag, "YES"))
                                    ValRiga.FlagBFD = "S";
                                if (Funct.ParseFlag(flag, "NO"))
                                    ValRiga.FlagBFD = "N";
                            }
                            else
                                ValRiga.FlagBFD = "N";

                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim()))
                                ValRiga.DatabaseName = worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Database"]].Text.Trim();

                            listaFile.Add(ValRiga);

                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2).Width = 100;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_TABELLE + ConfigFile.TABELLE_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Nome Tabella"]].Text) && string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Flag BFD"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Table '" + worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Table '" + worksheet.Cells[RowPos, ConfigFile._TABELLE["Nome Tabella"]].Text + "'. Validation OK", 3, ConfigFile.INFO);
                        }
                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(new FileInfo(Funct.GetFolderDestination2(fileDaAprire.FullName, fileDaAprire.Name)));
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }
        */

        /// <summary>
        /// Reads and processes Table data from excel's 'TABELLE' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /*
        public static List<RelationT> ReadXFileRelation(FileInfo fileDaAprire, string db, List<AttributeT> attributeList, string sheet = ConfigFile.RELAZIONI)
        {
            string file = fileDaAprire.FullName;
            List<RelationT> listaFile = new List<RelationT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Tables. File " + fileDaAprire.Name + " doesn't exist.", 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Reading Relation. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 3, ConfigFile.ERROR);
                return listaFile = null;
            }
            
            bool FilesEnd = false;
            int EmptyRow = 0;
            
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string error = null;
                        string datatypeError = null;
                        string identificativoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativo relazione"]].Text.ToUpper().Trim();
                        string tabellaPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Padre"]].Text.ToUpper().Trim();
                        string tabellaFiglia = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tabella Figlia"]].Text.ToUpper().Trim();
                        string cardinalita = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Cardinalità"]].Text.ToUpper().Trim();
                        string campoPadre = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Padre"]].Text.ToUpper().Trim();
                        string campoFiglio = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Campo Figlio"]].Text.ToUpper().Trim();
                        string identificativa = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Identificativa"]].Text.ToUpper().Trim();
                        string eccezione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text.ToUpper().Trim();
                        string tipoRelazione = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Tipo Relazione"]].Text.ToUpper().Trim();
                        string note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text.ToUpper().Trim();
                        // ##########################################################
                        // SEZIONE per confronto corrispondenza DATATYPE Padre-Figlio
                        DataTypeT dtPadre = new DataTypeT();
                        DataTypeT dtFiglio = new DataTypeT();
                        try
                        {
                            dtPadre = Funct.ParseDataType(attributeList.Where(x => x.NomeTabellaLegacy == tabellaPadre && x.NomeCampoLegacy == campoPadre).FirstOrDefault().DataType, db);
                        }
                        catch
                        {
                            Logger.PrintLC("Could not parse DataType of field " + campoPadre + " in table " + tabellaPadre, 3, ConfigFile.ERROR);
                            if (datatypeError == null)
                                datatypeError = "ERROR: ";
                            datatypeError += "DATATYPE del Campo Padre è null. ";
                        }
                        try
                        {
                            dtFiglio = Funct.ParseDataType(attributeList.Where(x => x.NomeTabellaLegacy == tabellaFiglia && x.NomeCampoLegacy == campoFiglio).FirstOrDefault().DataType, db);
                        }
                        catch
                        {
                            Logger.PrintLC("Could not parse DataType of field " + campoFiglio + " in table " + tabellaFiglia, 3, ConfigFile.ERROR);
                            if (datatypeError == null)
                                datatypeError = "ERROR: ";
                            datatypeError += "DATATYPE del Campo Figlio è null. ";
                        }
                        if((dtPadre.Type == null || dtFiglio.Type == null) && datatypeError == null)
                        {
                            if (datatypeError == null)
                                datatypeError = "WARNING: ";
                            if(dtPadre.Type == null)
                                datatypeError += "DATATYPE del Campo Padre non recuperabile. ";
                            if(dtFiglio.Type == null)
                                datatypeError += "DATATYPE del Campo Figlio non recuperabile. ";
                        }
                        if(datatypeError == null && (dtPadre.Type != dtFiglio.Type))
                        {
                            if (datatypeError == null)
                                datatypeError = "WARNING: ";
                            datatypeError += "DATATYPE del Campo Padre ("+ dtPadre.Type +") e del Campo Figlio ("+ dtFiglio.Type +") non corrispondenti. ";
                        }
                        // ##########################################################
                        if (listaFile.Exists(x => x.IdentificativoRelazione == identificativoRelazione &&
                                                  x.TabellaPadre == tabellaPadre &&
                                                  x.TabellaFiglia == tabellaFiglia &&
                                                  x.CampoPadre == campoPadre &&
                                                  x.CampoFiglio == campoFiglio)
                                                  )
                        {
                            incorrect = true;
                            error += "Relazione già presente con ID: " + identificativoRelazione + " Tabella Padre: " + tabellaPadre + " Tabella Figlia: " + tabellaFiglia + " Campo Padre: " + campoPadre + " Campo Figlia: " + campoFiglio;
                        }
                            if (string.IsNullOrWhiteSpace(identificativoRelazione))
                        {
                            incorrect = true;
                            error += "IDENTIFICATIVO RELAZIONE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(tabellaPadre))
                        {
                            incorrect = true;
                            error += "TABELLA PADRE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(tabellaFiglia))
                        {
                            incorrect = true;
                            error += "TABELLA FIGLIA mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(cardinalita))
                        {
                            incorrect = true;
                            error += "CARDINALITA mancante. ";
                        }
                        else
                        {
                            switch (cardinalita.ToUpper())
                            { 
                                case "1:1":
                                    break;
                                case "1:N":
                                    break;
                                case "N:N":
                                    break;
                                case "(0,1) A (0,1)":
                                    break;
                                case "(0,1) A (1,M)":
                                    break;
                                case "(0,1) A (0,1,M)":
                                    break;
                                case "1 A (0,1)":
                                    break;
                                case "1 A (1,M)":
                                    break;
                                case "1 A (0,1,M)":
                                    break;
                                default:
                                    incorrect = true;
                                    error += "CARDINALITA non conforme. ";
                                    break;
                            }
                        }
                        if (string.IsNullOrWhiteSpace(campoPadre))
                        {
                            incorrect = true;
                            error += "CAMPO PADRE mancante. ";
                        }
                        if (string.IsNullOrWhiteSpace(campoFiglio))
                        {
                            incorrect = true;
                            error += "CAMPO FIGLIO mancante. ";
                        }
                        if (!string.IsNullOrWhiteSpace(identificativa))
                        {
                            if(!(Funct.ParseFlag(identificativa.ToUpper(),"YES") || Funct.ParseFlag(identificativa.ToUpper(),"NO")))
                            {
                                incorrect = true;
                                error += "IDENTIFICATIVA non conforme. ";
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(tipoRelazione))
                        {
                            string upperTipoRelazione = tipoRelazione.ToUpper();
                            if (!(upperTipoRelazione.Equals("L") || upperTipoRelazione.Equals("LOGICA") ||
                                upperTipoRelazione.Equals("F") || upperTipoRelazione.Equals("FISICA")))
                            {
                                incorrect = true;
                                error += "TIPO RELAZIONE non conforme";
                            } 
                        }

                        if (incorrect == false)
                        { 
                            EmptyRow = 0;
                            RelationT ValRiga = new RelationT(row: RowPos, db: db);
                            ValRiga.IdentificativoRelazione = identificativoRelazione;
                            ValRiga.TabellaPadre = tabellaPadre;
                            ValRiga.TabellaFiglia = tabellaFiglia;
                            switch (cardinalita.ToUpper())
                            {
                                case "1:1":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1:N":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "N:N":
                                    ValRiga.History = "CARDINALITA non gestita dall'applicazione";
                                    ValRiga.NullOptionType = null;
                                    break;
                                case "(0,1) A (0,1)":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "(0,1) A (1,M)":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "(0,1) A (0,1,M)":
                                    ValRiga.Cardinalita = -3;
                                    ValRiga.NullOptionType = 100;
                                    break;
                                case "1 A (0,1)":
                                    ValRiga.Cardinalita = -1;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1 A (1,M)":
                                    ValRiga.Cardinalita = -2;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                case "1 A (0,1,M)":
                                    ValRiga.Cardinalita = -3;
                                    ValRiga.NullOptionType = 101;
                                    break;
                                default:
                                    ValRiga.History = "CARDINALITA non conforme";
                                    ValRiga.NullOptionType = null;
                                    break;
                            }
                            ValRiga.CampoPadre = campoPadre;
                            ValRiga.CampoFiglio = campoFiglio;
                            if (Funct.ParseFlag(identificativa.ToUpper(),"YES"))
                                ValRiga.Identificativa = 2;
                            else
                                ValRiga.Identificativa = 7;

                            if (string.IsNullOrEmpty(tipoRelazione))
                            {
                                ValRiga.TipoRelazione = true;
                            }
                            else
                            {
                                switch (tipoRelazione.ToUpper())
                                {
                                    case "L":
                                        ValRiga.TipoRelazione = true;
                                        break;
                                    case "LOGICA":
                                        ValRiga.TipoRelazione = true;
                                        break;
                                    case "F":
                                        ValRiga.TipoRelazione = false;
                                        break;
                                    case "FISICA":
                                        ValRiga.TipoRelazione = false;
                                        break;
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text))
                                ValRiga.Eccezioni = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Eccezioni"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text))
                                ValRiga.Note = worksheet.Cells[RowPos, ConfigFile._RELAZIONI["Note"]].Text;
                            listaFile.Add(ValRiga);
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Value = (datatypeError != null ? datatypeError : string.Empty);
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Value = (datatypeError != null ? datatypeError : string.Empty) + error;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_RELAZIONI + ConfigFile.RELAZIONI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._TABELLE["Nome Tabella"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Relation '" + identificativoRelazione + "' between Table '" + tabellaPadre + "' and Table '"+ tabellaFiglia + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Relation '" + identificativoRelazione + "' between Table '" + tabellaPadre + "' and Table '" + tabellaFiglia + "'. Validation OK", 3, ConfigFile.INFO);
                        }

                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(fileDaAprire);
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }
        */

        /// <summary>
        /// Reads and processes Attributes data from excel's 'ATTRIBUTI' sheet
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /*
        public static List<AttributeT> ReadXFileAttribute(FileInfo fileDaAprire, string db, string sheet = ConfigFile.ATTRIBUTI)
        {
            string file = fileDaAprire.FullName;
            List<AttributeT> listaFile = new List<AttributeT>();

            if (!File.Exists(file))
            {
                Logger.PrintLC("Reading Attributes. File " + fileDaAprire.Name + " doesn't exist.", 2, ConfigFile.ERROR);
                return listaFile = null;
            }
            FileOps.RemoveAttributes(file);

            if (fileDaAprire.Extension.ToUpper() == ".XLS")
            {
                if (!ConvertXLStoXLSX(file))
                    return listaFile = null;
                file = Path.ChangeExtension(file, ".xlsx");
                fileDaAprire = new FileInfo(file);
            }
            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Reading Attributes. Could not open file " + fileDaAprire.Name + "in location " + fileDaAprire.DirectoryName, 2, ConfigFile.ERROR);
                return listaFile = null;
            }

            bool FilesEnd = false;
            int EmptyRow = 0;
            foreach (var worksheet in ws)
            {
                if (worksheet.Name == sheet)
                {
                    FilesEnd = false;
                    for (int RowPos = ConfigFile.HEADER_RIGA + 1;
                            FilesEnd != true;
                            RowPos++)
                    {
                        bool incorrect = false;
                        string nomeTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text.ToUpper().Trim();
                        string nomeCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text.ToUpper().Trim();
                        if (nomeCampo.Contains("-"))
                        {
                            nomeCampo = nomeCampo.Replace("-", "_");
                            Logger.PrintLC("Field '" + worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Nome  Campo Legacy"]].Text + "' of Table '" + nomeTabella + "' has been renamed as " + nomeCampo + ". This value will be used to produce the Erwin file", 3, ConfigFile.WARNING);
                        }
                        string dataType = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Datatype"]].Text.Trim();
                        dataType = Funct.RemoveWhitespace(dataType);
                        string chiave = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave"]].Text.ToUpper().Trim();
                        string unique = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Unique"]].Text.ToUpper().Trim();
                        string chiaveLogica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Chiave Logica"]].Text.ToUpper().Trim();
                        string mandatoryFlag = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Mandatory Flag"]].Text.ToUpper().Trim();
                        string dominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dominio"]].Text.ToUpper().Trim();
                        string storica = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Storica"]].Text.ToUpper().Trim();
                        string datoSensibile = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Dato Sensibile"]].Text.ToUpper().Trim();
                        string definizoneCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text.ToUpper().Trim();
                        int tempInt = 0;
                        int? Integer = int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text.ToUpper().Trim(), out tempInt) ? (int?)tempInt : null;
                        int? Decimal = int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text.ToUpper().Trim(), out tempInt) ? (int?)tempInt : null;

                        worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "";
                        worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = "";

                        string error = "";
                        //Check Nome Tabella Legacy
                        if (string.IsNullOrWhiteSpace(nomeTabella))
                        {
                            incorrect = true;
                            error += "NOME TABELLA LEGACY mancante.";

                        }
                        //Check Nome Campo Legacy
                        if (string.IsNullOrWhiteSpace(nomeCampo))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "NOME CAMPO LEGACY mancante.";
                        }
                        //Check DataType
                        if (string.IsNullOrWhiteSpace(dataType))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DATATYPE mancante.";
                        }
                        else
                        {
                            DataTypeT dt = Funct.ParseDataType(dataType, db);
                            if (!dt.Correct)
                            {
                                incorrect = true;
                                if (!string.IsNullOrWhiteSpace(error))
                                    error += " ";
                                error += "DATATYPE non conforme.";
                            }
                            else
                            {
                                if (dt.Integer == null && Integer != null)
                                {
                                    List<string> DBlist = new List<string>();
                                        dt.DBType.ToList().ForEach(x => x.ToUpper());
                                    DBlist.ForEach(x => x.ToLower());
                                    dt.DBType = DBlist.ToArray();
                                    if (dt.DBType.Contains(dt.Type.ToLower() + "()"))
                                    {
                                        incorrect = true;
                                        if (!string.IsNullOrWhiteSpace(error))
                                            error += " ";
                                        error += "Valori del DATATYPE errati nella colonna 'Datatype'.";
                                    }
                                }
                            }
                        }
                        //Check Chiave
                        if (!(string.IsNullOrWhiteSpace(chiave)) && (!(Funct.ParseFlag(chiave, "YES") || Funct.ParseFlag(chiave, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "CHIAVE non conforme.";
                        }


                        //Check Unique
                        if (!(string.IsNullOrWhiteSpace(unique)) && (!(Funct.ParseFlag(unique, "YES") || Funct.ParseFlag(unique, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "UNIQUE non conforme.";
                        }
                        //Check Chiave Logica
                        if (!(string.IsNullOrWhiteSpace(chiaveLogica)) && (!(Funct.ParseFlag(chiaveLogica, "YES") || Funct.ParseFlag(chiaveLogica, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "CHIAVE LOGICA non conforme.";
                        }
                        //Check Mandatory Flag
                        if (!(string.IsNullOrWhiteSpace(mandatoryFlag)) && (!(Funct.ParseFlag(mandatoryFlag, "YES") || Funct.ParseFlag(mandatoryFlag, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "MANDATORY FLAG non conforme.";
                        }
                        //Check Dominio
                        if (!(string.IsNullOrWhiteSpace(dominio)) && (!(Funct.ParseFlag(dominio, "YES") || Funct.ParseFlag(dominio, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DOMINIO non conforme.";
                        }
                        if (!(string.IsNullOrWhiteSpace(datoSensibile)) && (!(Funct.ParseFlag(datoSensibile, "YES") || Funct.ParseFlag(datoSensibile, "NO"))))
                        {
                            incorrect = true;
                            if (!string.IsNullOrWhiteSpace(error))
                                error += " ";
                            error += "DATO SENSIBILE non conforme.";
                        }

                        if (string.IsNullOrWhiteSpace(definizoneCampo))
                        {
                            incorrect = true;
                            error += "DEFINIZIONE CAMPO non conforme.";

                        }
                        if (incorrect == false)
                        {
                            EmptyRow = 0;
                            AttributeT ValRiga = new AttributeT(row: RowPos, db: db, nomeTabellaLegacy: nomeTabella);
                            // Assegnazione valori checkati
                            ValRiga.NomeTabellaLegacy = nomeTabella;
                            ValRiga.NomeCampoLegacy = nomeCampo;
                            ValRiga.DataType = dataType;

                            if (Funct.ParseFlag(chiave, "YES"))
                                ValRiga.Chiave = 0;
                            else
                                ValRiga.Chiave = 100;

                            if (Funct.ParseFlag(unique, "YES"))
                                ValRiga.Unique = unique;
                            else
                                ValRiga.Unique = "N";

                            if (Funct.ParseFlag(chiaveLogica, "YES"))
                                ValRiga.ChiaveLogica = chiaveLogica;
                            else
                                ValRiga.ChiaveLogica = "N";

                            if (Funct.ParseFlag(mandatoryFlag, "YES"))
                                ValRiga.MandatoryFlag = 1;
                            else
                                ValRiga.MandatoryFlag = 0;

                            if (Funct.ParseFlag(dominio, "YES"))
                                ValRiga.Dominio = dominio;
                            else
                                ValRiga.Dominio = "N";

                            ValRiga.Storica = storica;

                            if (Funct.ParseFlag(datoSensibile, "YES"))
                                ValRiga.DatoSensibile = datoSensibile;
                            else
                                ValRiga.DatoSensibile = "N";
                            
                            //Assegnazione valori opzionali
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["SSA"]].Text))
                                ValRiga.SSA = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["SSA"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Area"]].Text))
                                ValRiga.Area = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Area"]].Text;
                            //if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text))
                            //    ValRiga.DefinizioneCampo = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Definizione Campo"]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati"]].Text))
                                ValRiga.TipologiaTabella = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Tipologia Tabella \n(dal DOC. LEGACY) \nEs: Dominio,Storica,\nDati"]].Text;
                            int t;  //Funzionale all'assegnazione di 'Lunghezza' e 'Decimali'
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text))
                                if (int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Lunghezza"]].Text, out t))
                                    ValRiga.Lunghezza = t;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text))
                                if(int.TryParse(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Decimali"]].Text, out t))
                                    ValRiga.Decimali = t;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Provenienza dominio "]].Text))
                                ValRiga.ProvenienzaDominio = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Provenienza dominio "]].Text;
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Note"]].Text))
                                ValRiga.Note = worksheet.Cells[RowPos, ConfigFile._ATTRIBUTI["Note"]].Text;
                            listaFile.Add(ValRiga);
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(34, 255, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "OK";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = "";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                            AttributeT ValRiga = new AttributeT(row: RowPos, db: db, nomeTabellaLegacy: nomeTabella);
                            // Assegnazione valori checkati
                            ValRiga.NomeTabellaLegacy = nomeTabella;
                            ValRiga.NomeCampoLegacy = nomeCampo;
                            ValRiga.DataType = dataType;
                            ValRiga.History = error;
                            ValRiga.Step = 0;
                            listaFile.Add(ValRiga);
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1).Width = 10;
                            worksheet.Column(ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2).Width = 50;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.Font.Bold = true;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Value = "KO";
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Value = error;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[RowPos, ConfigFile.HEADER_COLONNA_MAX_ATTRIBUTI + ConfigFile.ATTRIBUTI_EXCEL_COL_OFFSET2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        //******************************************
                        // Verifica lo stato delle successive 10 righe per determinare la fine della tabella.
                        int prossime = 0;
                        for (int i = 1; i < 11; i++)
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[RowPos + i, ConfigFile._ATTRIBUTI["Nome Tabella Legacy"]].Text))
                                prossime++;
                        }
                        if (prossime == 10)
                            FilesEnd = true;
                        //******************************************

                        if (incorrect)
                        {
                            Logger.PrintLC("Checked Field '" + nomeCampo  + "' of Table '" + nomeTabella + "'. Validation KO. Error: " + error, 3, ConfigFile.WARNING);
                        }
                        else
                        {
                            Logger.PrintLC("Checked Field '" + nomeCampo + "' of Table '" + nomeTabella + "'. Validation OK", 3, ConfigFile.INFO);
                        }
                    }
                    if (ConfigFile.DEST_FOLD_UNIQUE)
                    {
                        p.SaveAs(new FileInfo(Path.Combine(ConfigFile.FOLDERDESTINATION, fileDaAprire.Name)));
                    }
                    else
                    {
                        p.SaveAs(fileDaAprire);
                    }
                    return listaFile;
                }
            }
            return listaFile = null;
        }
        */


        /// Writes text in Excel's Table table
        /*
        private static bool XLSXWriteErrorInCell(FileInfo fInfo, List<EntityT> list, int col, int v, string tABELLE)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, tABELLE);
            }
            return true;
        }
        */
        /// Writes text in Excel's Relation table
        /*
        private static bool XLSXWriteErrorInCell(FileInfo fInfo, List<RelationT> list, int col, int v, string rELAZIONI)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, rELAZIONI);
            }
            return true;
        }
        */

        /// Writes text in Excel's Attributes table
        /*
        internal static bool XLSXWriteErrorInCell(FileInfo fInfo, List<AttributeT> list, int col, int v, string aTTRIBUTI)
        {
            if (list.Count > 0)
            {
                List<GenericTypeT> genericList = new List<GenericTypeT>();
                foreach (var entity in list)
                {
                    genericList.Add((GenericTypeT)entity);
                }
                return XLSXWriteErrorInCell(fInfo, genericList, col, v, aTTRIBUTI);
            }
            return true;
        }
        */

        /// Writes text in a specific Excel's cell, in a specific sheet
        /*
        public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<GenericTypeT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        {
            try
            {
                string file = fileDaAprire.FullName;
                if (!File.Exists(file))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
                    return false;
                }
                FileOps.RemoveAttributes(file);
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                        return false;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
                    return false;
                }

                foreach (var worksheet in ws)
                {
                    if (worksheet.Name == sheet)
                    {
                        try
                        {
                            foreach (var dati in Rows)
                            {
                                worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
                                worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
                                worksheet.Cells[dati.Row, column].Value = "KO";
                                string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
                                if (mystring == null)
                                    mystring = "";
                                if (!(mystring.Contains(dati.History)))
                                {
                                    worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
                                }
                                worksheet.Column(column + 1).Width = 100;
                                worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                worksheet.Column(column + 1).Style.WrapText = true;
                                Logger.PrintLC("Updating excel file for error: " + dati.History, 3);
                            }
                            p.SaveAs(fileDaAprire);
                            return true;
                        }
                        catch (Exception exp)
                        {
                            Logger.PrintLC("Error while writing on file " +
                                            fileDaAprire.Name +
                                            ". Description: " +
                                            exp.Message, 1, ConfigFile.ERROR);
                            return false;
                        }
                    }
                }
                Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
                return false;
            }
            catch
            {
                return false;
            }
        }
        */

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<RelationT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (mystring == null)
        //                            mystring = "";
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message, 1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<AttributeT> Rows,int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (mystring == null)
        //                            mystring = "";
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message,1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //public static bool XLSXWriteErrorInCell(FileInfo fileDaAprire, List<EntityT> Rows, int column, int priorityWrite, string sheet = ConfigFile.ATTRIBUTI)
        //{
        //    try
        //    {
        //        string file = fileDaAprire.FullName;
        //        if (!File.Exists(file))
        //        {
        //            Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }
        //        FileOps.RemoveAttributes(file);
        //        if (fileDaAprire.Extension.ToUpper() == ".XLS")
        //        {
        //            if (!ConvertXLStoXLSX(file))
        //                return false;
        //            file = Path.ChangeExtension(file, ".xlsx");
        //            fileDaAprire = new FileInfo(file);
        //        }
        //        ExApp = new Excel.ApplicationClass();
        //        ExcelPackage p = null;
        //        ExcelWorkbook WB = null;
        //        ExcelWorksheets ws = null;
        //        try
        //        {
        //            ExApp.DisplayAlerts = false;
        //            p = new ExcelPackage(fileDaAprire);
        //            ExApp.DisplayAlerts = true;
        //            WB = p.Workbook;
        //            ws = WB.Worksheets;
        //        }
        //        catch (Exception exp)
        //        {
        //            Logger.PrintLC("Reading file: " + fileDaAprire.Name + ": could not open file in location " + fileDaAprire.DirectoryName, priorityWrite, ConfigFile.ERROR);
        //            return false;
        //        }

        //        foreach (var worksheet in ws)
        //        {
        //            if (worksheet.Name == sheet)
        //            {
        //                try
        //                {
        //                    foreach (var dati in Rows)
        //                    {
        //                        worksheet.Cells[dati.Row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                        worksheet.Cells[dati.Row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
        //                        worksheet.Cells[dati.Row, column].Style.Font.Bold = true;
        //                        worksheet.Cells[dati.Row, column].Value = "KO";
        //                        string mystring = (string)worksheet.Cells[dati.Row, column + 1].Value;
        //                        if (!(mystring.Contains(dati.History)))
        //                        {
        //                            worksheet.Cells[dati.Row, column + 1].Value = mystring + dati.History;
        //                        }
        //                        worksheet.Column(column + 1).Width = 100;
        //                        worksheet.Cells[dati.Row, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //                        worksheet.Column(column + 1).Style.WrapText = true;
        //                        Logger.PrintLC("Updating excel file for error " + dati.History, 3);
        //                    }
        //                    p.SaveAs(fileDaAprire);
        //                    return true;
        //                }
        //                catch (Exception exp)
        //                {
        //                    Logger.PrintLC("Error while writing on file " +
        //                                    fileDaAprire.Name +
        //                                    ". Description: " +
        //                                    exp.Message, 1, ConfigFile.ERROR);
        //                    return false;
        //                }
        //            }
        //        }
        //        Logger.PrintLC("File writing. Sheet " + sheet + "could not be found in file " + fileDaAprire.Name, priorityWrite, ConfigFile.ERROR);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        /// Writes Excel stats for Entity
        /*
        public static bool WriteExcelStatsForEntity(FileInfo fileDaAprire, Dictionary<string, List<String>> CompareResults)
        {
            try
            {
                string file = fileDaAprire.FullName;
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage();
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets; 
                    ws.Add(ConfigFile.TABELLE_DIFF);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Errore durante la scrittura di: " + fileDaAprire.Name + ": impossibile aprire il file " + fileDaAprire.DirectoryName, 1, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.TABELLE_DIFF];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Column(1).Width = 50;
                worksheet.Column(2).Width = 50;
                worksheet.Cells[1, 1].Value = "Tabelle Documento Di Ricognizione Caricate In Erwin";
                worksheet.Cells[1, 2].Value = "Tabelle Documento DDL";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(1).Style.WrapText = true;
                worksheet.Column(2).Style.WrapText = true;
                worksheet.View.FreezePanes(2, 1);
                //ExcelRange firstRow = (ExcelRange)worksheet.Row(1);
                //firstRow.f
                //firstRow.Select();
                //firstRow.Application.ActiveWindow.FreezePanes = true;

                int row = 2;
                bool pair = true;
                bool ExcelVuoto = true;
                foreach (var result in CompareResults)
                {
                    foreach (var element in result.Value)
                    {
                        worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        if ((result.Key == "CollezioneTrovati") && ConfigFile.DDL_Show_Right_Rows)
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = element;
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneNonTrovatiSQL")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = "KO: Entity non trovata sul DDL";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneNonTrovatiXLS")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 2].Value = element;
                            worksheet.Cells[row, 1].Value = "KO: Entity non caricata su Erwin";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        
                    }
                    
                }
                if (ExcelVuoto)
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Differenza Riscontrata";
                    worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);
                }

                Logger.PrintLC("Fine compilazione file excel", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file. Errore: " + exp.Message , 4, ConfigFile.ERROR);
                return false;
            }
        }
        */

        /// Writes Excel stats for Attributes
        /*
        public static bool WriteExcelStatsForAttribute(FileInfo fileDaAprire, Dictionary<string, List<String>> CompareResults, List<AttributeT> Attributi)
        {
            try
            {
                string file = fileDaAprire.FullName;

                if (!File.Exists(file))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", 1, ConfigFile.ERROR);
                    return false;
                }
                FileOps.RemoveAttributes(file);
                if (fileDaAprire.Extension.ToUpper() == ".XLS")
                {
                    if (!ConvertXLStoXLSX(file))
                        return false;
                    file = Path.ChangeExtension(file, ".xlsx");
                    fileDaAprire = new FileInfo(file);
                }
                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets; //.Add(wsName + wsNumber.ToString());
                    ws.Add(ConfigFile.ATTRIBUTI_DIFF);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Lettura file: " + fileDaAprire.Name + ": impossibile aprire il percorso " + fileDaAprire.DirectoryName, 1, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.ATTRIBUTI_DIFF];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                worksheet.Column(1).Width = 45;
                worksheet.Column(2).Width = 45;
                worksheet.Column(3).Width = 25;
                worksheet.Column(4).Width = 25;
                worksheet.Column(5).Width = 25;
                worksheet.Cells[1, 1].Value = "Attributi Documento Di Ricognizione Caricati In Erwin";
                worksheet.Cells[1, 2].Value = "Attributi Documento DDL";
                worksheet.Cells[1, 3].Value = "Differenze Campo Datatype";
                worksheet.Cells[1, 4].Value = "Differenze Campo Chiave";
                worksheet.Cells[1, 5].Value = "Differenze Campo Mandatory";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                //worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 3].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 4].Style.Font.Color.SetColor(Color.Red);
                //worksheet.Cells[1, 5].Style.Font.Color.SetColor(Color.Red);
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(1).Style.WrapText = true;
                worksheet.Column(2).Style.WrapText = true;
                worksheet.Column(3).Style.WrapText = true;
                worksheet.Column(4).Style.WrapText = true;
                worksheet.Column(5).Style.WrapText = true;
                //Excel.Range firstRow = (Excel.Range)worksheet.Row(1);
                //firstRow.Activate();
                //firstRow.Select();
                //firstRow.Application.ActiveWindow.FreezePanes = true;
                worksheet.View.FreezePanes(2, 1);

                bool ExcelVuoto = true;

                int row = 2;
                bool pair = true;
                foreach (var result in CompareResults)
                {
                    foreach (var element in result.Value)
                    {
                        worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        if (result.Key == "CollezioneAttributiTrovati")
                        {
                            string[] elementi = element.Split('|');
                            if (elementi.Count() != 4)
                            {
                                worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                                ExcelVuoto = false;
                                continue;
                            }
                            
                            
                            AttributeT AttributoRif = Attributi.Find(x => elementi[0] == x.NomeTabellaLegacy + "." + x.NomeCampoLegacy);
                            bool datatypeOK = true;
                            bool mandatoryOK = true;
                            bool keyOK = true;
                            string mandatoryXLS = string.Empty;
                            string mandatoryDDL = string.Empty;
                            string keyXLS = string.Empty;
                            string keyDDL = string.Empty;

                            mandatoryDDL = elementi[2] == "true" ? "NOT NULL" : "NULL";
                            mandatoryXLS = AttributoRif.MandatoryFlag == 1 ? "NOT NULL" : "NULL";
                            keyXLS = elementi[3] == "true" ? "CHIAVE PRIMARIA" : "";
                            keyDDL = AttributoRif.Chiave == 0 ? "CHIAVE PRIMARIA" : "";

                            if (AttributoRif.DataType != elementi[1])
                                datatypeOK = false;
                            if (mandatoryDDL != mandatoryXLS)
                                mandatoryOK = false;
                            if (keyDDL != keyXLS)
                                keyOK = false;

                            if ((!ConfigFile.DDL_Show_Right_Rows) && datatypeOK && mandatoryOK && keyOK) 
                            {
                              // se tutte e 4 le condizioni sono vere non scrive. Se anche solo una è falsa scrive.  
                            }
                            else
                            { 
                                ExcelVuoto = false;
                                worksheet.Cells[row, 1].Value = elementi[0];
                                worksheet.Cells[row, 2].Value = elementi[0];
                                worksheet.Cells[row, 3].Value = "XLS: " + AttributoRif.DataType + "\n" + "DDL: " + elementi[1];
                                worksheet.Cells[row, 4].Value = "XLS: " + keyXLS + "\n" + "DDL: " + keyDDL;
                                worksheet.Cells[row, 5].Value = "XLS: " + mandatoryXLS + "\n" + "DDL: " + mandatoryDDL;


                                if (pair)
                                {
                                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.White);
                                    if (datatypeOK)
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                    if (mandatoryOK)
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                    if (keyOK)
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                                    else
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);

                                }
                                else
                                {
                                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                    if (datatypeOK)
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    if (mandatoryOK)
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    if (keyOK)
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                    else
                                        worksheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                }
                                row += 1;
                                pair = !pair;
                            }

                        }
                        if (result.Key == "CollezioneAttributiNonTrovatiSQL")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 1].Value = element;
                            worksheet.Cells[row, 2].Value = "KO: Attributo non trovato sul DDL";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        if (result.Key == "CollezioneAttributiNonTrovatiXLS")
                        {
                            ExcelVuoto = false;
                            worksheet.Cells[row, 2].Value = element;
                            worksheet.Cells[row, 1].Value = "KO: Attributo non caricato su Erwin";
                            if (pair)
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            else
                            {
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                                worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.White);
                            }
                            row += 1;
                            pair = !pair;
                        }
                        Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
                        
                    }

                }

                if (ExcelVuoto)
                {
                    worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Differenza Riscontrata";
                    worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 3].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 4].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 4].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[2, 5].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    worksheet.Cells[2, 5].Style.Font.Color.SetColor(Color.White);

                }

                Logger.PrintLC("Fine compilazione file excel", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file. Errore: " + exp.Message, 4, ConfigFile.ERROR);
                return false;
            }
        }
        */

        /// <summary>
        /// Write on specific 'ControlliTempistiche' file a
        /// list of value.
        /// </summary>
        /// <param name="fileDaAprire"></param>
        /// <param name="ListCodLocaleControllo"></param>
        /// <returns></returns>
        /*
        public static bool WriteDocExcelControlliTempistiche(FileInfo fileDaAprire, List<string> ListCodLocaleControllo)
        {
            string TemplateFile = null;
            if (!string.IsNullOrWhiteSpace(ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE))
            {
                TemplateFile = ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE;
            }
            else
            {
                Logger.PrintLC("Value of 'ControlliTempistiche Template' is not valid. Will not valorize its content.", 2, ConfigFile.ERROR);
                return false;
            }
            string file = fileDaAprire.FullName;
            try
            {
                if (!File.Exists(TemplateFile))
                {
                    Logger.PrintLC("Reading File " + fileDaAprire.Name + ": doesn't exist.", 1, ConfigFile.ERROR);
                    return false;
                }
                else
                {
                    File.Copy(TemplateFile, file);
                }
                FileOps.RemoveAttributes(file);
            }
            catch
            {
            }

            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = null;
            ExcelWorkbook WB = null;
            ExcelWorksheets ws = null;
            try
            {
                ExApp.DisplayAlerts = false;
                p = new ExcelPackage(fileDaAprire);
                ExApp.DisplayAlerts = true;
                WB = p.Workbook;
                ws = WB.Worksheets;
                //ws.Add(ConfigFile.CONTROLLI);
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Error while opening: " + fileDaAprire.FullName + ". Will not valorize its content.", 2, ConfigFile.ERROR);
                return false;
            }



            
            //ExcelWorksheet worksheet = null;
            //try
            //{
            //    worksheet = ws[ConfigFile.CONTROLLI_TEMPISTICHE];
            //}
            //catch
            //{
            //    Logger.PrintLC("Could not find sheet \"" + ConfigFile.CONTROLLI_TEMPISTICHE + "\" in " + file +
            //        ". Will not valorize its content.", 4, ConfigFile.INFO);
            //}
            //Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);
            


            //bool ExcelVuoto = true;
            foreach (var worksheet in ws)
            {
                if(worksheet.Name == ConfigFile.CONTROLLI_TEMPISTICHE) { 
                int row = 2;
                bool pair = true;
                    foreach (var element in ListCodLocaleControllo)
                    {
                        //worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                        string[] elementi = element.Split('|');
                        if (elementi.Count() == 4)
                        {
                            worksheet.Cells[row, 12].Value = "errore nella comparazione dell'elemento: " + element;
                            //ExcelVuoto = false;
                            continue;
                        }
                        //ExcelVuoto = false;
                        worksheet.Cells[row, 12].Value = element.ToString();

                        //if (pair)
                        //{
                        //    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                        //}
                        //else
                        //{
                        //    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                        //}
                        Logger.PrintLC("Riga " + row + " scritta nel file excel", 6, ConfigFile.INFO);
                        row += 1;
                        pair = !pair;
                    }
                }
            }
            fileDaAprire.IsReadOnly = false;

            try
            {
                p.Save();
                //p.SaveAs(fileDaAprire);

                if (ws != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(ws);
                    }
                    catch { }
                }
                if (WB != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(WB);
                    }
                    catch { }
                }
                if (ExApp != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(ExApp);
                    }
                    catch { }
                }
                if (p != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(p);
                    }
                    catch { }
                }
            }
            catch(Exception exp)
            {
                Logger.PrintLC("Error while closing Excel application. Will continue without notice.", 5, ConfigFile.ERROR);
            }
            return true;
        }
        */


        /// <summary>
        /// Write on specific 'ControlliTempistiche' file a
        /// list of value.
        /// </summary>
        /*
        public static bool WriteDocExcelControlliCampiX(FileInfo fileDaAprire, List<string> ExcelControlli)
        {
            string TemplateFile = ConfigFile.CONTROLLI_TEMPISTICHE_TEMPLATE;

            string file = fileDaAprire.FullName;
            if (!File.Exists(TemplateFile))
            {
                Logger.PrintLC("Trying to find File " + fileDaAprire.Name + ": doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
            else
            {
                File.Copy(TemplateFile, file);
            }
            fileDaAprire = new FileInfo(file);
            //FileOps.RemoveAttributes(file);

            ExApp = new Excel.ApplicationClass();
            ExcelPackage p = new ExcelPackage(fileDaAprire);
            ExcelWorkbook WB = p.Workbook;
            ExcelWorksheets ws = WB.Worksheets;

            //ExcelWorkbook WB = null;
            //ExcelWorksheets ws = null;
            //try
            //{
            //ExApp.DisplayAlerts = false;
            //p = new ExcelPackage(fileDaAprire);
            //ExApp.DisplayAlerts = true;
            //ws.Add(ConfigFile.CONTROLLI);
            //}

            var worksheet = ws[ConfigFile.CONTROLLI_TEMPISTICHE];
            bool isProtect = worksheet.Protection.IsProtected;
            Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

            worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Column(4).Width = 45;
            worksheet.Column(5).Width = 45;
            worksheet.Column(6).Width = 45;
            worksheet.Column(7).Width = 45;
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Column(4).Style.WrapText = true;
            worksheet.Column(5).Style.WrapText = true;
            worksheet.Column(6).Style.WrapText = true;
            worksheet.Column(7).Style.WrapText = true;
            worksheet.View.FreezePanes(2, 1);

            bool ExcelVuoto = true;

            int row = 2;
            bool pair = true;
            foreach (var element in ExcelControlli)
            {
                worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                string[] elementi = element.Split('|');
                if (elementi.Count() == 2)
                {
                    worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                    ExcelVuoto = false;
                    continue;
                }
                ExcelVuoto = false;
                worksheet.Cells[row, 12].Value = elementi[0];

                if (pair)
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                }
                else
                {
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                }
                row += 1;
                pair = !pair;

                Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
            }

            if (ExcelVuoto)
            {
                worksheet.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Cells[2, 1].Value = "Nessuna Controllo Riscontrato";

            }

            Logger.PrintLC("Fine compilazione file excel controlli", 4, ConfigFile.INFO);
            ExcelProtectedRangeCollection range = worksheet.ProtectedRanges;
            p.Save();
            p.Dispose();
            Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
            return true;
        }
        */

        /// <summary>
        /// Write on specific 'ControlliTempistiche' file a
        /// list of value.
        /// </summary>
        /*
        public static bool WriteDocExcelControlliCampi(FileInfo fileDaAprire, List<string> ExcelControlli)
        {
            string TemplateFile = ConfigFile.CONTROLLI_CAMPI_TEMPLATE;
            
            try
            {
                string file = fileDaAprire.FullName;
                if (!File.Exists(TemplateFile))
                {
                    Logger.PrintLC("Trying to find File " + fileDaAprire.Name + ": doesn't exist.", 2, ConfigFile.ERROR);
                    return false;
                }
                else
                {
                    File.Copy(TemplateFile, file);
                }
                FileOps.RemoveAttributes(file);

                ExApp = new Excel.ApplicationClass();
                ExcelPackage p = null;
                ExcelWorkbook WB = null;
                ExcelWorksheets ws = null;
                try
                {
                    ExApp.DisplayAlerts = false;
                    p = new ExcelPackage(fileDaAprire);
                    ExApp.DisplayAlerts = true;
                    WB = p.Workbook;
                    ws = WB.Worksheets;
                    //ws.Add(ConfigFile.CONTROLLI);
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Errore durante la scrittura di: " + fileDaAprire.Name + ": impossibile aprire il file in " + fileDaAprire.DirectoryName, 2, ConfigFile.ERROR);
                    return false;
                }

                var worksheet = ws[ConfigFile.CONTROLLI_CAMPI];

                Logger.PrintLC("Inizio compilazione file excel", 4, ConfigFile.INFO);

                worksheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                worksheet.Column(4).Width = 45;
                worksheet.Column(5).Width = 45;
                worksheet.Column(6).Width = 45;
                worksheet.Column(7).Width = 45;
                //worksheet.Cells[1, 1].Value = "Nome Struttura Informativa";
                //worksheet.Cells[1, 2].Value = "Nome Campo";
                //worksheet.Cells[1, 3].Value = "Cod Locale Controllo";
                //worksheet.Cells[1, 4].Value = "Ruolo Campo";
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Column(4).Style.WrapText = true;
                worksheet.Column(5).Style.WrapText = true;
                worksheet.Column(6).Style.WrapText = true;
                worksheet.Column(7).Style.WrapText = true;
                worksheet.View.FreezePanes(2, 1);

                bool ExcelVuoto = true;

                int row = 2;
                bool pair = true;
                foreach (var element in ExcelControlli)
                {
                    worksheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;

                    string[] elementi = element.Split('|');
                    if (elementi.Count() != 4)
                    {
                        worksheet.Cells[row, 1].Value = "errore nella comparazione dell'elemento: " + element;
                        ExcelVuoto = false;
                        continue;
                    }
                    ExcelVuoto = false;
                    worksheet.Cells[row, 4].Value = elementi[0];
                    worksheet.Cells[row, 5].Value = elementi[1];
                    worksheet.Cells[row, 6].Value = elementi[2];
                    worksheet.Cells[row, 7].Value = elementi[3];
 
                    if (pair)
                    {
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                    else
                    {
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                    }
                    row += 1;
                    pair = !pair;

                    Logger.PrintLC("Riga " + row + " scritta nel file excel", 5, ConfigFile.INFO);
                }

                if (ExcelVuoto)
                {
                    worksheet.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.White);
                    worksheet.Cells[2, 1].Value = "Nessuna Controllo Riscontrato";
                    
                }

                Logger.PrintLC("Fine compilazione file excel controlli", 4, ConfigFile.INFO);

                p.SaveAs(fileDaAprire);
                Logger.PrintLC(fileDaAprire + " Salvato", 4, ConfigFile.INFO);
                return true;
            }
            catch (Exception exp)
            {
                Logger.PrintLC("Errore durante la scrittura del file excel ControlliCampi. Errore: " + exp.Message, 4, ConfigFile.ERROR);
                return false;
            }
        }
        */
    }
}
