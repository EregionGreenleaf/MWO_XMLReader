using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class FileOps
    {
        private static string newExcel = ".xlsx";
        private static string oldExcel = ".xls";
        private static string all12Name = "allegato12";
        private static string versionPref = "V";

        /// <summary>
        /// Copies Erwin model to the local backup directory
        /// </summary>
        /// <param name="connessione_fileInfoERwin"></param>
        /// <param name="suffix"></param>
        public static void CopyErwinModel(string connessione_fileInfoERwin, string suffix)
        {
            string PercorsoCopieErwin = ConfigFile.PERCORSOCOPIEERWINDESTINATION;
            string DESTINATION = Path.Combine(PercorsoCopieErwin, Path.GetFileNameWithoutExtension(connessione_fileInfoERwin) + suffix + Path.GetExtension(connessione_fileInfoERwin));
            FileOps.CopyFile(connessione_fileInfoERwin, DESTINATION);
            Logger.PrintLC("Created copy of Erwin model file.", 4, ConfigFile.INFO);
        }


        /// <summary>
        /// Gets files from remote location, copying them locally
        /// </summary>
        /// <param name="FileDaElaborare"></param>
        /// <returns></returns>

        // Tuple< List<string>, int >
        public static ValueTuple<List<FileAll12T>, int> GetFilesFromRemote(List<FileAll12T> FileDaElaborare)
        {
            int returnValue = 0;
            if (FileDaElaborare.Count > 0)
            {
                if (ConfigFile.COPY_LOCAL)
                {
                    FileDaElaborare = Funct.LakeRemoteGet(FileDaElaborare);
                    if (FileDaElaborare == null)
                        returnValue = 5;
                    if (FileDaElaborare.Count == 0)
                        returnValue = 2;
                }
            }
            else
            {
                return ValueTuple.Create(FileDaElaborare, 2);
            }
            return ValueTuple.Create(FileDaElaborare, returnValue);
        }

        /// <summary>
        /// Checks if file is openable
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool isFileOpenable(string file)
        {
            return ExcelOps.isFileOpenable(file);
        }

        // ##########################################################
        // ##########################################################
        // ##########################################################

        /// <summary>
        /// Searches for all 'Allegato 12' files that have a valid DCP file to relay on
        /// </summary>
        /// <param name="allList"></param>
        /// <param name="dcpList"></param>
        /// <returns></returns>
        public static Tuple< List<FileAll12T>, List<FileAll12T>, List<FileDCPT>> GetValidAll12(List<FileAll12T> allList, List<FileDCPT> dcpList)
        {
            // We set all 'DCP' property to Null value
            allList.ForEach(x => x.SetDCPNull());
            // Search correspondences between All12 and DCPs
            allList = SearchCorrespondence(allList, dcpList);
            List<FileAll12T> trueAll12List = allList.Where(x => x.DCP != null).ToList();
            List<FileAll12T> missingCorrespondence = DetermineMissingCorrespondence(allList);
            // DEBUG
            //foreach (var file in missingCorrespondence)
            //{
            //    Logger.PrintLC(" File without corresponding DCP: " + file.FullName, 2);
            //}
            // ######
            List<FileDCPT> trueDCPList = GetDCPtoProcess(trueAll12List);
            return Tuple.Create(trueAll12List, missingCorrespondence, trueDCPList);
        }

        /// <summary>
        /// Searches and returns all valid DCPs related to valid All12s
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        public static List<FileDCPT> GetDCPtoProcess(List<FileAll12T> all12List)
        {
            List<FileDCPT> validDCPs = all12List.Select(x => x.DCP).ToList();
            return validDCPs;
        }

        /// <summary>
        /// Lists all 'Allegato 12' files where a correspondig DCP is missing (prints in log and relative OK/KO files)
        /// </summary>
        /// <param name="allFile"></param>
        /// <returns></returns>
        public static List<FileAll12T> DetermineMissingCorrespondence(List<FileAll12T> allFile)
        {
            List<FileAll12T> missingList = allFile.Where(x => x.DCP == null).ToList();
            List<FileAll12T> todoList = allFile.Where(x => x.DCP != null).ToList();
            // PRINT LOG AND 'OK/KO' FILE
            Funct.PrintProcessFileData(missingList, todoList);
            // ##########################
            return missingList;
        }




        /// <summary>
        /// Searches and sets correspondences between All12 and DCP files
        /// </summary>
        /// <param name="allList"></param>
        /// <param name="dcpList"></param>
        /// <returns></returns>
        public static List<FileAll12T> SearchCorrespondence(List<FileAll12T> allList, List<FileDCPT> dcpList)
        {
            // Clean the DCP List of duplicates
            dcpList = CleanDCPDuplicates(dcpList);
            // Calculate names of file (different versions) and directories
            allList.ForEach(x => x.SetAllNames());
            // Assign the correct DCP to each 'Allegato 12' file, if there's any
            foreach(FileAll12T allFile in allList)
            {
                dcpList.ForEach(dcp =>
                                {
                                    if (dcp.SimpleName.ToUpper() == allFile.RelatedDCPName.ToUpper())
                                        allFile.DCP = dcp;
                                }
                );
            }
            return allList;
        }

        /// <summary>
        /// Returns a list cleared of duplicates
        /// </summary>
        /// <param name="dcpList"></param>
        /// <returns></returns>
        public static List<FileDCPT> CleanDCPDuplicates(List<FileDCPT> dcpList)
        {
            return dcpList.GroupBy(x => x.SimpleName).Select(x => x.FirstOrDefault()).ToList();
        }

        // ##########################################################
        // ##########################################################
        // ##########################################################

        /// <summary>
        /// Filters a list of available DCPs
        /// </summary>
        /// <param name="where"></param>
        /// <returns></returns>
        public static List<FileDCPT> GetDCPFilesList(string where)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(where);
            if (!dirInfo.Exists)
            {
                Logger.PrintLC("Could not find the specified 'DCP''s directory (" + dirInfo.FullName + ").", 2, ConfigFile.ERROR);
                return null;
            }
            List<string> allFiles = DirOps.GetFilesToProcess(dirInfo.FullName, "*.xls|.xlsx").ToList();
            List <FileDCPT> allDCPFiles = SearchDCPFiles(allFiles);
            return allDCPFiles;
        }


        /// <summary>
        /// Filters the complete list of files to only get valid DCPs
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        /// **************************************************************************************************************************
        /// **************************************************************************************************************************
        public static List<FileDCPT> SearchDCPFiles(List<string> fileList)
        {
            List<string> listaFiltrata = new List<string>();
            List<string> fileProcessati = new List<string>();
            List<FileDCPT> listDef = new List<FileDCPT>();
            string parentDir = string.Empty;

            foreach (string file in fileList)
            {
                FileInfo fileInfo = new FileInfo(file);
                parentDir = fileInfo.DirectoryName.Split('\\').ToList().Last();
                string element = fileInfo.DirectoryName.Replace("\\" + parentDir, string.Empty);

                if (element.Split('\\').ToList().Last().ToUpper() == ConfigFile.INPUT_DCP_NAME.ToUpper())
                {
                    string[] TS = parentDir.Split('_');
                    if (TS.Count() == 2)
                    {
                        if (TS[0].Length == 8 && TS[1].Length == 4)
                        {
                            int tempInt = 0;
                            if (int.TryParse(TS[0], out tempInt) && int.TryParse(TS[1], out tempInt))
                            {
                                listaFiltrata.Add(file);
                            }
                        }
                    }
                }
            }

            foreach (var elemento in listaFiltrata)
            {
                if (!fileProcessati.Contains(Path.GetFileNameWithoutExtension(elemento)))
                {
                    List<string> listaUguali = listaFiltrata.Where
                                               (x => Path.GetFileNameWithoutExtension(x) == Path.GetFileNameWithoutExtension(elemento)).ToList();

                    string fileUltimaRelease = listaUguali.Count == 1 ? listaUguali.FirstOrDefault() : null;

                    if (listaUguali.Count > 1)
                    {
                        string DataOraOrdinato = string.Empty;
                        List<string> listaUgualiOrdinatasoloData = new List<string>();

                        foreach (var elemOrdinato in listaUguali)
                        {
                            FileInfo fileInfo_ordinato = new FileInfo(elemOrdinato);
                            DataOraOrdinato = fileInfo_ordinato.DirectoryName.Split('\\').ToList().Last();
                            DataOraOrdinato = DataOraOrdinato.Replace("_", string.Empty);
                            listaUgualiOrdinatasoloData.Add(DataOraOrdinato);
                        }

                        listaUgualiOrdinatasoloData.Sort();
                        string ReturnDataOra = listaUgualiOrdinatasoloData.Last().Substring(0, 8) + "_" +
                                               listaUgualiOrdinatasoloData.Last().Substring(8, 4);

                        fileUltimaRelease = listaUguali.Where(x => x.Contains(ReturnDataOra)).FirstOrDefault();
                    }

                    if (fileUltimaRelease != null)
                    {
                        FileDCPT tempFile = ParseDCP(fileUltimaRelease);
                        if (ValidateDCP(tempFile))
                        {
                            listDef.Add(tempFile);
                            fileProcessati.Add(Path.GetFileNameWithoutExtension(fileUltimaRelease));
                        }
                    }
                }
            }

            //to control items in listaDef
            //List<string> AAAAA = listDef.Select(x => x.FullName).ToList();

            return listDef;
            //**************************************************************************************************************************
            // **************************************************************************************************************************
        }






        /// <summary>
        /// Filters the complete list of files to only get valid DCPs
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        //public static List<FileDCPT> SearchDCPFiles(List<string> fileList)
        //{
        //    List<FileDCPT> returnList = new List<FileDCPT>();
        //    foreach (string file in fileList)
        //    {
        //        FileInfo fileInfo = new FileInfo(file);
        //        string parentDir = fileInfo.DirectoryName.Split('\\').ToList().Last();
        //        // Conditions to determine a valid "DCP" file
        //        if ((parentDir.ToUpper() == ConfigFile.INPUT_DCP_NAME.ToUpper()))
        //        {
        //            FileDCPT tempFile = ParseDCP(fileInfo.FullName);
        //            if (ValidateDCP(tempFile))
        //                returnList.Add(tempFile);
        //        }
        //    }

        //    return returnList;
        //}

        /// <summary>
        /// Validates the specific DCP file's name formatting
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool ValidateDCP(FileDCPT file)
        {
            file.Correct = false;
            if (file.Acronimo != null &&
               file.NomeModello != null &&
               file.SSA != null &&
               file.TipoDBMS != null)
                if (ConfigFile.DBS.Contains(file.TipoDBMS.ToUpper()))
                    file.Correct = true;
            return file.Correct;
        }

        /// <summary>
        /// Gets all the information from the DCPs file name
        /// </summary>
        /// <param name="fileI"></param>
        /// <returns></returns>
        public static FileDCPT ParseDCP(string fileI)
        {
            FileDCPT file = new FileDCPT();
            file.Estensione = new FileInfo(fileI).Extension;
            string fileWOE = Path.GetFileNameWithoutExtension(fileI);
            string[] composition = fileWOE.Split('_');
            if (composition.Count() == 4)
            {
                file.SSA = composition[0].Trim();
                file.Acronimo = composition[1].Trim();
                file.NomeModello = composition[2].Trim();
                file.TipoDBMS = composition[3].Trim().ToUpper();
                if (!ConfigFile.DBS.Contains(file.TipoDBMS.ToUpper()))
                    return file.SetAllNull();
                file.DirectoryName = new FileInfo(fileI).DirectoryName;
            }
            else
                return file.SetAllNull();
            file.SetFullName();
            file.SetSimpleName();
            if (fileWOE.ToUpper() == file.SimpleName.ToUpper())
                return file;
            else
                return file.SetAllNull();
        }


        // ##########################################################
        // ##########################################################
        // ##########################################################

        /// <summary>
        /// Searches all 'Allegato 12' files to elaborate
        /// </summary>
        /// <returns>List of unique 'Allegato 12' files at their latest version</returns>
        public static List<FileAll12T> GetAllegato12List(string where, bool print = true)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(where);
            if (!dirInfo.Exists)
            {
                Logger.PrintLC("Could not find the specified 'Allegato 12''s directory (" + dirInfo.FullName + ").", 2, ConfigFile.ERROR);
                return null;
            }
            List<string> allFiles = DirOps.GetFilesToProcess(dirInfo.FullName, "*.xls|.xlsx").ToList();
            List<FileAll12T> allAll12Files = SearchAll12Files(allFiles);
            List<FileAll12T> latestVersionList = GetLatestVersionAll12(allAll12Files);
            List<FileAll12T> purgedList = RemoveElaborated(latestVersionList);
            if (print)
                PrintFiles(allFiles, allAll12Files, latestVersionList, purgedList);
            return purgedList;
        }


        public static bool PrintFiles(List<string> allFiles, List<FileAll12T> allAll12Files, List<FileAll12T> latestVersionList, List<FileAll12T> purgedList)
        {
            Logger.PrintLC("List of all Excel files found:", 2, ConfigFile.INFO);
            if (allFiles.Count > 0)
            {
                allFiles.ForEach(file =>
                {
                    Logger.PrintLC(file, 3);
                });
            }
            else
                Logger.PrintLC("- NONE", 3);
            Funct.PrintAll12List("Lists of all 'Allegato 12' files:", 2, allAll12Files);
            Funct.PrintAll12List("List of the latest version of 'Allegato 12' files to elaborate:", 2, latestVersionList);
            if (ConfigFile.REMEMBER_ELABORATE)
            {
                Funct.PrintAll12List("List of 'Allegato 12' files purged from previously elaborated:", 2, purgedList);
            }
            return true;
        }


        /// <summary>
        /// Purges the input list of all elements previously elaborated
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<FileAll12T> RemoveElaborated (List<FileAll12T> list)
        {
            List<FileAll12T> returnList = new List<FileAll12T>();
            try
            {
                if (ConfigFile.REMEMBER_ELABORATE)
                {
                    list.ForEach(element =>
                    {
                        if (!ConfigFile.FILE_ELABORATED_LIST.Contains(Path.GetFileNameWithoutExtension(element.FullName)))
                            returnList.Add(element);
                    });
                }
                else
                {
                    returnList = list;
                }
            }
            catch
            {
                returnList = list;
            }
            return returnList;
        }


        /// <summary>
        /// Filters a list of files to return those 
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        public static List<FileAll12T> SearchAll12Files(List<string> fileList)
        {
            List<FileAll12T> returnList = new List<FileAll12T>();
            foreach(string file in fileList)
            {
                FileInfo fileInfo = new FileInfo(file);
                if(fileInfo.Extension.ToUpper() == newExcel.ToUpper())  // Check extension file (XLS/XLSX?)
                {
                    string parentDir = fileInfo.DirectoryName.Split('\\').ToList().Last();
                    // Conditions to determine a valid "Allegato 12" file
                    if((parentDir == ConfigFile.INPUT_ALLEGATO12_NAME) &&
                        (fileInfo.Name.ToUpper().Contains(all12Name.ToUpper())))
                    {
                        FileAll12T tempFile = ParseAll12(fileInfo.FullName);
                        if (tempFile.Correct)
                            returnList.Add(tempFile);
                    }
                }
            }
            return returnList;
        }


        /// <summary>
        /// Validates file name composition -> is an 'Allegato 12' file?
        /// </summary>
        /// <param name="fileI"></param>
        /// <returns>File in FileAll12T format</returns>
        public static FileAll12T ParseAll12(string fileI)
        {
            FileAll12T file = new FileAll12T();
            file.Correct = false;
            file.FullName = fileI;
            file.Estensione = new FileInfo(fileI).Extension;
            string fileWOE = Path.GetFileNameWithoutExtension(fileI);
            string[] composition =  fileWOE.Split('_');
            if (composition.Count() == 6)
            {
                file.SSA = composition[0].Trim();
                file.Acronimo = composition[1].Trim();
                file.NomeModello = composition[2].Trim();
                file.TipoDBMS = composition[3].Trim().ToUpper();
                file.All12Position = composition[4].Trim().ToUpper();
                file.Version = composition[5].Trim();
                file.SimpleName = file.SetNameWOVersion();
                file.SetDirectory();
                file.Correct = false;
                if (!ConfigFile.DBS.Contains(file.TipoDBMS))
                    return file;
                if (!(file.All12Position.ToUpper() == all12Name.ToUpper()))
                    return file;
                if (file.Version.ToUpper().Contains(versionPref))
                {
                    file.Version = file.Version.ToUpper().Replace(versionPref, "");
                    int temp;
                    if (!int.TryParse(file.Version, out temp))
                        return file;
                }
                else
                    return file;
            }
            else
                return file;
            file.Correct = true;
            return file;
        }

        /// <summary>
        /// Searches latest version of each file given as parameter.
        /// </summary>
        /// <param name="fileListOriginal"></param>
        /// <returns>List of unique 'Allegato 12' files</returns>
        public static List<FileAll12T> GetLatestVersionAll12(List<FileAll12T> fileListOriginal)
        {
            List<FileAll12T> fileList = fileListOriginal;
            List<FileAll12T> returnList = new List<FileAll12T>();
            foreach(FileAll12T file in fileList)
            {
                string nameWOV = file.SetNameWOVersion();
                if (returnList.Where(x => x.SimpleName == nameWOV).Count() > 0)
                    continue;
                List<FileAll12T> tempList = fileList.Where(x => (x.SetNameWOVersion() == nameWOV)).ToList();
                switch (tempList.Count)
                {
                    // Strange error (should not be possible)
                    case 0:
                        //fileList.Remove(file); // Cannot "remove" without breaking 'foreach'
                        break;
                    // It finds itself, so I add it
                    case 1:
                        returnList.Add(file);
                        break;
                    // It finds at least a duplicate. We need to find the last version.
                    default:
                        if (tempList.Count > 1)
                        {
                            FileAll12T lastVersion = new FileAll12T();
                            lastVersion.Version = "-1";
                            int actualLast = -1;
                            foreach(FileAll12T tFile in tempList)
                            {
                                int tempInt = -1;
                                int.TryParse(tFile.Version, out tempInt);
                                if(tempInt > actualLast)
                                {
                                    actualLast = tempInt;
                                    lastVersion = tFile;
                                }
                            }
                            returnList.Add(lastVersion);
                        }
                        break;
                }
            }
            return returnList;
        }

        /// <summary>
        /// NOT USED
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static FileT ParseFileName(string fileName)
        {
            FileT file = new FileT();
            FileInfo fileNameInfo = new FileInfo(fileName);
            string[] fileComponents;
            fileComponents = fileNameInfo.Name.Split(ConfigFile.DELIMITER_NAME_FILE);
            int length = fileComponents.Count();
            string correct = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + "_OK.txt");
            string error = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + "_KO.txt");

            if (length != 5)
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension>.", 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "Formattazione del nome file errata.", true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
                return file = null;
            }
            if (!ConfigFile.DBS.Contains(fileComponents[3].ToUpper()))
            {
                Logger.PrintLC(fileName + " file name doesn't conform to the formatting standard <SSA>_<ACRONYM>_<MODELNAME>_<DBMSTYPE>.<extension> . DB specified not present.", 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "DB specificato nel nome file non previsto.", true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
                return file = null;
            }

            try
            {
                file.SSA = fileComponents[0];
                file.Acronimo = fileComponents[1];
                file.NomeModello = fileComponents[2];
                file.TipoDBMS = fileComponents[3].ToUpper();
                file.Estensione = fileComponents[4];
            }
            catch (Exception exp)
            {
                Logger.PrintLC(fileName + "produced an error while parsing its name: " + exp.Message, 2, ConfigFile.ERROR);
                if (File.Exists(correct))
                {
                    File.Delete(correct);
                    Logger.PrintF(error, "er_driveup – Caricamento Excel su ERwin", true);
                    Logger.PrintF(error, "Colonne e Fogli formattati corretamente.", true);
                    Logger.PrintF(error, "Errore: " + exp.Message, true);
                }
                if (fileNameInfo.Extension.ToUpper() == ".XLS")
                {
                    string fXLSX = Path.Combine(fileNameInfo.DirectoryName, Path.GetFileNameWithoutExtension(fileNameInfo.FullName) + ".xlsx");
                    if (File.Exists(fXLSX))
                        File.Delete(fXLSX);
                }
                return file = null;
            }
            return file;
        }



        public static List<string> GetTrueFilesToProcess(string[] list)
        {
            List<string> nlist = new List<string>();
            List<string> Direct = new List<string>();

            if (list != null)
            {
                bool notFullRecursive = true;
                if (!string.IsNullOrEmpty(ConfigFile.INPUT_DCP_NAME))
                {
                    notFullRecursive = true;
                    nlist = (from c in list
                             where c.Contains(ConfigFile.INPUT_DCP_NAME)
                             select c).ToList();
                }
                else
                {
                    notFullRecursive = false;
                    nlist = (from c in list
                             where c.Contains(ConfigFile.ROOT)
                             select c).ToList();
                }
                if (notFullRecursive)
                {
                    int pathLenght = ConfigFile.INPUT_DCP_NAME.Length;
                    foreach (string file in nlist)
                    {

                        try
                        {

                            FileInfo fileI = new FileInfo(file);
                            DirectoryInfo dir = fileI.Directory;
                            int dirLenght = dir.FullName.Length;
                            string padre = dir.FullName.Substring(dirLenght - pathLenght);
                            if (padre == ConfigFile.INPUT_DCP_NAME)
                                Direct.Add(file);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                else
                {
                    Direct = nlist;
                }
                if(Direct != null)
                    Direct = CleanDuplicates(Direct);
            }
            return Direct;
        }

        public static List<string> CleanDuplicates(List<string> list)
        {
            List<string> nlist = new List<string>();
            List<string> trueList = new List<string>();
            if (list != null)
            {
                foreach(var x in list)
                {
                    string XLS = Path.Combine(Path.GetDirectoryName(x), Path.GetFileNameWithoutExtension(x) + ".xls");
                    string XLSX = Path.Combine(Path.GetDirectoryName(x), Path.GetFileNameWithoutExtension(x) + ".xlsx");
                    if (!nlist.Contains(XLS) && !nlist.Contains(XLSX))
                    {
                        nlist.Add(x);
                    }
                }
                List<string> nameList = new List<string>(); //da aggiungere fuori dall'IF
                foreach (var elemento in nlist)
                {
                    if (!(nameList.Contains(Path.GetFileNameWithoutExtension(elemento))))
                    {
                        nameList.Add(Path.GetFileNameWithoutExtension(elemento));
                        trueList.Add(elemento);
                    }
                }
            }
            return trueList;
        }

        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }
        /// <summary>
        /// Removes a specific Attribute from a file.
        /// </summary>
        /// <param name="filePath">Path and file name to be elaborated</param>
        /// <param name="attribute">Attribute to be removed. 'ReadOnly' by default.</param>
        public static void RemoveAttributes(string filePath, FileAttributes attribute = FileAttributes.ReadOnly)
        {
            if (File.Exists(filePath))
            {
                FileAttributes attributes = File.GetAttributes(filePath);

                if ((attributes & attribute) == attribute)
                {
                    // Make the file RW
                    attributes = RemoveAttribute(attributes, attribute);
                    File.SetAttributes(filePath, attributes);
                    Logger.PrintLC(filePath + " is no longer RO.", 6, ConfigFile.INFO);
                }
            }
        }

        public static bool CopyFile(string originFile, string destinationFile)
        {
            if (File.Exists(originFile))
            {
                FileInfo fileOriginInfo = new FileInfo(originFile);
                FileInfo fileDestinationInfo = new FileInfo(destinationFile);
                try
                {
                    if (!Directory.Exists(fileDestinationInfo.DirectoryName))
                    {
                        Directory.CreateDirectory(fileDestinationInfo.DirectoryName);
                    }
                    RemoveAttributes(originFile);
                    if (File.Exists(destinationFile))
                        RemoveAttributes(destinationFile);
                    File.Copy(originFile, destinationFile, true);
                    Logger.PrintLC(originFile + " copied to " + 
                                   fileDestinationInfo.DirectoryName + " with the name: " + 
                                   fileDestinationInfo.Name, 2, ConfigFile.INFO);
                    return true;
                }
                catch(Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
        }


        public static bool CopyFile(string originFile, string destinationFile, bool bloccante)
        {
            if (File.Exists(originFile))
            {
                FileInfo fileOriginInfo = new FileInfo(originFile);
                FileInfo fileDestinationInfo = new FileInfo(destinationFile);
                try
                {
                    if (!Directory.Exists(fileDestinationInfo.DirectoryName))
                    {
                        Directory.CreateDirectory(fileDestinationInfo.DirectoryName);
                    }
                    RemoveAttributes(originFile);
                    if (File.Exists(destinationFile))
                        RemoveAttributes(destinationFile);
                    File.Copy(originFile, destinationFile, true);
                    Logger.PrintLC(originFile + " copied to " +
                                   fileDestinationInfo.DirectoryName + " with the name: " +
                                   fileDestinationInfo.Name, 2, ConfigFile.INFO);
                    return true;
                }
                catch (Exception exp)
                {
                    Logger.PrintLC("Could not copy file " + fileOriginInfo.FullName + " - Error: " + exp.Message, 2, ConfigFile.ERROR);
                    return false;
                }
            }
            else
            {
                Logger.PrintLC("Error recovering " + originFile + ". File doesn't exist.", 2, ConfigFile.ERROR);
                return false;
            }
        }




        /// <summary>
        /// Legge tutte le righe del file specificato e restituisce una collezione di righe
        /// </summary>
        /// <param name="File"></param>
        /// <param name="ListaRigheSqlFile"></param>
        /// <returns></returns>
        public static bool LeggiFile(string File, ref List<string> ListaRigheSqlFile)
        {
            try
            {
                int counter = 0;
                string line;

                // Read the file and display it line by line.  
                System.IO.StreamReader file =
                    new System.IO.StreamReader(File);
                while ((line = file.ReadLine()) != null)
                {
                    ListaRigheSqlFile.Add(line);
                    counter++;
                }

                file.Close();
                
            }
            catch
            {
                return false;
            }
            return true;
        }


    }
}
