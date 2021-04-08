using System;
using System.IO;
using System.Collections.Generic;
using SwiftExcel;

namespace MARC
{

    public class LibraryHelper
    {
        private FileMARC marcRecords = new FileMARC();
        public string[] placementDataFilter;
        public string sourcePath;
        public string destinationPath;
        public bool archiveSourceFile = true;
        public string archivePath;
        public string sourceFileFilter;


        public LibraryHelper()
        {
            //nothing special in the constructor so far... :)
        }

        public void Process()
        {

            DirectoryInfo DI = new DirectoryInfo(sourcePath);
            var Files = DI.GetFiles(sourceFileFilter);
            foreach (var file in Files)
            {
                ConvertToXLS(file.Name);
            }

        }

        private void ConvertToXLS(string fileName)
        {
            var sourceFile = Path.Combine(sourcePath, fileName);
            var dateString = DateTime.Now.ToShortDateString() + "-" + DateTime.Now.ToShortTimeString().Replace(":", string.Empty);
            var destFile = Path.Combine(destinationPath, $"{dateString}_{fileName}.xlsx");
            marcRecords.ImportMARC(sourceFile);
            System.Console.WriteLine($"Imported file '{sourceFile}'");
            ExportXLS(destFile);
            if (archiveSourceFile)
            {
                var archiveFile = Path.Combine(archivePath, $"{dateString}_{fileName}");
                File.Move(sourceFile, archiveFile);
            }
            else
            {
                File.Delete(sourceFile);
            }
            System.Console.WriteLine($"Exported file '{destFile}'");
        }

        private void ExportXLS(string toPath)
        {
            /** Fields:
                100 a (author)
                245 a (title)
                245 b (title)
                942 c (item type)
                952 a (permanent location)
                952 c (shelving location)
            **/
            int xlRow = 1;

            using (var ew = new ExcelWriter(toPath))
            {
                ew.Write($"Author", 1, xlRow);
                ew.Write($"Title", 2, xlRow);
                ew.Write($"Type", 3, xlRow);
                if (placementDataFilter != null)
                {
                    ew.Write($"Shelving location", 4, xlRow);
                }

                xlRow++;
                foreach (Record record in marcRecords)
                {
                    Field authorField = record["100"];
                    Field titleField = record["245"];
                    Field typeField = record["942"];
                    List<Field> locationFields = record.GetFields("952");

                    string Author;
                    string CoAuthors = "";
                    string Title;
                    string Type;
                    string Placement;

                    #region authorfield
                    if (authorField == null)
                    {
                        Author = "";
                    }
                    else
                    {
                        if (authorField.IsDataField())
                        {
                            DataField authorDataField = (DataField)authorField;
                            Subfield authorName = authorDataField['a'];
                            Author = authorName.Data;
                        }
                        else
                        {
                            Author = "";
                        }
                    }
                    #endregion

                    #region titleField
                    Title = "";
                    if (titleField == null)
                    {
                        Title = "";
                    }
                    else
                    {
                        if (titleField.IsDataField())
                        {
                            DataField titleDataField = (DataField)titleField;
                            Subfield titleA = titleDataField['a'];
                            Subfield titleB = titleDataField['b'];
                            Subfield coAuthorsField = titleDataField['c'];

                            if (titleA != null)
                            {
                                Title += titleA.Data;
                                if (Title.EndsWith("/"))
                                {
                                    Title = Title.Substring(0, Title.Length - 1);
                                }
                            }
                            if (titleB != null)
                            {
                                Title = Title.TrimEnd();
                                Title += " ";
                                Title += titleB.Data;
                                if (Title.EndsWith("/"))
                                {
                                    Title = Title.Substring(0, Title.Length - 1);
                                }
                            }
                            if (coAuthorsField != null)
                            {
                                CoAuthors = coAuthorsField.Data;
                            }
                        }
                        else
                        {
                            Title = "";
                        }
                    }
                    if (Title.EndsWith('.'))
                    {
                        Title = Title.Substring(0, Title.Length - 1);
                    }
                    #endregion

                    #region typeField
                    Type = "";
                    if (typeField != null)
                    {
                        if (typeField.IsDataField())
                        {
                            DataField typeDataField = (DataField)typeField;
                            Subfield typeData = typeDataField['c'];
                            if (typeData != null)
                            {
                                Type = typeData.Data;
                            }

                        }
                        else
                        {
                            Type = "";
                        }
                    }
                    #endregion

                    #region locationField
                    Placement = "";
                    if (locationFields != null)
                    {
                        if (locationFields[0].IsDataField())
                        {
                            if (placementDataFilter != null)
                            {
                                foreach (var filter in placementDataFilter)
                                {
                                    if (Placement == "")
                                    {
                                        foreach (DataField locationDataField in locationFields)
                                        {

                                            Subfield libraryData = locationDataField['a'];
                                            Subfield placementData = locationDataField['c'];
                                            if (libraryData.Data == filter)
                                            {
                                                Placement = placementData.Data;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Placement = "";
                        }
                    }
                    #endregion


                    if (Author.EndsWith(','))
                    {
                        Author = Author.Substring(0, Author.Length - 1);
                    }


                    ew.Write($"{Author}", 1, xlRow);
                    ew.Write($"{Title}", 2, xlRow);
                    ew.Write($"{Type}", 3, xlRow);
                    if (placementDataFilter != null)
                    {
                        ew.Write($"{Placement}", 4, xlRow);
                    }


                    xlRow++;


                }
            }
        }
    }
}