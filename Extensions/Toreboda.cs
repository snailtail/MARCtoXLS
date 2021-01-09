using System;
using System.Collections.Generic;
using SwiftExcel;

namespace MARC
{

    public class Toreboda
    {
        private FileMARC marcRecords = new FileMARC();

        public Toreboda()
        {

        }

        public void ConvertToXLS(string fromPath, string toPath)
        {
            marcRecords.ImportMARC(fromPath);
            System.Console.WriteLine($"Imported file '{fromPath}'");
            ExportXLS(toPath);
            System.Console.WriteLine($"Exported file '{toPath}'");
        }

        private void ExportXLS(string toPath)
        {
            /** Fields:
                100 a (author)
                245 a (title)
                245 b (title)
                942 c (item type)
                952 a (permanent location) (TORE = Töreboda, 8BYI=Älgarås, 8BYS=Moholm)
                952 c (shelving location)
            **/
            int xlRow = 1;

            using (var ew = new ExcelWriter(toPath))
            {
                ew.Write($"Författare", 1, xlRow);
                ew.Write($"Titel", 2, xlRow);
                ew.Write($"Typ", 3, xlRow);
                ew.Write($"Placering", 4, xlRow);
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

                            // Tre steg för placeringslogiken
                            // 1. Om placering finns i Töreboda (TORE) används den
                            foreach (DataField locationDataField in locationFields)
                            {

                                Subfield libraryData = locationDataField['a'];
                                Subfield placementData = locationDataField['c'];
                                if (libraryData.Data == "TORE")
                                {
                                    Placement = placementData.Data;
                                }
                            }

                            // 2. Om placering är tom, kolla om den finns i Älgarås.
                            if (Placement == "")
                            {
                                foreach (DataField locationDataField in locationFields)
                                {

                                    Subfield libraryData = locationDataField['a'];
                                    Subfield placementData = locationDataField['c'];
                                    if (libraryData.Data == "8BYI")
                                    {
                                        Placement = placementData.Data;
                                    }
                                }
                            }

                            // 3. Om placering fortfarande är tom, kolla om den finns i Moholm.
                            if (Placement == "")
                            {
                                foreach (DataField locationDataField in locationFields)
                                {

                                    Subfield libraryData = locationDataField['a'];
                                    Subfield placementData = locationDataField['c'];
                                    if (libraryData.Data == "8BYS")
                                    {
                                        Placement = placementData.Data;
                                    }
                                }
                            }


                            /**
                            if(Placement=="") //om inte placeringen TORE, 8BYI eller 8BYS hittas, använd första bästa bara.
                            {
                                DataField locationDataField = (DataField)locationFields[0];
                                Subfield placementData = locationDataField['c'];
                                Placement = placementData.Data;
                            }
                            **/

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
                    ew.Write($"{Placement}", 4, xlRow);

                    xlRow++;


                }
            }
        }
    }
}