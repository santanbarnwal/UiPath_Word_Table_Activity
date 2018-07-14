using System;
using System.Activities;
using System.ComponentModel;
using System.Data;
using word = Microsoft.Office.Interop.Word;
using wordnm = Microsoft.Office.Interop.Word;
using colnm = System.Drawing.Color;
using System.IO;

namespace App_Integration
{
    namespace Word
    {
        namespace Table
        {
            [Designer(typeof(ActivityDesigner1))]
            public class Insert_Table : CodeActivity
            {
                [DisplayName("Path"),Category("Input"), RequiredArgument, Description("Word File Path")]
                public InArgument<String> Path1 { get; set; }
                [DisplayName("Data Table"), Category("Input"), RequiredArgument, Description("Data Table")]
                public InArgument<DataTable> dataTable { get; set; }
                [Category("Input"), DisplayName("Background Run"), Description("True: Task Performed in Background, False: Task Performed will be visible")]
                public InArgument<Boolean> Background { get; set; }

                [Category("Optional Color"),DisplayName("Header Background"), Description("Set Backgrond Color of Header"), DefaultValue(typeof(colnm))]
                public InArgument<colnm> BackHeaderColor { get; set; }
                [Category("Optional Color"), DisplayName("Header Text"), Description("Set Text Color of Header")]
                public InArgument<colnm> FontHeaderColor { get; set; }
                [Category("Optional Color"), DisplayName("Even Rows Background"), Description("Set Backgrond Color of Even Rows")]
                public InArgument<colnm> BackAlt1Color { get; set; }
                [Category("Optional Color"), DisplayName("Even Rows Text"), Description("Set Text Color of Even Rows")]
                public InArgument<colnm> FontAlt1Color { get; set; }
                [Category("Optional Color"), DisplayName("Odd Rows Background"), Description("Set Background Color of Odd Rows")]
                public InArgument<colnm> BackAlt2Color { get; set; }
                [Category("Optional Color"), DisplayName("Odd Rows Text"), Description("Set Text Color of Odd Rows")]
                public InArgument<colnm> FontAlt2Color { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    colnm cBackHeaderColor = BackHeaderColor.Get(context);
                    colnm cFontHeaderColor = FontHeaderColor.Get(context);
                    colnm cBackAlt1Color = BackAlt1Color.Get(context);
                    colnm cBackAlt2Color = BackAlt2Color.Get(context);
                    colnm cFontAlt1Color = FontAlt1Color.Get(context);
                    colnm cFontAlt2Color = FontAlt2Color.Get(context);
                    



                    if (cBackHeaderColor== colnm.Empty)
                    {
                        cBackHeaderColor = colnm.White;
                    }
                    if (cFontHeaderColor == colnm.Empty)
                    {
                        cFontHeaderColor = colnm.Black;
                    }
                    if (cBackAlt1Color == colnm.Empty)
                    {
                        cBackAlt1Color = colnm.White;
                    }
                    if (cBackAlt2Color == colnm.Empty)
                    {
                        cBackAlt2Color = colnm.White;
                    }
                    if (cFontAlt1Color == colnm.Empty)
                    {
                        cFontAlt1Color = colnm.Black;
                    }
                    if (cFontAlt2Color == colnm.Empty)
                    {
                        cFontAlt2Color = colnm.Black;
                    }
                   

                    WriteTableInWord(dataTable.Get(context),Background.Get(context), Path1.Get(context), cBackHeaderColor, cFontHeaderColor, cBackAlt1Color, cFontAlt1Color, cBackAlt2Color, cFontAlt2Color);
                }


                private static void WriteTableInWord(DataTable dt,Boolean Background, String Path1, colnm BackHeaderColor, colnm FontHeaderColor, colnm BackAlt1Color, colnm FontAlt1Color, colnm BackAlt2Color, colnm FontAlt2Color)
                {
                    wordnm.Application word = null;
                    wordnm.Document doc = null;

                    object oMissing = System.Reflection.Missing.Value;
                    object oEndOfDoc = "\\endofdoc";

                    try
                    {
                        word = new wordnm.Application();
                        if (Background)
                        {
                            word.Visible = false;
                        }
                        else
                        {
                            word.Visible = true;
                        }
                        if (File.Exists(Path1))
                        {
                            doc = word.Documents.Open(Path1);
                        }
                        else
                        {
                            doc = word.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (word != null && doc != null)
                    {
                        wordnm.Table newTable;
                        wordnm.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                        newTable = doc.Tables.Add(wrdRng, 1, dt.Columns.Count, ref oMissing, ref oMissing);
                        newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                        newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                        newTable.AllowAutoFit = true;

                        int count = 1;

                        foreach (var cell in dt.Rows[0].ItemArray)
                        {
                            newTable.Cell(newTable.Rows.Count, count).Range.Text = dt.Columns[count - 1].ColumnName;
                            newTable.Cell(newTable.Rows.Count, count).Shading.BackgroundPatternColor = convertColorToWdColor(BackHeaderColor);
                            newTable.Cell(newTable.Rows.Count, count).Range.Font.Color = convertColorToWdColor(FontHeaderColor);
                            count++;
                        }


                        int alt = 1;

                        foreach (DataRow row in dt.Rows)
                        {
                            count = 1;
                            newTable.Rows.Add();
                            foreach (var cell in row.ItemArray)
                            {
                                if (alt % 2 == 0)
                                {
                                    newTable.Cell(newTable.Rows.Count, count).Shading.BackgroundPatternColor = convertColorToWdColor(BackAlt1Color);
                                    newTable.Cell(newTable.Rows.Count, count).Range.Font.Color = convertColorToWdColor(FontAlt1Color);
                                }
                                else
                                {
                                    newTable.Cell(newTable.Rows.Count, count).Shading.BackgroundPatternColor = convertColorToWdColor(BackAlt2Color);
                                    newTable.Cell(newTable.Rows.Count, count).Range.Font.Color = convertColorToWdColor(FontAlt2Color);
                                }
                                newTable.Cell(newTable.Rows.Count, count).Range.Text = cell.ToString();
                                count++;
                            }
                            alt++;
                        }
                    }

                    object filename = Path1;

                    word.ActiveDocument.SaveAs(ref filename, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                               ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                               ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                               ref oMissing, ref oMissing, ref oMissing);
                    word.Quit();
                }


                private static wordnm.WdColor convertColorToWdColor(colnm col)
                {
                    wordnm.WdColor c = (wordnm.WdColor)(col.R + 0x100 * col.G + 0x10000 * col.B);
                    return c;
                }


            }

        }
    }
}
