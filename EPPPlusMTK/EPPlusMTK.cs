using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
/* To work eith EPPlus library */
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
/* For I/O purpose */
using System.IO;
/* For Diagnostics */
using System.Diagnostics;

namespace EPPPlusMTK
{
    public class EPPlusMTK
    {
        ExcelPackage package = new ExcelPackage();
        ExcelWorksheet worksheet;
        //Crear Archivo
        //public void CreateFile(string Directory,string Name,string Extension)       
        public void Open(string Directory, string Name, string Extension)
        {
            string File = Directory + '/' + Name + "." + Extension;

            package = new ExcelPackage(new FileInfo(File));
        }
        //Crear Hoja
        public void CreateSheet(string NameSheet)
        {
            worksheet = package.Workbook.Worksheets.Add(NameSheet);
        }

        //Ancho de la columna
        public void CellsHorizontalSize( Int32 ColumnStart, Int32 Size)
        {
            worksheet.Column(ColumnStart).Width = Size;
        }

        //Alto de la Fila
        public void CellsVerticalSize(Int32 RowStart, Int32 Size)
        {
            worksheet.Row(RowStart).Height = Size;
        }

        //Imprime En la Celda Indicada
        //PrintCell
        public void Cells(Int32 RowStart, Int32 ColumnStart, string Text)
        {
            worksheet.Cells[RowStart, ColumnStart].Value = Text;
        }

        //Imprime en varias celdas
        public void Cells(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, string Text)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = Text;
        }

        //Imprime Numero En la Celda Indicada
        public void CellsNumber(Int32 RowStart, Int32 ColumnStart, double Numero)
        {
            worksheet.Cells[RowStart, ColumnStart].Value = Numero;
        }

        public void CellsDate(Int32 RowStart, Int32 ColumnStart, String Date)
        {
            try
            {
                if (Date != "/  /" && Date != "")
                {
                    worksheet.Cells[RowStart, ColumnStart].Value = Convert.ToDateTime(Date);
                    worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "dd/mm/yyyy";
                }
                else
                {
                    worksheet.Cells[RowStart, ColumnStart].Value = "";
                }
            }
            catch(Exception ex){
                worksheet.Cells[RowStart, ColumnStart].Value = "";
            }
        }

        public void CellsDate(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, String Date)
        {
            try
            {
                if (Date != "/  /" && Date != "")
                {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = Convert.ToDateTime(Date);
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "dd/mm/yy";
                }
                else
                {
                    worksheet.Cells[RowStart, ColumnStart].Value = "";
                }
            }
            catch (Exception ex)
            {
                worksheet.Cells[RowStart, ColumnStart].Value = "";
            }
        }

        public void CellsDateTime(Int32 RowStart, Int32 ColumnStart, String Date)
        {
            try
            {
                if (Date != "/  /" && Date != "")
                {
            worksheet.Cells[RowStart, ColumnStart].Value = Convert.ToDateTime(Date);
            worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                 }
                else
                {
                    worksheet.Cells[RowStart, ColumnStart].Value = "";
                }
            }
            catch (Exception ex)
            {
                worksheet.Cells[RowStart, ColumnStart].Value = "";
            }
        }

        public void CellsDateTime(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, String Date)
        {
             try
             {
                 if (Date != "/  /" && Date != "")
                {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = Convert.ToDateTime(Date);
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "dd/mm/yy hh:mm:ss";
             }
                else
                {
                    worksheet.Cells[RowStart, ColumnStart].Value = "";
                }
            }
            catch (Exception ex)
            {
                worksheet.Cells[RowStart, ColumnStart].Value = "";
            }
        }


        //Imprime en varias celdas
        public void CellsNumber(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, double Numero)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = Numero;
        }

        //Alineacion al centro
        public void CellsHorizontalCenter(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        }

        //Alineacion al centro Varias
        public void CellsHorizontalCenter(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        }

        //Alineacion a Izquierda
        public void CellsHorizontalLeft(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        }

        //Alineacion a Izquierda Varias
        public void CellsHorizontalLeft(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        }

        //Alineacion a Derecha
        public void CellsHorizontalRight(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
        }

        //Alineacion a Derecha Varias
        public void CellsHorizontalRight(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
        }

        //Alineacion Justificada
        public void CellsHorizontalJustify(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify;
        }

        //Alineacion a Justificada Varias
        public void CellsHorizontalJustify(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify;
        }

        //Alineacion Arriba
        public void CellsVerticalTop(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        }

        //Alineacion a Arriba Varias
        public void CellsVerticalTop(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        }

        //Alineacion Vertical Centro
        public void CellsVerticalCenter(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }

        //Alineacion a Vertical Centro Varias
        public void CellsVerticalCenter(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }

        //Alineacion Vertical Base
        public void CellsVerticalBottom(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
        }

        //Alineacion a Vertical Base Varias
        public void CellsVerticalBottom(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
        }

        //Alineacion Vertical Justificado
        public void CellsVerticalJustify(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Justify;
        }

        //Alineacion a Vertical Justificado Varias
        public void CellsVerticalJustify(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Justify;
        }

        //Imprimir Formula En La Celda
        //PrintFormula
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, string Formula)
        {
            worksheet.Cells[RowStart, ColumnStart].Formula = Formula;
        }

        //Imprimir Formula En La Celda
        //PrintFormula
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, Int32 RowFirstCell, Int32 ColumnFirstCell, Int32 RowSecundCell, Int32 ColumnSecundCell, string Operation)
        {
            worksheet.Cells[RowStart, ColumnStart].Formula = "(" + worksheet.Cells[RowFirstCell, ColumnFirstCell].Address + Operation + worksheet.Cells[RowSecundCell, ColumnSecundCell].Address + ")";
        }

        //Imprimir Formula En Varias Celdas
        //PrintFormula
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, string Formula, Int32 RowStartRead, Int32 ColumnStartRead, Int32 RowLastRead, Int32 ColumnLastRead)
        {

            worksheet.Cells[RowStart, ColumnStart].Formula = Formula + "(" + worksheet.Cells[RowStartRead, ColumnStartRead].Address + ":" + worksheet.Cells[RowLastRead, ColumnLastRead].Address + ")";
        }

        //Igualar Celda
        //PrintFormula
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, Int32 RowRead, Int32 ColumnRead)
        {
            worksheet.Cells[RowStart, ColumnStart].Formula = worksheet.Cells[RowRead, ColumnRead].Address;
        }

        //Imprimir Formula En Varias Celdas
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, string Formula, string[] ArrayRow, string[] ArrayCol)
        {
            string arreglo = "";
            for (int x = 0; x < ArrayRow.Length; x++)
            {
                arreglo += worksheet.Cells[Convert.ToInt32(ArrayRow[x].ToString()), Convert.ToInt32(ArrayCol[x].ToString())].Address;
                if (ArrayRow.Length > x + 1)
                {
                    arreglo += ",";
                }
            }
            worksheet.Cells[RowStart, ColumnStart].Formula = Formula + "(" + arreglo.ToString() + ")";

        }

        //Imprimir Formula En Varias Celdas
        public void CellsFormula(Int32 RowStart, Int32 ColumnStart, string Formula, string Cells)
        {
            worksheet.Cells[RowStart, ColumnStart].Formula = Formula + "(" + Cells + ")";
        }

        //Imprimir Borde De Celda
        public void CellsBord(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        //Imprimir Borde de Celdas
        public void CellsBord(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        //Fondo De La Celda
        public void CellsColor(Int32 RowStart, Int32 ColumnStart, string Color)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[RowStart, ColumnStart].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(Color));
        }

        //Fondo De Varias Celdas
        public void CellsColor(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, string Color)
        {
            worksheet.Cells[RowStart, ColumnStart,RowLast,ColumnLast].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(Color));
        }

        //Letra En Negrita
        public void CellsBold(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Font.Bold = true;
        }

        //Negrita En Varias Celdas
        public void CellsBold(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Bold = true;
        }

        //Letra En Italica
        public void CellsItalic(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Font.Italic = true;
        }

        //Letra En Italica 
        public void CellsItalic(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Italic = true;
        }

        //Letra Subrayada 
        public void CellsUnderLine(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Font.UnderLine = true;
        }

        //Letra Subrayado Rango
        public void CellsUnderLine(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.UnderLine = true;
        }

        //Color de la Letra 
        public void CellsFontColor(Int32 RowStart, Int32 ColumnStart, string Color)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(Color));
        }

        //Color de la Letra Varias
        public void CellsFontColor(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, string Color)
        {
            worksheet.Cells[RowStart, ColumnStart,RowLast,ColumnLast].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(Color));
        }

        //Rotacion del Texto
        public void CellsRotation(Int32 RowStart, Int32 ColumnStart, Int32 Rotacion)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.TextRotation = Rotacion;
        }

        // Tamaño De Fuente 
        public void CellsFontSize(Int32 RowStart, Int32 ColumnStart, int Size)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Font.Size = Size;
        }

        //Tamaño De Funete Por Rango
        public void CellsFontSize(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast, int Size)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Size = Size;
        }

        //Combinacion de Celdas
        public void CellsCombine(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Merge = true;
        }

        // Ajustar Texto
        public void CellsWrapText(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.WrapText = true;
        }

        //Formato de Celda Numero
        public void CellsFormatNum(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "0.00";
        }

        //Formato de Celda Numero Rango
        public void CellsFormatNum(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart,RowLast,ColumnLast].Style.Numberformat.Format = "0.00";
        }

        //Formato de Celda fecha
        public void CellsFormatDate(Int32 RowStart, Int32 ColumnStart)
        {
            worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "mm-dd-yy";            
        }

        //Formato de Celda fecha Rango
        public void CellsFormatDate(Int32 RowStart, Int32 ColumnStart, Int32 RowLast, Int32 ColumnLast)
        {
            worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "mm-dd-yy";
        }

        //Grafica
        public void Grafica(Int32 RowStartDesign, Int32 ColumnStartDesign, Int32 RowStartRead, Int32 ColumnStartRead, Int32 RowLastRead, Int32 ColumnLastRead, int SizeWidth, int Height, string Name, string Title)
        {
            var chart = worksheet.Drawings.AddChart(Name, OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered);
            chart.Title.Text = Title;
            chart.SetPosition(RowStartDesign, 0, ColumnStartDesign, 0);
            chart.SetSize(SizeWidth, Height); // Tamaño de la gráfica
            chart.Legend.Remove(); // Si desea eliminar la leyenda

            // Define donde está la información de la gráfica.
            // Entiendase el nombre de la serie y los valores.
            var serie = chart.Series.Add(worksheet.Cells[ColumnStartRead+":" + ColumnLastRead], worksheet.Cells[RowStartRead + ":" + RowLastRead]);
        }

        public void Image(Int32 RowStart, Int32 ColumnStart, string imagePath, Int32 WidthSize, Int32 HeightSize)
        {
            Random random = new Random();
            int randomNumber = random.Next(0, 100);

            Bitmap image = new Bitmap(imagePath);
            ExcelPicture excelImage = null;
            if (image != null)
            {
                excelImage = worksheet.Drawings.AddPicture("Imagen" + randomNumber, image);
                excelImage.From.Column = ColumnStart;
                excelImage.From.Row = RowStart;
                excelImage.SetSize(WidthSize, HeightSize);
                // 2x2 px space for better alignment
                excelImage.From.ColumnOff = Pixel2MTU(2);
                excelImage.From.RowOff = Pixel2MTU(2);
            }
        }
        public int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }
        //Guardar Archivo
        public void Save()
        {           
            package.Save();
        }
    }
}
