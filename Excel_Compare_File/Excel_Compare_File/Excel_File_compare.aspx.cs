using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace Excel_Compare_File
{
    public partial class Excel_File_compare : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            cardresult.Visible = false;
        }
        protected void ButtonCompare_Click(object sender, EventArgs e)
        {
            try
            {
                if (FileUpload1.HasFile && FileUpload2.HasFile)
                {
                    DataTable dt1 = ReadExcelFile(FileUpload1.PostedFile.InputStream);
                    DataTable dt2 = ReadExcelFile(FileUpload2.PostedFile.InputStream);

                    if (!ColumnNamesMatch(dt1, dt2))
                    {
                        LabelResult.Text = "Column names do not match. Please re-upload Excel files with matching column names.";
                        GridViewResult.Visible = false;
                        return;
                    }

                    var comparisonResult = CompareDataTables(dt1, dt2);
                    int totalDifferences = comparisonResult.Rows.Count;

                    if (totalDifferences > 0)
                    {
                        cardresult.Visible = true;
                        GridViewResult.DataSource = comparisonResult;
                        GridViewResult.DataBind();
                        GridViewResult.Visible = true;
                        LabelResult.Text = $"Differences found: {totalDifferences}";
                    }
                    else
                    {
                        LabelResult.Text = "No differences found.";
                        ///hello
                        GridViewResult.Visible = false;
                    }
                }
                else
                {
                    LabelResult.Text = "Please upload both files.";
                }
            }
            catch (Exception ex)
            {
                LabelResult.Text = $"An error occurred: {ex.Message}";
            }
        }
        private bool ColumnNamesMatch(DataTable dt1, DataTable dt2)
        {
            if (dt1.Columns.Count != dt2.Columns.Count)
            {
                return false;
            }

            for (int i = 0; i < dt1.Columns.Count; i++)
            {
                if (!dt1.Columns[i].ColumnName.Equals(dt2.Columns[i].ColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }
        private DataTable ReadExcelFile(Stream fileStream)
        {
            DataTable dt = new DataTable();
            using (ExcelPackage package = new ExcelPackage(fileStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                bool hasHeader = true;

                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dt.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }

                var startRow = hasHeader ? 2 : 1;

                for (int rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dt.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        private DataTable CompareDataTables(DataTable dt1, DataTable dt2)
        {

            DataTable dtResult = dt1.Clone();


            dtResult.Columns.Add("Line Number", typeof(int));

            int rowCount = Math.Min(dt1.Rows.Count, dt2.Rows.Count);


            for (int i = 0; i < rowCount; i++)
            {
                DataRow row1 = dt1.Rows[i];
                DataRow row2 = dt2.Rows[i];
                DataRow resultRow = dtResult.NewRow();
                bool isMismatch = false;


                for (int j = 0; j < dt1.Columns.Count; j++)
                {
                    if (!row1[j].Equals(row2[j]))
                    {
                        resultRow[j] = $"Mismatch: {row1[j]} vs {row2[j]}";
                        isMismatch = true;
                    }
                    else
                    {
                        resultRow[j] = row1[j];
                    }
                }


                if (isMismatch)
                {
                    resultRow["Line Number"] = i + 2;
                    dtResult.Rows.Add(resultRow);
                }
            }


            if (dt1.Rows.Count > rowCount)
            {
                AddRemainingRows(dt1, dtResult, rowCount);
            }
            if (dt2.Rows.Count > rowCount)
            {
                AddRemainingRows(dt2, dtResult, rowCount);
            }

            return dtResult;
        }


        private void AddRemainingRows(DataTable sourceDt, DataTable dtResult, int startIndex)
        {
            for (int i = startIndex; i < sourceDt.Rows.Count; i++)
            {
                DataRow row = sourceDt.Rows[i];
                DataRow resultRow = dtResult.NewRow();

                for (int j = 0; j < sourceDt.Columns.Count; j++)
                {
                    resultRow[j] = row[j];
                }

                resultRow["Line Number"] = i + 1;
                dtResult.Rows.Add(resultRow);
            }
        }


        protected void GridViewResult_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                foreach (TableCell cell in e.Row.Cells)
                {
                    if (cell.Text.Contains("Mismatch"))
                    {
                        cell.ForeColor = Color.Red;
                    }
                }
            }
        }

        protected void ButtonExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (GridViewResult.Rows.Count > 0)
                {
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Differences");


                        for (int i = 0; i < GridViewResult.HeaderRow.Cells.Count; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = GridViewResult.HeaderRow.Cells[i].Text;
                            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                            worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        for (int i = 0; i < GridViewResult.Rows.Count; i++)
                        {
                            for (int j = 0; j < GridViewResult.Rows[i].Cells.Count; j++)
                            {
                                if (GridViewResult.Rows[i].Cells[j].Controls.Count > 0)
                                {
                                    var cellValue = (GridViewResult.Rows[i].Cells[j].Controls[0] as Label)?.Text;
                                    worksheet.Cells[i + 2, j + 1].Value = HttpUtility.HtmlDecode(cellValue ?? string.Empty);
                                }
                                else
                                {
                                    string decodedText = HttpUtility.HtmlDecode(GridViewResult.Rows[i].Cells[j].Text);
                                    worksheet.Cells[i + 2, j + 1].Value = decodedText;
                                }
                            }
                        }


                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();


                        using (MemoryStream stream = new MemoryStream())
                        {
                            package.SaveAs(stream);
                            var fileName = "Differences_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                            Response.Clear();
                            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            Response.AddHeader("content-disposition", "attachment; filename=" + fileName);
                            Response.BinaryWrite(stream.ToArray());
                            Response.End();
                        }
                    }
                }
                else
                {
                    LabelResult.Text = "No data available to export.";
                }
            }
            catch (Exception ex)
            {
                LabelResult.Text = $"An error occurred during export: {ex.Message}";
            }
        }

        protected void ButtonRefresh_Click(object sender, EventArgs e)
        {
            FileUpload1.Attributes.Clear();
            FileUpload2.Attributes.Clear();
            GridViewResult.Visible = false;
            LabelResult.Text = string.Empty;
        }

    }
}