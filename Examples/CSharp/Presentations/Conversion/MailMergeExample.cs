using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using DataTable = System.Data.DataTable;

namespace CSharp.Presentations.Conversion
{
    // In this example, based on a simple presentation template and a simplified database. We demonstrate the
    // possibility of creating a set of presentations for each of the departments of an imaginary organization.
    // Each of the resulting presentations will include the name of the department, the name of the manager,
    // the staff of the department, and the chart for the schedule of the plan.

    public class MailMergeExample
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();
            string presTemplatePath = Path.Combine(dataDir, "PresentationTemplate.pptx");
            string resultPath = Path.Combine(RunExamples.OutPath, "MailMergeResult");

            // Path to the data.
            // XML data is one of the examples of the possible MailMerge data sources (among RDBMS and other types of data sources). 
            string dataPath = Path.Combine(dataDir, "TestData.xml");

            // Check if result path exists
            if (!Directory.Exists(resultPath))
                Directory.CreateDirectory(resultPath);

            // Creating DataSet using XML data
            using (DataSet dataSet = new DataSet())
            {
                dataSet.ReadXml(dataPath);

                DataTableCollection dataTables = dataSet.Tables;
                DataTable usersTable = dataTables["TestTable"];
                DataTable staffListTable = dataTables["StaffList"];
                DataTable planFactTable = dataTables["Plan_Fact"];

                // For all records in main table we will create a separate presentation
                foreach (DataRow userRow in usersTable.Rows)
                {
                    // create result (individual) presentation name
                    string presPath = Path.Combine(resultPath, "PresFor_" + userRow["Name"] + ".pptx");

                    // Load presentation template
                    using (Presentation pres = new Presentation(presTemplatePath))
                    {
                        // Fill text boxes with data from data base main table
                        ((AutoShape)pres.Slides[0].Shapes[0]).TextFrame.Text =
                            "Chief of the department - " + userRow["Name"];
                        ((AutoShape)pres.Slides[0].Shapes[4]).TextFrame.Text = userRow["Department"].ToString();

                        // Get image from data base
                        byte[] bytes = Convert.FromBase64String(userRow["Img"].ToString());

                        // insert image into picture frame of presentation
                        IPPImage image = pres.Images.AddImage(bytes);
                        IPictureFrame pf = pres.Slides[0].Shapes[1] as PictureFrame;
                        pf.PictureFormat.Picture.Image.ReplaceImage(image);

                        // Get abd prepare text frame for filling it with datas
                        IAutoShape list = pres.Slides[0].Shapes[2] as IAutoShape;
                        ITextFrame textFrame = list.TextFrame;

                        textFrame.Paragraphs.Clear();
                        Paragraph para = new Paragraph();
                        para.Text = "Department Staff:";
                        textFrame.Paragraphs.Add(para);

                        // fill staff data
                        FillStaffList(textFrame, userRow, staffListTable);

                        // fill plan fact data
                        FillPlanFact(pres, userRow, planFactTable);

                        pres.Save(presPath, SaveFormat.Pptx);
                    }
                }
            }
        }

        /// <summary>
        /// Fill text frame with datas from slave table as a list with bullet
        /// </summary>
        static void FillStaffList(ITextFrame textFrame, DataRow userRow, DataTable staffListTable)
        {
            foreach (DataRow listRow in staffListTable.Rows)
            {
                if (listRow["UserId"].ToString() == userRow["Id"].ToString())
                {
                    Paragraph para = new Paragraph();
                    para.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                    para.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
                    para.Text = listRow["Name"].ToString();
                    para.ParagraphFormat.Bullet.Color.ColorType = ColorType.RGB;
                    para.ParagraphFormat.Bullet.Color.Color = Color.Black;
                    para.ParagraphFormat.Bullet.IsBulletHardColor = NullableBool.True;
                    para.ParagraphFormat.Bullet.Height = 100;
                    textFrame.Paragraphs.Add(para);
                }
            }
        }

        /// <summary>
        /// Fills data chart from the secondary planFact table  
        /// </summary>
        static void FillPlanFact(Presentation pres, DataRow row, DataTable planFactTable)
        {
            IChart chart = pres.Slides[0].Shapes[3] as Chart;
            IChartTitle chartTitle = chart.ChartTitle;
            chartTitle.TextFrameForOverriding.Text = row["Name"] + " : Plan / Fact";

            DataRow[] selRows = planFactTable.Select("UserId = " + row["Id"]);
            string range = chart.ChartData.GetRange();

            IChartDataWorkbook cellsFactory = chart.ChartData.ChartDataWorkbook;
            int worksheetIndex = 0;

            chart.ChartData.Series[0].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 1, 1,
                    double.Parse(selRows[0]["PlanData"].ToString())));
            chart.ChartData.Series[1].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 1, 2,
                    double.Parse(selRows[0]["FactData"].ToString())));

            chart.ChartData.Series[0].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 2, 1,
                    double.Parse(selRows[1]["PlanData"].ToString())));
            chart.ChartData.Series[1].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 2, 2,
                    double.Parse(selRows[1]["FactData"].ToString())));

            chart.ChartData.Series[0].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 3, 1,
                    double.Parse(selRows[2]["PlanData"].ToString())));
            chart.ChartData.Series[1].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 3, 2,
                    double.Parse(selRows[2]["FactData"].ToString())));

            chart.ChartData.Series[0].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 3, 1,
                    double.Parse(selRows[3]["PlanData"].ToString())));
            chart.ChartData.Series[1].DataPoints.AddDataPointForLineSeries(
                cellsFactory.GetCell(worksheetIndex, 3, 2,
                    double.Parse(selRows[3]["FactData"].ToString())));

            chart.ChartData.SetRange(range);
        }
    }
}
