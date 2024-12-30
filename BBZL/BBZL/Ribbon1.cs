using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace BBZL
{
    public partial class Ribbon1
    {
        PowerPoint.Application app;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }
        //点击按钮时，将化学式中的数字替换为下标
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ReplaceWithSubscript(app.ActivePresentation);
        }

        // 点击按钮时，设置选中的文本框的轮廓样式
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (shape.Type == Office.MsoShapeType.msoTextBox || shape.Type == Office.MsoShapeType.msoPlaceholder)
                    {
                        // 设置形状轮廓的粗细
                        shape.Line.Weight = 1.5f;

                        // 设置形状轮廓的虚线样式
                        shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                        // 设置形状轮廓的颜色
                        shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#C00000"));
                    }
                }
            }
        }

        //点击按钮时，裁剪选中的形状
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            float cropWidth, cropHeight;

            if (float.TryParse(editBox1.Text, out cropWidth) && float.TryParse(editBox2.Text, out cropHeight))
            {
                CropSelectedShapes(cropWidth, cropHeight);
            }
            else
            {
                MessageBox.Show("请输入有效的宽度和高度。", "无效输入", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //点击按钮时，将选中的表格格式化为三线表
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSelectedTableAsThreeLineTable();
        }

        //点击按钮时，使用word功能识别错别字及语法问题
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            CheckSpellingAndHighlightCurrentSlide(app.ActiveWindow.View.Slide);
        }




        // 将化学式中的数字替换为下标的方法
        private void ReplaceWithSubscript(PowerPoint.Presentation pres)
        {
            string pattern = @"([A-Z][a-z]*)(\d*)"; // 匹配化学式的正则表达式模式

            foreach (PowerPoint.Slide slide in pres.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        ProcessTextRange(shape.TextFrame.TextRange, pattern);
                    }
                    else if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        foreach (PowerPoint.Row row in shape.Table.Rows)
                        {
                            foreach (PowerPoint.Cell cell in row.Cells)
                            {
                                ProcessTextRange(cell.Shape.TextFrame.TextRange, pattern);
                            }
                        }
                    }
                }
            }
        }

        //将形状、文本框及表格内的文本都遍历进去的方法
        private void ProcessTextRange(PowerPoint.TextRange textRange, string pattern)
        {
            var matches = Regex.Matches(textRange.Text, pattern);

            foreach (Match match in matches)
            {
                string element = match.Groups[1].Value;
                string subscript = match.Groups[2].Value;

                if (!string.IsNullOrEmpty(subscript))
                {
                    int startIndex = match.Index + element.Length + 1;
                    for (int i = 0; i < subscript.Length; i++)
                    {
                        textRange.Characters(startIndex + i, 1).Font.Subscript = Office.MsoTriState.msoTrue;
                    }
                }
            }
        }

        //裁剪选中的形状的方法
        private void CropSelectedShapes(float cropWidth, float cropHeight)
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    float centerX = shape.Left + shape.Width / 2;
                    float centerY = shape.Top + shape.Height / 2;

                    // 将厘米转换为磅，确保输入与输出的单位一致
                    float cropWidthInPoints = cropWidth * 72f / 2.54f;
                    float cropHeightInPoints = cropHeight * 72f / 2.54f;

                    shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    shape.Width = cropWidthInPoints;
                    shape.Height = cropHeightInPoints;
                    shape.Left = centerX - cropWidthInPoints / 2;
                    shape.Top = centerY - cropHeightInPoints / 2;
                }
            }
        }

        //将选中的表格格式化为三线表的方法
        private void FormatSelectedTableAsThreeLineTable()
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.Table table = shape.Table;

                        // 清除表格的底纹颜色并设置文字颜色为黑色，文字居中
                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                table.Cell(i, j).Shape.Fill.BackColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                table.Cell(i, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            }
                        }

                        // 设置表格的上边框、第一排单元格的底边框和最后一排单元格的下边框
                        for (int j = 1; j <= table.Columns.Count; j++)
                        {
                            // 上边框
                            table.Cell(1, j).Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 1.5f;
                            table.Cell(1, j).Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                            // 第一排单元格的底边框
                            table.Cell(1, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 1.5f;
                            table.Cell(1, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                            // 最后一排单元格的下边框
                            table.Cell(table.Rows.Count, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 1.5f;
                            table.Cell(table.Rows.Count, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        }

                        // 清除其他边框
                        for (int i = 2; i < table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 0f;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 0f;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0f;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0f;
                            }
                        }
                    }
                }
            }
        }

        //使用word识别错别字及语法问题的方法
        private void CheckSpellingAndHighlightCurrentSlide(PowerPoint.Slide slide)
        {
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            Word.Document tempDoc = wordApp.Documents.Add();
            Word.Range wordRange = tempDoc.Content;

            // 将当前幻灯片的文本复制到临时Word文档中
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    string text = textRange.Text;

                    wordRange.Text += text + "\n";
                }
            }

            wordRange.CheckSpelling();
            wordRange.CheckGrammar();

            // 将错误标记回当前幻灯片
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    string text = textRange.Text;

                    foreach (Word.Range error in wordRange.SpellingErrors)
                    {
                        int start = error.Start;
                        int length = error.End - error.Start;

                        PowerPoint.TextRange errorRange = textRange.Characters(start + 1, length);
                        errorRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                    }

                    foreach (Word.Range error in wordRange.GrammaticalErrors)
                    {
                        int start = error.Start;
                        int length = error.End - error.Start;

                        PowerPoint.TextRange errorRange = textRange.Characters(start + 1, length);
                        errorRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                    }
                }
            }

            tempDoc.Close(false);
            wordApp.Quit();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Form infoForm = new Form();
            infoForm.Text = "插件信息";
            infoForm.Size = new System.Drawing.Size(400, 200);

            Label infoLabel = new Label();
            infoLabel.Text = "免责声明：\n本插件仅供学习和交流使用，\n作者不对因使用本插件而产生的任何后果负责，\n若您认为侵犯了您的权益，\n请尽快与我们联系。\n\n作者信息：\n作者：一炙穿云箭\n邮箱：2321551492@qq.com";
            infoLabel.AutoSize = true;
            infoLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            infoLabel.Dock = DockStyle.Fill;

            infoForm.Controls.Add(infoLabel);
            infoForm.ShowDialog();
        }
    }
}










