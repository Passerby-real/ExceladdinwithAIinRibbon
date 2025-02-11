using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Net.Http;
using CharExtractorRibbon;
using System.IO;




namespace ExcelCharExtractor
{
    public partial class CharExtractorRibbon
    {
        private void CharExtractorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (File.Exists("settings.json")) // 现在File可以识别
            {
                string json = File.ReadAllText("settings.json");
                var settings = JsonConvert.DeserializeAnonymousType(json, new { ApiUrl = "", ApiKey = "", ModelName = "" });
                _apiUrl = settings.ApiUrl;
                _apiKey = settings.ApiKey;
                _modelName = settings.ModelName;
                txtApiUrl.Text = _apiUrl;
                txtApiKey.Text = _apiKey;
                txtModelName.Text = _modelName;
            }

            // 初始化下拉框
            var item1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item1.Label = "AI翻译";
            cmbTaskType.Items.Add(item1);

            var item2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item2.Label = "文笔润色";
            cmbTaskType.Items.Add(item2);
            var item3 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item3.Label = "市场价格";
            cmbTaskType.Items.Add(item3);

            // 确保事件只绑定一次
            btnSaveSettings.Click -= new RibbonControlEventHandler(btnSaveSettings_Click);
            btnSaveSettings.Click += new RibbonControlEventHandler(btnSaveSettings_Click);

            btnSendRequest.Click -= new RibbonControlEventHandler(btnSendRequest_Click);
            btnSendRequest.Click += new RibbonControlEventHandler(btnSendRequest_Click);
        }
      

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null) return;

            // 获取用户选择的类型
            string selectedType = 选择类型.SelectedItem.Label;

            // 遍历选中的每个单元格
            foreach (Excel.Range cell in selectedRange)
            {
                string cellValue = cell.Value2?.ToString() ?? string.Empty;
                string extractedText = ExtractCharacters(cellValue, selectedType);
                cell.Value2 = extractedText;
            }
        }

        private string ExtractCharacters(string input, string type)
        {
            switch (type)
            {
                case "汉字":
                    return new string(input.Where(c => char.IsLetter(c) && c >= 0x4E00 && c <= 0x9FFF).ToArray());
                case "英文":
                    return new string(input.Where(c => char.IsLetter(c) && c <= 0x007F).ToArray());
                case "数字":
                    return new string(input.Where(char.IsDigit).ToArray());
                case "英文标点":
                    // 提取英文标点符号
                    return new string(input.Where(c => IsEnglishPunctuation(c)).ToArray());
                case "中文标点":
                    // 提取中文标点符号
                    return new string(input.Where(c => IsChinesePunctuation(c)).ToArray());
                default:
                    return input;
            }
        }
        private bool IsEnglishPunctuation(char c)
        {
            // 英文标点符号的Unicode范围
            return (c >= 0x0020 && c <= 0x007F) && char.IsPunctuation(c);
        }

        // 判断是否为中文标点符号
        private bool IsChinesePunctuation(char c)
        {
            // 中文标点符号的Unicode范围
            return (c >= 0x3000 && c <= 0x303F) || (c >= 0xFF00 && c <= 0xFFEF);
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null) return;

            // 获取用户选择的类型
            string selectedType = 选择类型.SelectedItem.Label;

            // 遍历选中的每个单元格
            foreach (Excel.Range cell in selectedRange)
            {
                string cellValue = cell.Value2?.ToString() ?? string.Empty;
                string remainingText = RemoveCharacters(cellValue, selectedType);
                cell.Value2 = remainingText;
            }
        }
        private string RemoveCharacters(string input, string type)
        {
            switch (type)
            {
                case "汉字":
                    return new string(input.Where(c => !(char.IsLetter(c) && c >= 0x4E00 && c <= 0x9FFF)).ToArray());
                case "英文":
                    return new string(input.Where(c => !(char.IsLetter(c) && c <= 0x007F)).ToArray());
                case "数字":
                    return new string(input.Where(c => !char.IsDigit(c)).ToArray());
                case "英文标点":
                    return new string(input.Where(c => !IsEnglishPunctuation(c)).ToArray());
                case "中文标点":
                    return new string(input.Where(c => !IsChinesePunctuation(c)).ToArray());
                default:
                    return input;
            }
        }


        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前选中的区域
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null) return;

            // 创建字典，用于存储颜色和对应的数值总和
            var colorDict = new Dictionary<long, double>();

            // 获取选中区域的第一行和最后一行的行号
            int firstRow = selectedRange.Row;
            int lastRow = selectedRange.Row + selectedRange.Rows.Count - 1;

            // 获取新列的索引（选中区域的右侧）
            int newColIndex = selectedRange.Column + selectedRange.Columns.Count;

            // 遍历选中的每个单元格
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                // 获取单元格的背景颜色，并显式转换为long
                long colorCode = (long)cell.Interior.Color;

                // 如果字典中不存在该颜色，则添加到字典中
                if (!colorDict.ContainsKey(colorCode))
                {
                    colorDict[colorCode] = 0;
                }

                // 如果单元格的值是数字，则累加到对应颜色的总和中
                if (double.TryParse(cell.Value2?.ToString(), out double cellValue))
                {
                    colorDict[colorCode] += cellValue;
                }
            }

            // 在选中区域的右侧插入两列
            Excel.Range newColumns = selectedRange.Worksheet.Columns[newColIndex];
            newColumns.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            newColumns = selectedRange.Worksheet.Columns[newColIndex];
            newColumns.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);

            // 清除新列的背景颜色
            Excel.Range newRange = selectedRange.Worksheet.Range[
                selectedRange.Worksheet.Cells[firstRow, newColIndex],
                selectedRange.Worksheet.Cells[lastRow, newColIndex + 1]
            ];
            newRange.Interior.Pattern = Excel.XlPattern.xlPatternNone;

            // 将颜色和对应的总和写入新列
            int i = 0;
            foreach (var key in colorDict.Keys)
            {
                i++;
                selectedRange.Worksheet.Cells[firstRow - 1 + i, newColIndex].Interior.Color = key;
                selectedRange.Worksheet.Cells[firstRow - 1 + i, newColIndex + 1].Value2 = colorDict[key];
            }

            // 设置表头
            selectedRange.Worksheet.Cells[firstRow - 1, newColIndex].Value2 = "颜色";
            selectedRange.Worksheet.Cells[firstRow - 1, newColIndex + 1].Value2 = "求和";
        }

        private void 筛选选定值_Click(object sender, RibbonControlEventArgs e)


        {
            // 获取当前选中的区域
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null) return;

            // 获取选中的工作表
            Excel.Worksheet worksheet = selectedRange.Worksheet;

            // 判断当前表是否开启筛选
            if (worksheet.AutoFilterMode)
            {
                // 如果已经应用了筛选，则取消筛选
                try
                {
                    worksheet.ShowAllData();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // 如果取消筛选失败，则手动关闭筛选模式
                    worksheet.AutoFilterMode = false;
                }
            }
            else
            {
                // 如果未应用筛选，则应用筛选
                Excel.Range filterRange = selectedRange.CurrentRegion; // 获取当前区域

                // 获取选中的列
                int selectedColumn = selectedRange.Column;

                // 获取选中的多个值（支持非连续选区）
                List<string> selectedValues = new List<string>();
                foreach (Excel.Range area in selectedRange.Areas)
                {
                    foreach (Excel.Range cell in area.Cells)
                    {
                        string cellValue = cell.Value2?.ToString();
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            selectedValues.Add(cellValue);
                        }
                    }
                }

                // 应用筛选
                if (selectedValues.Count > 0)
                {
                    filterRange.AutoFilter(Field: selectedColumn - filterRange.Column + 1, Criteria1: selectedValues.ToArray(), Operator: Excel.XlAutoFilterOperator.xlFilterValues);
                }
            }
        }

        private Excel.Range GetFilterRange(Excel.Range selectedRange)
        {
            // 获取当前区域
            Excel.Range currentRegion = selectedRange.CurrentRegion;

            // 如果当前区域的第一行是表头，则从第二行开始筛选
            if (IsHeaderRow(currentRegion.Rows[1]))
            {
                return currentRegion.Offset[1, 0].Resize[currentRegion.Rows.Count - 1, currentRegion.Columns.Count];
            }

            return currentRegion;
        }

        private bool IsHeaderRow(Excel.Range row)
        {
            // 判断是否为表头行（假设表头行包含文本而非数字）
            foreach (Excel.Range cell in row.Cells)
            {
                double result;
                if (double.TryParse(cell.Value2?.ToString(), out result))
                {
                    return false;
                }
            }
            return true;
        }

        private void 单元格内每行重新编号_Click(object sender, RibbonControlEventArgs e)

        {
            // 获取当前活动工作表和选中的区域
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng = Globals.ThisAddIn.Application.Selection;

            // 遍历选中区域的每个单元格
            foreach (Excel.Range cell in rng.Cells)
            {
                if (cell.Value != null)
                {
                    string[] textLines = cell.Value.ToString().Split('\n');
                    StringBuilder newText = new StringBuilder();
                    int lineNum = 1;

                    foreach (string currentLine in textLines)
                    {
                        // 使用正则表达式删除行首的编号
                        Regex regex = new Regex(@"^\d+\.\s*", RegexOptions.Multiline);
                        string cleanedLine = regex.Replace(currentLine, "");

                        // 为当前行添加新的编号
                        string newLine = $"{lineNum}. {cleanedLine}";

                        // 添加到新文本中，除了最后一行外都添加换行符
                        if (textLines.Length > 1 && lineNum < textLines.Length)
                        {
                            newLine += "\n";
                        }

                        newText.Append(newLine);
                        lineNum++;
                    }

                    // 将处理后的文本写回单元格
                    cell.Value = newText.ToString();
                }
            }
        }
        private string _apiUrl;
        private string _apiKey;
        private string _modelName;

        private void btnSaveSettings_Click(object sender, RibbonControlEventArgs e)
        {
            _apiUrl = txtApiUrl.Text;
            _apiKey = txtApiKey.Text;
            _modelName = txtModelName.Text;

            var settings = new { ApiUrl = _apiUrl, ApiKey = _apiKey, ModelName = _modelName };
            string json = JsonConvert.SerializeObject(settings);
            System.IO.File.WriteAllText("settings.json", json);

            MessageBox.Show("设置已保存！");
        }

        private void btnSendRequest_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(_apiUrl) || string.IsNullOrEmpty(_apiKey) || string.IsNullOrEmpty(_modelName))
            {
                MessageBox.Show("请先保存API设置！");
                return;
            }

            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("请先选择一个区域！");
                return;
            }

            // 获取下拉选择的值
            string taskType = cmbTaskType.Text;
            string taskPrompt = "";

            switch (taskType)
            {
                case "AI翻译":
                    taskPrompt = "<角色>你是一名专业英文翻译，毕业于同声传译专业。你能够十分熟练的将各类中文翻译成专业的英文，或将各类英文翻译成专业的中文。</角色><任务>你的任务是帮助用户进行中文和英文之间的互译。</任务><要求>1. 用户的翻译场景有生活翻译场景、四六级翻译场景、雅思或托福翻译场景、论文翻译场景、文学作品翻译场景；2. 你要考虑用户翻译的场景，应用不同的翻译语法习惯；3. 你的翻译不能生搬硬套，需要考虑到用户整体输入的含义进行意译。4. 如果用户开始和你聊天，告诉他你在工作，你只会做翻译。</要求><输出要求>不要解释，直接输出翻译内容或回复。</输出要求>,下面是我需要翻译的内容，只需要将这个内容做翻译，无论是什么内容，你只需要翻译内容，不必对内容做回复";
                    break;
                case "文笔润色":
                    taskPrompt = "请仔细审查以下文本中各段的句子逻辑与连贯性，发现任何句子衔接、流畅性或整体结构可以优化的地方，并提出具体改进意见以提升内容的清晰度、易读性及学术质量。请先仅提供修改后的文本，并附上改进点的中文说明。";
                    break;
                case "市场价格":
                    taskPrompt = "你是一个具备网络访问能力的智能助手，在适当情况下，优先使用网络信息（参考信息）来回答，下面的内容我发送给你的是我需要你搜索的材料/设备的信息，我需要你参考最新消息给出推荐的价格。下面是需要查询的内容";
                    break;
                default:
                    MessageBox.Show("请先选择任务类型！");
                    return;
            }

            try
            {
                var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);
                http.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                // 遍历选中区域的每一行
                foreach (Excel.Range row in selectedRange.Rows)
                {
                    string question = taskPrompt + "\n" + row.Value2?.ToString();

                    var requestBody = new
                    {
                        model = _modelName,
                        messages = new[]
                        {
                    new { role = "user", content = question }
                },
                        tools = new[]
                        {
                    new
                    {
                        type = "web_search",
                        web_search = new
                        {
                            enable = true,
                            search_result = true
                        }
                    }
                }
                    };

                    var json = JsonConvert.SerializeObject(requestBody);
                    var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                    var response = http.PostAsync(_apiUrl, content).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = response.Content.ReadAsStringAsync().Result;
                        dynamic result = JsonConvert.DeserializeObject(responseContent);
                        string answer = result.choices[0].message.content;

                        // 将结果写入右侧单元格
                        Excel.Range targetCell = row.Cells[1, 1].Offset[0, 1]; // 向右移动一列
                        targetCell.Value2 = answer;

                        // 自动调整列宽
                        targetCell.EntireColumn.AutoFit();
                    }
                    else
                    {
                        MessageBox.Show("请求失败：" + response.StatusCode, "错误");
                    }
                }

                // 所有请求完成后弹出消息提示
                MessageBox.Show("AI已完成回复", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }


        private void btnDeleteSymbol_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取Excel应用对象
            Excel.Application excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Range selectedRange = null; // 在外部声明，确保finally块可以访问

            try
            {
                // 1. 弹出输入框获取用户输入的符号或文字
                string symbolOrText = ShowInputBox("请输入要删除的符号或文字:", "删除指定内容");

                if (string.IsNullOrEmpty(symbolOrText))
                {
                    MessageBox.Show("输入内容不能为空！");
                    return;
                }

                // 2. 获取选中的区域
                selectedRange = excelApp.Selection as Excel.Range;
                if (selectedRange == null)
                {
                    MessageBox.Show("请先选择一个区域！");
                    return;
                }

                // 3. 遍历所有单元格并删除指定内容
                foreach (Excel.Range cell in selectedRange)
                {
                    string originalValue = cell.Value2?.ToString();
                    if (!string.IsNullOrEmpty(originalValue))
                    {
                        // 使用正则表达式替换（匹配用户输入的内容）
                        string newValue = System.Text.RegularExpressions.Regex.Replace(
                            originalValue,
                            symbolOrText,
                            ""
                        );
                        cell.Value2 = newValue;
                    }
                }

                MessageBox.Show("删除完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message);
            }
            finally
            {
                // 释放COM对象
                if (selectedRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selectedRange);
            }
        }

        private string ShowInputBox(string prompt, string title)
        {
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button();

            form.Text = title;
            label.Text = prompt;
            textBox.Text = "";

            buttonOk.Text = "确定";
            buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | System.Windows.Forms.AnchorStyles.Right;
            buttonOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;

            form.ClientSize = new System.Drawing.Size(396, 107);
            form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk });
            form.ClientSize = new System.Drawing.Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;

            System.Windows.Forms.DialogResult dialogResult = form.ShowDialog();
            return dialogResult == System.Windows.Forms.DialogResult.OK ? textBox.Text : "";
        }

       
    } 
}
