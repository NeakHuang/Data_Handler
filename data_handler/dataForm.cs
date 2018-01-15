using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;

namespace data_handler {


    public partial class dataForm : Form {
        // 首行
        private List<string> firstRow = new List<string>();

        // 绑定实名认证
        private List<ExcelObj> list1 = new List<ExcelObj>();
        // 绑定账号
        private List<ExcelObj> list2 = new List<ExcelObj>();
        // 道具补发
        private List<ExcelObj> list3 = new List<ExcelObj>();
        // 后台功能需求
        private List<ExcelObj> list4 = new List<ExcelObj>();
        // 默认类型
        private List<ExcelObj> list5 = new List<ExcelObj>();
        // 其他问题
        private List<ExcelObj> list6 = new List<ExcelObj>();
        // 实名查询账号
        private List<ExcelObj> list7 = new List<ExcelObj>();
        // 数据关联
        private List<ExcelObj> list8 = new List<ExcelObj>();
        // 系统bug反馈
        private List<ExcelObj> list9 = new List<ExcelObj>();
        // 修改密码
        private List<ExcelObj> list10 = new List<ExcelObj>();
        // 已禁平台添加设备号
        private List<ExcelObj> list11 = new List<ExcelObj>();
        // 账号绑定信息修改
        private List<ExcelObj> list12 = new List<ExcelObj>();

        // 充值未到账-异常
        private List<ExcelObj> listOrderExcp = new List<ExcelObj>();
        // 充值未到账-漏单
        private List<ExcelObj> listOrderMiss = new List<ExcelObj>();

        private string _strFilePath;
        public dataForm() {
            InitializeComponent();
            // 关闭分析按钮，当选择文件后才可以点击
            button2.Enabled = false;

            // 关闭textBox1的输入
            textBox1.Enabled = false;
            // 关闭textBox2的输入
            textBox2.Enabled = false;
        }
        /// <summary>
        /// 分析文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e) {
            OpenFileDialog ofDiag = new OpenFileDialog();
            ofDiag.Filter = "文本文件|*.xls;*.xlsx";
            // 重置所有list
            resetList();

            if (ofDiag.ShowDialog() == DialogResult.OK) {
                this.textBox1.Text = "文件选取成功";
                _strFilePath = ofDiag.FileName;
                // 分析处理
                process_analyze();

                button2.Enabled = true;

            } else {
                this.textBox1.Text = "文件选取失败";
            }
        }
        /// <summary>
        /// 导出结果按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e) {
            string path = @"..\data\";
            string excPath = path + "分析结果.xlsx";
            if (File.Exists(excPath)) {
                File.Delete(excPath);
            }
            FileInfo fileInfo = new FileInfo(excPath);
            ExcelPackage package = new ExcelPackage(fileInfo);

            WriteTxt(path, "绑定实名认证", list1, package);
            WriteTxt(path, "绑定账号", list2, package);
            WriteTxt(path, "道具补发", list3, package);
            WriteTxt(path, "后台功能需求", list4, package);
            WriteTxt(path, "默认类型", list5, package);
            WriteTxt(path, "其他问题", list6, package);
            WriteTxt(path, "实名查询账号", list7, package);
            WriteTxt(path, "数据关联", list8, package);
            WriteTxt(path, "系统bug反馈", list9, package);
            WriteTxt(path, "修改密码", list10, package);
            WriteTxt(path, "已禁平台添加设备号", list11, package);
            WriteTxt(path, "账号绑定信息修改", list12, package);

            WriteOrders(@"..\data\", "充值未到账", listOrderMiss, listOrderExcp, package);

            this.textBox2.Text = "分析结果完毕";
        }

        /// <summary>
        /// 通用
        /// </summary>
        private void WriteTxt(string path, string name, List<ExcelObj> list, ExcelPackage package) {
            string txtPath = path + name + ".txt";
            if (File.Exists(txtPath)) {
                File.Delete(txtPath);
            }
            if (list.Count == 0) {
                return;
            }
            int index = 1;
            // 新增excel标签
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(name);
            // 设置标题列宽为
            sheet.Cells[index, 1].Value = "工单号";
            sheet.Cells[index, 2].Value = "标题";
            sheet.Column(2).Width = 100;

            // 新增文本
            System.IO.StreamWriter fs = new System.IO.StreamWriter(txtPath, true);
            
            foreach (ExcelObj obj in list) {
                index++;
                string line = index-1 + "\t工单号:" + obj.id + "\t" + obj.title;
                fs.WriteLine(line);// 直接追加文件末尾，换行 

                sheet.Cells[index, 1].Value = obj.id;
                sheet.Cells[index, 2].Value = obj.title;
            }
            fs.Close();

            //string excPath = path + "分析结果.xlsx";
            //Stream stream = new FileStream(path, FileMode.Create);分析结果.xlsx
            package.Save();
        }

        /// <summary>
        /// 订单相关写入
        /// </summary>
        private void WriteOrders(string path, string name, 
            List<ExcelObj> listMiss, List<ExcelObj> listExcp, ExcelPackage package) {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(name);

            string txtPath = path + name + ".txt";
            if (File.Exists(txtPath)) {
                File.Delete(txtPath);
            }
            // 设置标题列宽为
            sheet.Cells[1, 1].Value = "平台";
            sheet.Cells[1, 2].Value = "数量";
            sheet.Column(1).Width = 21;

            System.IO.StreamWriter fs = new System.IO.StreamWriter(txtPath, true);
            string[] channelName = { "APP", "天命APP", "天命安卓", "国际APP", "国际安卓", "应用宝",
                "VIVO", "多酷", "华为", "九游", "安卓360", "xy助手", "魅族", "oppo", "联想", "小米", "安卓91" };
            int[] channelNum = new int[channelName.Length];
            foreach (ExcelObj obj in listMiss) {
                //string line = obj.id + "\t" + obj.mod + "\t" + obj.missType;
                //fs.WriteLine(line);// 直接追加文件末尾，换行 
                if (obj.platform.Contains("APP区")) {
                    channelNum[0]++;
                } else if (obj.platform.Contains("天命")) {
                    if (obj.missChannel.Contains("APP")) {
                        channelNum[1]++;
                    } else {
                        channelNum[2]++;
                    }
                } else if (obj.platform.Contains("国际")) {
                    if (obj.missChannel.Contains("APP")) {
                        channelNum[3]++;
                    } else {
                        channelNum[4]++;
                    }
                } else if (obj.platform.Contains("安卓")) {
                    if (obj.missChannel.Contains("应用宝")) {
                        channelNum[5]++;
                    } else if (obj.missChannel.Contains("VIVO")) {
                        channelNum[6]++;
                    } else if (obj.missChannel.Contains("多酷")) {
                        channelNum[7]++;
                    } else if (obj.missChannel.Contains("华为")) {
                        channelNum[8]++;
                    } else if (obj.missChannel.Contains("九游")) {
                        channelNum[9]++;
                    } else if (obj.missChannel.Contains("360")) {
                        channelNum[10]++;
                    } else if (obj.missChannel.Contains("xy助手")) {
                        channelNum[11]++;
                    } else if (obj.missChannel.Contains("魅族")) {
                        channelNum[12]++;
                    } else if (obj.missChannel.Contains("oppo")) {
                        channelNum[13]++;
                    } else if (obj.missChannel.Contains("联想")) {
                        channelNum[14]++;
                    } else if (obj.missChannel.Contains("小米")) {
                        channelNum[15]++;
                    } else if (obj.missChannel.Contains("91")) {
                        channelNum[16]++;
                    }
                }
            }
            string channel;
            int num;
            for (int i = 0; i < channelName.Length; i++) {
                channel = channelName[i];
                num = channelNum[i];
                fs.WriteLine(channel + ":\t" + num);

                sheet.Cells[i+2, 1].Value = channel;
                sheet.Cells[i+2, 2].Value = num;
            }
            fs.WriteLine("漏单订单总数:\t" + listMiss.Count);
            fs.WriteLine("----------");
            fs.WriteLine("异常订单总数\t" + listExcp.Count);
            fs.WriteLine("----------");
            fs.WriteLine("充值未到账订单总数:\t" + (listMiss.Count + listExcp.Count));

            int maxRowNum = sheet.Dimension.End.Row;
            sheet.Cells[maxRowNum + 1, 1].Value = "漏单订单总数";
            sheet.Cells[maxRowNum + 1, 2].Value = listMiss.Count;
            sheet.Cells[maxRowNum + 2, 1].Value = "异常订单总数";
            sheet.Cells[maxRowNum + 2, 2].Value = listExcp.Count;
            sheet.Cells[maxRowNum + 3, 1].Value = "充值未到账订单总数";
            sheet.Cells[maxRowNum + 3, 2].Value = listMiss.Count + listExcp.Count;

            fs.Close();

            package.Save();
        }

        /// <summary>
        /// 处理分析
        /// </summary>
        private void process_analyze() {
            FileStream stream = File.Open(_strFilePath, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);

            //结构
            //row1 中文注释
            //row2 开始常规数据
            int row = 0;
            do {
                while (reader.Read()) {
                    row++;
                    ExcelObj excelObj = new ExcelObj();
                    for (int i = 0; i < reader.FieldCount; i++) {
                        string content = reader.GetValue(i).ToString();
                        if (row == 1) {
                            firstRow.Add(content.Trim());
                            continue;
                        }
                        if (i == firstRow.IndexOf("工单号")) {            // 工单号
                            excelObj.id = Convert.ToInt32(content);
                        } else if (i == firstRow.IndexOf("优先级")) {     // 优先级
                            excelObj.order = content;
                        } else if (i == firstRow.IndexOf("标题")) {       // 标题
                            excelObj.title = content;
                        } else if (i == firstRow.IndexOf("状态")) {       // 状态
                            excelObj.state = content;
                        } else if (i == firstRow.IndexOf("模板类型")) {   // 模板类型
                            excelObj.mod = content;
                        } else if (i == firstRow.IndexOf("发起人")) {     // 发起人
                            excelObj.senderName = content;
                        } else if (i == firstRow.IndexOf("受理人 ")) {    // 受理人 
                            excelObj.acceptName = content;
                        } else if (i == firstRow.IndexOf("漏单类型")) {   // 漏单类型
                            excelObj.missType = content;
                        } else if (i == firstRow.IndexOf("平台")) {       // 平台
                            excelObj.platform = content;
                        } else if (i == firstRow.IndexOf("渠道号")) {     // 渠道号
                            excelObj.channel = content;
                        } else if (i == firstRow.IndexOf("漏单渠道")) {   // 漏单渠道
                            excelObj.missChannel = content;
                        }
                    }
                    if (row != 1) {
                        // 归类内容
                        classifyExcelObj(excelObj);
                    }
                }
            } while (reader.NextResult());
        }

        private void textBox1_TextChanged(object sender, EventArgs e) {

        }

        /// <summary>
        /// 归类内容
        /// </summary>
        /// <param name="obj"></param>
        private void classifyExcelObj(ExcelObj obj) {
            //if (obj.title.Contains("无需处理") || obj.title.Contains("無需處理")) {
            //    return;
            //}
            if (obj.title.Contains("重复提交") || obj.title.Contains("重複提交")) {
                return;
            }
            if (obj.title.Contains("重复记录") || obj.title.Contains("重複記錄")) {
                return;
            }
            if (obj.title.Contains("模板错误") || obj.title.Contains("模板錯誤")) {
                return;
            }
            //if (obj.title.Contains("信息不全") || obj.title.Contains("信息不全")) {
            //    return;
            //}
            //if (obj.title.Contains("待跟进") || obj.title.Contains("待跟進")) {
            //    return;
            //}
            
            switch (obj.mod) {
                case "绑定实名认证":
                    list1.Add(obj);
                    break;
                case "绑定账号":
                    list2.Add(obj);
                    break;
                case "充值未到账":
                    if (obj.missType.Contains("异常") || obj.missType.Contains("異常")) {
                        listOrderExcp.Add(obj);
                    } else {
                        listOrderMiss.Add(obj);
                    }
                    break;
                case "道具补发":
                    list3.Add(obj);
                    break;
                case "后台功能需求":
                    list4.Add(obj);
                    break;
                case "默认类型":
                    list5.Add(obj);
                    break;
                case "其他问题":
                    list6.Add(obj);
                    break;
                case "实名查询账号":
                    list7.Add(obj);
                    break;
                case "数据关联":
                    list8.Add(obj);
                    break;
                case "系统bug反馈":
                    list9.Add(obj);
                    break;
                case "修改密码":
                    list10.Add(obj);
                    break;
                case "已禁平台添加设备号":
                    list11.Add(obj);
                    break;
                case "账号绑定信息修改":
                    list12.Add(obj);
                    break;
            }
        }

        private void resetList() {
            firstRow.Clear();

            list1.Clear();
            list2.Clear();
            list3.Clear();
            list4.Clear();
            list5.Clear();
            list6.Clear();
            list7.Clear();
            list8.Clear();
            list9.Clear();
            list10.Clear();
            list11.Clear();
            list12.Clear();

            listOrderExcp.Clear();
            listOrderMiss.Clear();
        }
    }
} // namespace
