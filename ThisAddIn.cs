using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    
    public partial class ThisAddIn
    {
        public static class Global
        {
            //是否执行重命名、删除、移动文件等功能标识,0为关闭，1为打开
            public static int readFile;

            //聚光灯功能开关标识，0为关闭，1为打开
            public static int spotlight;
            public static int spotlightColorIndex;
            
            public static bool created_qr_sheet=false;
        }

        public static Excel.Application app;         //声明一个Excel的Application变量
        public static string hotKey;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = Globals.ThisAddIn.Application;
            Global.spotlight = 0;

            //监听选定单元格变化所触发事件
            app.SheetSelectionChange += app_SheetSelectionChange;
            
            app.SheetBeforeDelete += new Excel.AppEvents_SheetBeforeDeleteEventHandler(Application_SheetBeforeDelete);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        //选定单元格变化事件
        private void app_SheetSelectionChange(object sh, Excel.Range Target)
        {
            //当聚光灯功能打开时，变更选定单元格即触发
            if (Global.spotlight == 1)
            {
                Excel.Range selectedRange = Target;
                Excel.Worksheet activesheet = ThisAddIn.app.ActiveSheet;

                //如果选择多个单元格时退出事件，选择1个单元格时触发事件
                if (selectedRange.Count != 1)
                {
                    return;
                }
                else
                {
                    activesheet.Cells.Interior.ColorIndex = 0;
                    ThisAddIn.app.ScreenUpdating = false;
                    selectedRange.EntireRow.Interior.ColorIndex =Global.spotlightColorIndex;
                    selectedRange.EntireColumn.Interior.ColorIndex = Global.spotlightColorIndex;
                    //selectedRange.EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204)); ;
                    //selectedRange.EntireColumn.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204)); ;
                    ThisAddIn.app.ScreenUpdating = true;
                }
            }
        }

        //监听删除表事件
        private void Application_SheetBeforeDelete(object Sh)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            if (ws.Name == "_rename")
            {
                Global.readFile = 0;    
            }
            if(ws.Name== "_QR_Scan")
            {
                Global.created_qr_sheet=false;
            }
        }        
    }
}
