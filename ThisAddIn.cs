﻿using System.Windows.Media.Media3D;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        public static class Globle
        {
            public static int spotlight;             //聚光灯功能开关标识，0为关闭，1为打开
        }

        public static Excel.Application app;         //声明一个excel变量
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = Globals.ThisAddIn.Application;
            Globle.spotlight = 0;

            //监听选定单元格变化所触发事件
            app.SheetSelectionChange += app_SheetSelectionChange;

        }


        //选定单元格变化事件
        private void app_SheetSelectionChange(object sh, Excel.Range Target)
        {
            //当聚光灯功能打开时，变更选定单元格即触发
            if (Globle.spotlight == 1)
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
                    selectedRange.EntireRow.Interior.ColorIndex = 35;
                    selectedRange.EntireColumn.Interior.ColorIndex = 35;
                    ThisAddIn.app.ScreenUpdating = true;
                }
            }
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
    }
}
