using System;
using System.Windows.Forms;

public class ToExcel
{
	public ToExcel()
	{
	}

    public static void dataGVToExcel(DataGridView dGV)
    {
        try
        {
            //没有数据的话就不往下执行  
            if (dGV.Rows.Count == 0)
            {
                MessageBox.Show("Can't not save null record!");
                return;
            }
            //实例化一个Excel.Application对象  
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写  
            excel.Visible = true;

            //新增加一个工作簿，Workbook是直接保存，不会弹出保存对话框，加上Application会弹出保存对话框，值为false会报错  
            excel.Application.Workbooks.Add(true);
            //生成Excel中列头名称  
            for (int i = 0; i < dGV.Columns.Count; i++)
            {
                if (dGV.Columns[i].Visible == true)
                {
                    excel.Cells[1, i + 1] = dGV.Columns[i].HeaderText;
                }

            }
            //把DataGridView当前页的数据保存在Excel中  
            for (int i = 0; i < dGV.Rows.Count - 1; i++)
            {
                System.Windows.Forms.Application.DoEvents();
                for (int j = 0; j < dGV.Columns.Count; j++)
                {
                    if (dGV.Columns[j].Visible == true)
                    {
                        if (dGV[j, i].ValueType == typeof(string))
                        {
                            excel.Cells[i + 2, j + 1] = "'" + dGV[j, i].Value.ToString();
                        }
                        else
                        {
                            excel.Cells[i + 2, j + 1] = dGV[j, i].Value.ToString();
                        }
                    }

                }
            }

            //设置禁止弹出保存和覆盖的询问提示框  
            //excel.DisplayAlerts = false;
            //excel.AlertBeforeOverwriting = false;

            //保存工作簿  
            //excel.Application.Workbooks.Add(true).Save();
            //保存excel文件  
            //excel.Save("D:" + "\\KKHMD.xls");

            //确保Excel进程关闭  
            excel.Quit();
            excel = null;
            GC.Collect();//如果不使用这条语句会导致excel进程无法正常退出，使用后正常退出
            MessageBox.Show("Save success！", "MessagaBox");

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "error");
        }

    }
}
