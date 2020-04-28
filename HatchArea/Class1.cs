using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using DotNetARX;
using Autodesk.AutoCAD.Colors;

namespace HatchArea
{
    public class HatchArea
    {
      private readonly string[] _layArr = new string[] { "旱地", "果园", "有林地", "坑塘水面", "沟渠", "交通运输道路" };//要统计的图层

        [CommandMethod("JSMJ")]

        public void ComputeArea()
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //客户要求2010和2008均能运行
            string ver = acadApp.Version.ToString().Substring(0,2);//2010="18"
            ed.WriteMessage("\n支持的图层名：旱地，果园，有林地，坑塘水面，沟渠，交通运输道路\n其他图层不能被统计!");
            //对模型空间中的块参照进行统计
            var hatches = My_GetEntsInModelSpace<Hatch>(db, OpenMode.ForRead, false);
            //提示用户输入表格插入点
            PromptPointResult ppr = ed.GetPoint("\n请选择表格插入点:");
            if (ppr.Status != PromptStatus.OK) return;

            ObjectId currentTableId = createTable(db,ppr.Value);//创建表格
          
            double sum_mj=0;
            foreach (string layerName in _layArr)
            {
                using (Transaction ts = db.TransactionManager.StartTransaction())
                {
                    List<Hatch> currentHatches= hatches.Where(hatch => hatch.Layer.Contains(layerName)).ToList();
                    if (currentHatches.Count==0)
                        continue;//本图层无对象，换个图层。
                    double mj = currentHatches.Sum(h => h.Area);//正确
                    mj = mj / 10000; //换算为公顷
                    string mj2= Math.Round(mj, 4, MidpointRounding.AwayFromZero).ToString("f4");//保留4位
                    sum_mj+= Convert.ToDouble(mj2);//总面积
                    Table tb = currentTableId.GetObject(OpenMode.ForWrite, false) as Table;

                    string temp="";
                    switch (layerName)
                    {
                        case "旱地":
                            tb.SetTextString(3, 5, mj2);
                            tb.SetTextString(4, 5, mj2);
                            break;
                        case "果园":
                            tb.SetTextString(5, 5, mj2);
                            tb.SetTextString(6, 5, mj2);
                            break;
                        case "有林地":
                            tb.SetTextString(7, 5, mj2);
                            tb.SetTextString(8, 5, mj2);
                            break;
                        case "坑塘水面":
                            tb.SetTextString(9, 5, mj2);
                            temp = tb.Value(10, 5) as string;
                            temp = Math.Round(Convert.ToDouble(temp) + Convert.ToDouble(mj2), 4, MidpointRounding.AwayFromZero).ToString("f4");
                            tb.SetTextString(11, 5, temp);
                            break;
                        case "沟渠":
                            tb.SetTextString(10, 5, mj2);
                            temp = tb.Value(9, 5) as string;
                            temp = Math.Round(Convert.ToDouble(temp) + Convert.ToDouble(mj2), 4, MidpointRounding.AwayFromZero).ToString("f4");
                            tb.SetTextString(11, 5, temp);
                            break;
                        case "交通运输道路":
                            tb.SetTextString(12, 5, mj2);
                            tb.SetTextString(13, 5, mj2);
                            break;
                     }
                    ts.Commit();
                }
            }
            using (Transaction ts = db.TransactionManager.StartTransaction())
            {
                Table tb = currentTableId.GetObject(OpenMode.ForWrite, false) as Table;
                tb.SetTextString(14, 5, Math.Round(sum_mj, 4, MidpointRounding.AwayFromZero).ToString("f4"));//填写总面积
                //缩放到对象
                zoom_window(ed, tb);
                //窗口缩放命令
                //doc.SendCommand("ZOOM\nW\n"+ex.MinPoint.X+ "," +ex.MinPoint.Y + "\n" +ex.MaxPoint.X+","+ex.MaxPoint.Y+"\n");
                ts.Commit();
            }
            ed.WriteMessage("\n运行完毕!如有疑问请联系qq：985012864（CAD插件开发）");
        }
        //缩放到对象
        private void zoom_window (Editor ed ,Entity en)
        {
            Extents3d ex = en.GeometricExtents;
            Point3d p1 = ex.MinPoint;
            Point3d p2 = ex.MaxPoint;
            double h = p2.Y - p1.Y;
            double w = p2.X - p1.X;
            ViewTableRecord viewTableRecord = ed.GetCurrentView();
            viewTableRecord.CenterPoint = new Point2d(p1.X + (w / 2), p1.Y + (h / 2))
;            viewTableRecord.Height = h;
            viewTableRecord.Width = w;
            ed.SetCurrentView(viewTableRecord);
            Application.UpdateScreen();
        }
        private List<T> My_GetEntsInModelSpace<T>(Database db, OpenMode mode, bool openErased) where T : Entity
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //获取类型T代表的DXF代码名用于构建选择集过滤器
            string dxfname = RXClass.GetClass(typeof(T)).DxfName;
            //构建选择集过滤器   
            TypedValue[] values = { new TypedValue((int)DxfCode.Start, dxfname),
                new TypedValue((int)DxfCode.LayoutName,"Model")};
            SelectionFilter filter = new SelectionFilter(values);
            //选择符合条件的所有实体
            PromptSelectionOptions pso=new PromptSelectionOptions();
            pso.MessageForAdding = "\n请选择要统计的实体";
            PromptSelectionResult entSelected = ed.GetSelection(pso,filter);
            if (entSelected.Status != PromptStatus.OK) return null;
            SelectionSet ss = entSelected.Value;
            List<T> ents = new List<T>();
            using (Transaction ts = db.TransactionManager.StartTransaction())
            {
                ents.AddRange(ss.GetObjectIds().Select(id => ts.GetObject(id, mode, openErased)).OfType<T>());
            }
            return ents;
        }
        
        private ObjectId createTable(Database db,Point3d insertPt)
        {
            ObjectId objID = ObjectId.Null;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ObjectId styleId = AddTableStyle("MyTable");//get样式，若未定义则新建
                Table table = new Table();
                table.TableStyle = styleId;
                table.Position = insertPt;

                table.SetSize(15,6);//设定行/列总数
                table.SetRowHeight(3.75);     //设定行高
                table.SetRowHeight(0, 5);//标题行高
                table.SetColumnWidth(25);    //设定列宽

                //设定固定内容及合并单元格
                table.SetTextString(0,0, "复垦后土地利用结构调查表");//（为了与2008兼容）SetTextString为2010版本以下的方法
                table.SetTextString(1, 0, "一级类");
                table.SetTextString(1, 2, "二级类");
                table.SetTextString(1, 4, "复垦后面积/hm2");
                //合并单元格
                table.MergeCells(CellRange.Create(table, 1, 0, 1, 1));
                table.MergeCells(CellRange.Create(table, 1, 2, 1,3));
                table.MergeCells(CellRange.Create(table, 1, 4, 1, 5));
                List<string> btList = new string[] { "地类代码", "地类名称", "地类代码", "地类名称", "设计值","复核值" }.ToList();
                for (int i = 0; i < btList.Count; i++)
                {
                    //设置表头文字
                    table.Cells[2, i].TextString = btList[i];
                }
                table.SetTextString(3,0, "01");
                table.SetTextString(3, 1, "耕地");
                table.SetTextString(3, 2, "013");
                table.SetTextString(3, 3, "旱地");
                table.SetTextString(4, 2, "小计");
                table.MergeCells(CellRange.Create(table, 3, 0, 4, 0));
                table.MergeCells(CellRange.Create(table, 3, 1, 4, 1));
                table.MergeCells(CellRange.Create(table, 4, 2, 4, 3));

                table.SetTextString(5,0, "02");
                table.SetTextString(5, 1, "园地");
                table.SetTextString(5, 2, "021");
                table.SetTextString(5, 3, "果园");
                table.SetTextString(6, 2, "小计");
                table.MergeCells(CellRange.Create(table, 5, 0, 6, 0));
                table.MergeCells(CellRange.Create(table, 5, 1, 6, 1));
                table.MergeCells(CellRange.Create(table, 6, 2, 6, 3));

                table.SetTextString(7,0, "03");
                table.SetTextString(7, 1, "林地");
                table.SetTextString(7, 2, "031");
                table.SetTextString(7, 3, "有林地");
                table.SetTextString(8, 2, "小计");
                table.MergeCells(CellRange.Create(table, 7, 0, 8, 0));
                table.MergeCells(CellRange.Create(table, 7, 1, 8, 1));
                table.MergeCells(CellRange.Create(table, 8, 2, 8, 3));

                table.SetTextString(9,0, "11");
                table.SetTextString(9, 1, "水利设施用地");
                table.SetTextString(9, 2, "114");
                table.SetTextString(10, 2, "117");
                table.SetTextString(9, 3, "坑塘水面");
                table.SetTextString(10, 3, "沟渠");
                table.SetTextString(11, 2, "小计");
                table.MergeCells(CellRange.Create(table, 9, 0, 11, 0));
                table.MergeCells(CellRange.Create(table, 9, 1, 11, 1));
                table.MergeCells(CellRange.Create(table, 11, 2, 11, 3));

                table.SetTextString(12, 0, "10");
                table.SetTextString(12, 1, "交通运输用地");
                table.SetTextString(12, 2, "104");
                table.SetTextString(12, 3, "交通运输道路");
                table.SetTextString(13, 2, "小计");
                table.MergeCells(CellRange.Create(table, 12, 0, 13, 0));
                table.MergeCells(CellRange.Create(table, 12, 1, 13, 1));
                table.MergeCells(CellRange.Create(table, 13, 2, 13, 3));

                table.SetTextString(14, 0, "总计");
                table.MergeCells(CellRange.Create(table, 14, 0, 14, 3));
                set_DefaultFKHMJ(table);//设定面积默认值
                objID = db.AddToModelSpace(table);//添加到模型空间，table打开写模式后仍然可以修改
                trans.Commit();
            }
            return objID;
        }

        private void set_DefaultFKHMJ(Table table)
        {
            for (int c = 4; c < 6; c++)
            {
                for (int r = 3; r < 15; r++)
                {
                    if (c == 4)
                        table.SetTextString(r, c, "0");
                    else
                        table.SetTextString(r, c, "0.0000");
                }
            }
        }

        //为当前图形添加一个新的表格样式//CAD2010及以上版本可以用CAD2010版本以下的方法，但是不建议。
        public static ObjectId AddTableStyle(string style)
        {
            ObjectId styleId; // 存储表格样式的Id
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 打开表格样式字典
                DBDictionary dict = (DBDictionary)db.TableStyleDictionaryId.GetObject(OpenMode.ForRead);
                if (dict.Contains(style)) // 如果存在指定的表格样式
                    styleId = dict.GetAt(style); // 获取表格样式的Id
                else
                {
                    TableStyle ts = new TableStyle(); // 新建一个表格样式
                    // 设置表格的标题行为灰色
                    //ts.SetBackgroundColor(Color.FromColorIndex(ColorMethod.ByAci, 8), (int)RowType.TitleRow);
                    // 设置表格所有行的外边框的线宽为0.30mm
                    ts.SetGridLineWeight(LineWeight.LineWeight030, (int)GridLineType.OuterGridLines, TableTools.AllRows);
                    // 不加粗表格表头行的底部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalBottom, (int)RowType.HeaderRow);
                    // 不加粗表格数据行的顶部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalTop, (int)RowType.DataRow);
                    // 设置表格中所有行的文本高度为0.8
                    ts.SetTextHeight( 2.5, TableTools.AllRows);
                    // 设置表格中所有行的对齐方式为正中
                    ts.SetAlignment(CellAlignment.MiddleCenter, TableTools.AllRows);
                    dict.UpgradeOpen();//切换表格样式字典为写的状态
                    // 将新的表格样式添加到样式字典并获取其Id
                    styleId = dict.SetAt(style, ts);
                    // 将新建的表格样式添加到事务处理中
                    trans.AddNewlyCreatedDBObject(ts, true);
                    trans.Commit();
                }
            }
            return styleId; // 返回表格样式的Id
        }
    }
}
