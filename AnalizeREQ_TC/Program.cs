using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
//Create COM Objects.Create a COM object for everything that is referenced

//C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c dodać referencje z palca

namespace AnalizeREQ_TC
{
    class TestPlan
    {
        public List<string> TC_ID_List; // list of ID of TC
        public List<string> REQ_List; // list of REQ
        public List<List<string>> REQs_ = new List<List<string>>(); // [TCIndex][REQs]
        public List<TC> TC_List = new List<TC>();
        public List<REQ> REQ_List_Obj = new List<REQ>();
        public TestPlan(List<string> TC_ID_List_C, List<string> REQ_List_C)
        {
            TC_ID_List = TC_ID_List_C;
            REQ_List = REQ_List_C;
        }

        public void SplitREQ()
        {
            foreach (var item in REQ_List)
            {
                List<string> tmpListOfREQs = item.Split(',').ToList();
                List<string> tmpList = new List<string>();
                foreach (var item2 in tmpListOfREQs)
                {
                    string tmp = item2;
                    tmp = tmp.Replace("https://doors.dgs.com:8443/dwa/rm/urn:rational::1-4fa78175254e4271-O-", "");
                    tmp = tmp.Replace("-0000232f", "");
                    tmpList.Add(tmp);
                }
                REQs_.Add(tmpList);

                //REQs_.Add(item.Split(',').ToList());
            }
            REQ_List.Clear();
        }

        public void convertLinkToREQ_Number()
        {
            for (int i = 0; i < REQs_.Count(); i++)
            {
                int j = 0;
                foreach (var item in REQs_[i])
                {
                    string tmp = item;
                    tmp = tmp.Replace("https://doors.dgs.com:8443/dwa/rm/urn:rational::1-4fa78175254e4271-O-", "");
                    tmp = tmp.Replace("- 0000232f", "");
                    REQs_[i][j] = tmp;
                    j++;
                }
            }
        }

        public void getAllREQFromTCs()
        {
            foreach (var item in TC_List)
            {
                foreach (var REQfromTC in item.REQ_List)
                {
                    REQ_List.Add(REQfromTC);
                }                
            }
        }

        public void showAllREQ()
        {
            foreach (var item in REQ_List)
            {
                Console.WriteLine(item);
            }
        }

        public void showDoubleCoverage()
        {
            List<string> uniqueREQ = REQ_List.Select(x => x).Distinct().ToList();

            foreach (var item in uniqueREQ)
            {
                int count = 0;
                for (int i = 0; i < REQ_List.Count(); i++)
                {
                    if (REQ_List[i] == item)
                    {
                        count++;
                    }
                }
                //if (count > 1)
                //{
                    Console.Write("REQ: " + item + " Exist: " + count + " x in TC :\n");
                    for (int i = 0; i < TC_List.Count(); i++)
                    {
                        if (TC_List[i].checkIfREQ_Exist(item))
                        {
                            Console.WriteLine("\t"+TC_List[i].ID);
                        }
                    }
                //}
                //else
                //Console.WriteLine("REQ: " + item + " Exist: " + count + " x in TC :");
            }

        }

        public void CreateTCs()
        {
            for (int i = 0; i < TC_ID_List.Count(); i++)
            {
                TC_List.Add(new TC(TC_ID_List[i], REQs_[i]));
            }           
        }

        public void CreateREQs()
        {
            for (int i = 0; i < TC_List.Count(); i++)
            {
                for (int j = 0; j < TC_List[i].REQ_List.Count(); j++)
                {
                    REQ tmp = new REQ(TC_List[i].ID);
                    tmp.AddRef(TC_List[i].REQ_List[j]);
                    REQ_List_Obj.Add(tmp);
                }
            }            
        }

        public void printAll_REQ()
        {
            for (int i = 0; i < REQ_List_Obj.Count(); i++)
            {
                REQ_List_Obj[i].printREQ();
            }
        }

        public void printAll_TC_REQ()
        {
            for (int i = 0; i < TC_List.Count(); i++)
            {
                TC_List[i].printREQ();
            }
        }



    }

    class TC
    {
        public string ID;
        public List<string> REQ_List; // list of REQ
        public List<REQ> REQ_obj = new List<REQ>(); // lista obiektów REQ 

        public TC(string name, List<string> REQ_List_C)
        {
            ID = name;
            REQ_List = REQ_List_C;
            for (int i = 0; i < REQ_List_C.Count(); i++)
            {
                REQ tmp = new REQ(REQ_List_C[i]);
                tmp.AddRef(ID);
                REQ_obj.Add(tmp);
            }
        }

        public void printREQ()
        {
            Console.WriteLine("TC ID: " + ID + "\nREQ: ");
            foreach (var item in REQ_List)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine("\n");
        }

        public bool checkIfREQ_Exist(string name)
        {
            foreach (var item in REQ_List)
            {
                if (item == name)
                {
                    return true;
                }
            }
            return false;
        }

    }


    class REQ
    {
        string ID;
        List<string> ConnectedTC_ID = new List<string>();

        public REQ(string name)
        {
            ID = name;
        }

        public bool checkIfRefExist(string ID) // jeżeli istnieje na liscie to true jezeli nie to false
        {
            foreach (var item in ConnectedTC_ID)
            {
                if (item == ID)
                {
                    return true;
                }
            }
            return false;

        }

        public void AddRef(string ID)
        {
            if (!checkIfRefExist(ID))
            {
                ConnectedTC_ID.Add(ID);
            }
        }

        public void printREQ()
        {
            Console.WriteLine("TC ID: " + ID + "\n REQ: ");
            foreach (var item in ConnectedTC_ID)
            {
                Console.WriteLine(item + " ,");
            }
            Console.WriteLine("\n");
        }

    }
    class Program
    {
        
        static void Main(string[] args)
        {

            Console.WriteLine("Path to xlsx file");
            
           TestPlan MyTestPlan = getExcelFile(Console.ReadLine()); //oject with list of REQ and TC IDs

            MyTestPlan.SplitREQ();
           // MyTestPlan.convertLinkToREQ_Number();
            MyTestPlan.CreateTCs();
            MyTestPlan.CreateREQs();
            //MyTestPlan.printAll_REQ();
            MyTestPlan.printAll_TC_REQ();
            MyTestPlan.getAllREQFromTCs();
            //MyTestPlan.showAllREQ();
            MyTestPlan.showDoubleCoverage();

        }

        public static TestPlan getExcelFile(string path)
        {

            List<string> TC_ID_List = new List<string>(); // list of ID of TC
            List<string> REQ_List = new List<string>(); // list of REQ


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);//@"C:\Users\paze\Documents\GitHub\AnalizeREQ_TC\AnalizeREQ_TC\test.xlsx"
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    TC_ID_List.Add(xlRange.Cells[i, 1].Value2.ToString());


                for (int j = 2; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        REQ_List.Add(xlRange.Cells[i, j].Value2.ToString());
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return new TestPlan(TC_ID_List, REQ_List);
        }
    }
}
