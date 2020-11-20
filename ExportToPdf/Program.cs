using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExportToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            try 
            {
                DataTable objDataTable = new DataTable();

                //objDataTable.Columns.Add("ClientBidSno", typeof(long));
                //objDataTable.Columns.Add("ClientSno", typeof(long));
                //objDataTable.Columns.Add("ClientOfferSno", typeof(long));
                //objDataTable.Columns.Add("BidPrice", typeof(decimal));
                //objDataTable.Columns.Add("BidDescription", typeof(string));
                //objDataTable.Columns.Add("CompanySno", typeof(long));
                Console.WriteLine(" Building Data .. " + DateTime.Now.ToString());
                Console.WriteLine(); 

                for (int i = 0; i < 50; i++)
                {
                    objDataTable.Columns.Add("Col" + i.ToString(), typeof(string));
                }

                DataRow workRow;
                List<int> lstRepeatCol = new List<int> { 0 };

                for (int iRow = 0; iRow < 100; iRow++)
                {
                    workRow= objDataTable.NewRow();
                    for (int iCol = 0; iCol < 50; iCol++)
                    {
                        if (lstRepeatCol.Contains(iCol))
                        {
                            workRow[iCol] = "Row" + " " + iRow.ToString() + " RepCol " + iCol.ToString();                                               
                        }
                        else
                        {
                            workRow[iCol] = "Row" + " " + iRow.ToString();
                        }
 
                        
                    }
                    objDataTable.Rows.Add(workRow);
                }

                Console.WriteLine(" Exporting .. ");
                Console.WriteLine(); 

                ExportToPDFiTextSharp exp = new ExportToPDFiTextSharp();
                
                exp.ExportToPDF(objDataTable, "", null, lstRepeatCol);

                Console.WriteLine(" Finished .. ");
                Console.WriteLine(); 

                Console.ReadLine(); 


            }
            catch (Exception ex) {
                string msg = ex.Message;
            }


           
           
        }
    }
}
