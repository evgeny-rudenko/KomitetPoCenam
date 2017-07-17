using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel; 


namespace Ceny2014
{
    class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Worksheet ErSheet = null;

        /// <summary>
        /// Максимальные цены
        /// опт и розница
        /// </summary>
        private struct SCena
        {
            public double Opt;
            public double Rozn;
        }
        static SCena GetMaxPrice (float Reestr)
        {
            SCena Cena;

            Cena.Opt = 0;
            Cena.Rozn = 0;
            
            
            if (Reestr<50)
            {
                
                Cena.Rozn = (Reestr + Reestr * 0.4 + Reestr * 0.3)*1.1;
                Cena.Opt  = (Reestr + Reestr * 0.3) *1.1;
            }
            
            if (Reestr>=50 && Reestr <500)
            {
                Cena.Rozn = (Reestr  + Reestr*0.25+ Reestr*0.20)*1.1;
                Cena.Opt = (Reestr + Reestr * 0.20)*1.1;
            }

            if (Reestr >= 500)
            {
                Cena.Rozn = (Reestr  + Reestr*0.15+ Reestr*0.12)*1.1;
                Cena.Opt = (Reestr +  Reestr * 0.12)*1.1 ;
            }
            return Cena;
        }
        static void Main(string[] args)
        {

            if (args.Length == 0)
            {
                return;
            }

            String fName = args[0].ToString();


            int firstRow = 11;
            int lastRow = 25617;
            //int Errors = 0;
            int zavoderror = 0;
            //int lastRow = 16380;

            MyApp = new Excel.Application();
            MyApp.Visible = true;
            MyBook = MyApp.Workbooks.Open(fName);
            //MySheet = (Excel.Worksheet)MyBook.Sheets[3]; // Explicit cast is not required here
            
            
            ErSheet = (Excel.Worksheet)MyBook.Sheets["Проверка"];
            MySheet = (Excel.Worksheet)MyBook.Sheets["П2"]; // Explicit cast is not required here
                                                            //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; 


            /// страничка сосписком ошибок
           for (int index = 5; index <= 2500; index++)
            {
           
                String CellV = ErSheet.get_Range("C" + index.ToString(), "C" + index.ToString()).Cells.Value;
                String CellError = ErSheet.get_Range("D" + index.ToString(), "D" + index.ToString()).Cells.Value;
                
                if (CellV == null)
                {
                    break;
                }


                if (!CellV.StartsWith("П2"))
                {

                    continue;
                }

                CellV = CellV.Substring(4, CellV.Length-4);
                Console.WriteLine(CellV);
                //int nmb;
                Excel.Range rng4;

                if (Properties.Settings.Default.UpdatePrice == false)
                { 
                rng4 = MySheet.get_Range("K" + CellV, "K" + CellV);
                rng4.Value2 = null;
                
                rng4 = MySheet.get_Range("L" + CellV, "L" + CellV);
                rng4.Value2 = null;

                rng4 = MySheet.get_Range("O" + CellV, "O" + CellV);
                rng4.Value2 = null;
                
                rng4 = MySheet.get_Range("P" + CellV, "P" + CellV);
                rng4.Value2 = null;

                rng4 = MySheet.get_Range("R" + CellV, "R" + CellV);
                rng4.Value2 = null;

                rng4 = MySheet.get_Range("X" + CellV, "X" + CellV);
                rng4.Value2 = null;
                }
                else /// тут мы исправляем цены
                {
                    if (CellError.Contains ("Полученная средневзвешенная цена приобретения превышает предельную оптовую цену по ненаркотическим ЛП") == true)
                    {
                        Console.WriteLine(CellError + "  " + CellV.ToString());
                        rng4 = MySheet.get_Range("O" + CellV, "O" + CellV);
                        rng4.Value2 = rng4.Value2 * 0.95;

                       
                    }

                    if (CellError.Contains("Полученная средневзвешенная цена реализации превышает предельную розничную цену по ненаркотическим ЛП") == true)
                    {
                        Console.WriteLine(CellError + "  " + CellV.ToString());
                        rng4 = MySheet.get_Range("P" + CellV, "P" + CellV);
                        rng4.Value2 = rng4.Value2*0.95;
                    }

                    //                    

                }
            }

            for (int index = firstRow; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("K" + index.ToString(), "X" + index.ToString()).Cells.Value;
                if (MyValues.GetValue(1, 1) == null)
                {
                    continue;
                }


                if (MyValues.GetValue(1, 10) == null)
                {

                    Console.WriteLine("Дошли до конца таблицы");
                    Console.WriteLine(index.ToString());

                    break;
                }

                /// непонятная ситуация - реестровые цены без ндс. Цены звода типа как с НДС.
                /// если выгрузили цены завода с НДС, то исправляем, чтобы не было вопросов
                if (Properties.Settings.Default.ProducerPriceFix == true)
                {
                    Excel.Range rngProducer = MySheet.get_Range("R" + index.ToString(), "R" + index.ToString());
                    Excel.Range rngReestr = MySheet.get_Range("Q" + index.ToString(), "Q" + index.ToString());
                    Excel.Range rngKolvo = MySheet.get_Range("K" + index.ToString(), "K" + index.ToString());


                    if (rngProducer.Value2 > rngReestr.Value2)
                    {
                        zavoderror++;
                        //rngProducer.Value2 = rngProducer.Value2 / 1.1;
                        //rngKolvo.Value2 = rngKolvo.Value2 * 1.1;

                        //// удаляем все строки где кривая цена завода
                        /* Excel.Range rng4;
                         rng4 = MySheet.get_Range("K" + index.ToString(), "K" + index.ToString());
                         rng4.Value2 = null;

                         rng4 = MySheet.get_Range("L" + index.ToString(), "L" + index.ToString());
                         rng4.Value2 = null;

                         rng4 = MySheet.get_Range("O" + index.ToString(), "O" + index.ToString());
                         rng4.Value2 = null;

                         rng4 = MySheet.get_Range("P" + index.ToString(), "P" + index.ToString());
                         rng4.Value2 = null;

                         rng4 = MySheet.get_Range("R" + index.ToString(), "R" + index.ToString());
                         rng4.Value2 = null;

                         rng4 = MySheet.get_Range("X" + index.ToString(), "X" + index.ToString());
                         rng4.Value2 = null;
                         */



                        Console.WriteLine("Иправляем цену завода в строке " + index.ToString());
                    }

                    rngProducer = MySheet.get_Range("U" + index.ToString(), "U" + index.ToString());
                    rngReestr = MySheet.get_Range("T" + index.ToString(), "T" + index.ToString());

                    if (rngProducer.Value2 > rngReestr.Value2)
                    {
                        rngKolvo.Value2 = rngKolvo.Value2 * 1.1;
                        Console.WriteLine("Иправляем цену завода в строке " + index.ToString());
                        zavoderror++;
                    }

                }
            }


                /// в настройках можно поставть запрет обновления данных на следующий год
                if (Properties.Settings.Default.Update2015 == false )
            {
                
                Console.WriteLine("Исправили цену завода "  + zavoderror.ToString());
                Console.WriteLine("Закончили убирать ошибки. Тыц Тыц любая кнопочка");
                Console.ReadKey();
                return;
            }

            //read  

            //BindingList<Employee> EmpList = new BindingList<Employee>();




            Random x = new Random();
            
            float Reestr, CenaRozn, CenaOpt,  Kolvo2014, Summa2015, Summa2014, Pribyl2014, Pribyl2015, CenaFaktProizv2015; //Kolvo2015
            //int errors = 0;
            for (int index = firstRow; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("K" + index.ToString(), "X" + index.ToString()).Cells.Value;
                if (MyValues.GetValue(1, 1) == null)
                {
                    continue;
                }


                if (MyValues.GetValue(1, 10) == null)
                {
                    
                    Console.WriteLine("Дошли до конца таблицы");
                    Console.WriteLine(index.ToString());

                    break;
                }


                //Reestr = float.Parse(MyValues.GetValue(1, 9).ToString());
                Reestr = float.Parse(MyValues.GetValue(1, 10).ToString());
                CenaRozn = float.Parse(MyValues.GetValue(1, 4).ToString());
                CenaOpt = float.Parse(MyValues.GetValue(1, 3).ToString());
                CenaFaktProizv2015 = float.Parse(MyValues.GetValue(1, 11).ToString());

                //Kolvo2015 = 0;// float.Parse(MyValues.GetValue(1, 2).ToString());
                Kolvo2014 = float.Parse(MyValues.GetValue(1, 1).ToString());

                Summa2015 = float.Parse(MyValues.GetValue(1, 6).ToString());
                Summa2014 = float.Parse(MyValues.GetValue(1, 5).ToString());

                Pribyl2014 = float.Parse(MyValues.GetValue(1, 13).ToString());
                Pribyl2015 = Pribyl2014 * Properties.Settings.Default.KoefRosta;

                //SCena mc;
              //  mc = GetMaxPrice(Reestr);

                // Заполняем данные по 2015 году
                /// чтобы не тыкать ничего в екселе
                Excel.Range rng4 = MySheet.get_Range("L" + index.ToString(), "L" + index.ToString());
                rng4.Value2 = Kolvo2014 * 1.04;//1.12;
                Excel.Range rng3 = MySheet.get_Range("X" + index.ToString(), "X" + index.ToString());
                rng3.Value2 = Pribyl2015;
                //////


               


                //// неправильно проставляет впервом случае - дальше костыль
              /*  rng4 = MySheet.get_Range("W" + index.ToString(), "W" + index.ToString());
                Pribyl2015 = float.Parse( rng4.get .Value2);


                rng4 = MySheet.get_Range("X" + index.ToString(), "X" + index.ToString());
                rng4.Value2 = Pribyl2015* Properties.Settings.Default.KoefRosta;

            */
                Console.WriteLine(index.ToString());

                //if ((CenaOpt < mc.Opt / 2) || (CenaRozn < mc.Rozn / 2) || (CenaOpt > mc.Opt) || (CenaRozn > mc.Rozn)) // || (Reestr < CenaFaktProizv2015) || (Reestr/1.85 < CenaFaktProizv2015) )  // (((Reestr*1.1)/CenaOpt > 2) || ((Reestr*1.1)/CenaRozn>2)||(CenaRozn>(Reestr*1.1)) )
            //{


                  //String er="";
                #region
                /*
                if ((Reestr * 1.1) / CenaOpt > 2)
                {
                    er = " ((Reestr*1.1)/CenaOpt > 2)";
                }

                if ((Reestr * 1.1) / CenaRozn > 2)
                {
                    er = " ((Reestr*1.1)/CenaRozn>2)";

                }


                if (CenaRozn > (Reestr * 1.1))
                {
                    er = " (CenaRozn>(Reestr*1.1))";

                }
                errors++;*/
                #endregion

                #region
                ///цена производителя больше цены реестра - вероятно что просто не учтен НДС
          /*          if (CenaFaktProizv2015 > Reestr)
                    {
                        Console.WriteLine(errors.ToString() + " " + index.ToString() + " " + CenaOpt.ToString() + " " + Reestr.ToString() + er);
                        Errors++;
                        Excel.Range rng1 = MySheet.get_Range("O" + index.ToString(), "O" + index.ToString());
                        Excel.Range rng2 = MySheet.get_Range("P" + index.ToString(), "P" + index.ToString());
                        Excel.Range FaktCenProizv = MySheet.get_Range("R" + index.ToString(), "R" + index.ToString());

                        /// меняем 
                        FaktCenProizv.Value2 = Reestr * Kolvo2014 * 0.95 * (1 - x.Next(3) / 100) * 1.1;
                }

                
            */        
                /// проверяем - есть ли разница между ценой реестра и фактической отпускной ценой
      /*          if (Reestr/ CenaFaktProizv2015 >1.3)
                {
                    Console.WriteLine(errors.ToString() + " " + index.ToString() + " " + CenaOpt.ToString() + " " + Reestr.ToString() + er);
                    Errors++;
                    Excel.Range rng1 = MySheet.get_Range("O" + index.ToString(), "O" + index.ToString());
                    Excel.Range rng2 = MySheet.get_Range("P" + index.ToString(), "P" + index.ToString());
                    Excel.Range FaktCenProizv = MySheet.get_Range("R" + index.ToString(), "R" + index.ToString());

                    /// меняем 
                    FaktCenProizv.Value2 = Reestr * Kolvo2014 * 0.95 * (1 - x.Next(3) / 100) * 1.1;

                }*/

                    //MyValues.SetValue((Reestr * Kolvo2014),1,5);
                    //MyValues.SetValue((Reestr * Kolvo2014*1.1), 1, 6);

//                    rng1.Value2 = Reestr * Kolvo2014 * (1-x.Next(3)/100)*1.1;
  //                  rng2.Value2 = Reestr * Kolvo2014 * 1.1 * (1 - x.Next(3) / 100)*1.1;
    //                FaktCenProizv.Value2 = Reestr * Kolvo2014 * 0.95 * (1 - x.Next(3) / 100)*1.1;
                                                  

                //}

                //Excel.Range CenaSrednProizv = MySheet.get_Range("U" + index.ToString(), "U" + index.ToString());


            }


            //write
            /*
             lastRow += 1;
MySheet.Cells[lastRow, 1] = emp.Name;
MySheet.Cells[lastRow, 2] = emp.Employee_ID;
MySheet.Cells[lastRow, 3] = emp.Email_ID;
MySheet.Cells[lastRow, 4] = emp.Number;
EmpList.Add(emp);
MyBook.Save(); 
             * 
             */

            //   MyBook.SaveAs ("12345.xlsx");            
            //  MyBook.Close();

            #endregion

            Console.WriteLine("Исправили цену завода " + zavoderror.ToString());
            Console.WriteLine("Готово ! Осталось нажать любую кнопку и исправить оставшиеся ошибки вручную.");
            Console.ReadKey();
        }
    }
}
