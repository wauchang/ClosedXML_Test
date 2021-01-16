using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXMLTest
{
    //データを収めるもの
    public class Item
    {
        //通し番号・名前・値１・値２
        public int Number { get; private set; }
        public string Name { get; private set; }
        public int Value1 { get; set; }
        public int Value2 { get; set; }

        public Item(int number, string name, int value1, int value2)
        {
            Number = number;
            Name = name;
            Value1 = value1;
            Value2 = value2;
        }

    }

    public class main
    {
        List<Item> V1ComparedItems { get; set; } = new List<Item>();
        List<Item> V2ComparedItems { get; set; } = new List<Item>();

        public static void Main()
        {
            main obj = new main();

            //XLファイルのパスを入力
            Console.Write("Enter the Pass : ");
            string pass = @"F:\wau\Documents\todoList.xlsx";
            //string pass = Console.ReadLine();

            //ワークブック・シートを開く
            var workbook = new XLWorkbook(pass);
            var sheet = workbook.Worksheet("Main");

            //各要素の列挙
            List<Item> Items = new List<Item>();
            for (int index = 2; !string.IsNullOrEmpty(sheet.Cell(index, 3).Value.ToString()); index++)
            {
                var cell1 = sheet.Cell(index, 1);
                var cell2 = sheet.Cell(index, 3);
                var cell3 = sheet.Cell(index, 4);
                var cell4 = sheet.Cell(index, 5);

                int v3, v4;
                if (string.IsNullOrEmpty(cell3.Value.ToString())) v3 = 0; else v3 = (int)cell3.GetDouble();
                if (string.IsNullOrEmpty(cell4.Value.ToString())) v4 = 0; else v4 = (int)cell4.GetDouble();

                Item item = new Item((int)cell1.GetDouble(), cell2.GetString(), v3, v4);
                Items.Add(item);
            }
            //要素表示
            foreach (Item item in Items)
            {
                Console.WriteLine(item.Name);
            }

            Console.WriteLine("-------------------------");
            //比較するよ
            for (int index = 0; index < Items.Count; index++)
            {
                obj.V1ComparedItems.Add(obj.Comparing(Items, index, 1));
                obj.V1ComparedItems.Sort((a, b) => (a.Value1 - b.Value1) * (-1));
                obj.V2ComparedItems.Add(obj.Comparing(Items, index, 2));
                obj.V2ComparedItems.Sort((a, b) => (a.Value2 - b.Value2) * (-1));
                foreach (Item item in obj.V1ComparedItems)
                {
                    Console.WriteLine(item.Name);
                }
            }

            List<Item> ComparedItems = obj.V1ComparedItems;
            ComparedItems.Sort((a, b) => (a.Number - b.Number));

            //新規ブック・シート作成
            bool nameMatching = false;
            string matchedName = "";
            IXLWorksheet newSheet;
            foreach (IXLWorksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name == "NewList")
                {
                    nameMatching = true;
                    matchedName = worksheet.Name;
                }
            }

            if(nameMatching == false)
            {
                sheet.CopyTo("NewList");
                newSheet = workbook.Worksheet("NewList");
            }
            else
            {
                sheet.CopyTo(matchedName + "1");
                newSheet = workbook.Worksheet("matchedName" + "1");
            }

            
            newSheet.Cell(1, 1).Value = "No";
            newSheet.Cell(1, 3).Value = "Name";
            newSheet.Cell(1, 4).Value = "Necessity";
            newSheet.Cell(1, 5).Value = "Interest";

            for(int i = 0; i < ComparedItems.Count; i++)
            {
                newSheet.Cell(i + 2, 1).Value = ComparedItems[i].Number;
                newSheet.Cell(i + 2, 3).Value = ComparedItems[i].Name;
                newSheet.Cell(i + 2, 4).Value = ComparedItems[i].Value1;
                newSheet.Cell(i + 2, 5).Value = ComparedItems[i].Value2;
            }
            newSheet.Column(2).Delete();

            //newBook.SaveAs(@"F:\wau\Documents\NewBook.xlsx");
            workbook.Save();


        }

        //ValueNumberに寄らない一般化をしたい
        Item Comparing(List<Item> Items, int index, int valueNumber)
        {
            bool endOfCompare = false;                  //比較完了を表すbool
            List<Item> ComparedItems;                  //比較済みの要素
            string question, result;                                      //v1,v2で聞くことの切り替え

            if (valueNumber == 1)
            {
                ComparedItems = V1ComparedItems;
                question = " more necessary than ";
                result = "Necessity of ";
            }
            else
            {
                ComparedItems = V2ComparedItems;
                question = " more interesting than ";
                result = "Interest of ";
            }




            //比較済みの各要素と比較
            foreach (Item item in ComparedItems)
            {
                //比較済みのリストは降順に並んでいるので、大きいものから順に比較を進める
                //比較先より小さければ続行、大きいあるいは等しい場合は処理を終了する

                if (endOfCompare == false)
                {
                    char choice;

                    do
                    {
                        Console.Write("Is " + Items[index].Name + question + item.Name + " ? : Equal:0 Yes:1 No:2  ==> ");
                        do
                        {
                            choice = (char)Console.Read();
                        } while (choice == '\n' || choice == '\r');
                    } while (choice < '0' | choice > '2');

                    switch (choice)
                    {
                        case '0':　//比較先の要素と価値がイコールである場合、boolをtrueにして比較を終了する
                            Console.WriteLine(Items[index].Name + " == " + item.Name);
                            if (valueNumber == 1)
                            {
                                Items[index].Value1 = item.Value1;
                            }
                            else
                            {
                                Items[index].Value2 = item.Value2;
                            }
                            endOfCompare = true;
                            break;
                        case '1':　//比較先の要素より値が大きい場合、自信の値を比較先＋１して比較を終了する
                            Console.WriteLine(Items[index].Name + " > " + item.Name);
                            if (valueNumber == 1)
                            {
                                Items[index].Value1 = item.Value1 + 1;
                            }
                            else
                            {
                                Items[index].Value2 = item.Value2 + 1;
                            }
                            endOfCompare = true;
                            break;
                        case '2':　//比較先の要素より値が小さい場合、比較済みの各要素の値を＋１して進める
                            Console.WriteLine(Items[index].Name + " < " + item.Name);
                            if (valueNumber == 1)
                            {
                                item.Value1++;
                            }
                            else
                            {
                                item.Value2++;
                            }
                            break;
                    }
                }
            }

            //比較が終わったら値を表示して、返す
            if (valueNumber == 1)
            {
                foreach (Item ComparedItem in ComparedItems)
                {
                    Console.WriteLine(result + ComparedItem.Name + " : " + ComparedItem.Value1);
                }
                Console.WriteLine(result + Items[index].Name + " : " + Items[index].Value1);
            }
            else
            {
                foreach (Item ComparedItem in ComparedItems)
                {
                    Console.WriteLine(result + ComparedItem.Name + " : " + ComparedItem.Value2);
                }
                Console.WriteLine(result + Items[index].Name + " : " + Items[index].Value2);
            }

            Console.WriteLine();
            return Items[index];
        }
    }
}
