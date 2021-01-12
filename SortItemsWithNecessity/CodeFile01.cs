using ClosedXML.Excel;
using System;
using System.Collections.Generic;

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
            string pass = Console.ReadLine();

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
                Console.WriteLine("Compare " + Items[index].Name + " to others.");

                obj.V1ComparedItems.Add(obj.V1Comparing(Items, index));

            }
        }

        //実際の比較はここで行う
        Item V1Comparing(List<Item> Items, int index)
        {
            bool matchSomeValue = false;　                                                  //比較完了を表すbool
            List<Item> ComparedItems = V1ComparedItems;　　                //比較済みの要素
            string question = " more necessary than ";                                   //v1,v2で聞くことの切り替え


            //比較済みの各要素と比較
            foreach (Item item in ComparedItems)
            {
                //比較する要素の値が比較先以下であり、かつ、どの要素の値ともイコールじゃない間ループする
                //比較を進めて、全ての要素より大きくなるかどれかと価値が等しくなったらループを終了する
                if (Items[index].Value1 <= item.Value1 && matchSomeValue == false)
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
                            Items[index].Value1 = item.Value1;
                            matchSomeValue = true;
                            break;
                        case '1':　//比較先の要素より値が大きい場合、自信の値を比較先＋１して進める
                            Console.WriteLine(Items[index].Name + " > " + item.Name);
                            Items[index].Value1 = item.Value1 + 1;
                            Console.WriteLine("Debug : comparingValue is " + Items[index].Value1);
                            break;
                        case '2':　//比較先の要素より値が小さい場合、比較済みの各要素の値を全て＋１して進める
                            Console.WriteLine(Items[index].Name + " < " + item.Name);
                            foreach (Item ComparedItem in ComparedItems)
                            {
                                if (ComparedItem.Value1 >= Items[index].Value1) ComparedItem.Value1++;
                            }
                            break;
                    }
                }
            }

            //比較が終わったら値を表示して、返す
            foreach (Item ComparedItem in ComparedItems)
            {
                Console.WriteLine("Necessity of " + ComparedItem.Name + " : " + ComparedItem.Value1);
            }
            Console.WriteLine("Necessity of " + Items[index].Name + " : " + Items[index].Value1);
            Console.WriteLine();
            return Items[index];
        }

        Item Comparing(List<Item> Items, int index, int valueNumber)
        {
            bool matchSomeValue = false;　           //比較完了を表すbool
            List<Item> ComparedItems;　　            //比較済みの要素
            int comparingValue, comparedValue;     //比較する値の入れ物　値型じゃアカン気がする
            string question;                                       //v1,v2で聞くことの切り替え

            //V1,V2の切り替え
            if (valueNumber == 1)
            {
                ComparedItems = V1ComparedItems;
                comparingValue = Items[index].Value1;
                question = " more necessary than ";
            }
            else
            {
                ComparedItems = V2ComparedItems;
                comparingValue = Items[index].Value2;
                question = " more interesting than ";
            }

            //比較済みの各要素と比較
            foreach (Item item in ComparedItems)
            {
                //V1,V2の切り替え
                if (valueNumber == 1) comparedValue = item.Value1;
                else comparedValue = item.Value2;

                //比較する要素の値が比較先以下であり、かつ、どの要素の値ともイコールじゃない間ループする
                //比較を進めて、全ての要素より大きくなるかどれかと価値が等しくなったらループを終了する
                if (comparingValue <= comparedValue && matchSomeValue == false)
                {
                    char choice;

                    do
                    {
                        Console.Write("Is " + Items[index].Name + question + item.Name + "? : Equal:0 Yes:1 No:2  ==> ");
                        do
                        {
                            choice = (char)Console.Read();
                        } while (choice == '\n' || choice == '\r');
                    } while (choice < '0' | choice > '2');

                    switch (choice)
                    {
                        case '0':　//比較先の要素と価値がイコールである場合、boolをtrueにして比較を終了する
                            Console.WriteLine(Items[index].Name + " == " + item.Name);
                            comparingValue = comparedValue;
                            matchSomeValue = true;
                            break;
                        case '1':　//比較先の要素より値が大きい場合、自信の値を比較先＋１して進める
                            Console.WriteLine(Items[index].Name + " > " + item.Name);
                            comparingValue = comparedValue + 1;
                            Console.WriteLine("Debug : comparingValue is " + comparingValue);
                            break;
                        case '2':　//比較先の要素より値が小さい場合、比較済みの各要素の値を全て＋１して進める
                            Console.WriteLine(Items[index].Name + " < " + item.Name);
                            foreach (Item ComparedItem in ComparedItems)
                            {
                                if (valueNumber == 1 && ComparedItem.Value1 >= Items[index].Value1)
                                {
                                    ComparedItem.Value1++;
                                }
                                else if (valueNumber == 2 && ComparedItem.Value2 >= Items[index].Value2)
                                {
                                    ComparedItem.Value2++;
                                }
                            }
                            break;
                    }
                }
            }
            if (valueNumber == 1) Items[index].Value1 = comparingValue;
            else Items[index].Value2 = comparingValue;

            //比較が終わったら値を表示して、返す
            if (valueNumber == 1) Console.WriteLine("Necessity of " + Items[index].Name + " : " + Items[index].Value1);
            else Console.WriteLine("Interests of " + Items[index].Name + " : " + Items[index].Value1);
            Console.WriteLine();
            return Items[index];
        }
    }
}
