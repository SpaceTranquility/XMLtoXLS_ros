using System;
using System.Globalization;
using System.Collections.Generic;

using System.IO;
using System.IO.Compression;
using System.Linq;

using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;


namespace XmlToXls
{

    class Program
    {
        //запись о собственнике
        class Owner 
        {
            public string fio { get; set; }
            public string part { get; set; }
            public string partOf { get; set; }
            public string document { get; set; }
            public Owner(string FIO, string Part, string PartOf, string Document)
            {
                fio = FIO;
                part = Part;
                partOf = PartOf;
                document = Document;
            }
        }
        //Запись о квартире
        class Item_flat : IComparable, ICloneable 
        {
            public int number { get; set; }
            public string address { get; set; }
            public int numberOfFlat { get; set; }
            public string numFlat { get; set; }
            public double area { get; set; }
            public List<Owner> owners;

            public Item_flat(string Address, int NumFlat, double Area, List<Owner> Owners)
            {
                address = Address;
                numberOfFlat = NumFlat;
                area = Area;
                owners = Owners;
            }

            public Item_flat ()
            {
                number = 0;
                address = "";
                numberOfFlat = 0;
                area = 0;
                owners = new List<Owner>();
            }

            //---Определяем стандартные интерфейсы---
            public static bool operator <(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat < flat2.numberOfFlat);
            }
            public static bool operator >(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat > flat2.numberOfFlat);
            }
            public static bool operator <=(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat <= flat2.numberOfFlat);
            }
            public static bool operator >=(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat >= flat2.numberOfFlat);
            }
            public static bool operator ==(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat == flat2.numberOfFlat);
            }
            public static bool operator !=(Item_flat flat1, Item_flat flat2)
            {
                return (flat1.numberOfFlat != flat2.numberOfFlat);
            }

            public int CompareTo(object Flat)
            {
                const string eror = "Сравниваемый объект не пренадлежит классу Item_flat.";
                Item_flat flat = Flat as Item_flat;
                if (numberOfFlat < flat.numberOfFlat) return -1;
                if (numberOfFlat == flat.numberOfFlat) return 0;
                if (numberOfFlat > flat.numberOfFlat) return 1;

                throw new ArgumentException(eror);
            }

            public object Clone()
            {
                Item_flat flat = new Item_flat();
                flat.number = this.number;
                flat.address = this.address;
                flat.numberOfFlat = this.numberOfFlat;
                flat.area = this.area;
                flat.owners = this.owners;                
                return flat;
            }

            public override bool Equals(object obj)
            {
                var flat = obj as Item_flat;
                return flat != null &&
                       address == flat.address &&
                       numberOfFlat == flat.numberOfFlat &&
                       area == flat.area &&
                       EqualityComparer<List<Owner>>.Default.Equals(owners, flat.owners);
            }

            public override int GetHashCode()
            {
                var hashCode = 1608761099;
                hashCode = hashCode * -1521134295 + number.GetHashCode();
                hashCode = hashCode * -1521134295 + numberOfFlat.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(address);
                hashCode = hashCode * -1521134295 + area.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<List<Owner>>.Default.GetHashCode(owners);
                return hashCode;
            }
            //---Закончили стандартные интерфейсы---
        //КОнец класса Item_flat
        }
        //Запись об адресе
        class Buildding
        {
            public string address { get; set; }
            public List<Item_flat> flats { get; set; }

            public Buildding()
            {
                flats = new List<Item_flat>();
            }
            public Buildding(string Address)
            {
                address = Address;
                flats = new List<Item_flat>();
            }
            public Buildding(string Address, List<Item_flat> Flats)
            {
                address = Address;
                flats = Flats;
            }
        }

        //Создание шаблона Ecxel
        static void NewFormXls(Application application, string adress)
        {
            application.SheetsInNewWorkbook = 1;
            //XlsForm.Visible = true;

            application.Workbooks.Add();
            Worksheet sheet = (Worksheet)application.ActiveSheet;

            //Размеры            
            sheet.Columns[1].ColumnWidth = 4;
            sheet.Columns[2].ColumnWidth = 4;
            sheet.Columns[3].ColumnWidth = 30;
            sheet.Columns[4].ColumnWidth = 70;
            sheet.Columns[5].ColumnWidth = 10;
            sheet.Columns[6].ColumnWidth = 10;
            sheet.Columns[7].ColumnWidth = 10;

            sheet.Rows[2].RowHeight = 65;
            sheet.Rows[3].RowHeight = 20;
            sheet.Rows[4].RowHeight = 60;
            sheet.Rows[5].RowHeight = 40;

            sheet.Cells.Font.Size = 8;
            sheet.Cells.Font.Name = "Times New Roman";
            sheet.Cells.HorizontalAlignment = 3;
            sheet.Cells.VerticalAlignment = 2;

            //Первая строка
            sheet.Cells[2, 1].Value = "Приложение №__ к Протоколу №________\n" +
                                "общего собрания собственников помещений\n" +
                                "многоквартирного дома, расположенного, по адресу:\n" +
                                "__________________________________________, литера А от ______________";
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 7]].Merge();
            sheet.Cells[2, 1].HorizontalAlignment = 4;

            //Вторая строка
            sheet.Cells[3, 1].Value = "Реестр собственников помещений многоквартирного дома по адресу:  "+ adress;
            sheet.Cells[3, 1].Font.Bold = true;
            sheet.Cells[3, 1].Font.Size = 10;
            sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, 7]].Merge();

            //Третья строка
            sheet.Cells[4, 1].Value = "№п/п";
            sheet.Cells[4, 2].Value = "№ пом.";
            sheet.Cells[4, 3].Value = "Ф.И.О.";
            sheet.Cells[4, 4].Value = "Документ, подтверждающий право собственности";
            sheet.Cells[4, 5].Value = "Сведения о размере доли в праве общей собственности на общее имущество собственников помещений в МКД	";
            sheet.Cells[4, 5].WrapText = true;
            sheet.Cells[4, 7].Value = "Площадь";
            sheet.Range[sheet.Cells[4, 1], sheet.Cells[5, 1]].Merge();
            sheet.Range[sheet.Cells[4, 2], sheet.Cells[5, 2]].Merge();
            sheet.Range[sheet.Cells[4, 3], sheet.Cells[5, 3]].Merge();
            sheet.Range[sheet.Cells[4, 4], sheet.Cells[5, 4]].Merge();
            sheet.Range[sheet.Cells[4, 5], sheet.Cells[4, 6]].Merge();
            sheet.Range[sheet.Cells[4, 7], sheet.Cells[5, 7]].Merge();

            //Четвёртая строка
            sheet.Cells[5, 5].Value = "размер доли ";
            sheet.Cells[5, 6].Value = "кв.м от общей площади помещения";
            sheet.Cells[5, 6].WrapText = true;

            sheet.Range[sheet.Cells[6, 1], sheet.Cells[10000, 4]].HorizontalAlignment = 2;
            sheet.Cells[2, 1].HorizontalAlignment = 4;

            //Конец NewFormXls
        }

        //Заполнение документа и сохранение в directory
        static void MakeXls(List<Buildding> all_buildings, string directory)
        {
            Console.WriteLine("\nНачинаю работу с MS Excel. Придётся подождать. Но, если больше 5 минут, то я завис.\n");
            Application xlsApp = new Application();
            Worksheet sheet;
            Workbook book;

            foreach (Buildding buildding in all_buildings)
            {
                Console.WriteLine($"Пишем адрес: {buildding.address}.");
                buildding.flats.Sort();

                //Создаём шаблон
                NewFormXls(xlsApp, buildding.address);

                book = xlsApp.ActiveWorkbook;
                sheet = (Worksheet)xlsApp.ActiveSheet;

                //Ставим значения
                int roll = 6, n = 1;
                foreach (Item_flat flat in buildding.flats)
                {
                    Console.WriteLine("Пишем: " + flat.address + " кв. " + flat.numFlat);
                    if (flat.owners.Count == 0)
                        Console.WriteLine("     Не нашёл информации о владельцах.");

                    foreach (var owner in flat.owners)
                    {
                        sheet.Cells[roll, 1].Value = n;
                        sheet.Cells[roll, 2].Value = flat.numFlat;
                        sheet.Cells[roll, 3].Value = owner.fio;
                        sheet.Cells[roll, 4].Value = owner.document;
                        sheet.Cells[roll, 5].Value = "'" + owner.part;
                        sheet.Cells[roll, 6].Value = owner.partOf;
                        sheet.Cells[roll, 7].Value = flat.area;
                        n++;
                        roll++;
                    }
                }
                //Сохраняем и закрываем книгу
                try
                {
                    if (File.Exists(directory + "\\" + buildding.address + ".xls"))
                        File.Delete(directory + "\\" + buildding.address + ".xlsx");
                    if (File.Exists(directory + "\\" + buildding.address + ".xlsx"))
                        File.Delete(directory + "\\" + buildding.address + ".xlsx");
                }
                catch
                {
                    Console.WriteLine("Не удалось удалить старый файл, возможно он открыт в каком-то приложении.");
                }
                //book.SaveAs(directory + buildding.address
                try { book.SaveAs(directory + buildding.address); }
                catch { Console.WriteLine("Ошибка сохранения."); Console.ReadKey(); }
                book.Close(true);
            }
            xlsApp.Quit();
            Console.WriteLine("Финиш.");
        }

        //Обработка XML, создание записи по квартире
        static Item_flat XmlProcessing(XDocument xDoc)
        {
            //Разделитель точка (для перевода строк в Double)
            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = ".";

            //Данные по помещению
            string number_flat_string = "";
            string address = ""; //Адрес
            int number_flat = 0; //Номер помещения
            double area = 0; //Площадь
            List<Owner> Owners = new List<Owner>(); //Список владельцев

            //Получаем корень
            XElement root = xDoc.Root;

            //Все узлы в документе. Перебираем
            IEnumerable<XElement> all_nodes = root.Descendants();
            foreach (XElement node in all_nodes)
            {
                switch (node.Name.LocalName)
                {
                    // ---Адрес---
                    case ("Address"):
                        {
                            foreach (XElement node_adr in node.Nodes())
                            {
                                switch (node_adr.Name.LocalName)
                                {
                                    case ("Street"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Name").Value + ", ";
                                        break;
                                    case ("Level1"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value + ", ";
                                        break;
                                    case ("Level2"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value + ", ";
                                        break;
                                    case ("Level3"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value;
                                        break;
                                    case ("Apartment"): //номер квартиры
                                        number_flat_string = node_adr.Attribute("Value").Value;
                                        try {number_flat = Convert.ToInt32(node_adr.Attribute("Value").Value); }
                                        catch { number_flat = Int32.MaxValue; Console.Write(" Плохой номер помещения. "); }
                                        
                                        break;

                                    default:
                                        break;
                                }
                            }
                            break;
                        } //Конец блока адреса

                    //---Площадь---
                    case ("Area"):
                        {
                            if (node.Parent.Name.LocalName == "Flat" && node.Parent.Parent.Name.LocalName != "Flats"
                                || node.Parent.Name.LocalName == "Building" && node.Parent.Parent.Name.LocalName != "Buildings")
                            {
                                string s = node.Value;
                                area = Convert.ToDouble(s, provider);
                            }
                            break;
                        }

                    //---Собственники---
                    case ("Rights"):
                        {
                            //Определение собственности
                            foreach (XElement right in node.Elements())
                            {
                                //Документ на собственность
                                string document = "";
                                string part = ""; //Доля
                                double partOf = 0; //Размер доли
                                try
                                {
                                    foreach (XElement registration in right.Elements())
                                        if (registration.Name.LocalName == "Registration")
                                        {
                                            foreach (XElement reg in registration.Elements())
                                                if (reg.Name.LocalName == "RegNumber")
                                                    document += "№ " + reg.Value;
                                            foreach (XElement reg in registration.Elements())
                                                if (reg.Name.LocalName == "RegDate")
                                                    document += ", " + reg.Value;
                                        }
                                    foreach (XElement point in right.Elements())
                                        if (point.Name.LocalName == "Share")
                                        {
                                            int num = Convert.ToInt32(point.Attribute("Numerator").Value);
                                            int den = Convert.ToInt32(point.Attribute("Denominator").Value);
                                            part += num + " / " + den;
                                            partOf = area * (double)num / (double)den;
                                        }
                                    foreach (XElement point in right.Elements())
                                        if (point.Name.LocalName == "Name")
                                        {
                                            document += ", " + point.Value;
                                            if (part == "")
                                            {
                                                if (point.Value == "Собственность") {part = "1"; partOf = area;}
                                                if (point.Value == "Общая совместная собственность") part = "*"; 
                                            }
                                        }
                                }
                                catch { Console.WriteLine("Ошибка получения документа на собственность"); }

                                //Собственники                               
                                string FIO = ""; //Ф.И.О
                                try
                                {
                                    foreach (XElement owners in right.Elements())
                                        if (owners.Name.LocalName == "Owners")
                                        {
                                            foreach (XElement owner in owners.Elements())
                                                if (owner.Name.LocalName == "Owner")
                                                    foreach (XElement person in owner.Elements())
                                                        if (person.Name.LocalName == "Person")
                                                        {
                                                            foreach (XElement point in person.Elements())
                                                                if (point.Name.LocalName == "FamilyName")
                                                                    FIO += point.Value + " ";
                                                            foreach (XElement point in person.Elements())
                                                                if (point.Name.LocalName == "FirstName")
                                                                    FIO += point.Value + " ";
                                                            foreach (XElement point in person.Elements())
                                                                if (point.Name.LocalName == "Patronymic")
                                                                    FIO += point.Value + " ";

                                                            // Записываем собственника
                                                            Owners.Add(new Owner(FIO, part, Convert.ToString(partOf), document));
                                                            FIO = "";
                                                        }
                                        }
                                }
                                catch { Console.WriteLine("Ошибка получения имени собственника."); }
                            }
                            break;
                        } //Конец поиска блока Rights (собственники)

                    default:
                        break;
                }

            }//Окончание перебора узлов документа
            foreach (Owner owner in Owners)
            {
                if (owner.part == "*")
                {
                    owner.part = "1/" + Owners.Count() + "*";
                    owner.partOf = Convert.ToString(area * (double)1 / (double)Owners.Count()) + "*";
                }
            }
            Item_flat flat_temp = new Item_flat(address, number_flat, area, Owners);
            flat_temp.numFlat = number_flat_string;
            return (flat_temp);
        //Конец XmlProcessing
        }

        //Удаление файлов в директории
        static void clrDir(string adr)
        {
            //*** Рассмотреть внедрение DirectoryInfo
            DirectoryInfo dir = new DirectoryInfo(adr);
            foreach (var file in dir.GetFiles())
                file.Delete();
        }

        static void Main(string[] args)
        {
            //Приветствие
            Console.WriteLine("Start.");
            //Задаём Стартовую директорию
            string home = Directory.GetCurrentDirectory() + "\\";

            // * Выход в тестовую дерикторию 
            //home += "\\test\\";
            // -1-- Временная директория
            string temp = home + "__tmp__";
            Directory.CreateDirectory(temp);

            //Все квартиры из всех файлов
            List<Item_flat> All_flats = new List<Item_flat>();

            // -2-- Работа с XML
            //Создаём коллекцию нужных файлов в дериктории
            IEnumerable<string> All_Zip = Directory.EnumerateFiles(home, "*.zip", SearchOption.AllDirectories);
            IEnumerable<string> All_XML = Directory.EnumerateFiles(home, "*.xml", SearchOption.AllDirectories);
            //Выводим общие число файлов
            Console.WriteLine($"Found {All_Zip.Count()} archives and {All_XML.Count()} XMLs Please wait.");
                       
            //Проход по Xml файлам (если есть)
            if(All_XML.Count() > 0)
            {
                //XDocument xDoc = new XDocument();
                foreach (string xml_filename in All_XML)
                {
                    try { All_flats.Add(XmlProcessing(XDocument.Load(xml_filename))); }
                    catch { continue; };
                    Console.WriteLine("Прочитал: " + All_flats.Last().address + " кв. " + All_flats.Last().numFlat);
                }
            }

            //Проход по архивам (если есть)
            //Распаковка архивов в папку
            foreach (string arhive_1lvl in All_Zip)
            {
                try
                {
                    clrDir(temp); //чистим временную папку
                    ZipFile.ExtractToDirectory(arhive_1lvl, temp); //распаковываем

                    //Распаковка вложениых архивов
                    foreach (string arhive_1lv2 in Directory.EnumerateFiles(temp, "*.zip", SearchOption.AllDirectories))
                    {
                        //РАзархивация
                        ZipFile.ExtractToDirectory(arhive_1lv2, temp);
                        //Получили XML
                        foreach (string xml_filename in Directory.EnumerateFiles(temp, "*.xml", SearchOption.AllDirectories))
                        {
                            All_flats.Add(XmlProcessing(XDocument.Load(xml_filename)));
                            Console.WriteLine("Прочитал: " + All_flats.Last().address + " кв. " + All_flats.Last().numFlat);
                            //Выгрузка
                            //Directory.CreateDirectory(home + "\\Выгрузка\\");
                            //File.Copy(xml_filename, home + "\\Выгрузка\\" + All_flats.Last().address + " кв " + All_flats.Last().numFlat + ".xml");
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Ошибка при работе с архивами.");
                }
            }
            // --2- Конец рботы с XML

            clrDir(temp);
            Directory.Delete(temp);
            // --1- *Возможны исключения* удаление временной дериктории

            //Распределяем квартиры по адресам
            List <Buildding> All_buildings = new List<Buildding>();
            string active_address = "Ну это точно не попадётся, палка-копалка";
            int active_building = 0;

            //перебор квартир
            foreach (Item_flat flat in All_flats)
            {
                if(active_address == flat.address)
                {
                    All_buildings[active_building].flats.Add(flat);
                }
                else
                {
                    // Проверяем наличие дома
                    bool need_new = true;
                    for (int n=0; n< All_buildings.Count(); n++)
                    {
                        if (All_buildings[n].address == flat.address)
                        {
                            active_address = flat.address;
                            active_building = n;
                            n = All_buildings.Count();
                            need_new = false;
                        }
                    }
                    // Создаём новый дом
                    if(need_new)
                    {
                        All_buildings.Add(new Buildding(flat.address));
                        active_building = All_buildings.Count() - 1;
                        active_address = flat.address;
                    }
                    All_buildings[active_building].flats.Add(flat);
                }
            }
            // Получили список строений

            MakeXls(All_buildings, home);

            Console.WriteLine("Чтобы закрыть можно нажать Enter.");
            Console.Read();
        }
    }
}
