using System;
using System.Windows.Forms;
using System.Globalization;
using System.Collections.Generic;

using System.IO;
using System.IO.Compression;
using System.Linq;

using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace XmlToXls_3
{
    static class Source
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>  

        [STAThread]
        static void Main()
        {
            //home = Directory.GetCurrentDirectory() + "\\";

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }


        //запись о собственнике
        public class Owner
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
            public override bool Equals(object obj)
            {
                Owner owner = obj as Owner;
                return owner != null &&
                    fio == owner.fio &&
                    part == owner.part &&
                    partOf == owner.partOf &&
                    document == owner.document;
            }

        }
        //Запись о квартире
        public class Item_flat : IComparable, ICloneable
        {
            public int number { get; set; }
            public string address { get; set; }
            public string addressWithDot { get; set; }
            public int numberOfFlat { get; set; }
            public string numFlat { get; set; }
            public double area { get; set; }
            public List<Owner> owners;

            public Item_flat(string Address, int NumFlat, double Area, List<Owner> Owners)
            {
                addressWithDot = "";
                address = Address;
                numberOfFlat = NumFlat;
                area = Area;
                owners = Owners;
            }

            public Item_flat()
            {
                addressWithDot = "";
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
                flat.addressWithDot = this.addressWithDot;
                flat.number = this.number;
                flat.address = this.address;
                flat.numberOfFlat = this.numberOfFlat;
                flat.area = this.area;
                flat.owners = this.owners;
                return flat;
            }

            public override bool Equals(object obj)
            {
                Item_flat flat = obj as Item_flat;
                bool EqOwners = true;
                if (owners.Count == flat.owners.Count)
                    for (int i = 0; i < owners.Count; i++)
                    {
                        if (!owners[i].Equals(flat.owners[i]))
                            EqOwners = false;
                    }
                else EqOwners = false;

                return address == flat.address &&
                    numberOfFlat == flat.numberOfFlat &&
                    area == flat.area &&
                    EqOwners;
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
        public class Buildding
        {
            public string address { get; set; }
            public string addressWithDot { get; set; }
            public List<Item_flat> flats { get; set; }

            public Buildding()
            {
                flats = new List<Item_flat>();
            }
            public Buildding(string Address, string AddressWithDot = "")
            {
                addressWithDot = AddressWithDot;
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
        public static void NewFormXls(Excel.Application application, string adress)
        {
            application.SheetsInNewWorkbook = 1;
            //XlsForm.Visible = true;

            int firstRow = 0;

            application.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)application.ActiveSheet;

            //Размеры            
            sheet.Columns[1].ColumnWidth = 4;
            sheet.Columns[2].ColumnWidth = 9;
            sheet.Columns[3].ColumnWidth = 35;
            sheet.Columns[4].ColumnWidth = 70;
            sheet.Columns[5].ColumnWidth = 13;
            sheet.Columns[6].ColumnWidth = 13;
            sheet.Columns[7].ColumnWidth = 13;

            sheet.Rows[firstRow + 1].RowHeight = 65;
            sheet.Rows[firstRow + 2].RowHeight = 20;
            sheet.Rows[firstRow + 3].RowHeight = 60;
            sheet.Rows[firstRow + 4].RowHeight = 40;

            sheet.Cells.Font.Size = 10;
            sheet.Cells.Font.Name = "Times New Roman";
            sheet.Cells.HorizontalAlignment = 3;
            sheet.Cells.VerticalAlignment = 2;

            //Первая строка
            sheet.Cells[firstRow + 1, 1].Value = "Приложение №__ к Протоколу №________\n" +
                                "общего собрания собственников помещений\n" +
                                "многоквартирного дома, расположенного, по адресу:\n" +
                                "__________________________________________, от ______________";
            sheet.Range[sheet.Cells[firstRow + 1, 1], sheet.Cells[firstRow + 1, 7]].Merge();
            sheet.Cells[firstRow + 1, 1].HorizontalAlignment = 4;

            //Вторая строка
            sheet.Cells[firstRow + 2, 1].Value = "Реестр собственников помещений многоквартирного дома по адресу:  " + adress;
            sheet.Cells[firstRow + 2, 1].Font.Bold = true;
            sheet.Cells[firstRow + 2, 1].Font.Size = 12;
            sheet.Range[sheet.Cells[firstRow + 2, 1], sheet.Cells[firstRow + 2, 7]].Merge();

            //Третья строка
            sheet.Cells[firstRow + 3, 1].Value = "№п/п";
            sheet.Cells[firstRow + 3, 2].Value = "№ пом.";
            sheet.Cells[firstRow + 3, 3].Value = "Ф.И.О.";
            sheet.Cells[firstRow + 3, 4].Value = "Документ, подтверждающий право собственности";
            sheet.Cells[firstRow + 3, 5].Value = "Сведения о размере доли в праве общей собственности на общее имущество собственников помещений в МКД	";
            sheet.Cells[firstRow + 3, 5].WrapText = true;
            sheet.Cells[firstRow + 3, 7].Value = "Площадь";
            sheet.Range[sheet.Cells[firstRow + 3, 1], sheet.Cells[firstRow + 4, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 3, 2], sheet.Cells[firstRow + 4, 2]].Merge();
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 4, 3]].Merge();
            sheet.Range[sheet.Cells[firstRow + 3, 4], sheet.Cells[firstRow + 4, 4]].Merge();
            sheet.Range[sheet.Cells[firstRow + 3, 5], sheet.Cells[firstRow + 3, 6]].Merge();
            sheet.Range[sheet.Cells[firstRow + 3, 7], sheet.Cells[firstRow + 4, 7]].Merge();

            //Четвёртая строка
            sheet.Cells[firstRow + 4, 5].Value = "размер доли ";
            sheet.Cells[firstRow + 4, 6].Value = "кв.м от общей площади помещения";
            sheet.Cells[firstRow + 4, 6].WrapText = true;

            //Конец NewFormXls
        }

        //Заполнение документа и сохранение в directory
        public static void MakeXls(List<Buildding> all_buildings, bool fillEmty, string directory, ref ProgressBar progressBar)
        {
            //Console.WriteLine("\nНачинаю работу с MS Excel. Придётся подождать. Но, если больше 5 минут, то я завис.\n");
            Excel.Application xlsApp = new Excel.Application();
            Excel.Worksheet sheet;
            Excel.Workbook book;

            foreach (Buildding buildding in all_buildings)
            {
                //Console.WriteLine($"Пишем адрес: {buildding.address}.");
                buildding.flats.Sort();

                //Создаём шаблон
                NewFormXls(xlsApp, buildding.addressWithDot);

                book = xlsApp.ActiveWorkbook;
                sheet = (Excel.Worksheet)xlsApp.ActiveSheet;

                //Ставим значения
                int roll = 5, n = 1;
                foreach (Item_flat flat in buildding.flats)
                {
                    progressBar.Value++; // прогресс бар
                    //Console.WriteLine("Пишем: " + flat.address + ", " + flat.numFlat);
                    //if (flat.owners.Count == 0)
                        //Console.WriteLine("     Не нашёл информации о владельцах.");

                    foreach (Owner owner in flat.owners)
                    {
                        sheet.Range[sheet.Cells[roll, 1], sheet.Cells[roll, 4]].HorizontalAlignment = 2;
                        sheet.Cells[roll, 1].Value = n;
                        sheet.Cells[roll, 2].Value = flat.numFlat;
                        sheet.Cells[roll, 3].Value = owner.fio;
                        sheet.Cells[roll, 4].Value = owner.document;
                        sheet.Cells[roll, 5].Value = "'" + owner.part;
                        //sheet.Cells[roll, 6].NumberFormat = "@";
                        sheet.Cells[roll, 6].Value = owner.partOf;
                        sheet.Cells[roll, 7].Value = flat.area;
                        n++;
                        roll++;
                    }
                    //Если не найдены собственники
                    if (fillEmty && flat.owners.Count() == 0)
                    {
                        sheet.Range[sheet.Cells[roll, 1], sheet.Cells[roll, 4]].HorizontalAlignment = 2;
                        sheet.Cells[roll, 1].Value = n;
                        sheet.Cells[roll, 2].Value = flat.numFlat;
                        sheet.Cells[roll, 3].Value = " - нет данных - ";
                        sheet.Cells[roll, 4].Value = " --- ";
                        sheet.Cells[roll, 5].Value = " - ";
                        sheet.Cells[roll, 6].Value = " - ";
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
                    //Console.WriteLine("Не удалось удалить старый файл, возможно он открыт в каком-то приложении.");
                }
                //book.SaveAs(directory + buildding.address
                try { book.SaveAs(directory + buildding.address); }
                catch
                {
                    //Console.WriteLine("Ошибка сохранения xlsx. Возможно файл открыт в другой программе" +
                //"\nЧтобы продолжить можно нажать Enter."); Console.ReadKey();
                }
                book.Close(true);
            }
            xlsApp.Quit();
            //Console.WriteLine("Финиш.");
        }

        //Обработка XML, создание записи по квартире
        public static Item_flat XmlProcessing(XDocument xDoc)
        {
            //Разделитель точка (для перевода строк в Double)
            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = ".";

            //Данные по помещению
            string number_flat_string = "";
            string address = ""; //Адрес
            string addressWithDot = ""; //Дополнительно
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
                                        addressWithDot += node_adr.Attribute("Type").Value + ". ";
                                        addressWithDot += node_adr.Attribute("Name").Value + ", ";
                                        break;
                                    case ("Level1"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value + ", ";
                                        addressWithDot += node_adr.Attribute("Type").Value + ". ";
                                        addressWithDot += node_adr.Attribute("Value").Value + ", ";
                                        break;
                                    case ("Level2"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value + ", ";
                                        addressWithDot += node_adr.Attribute("Type").Value + ". ";
                                        addressWithDot += node_adr.Attribute("Value").Value + ", ";
                                        break;
                                    case ("Level3"):
                                        address += node_adr.Attribute("Type").Value + " ";
                                        address += node_adr.Attribute("Value").Value;
                                        addressWithDot += node_adr.Attribute("Type").Value + ". ";
                                        addressWithDot += node_adr.Attribute("Value").Value + ", ";
                                        break;
                                    case ("Apartment"): //номер квартиры
                                        number_flat_string = node_adr.Attribute("Value").Value + " кв.";
                                        try { number_flat = Convert.ToInt32(node_adr.Attribute("Value").Value); }
                                        catch { number_flat = Int32.MaxValue;/* //Console.WriteLine(" Плохой номер помещения. "); */}
                                        break;
                                    case ("Other"): //что-то другое

                                        string type = "", analise = ""; //Временные строки
                                        if (node_adr.Value.Length > 12 && node_adr.Value.Length < 30 && node_adr.Value.Substring(0, 12) == "машино-место")
                                        { type = "м/м"; analise = node_adr.Value.Substring(12); }
                                        else if (node_adr.Value.Substring(node_adr.Value.Length - 3) == "м/м")
                                        {
                                            type = "м/м";
                                            if (node_adr.Value.Length < 12) analise = node_adr.Value.Substring(0, node_adr.Value.Length - 3);
                                            else
                                            { analise = node_adr.Value.Substring(node_adr.Value.Length - 22); }
                                        }

                                        if (type == "м/м")
                                        {
                                            //Выдераем "a/b"
                                            bool notOne = false;
                                            string number = "";
                                            int num = 0, den = 0;
                                            foreach (char ch in analise)
                                                if (Char.IsNumber(ch))
                                                    number += ch;
                                                else if (num != 0) { den = Convert.ToInt32(number); break; }
                                                else if (ch == ',' && number != "") { number_flat_string += number + ","; number = ""; notOne = true; }
                                                else if (ch == '/' && number != "") { num = Convert.ToInt32(number); number = ""; }

                                            //Вносим в ячейку
                                            number_flat = 5000 + den * 400 + num;
                                            if (notOne) number_flat = 5000 + den * 400; //Переделываем
                                            number_flat_string += num + "/" + den + " " + type;
                                        }
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

                    //---Собственность---
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
                                                if (point.Value == "Собственность") { part = "1"; partOf = area; }
                                                if (point.Value == "Общая совместная собственность") part = "*";
                                            }
                                        }
                                }
                                catch {/* Console.WriteLine("Ошибка получения документа на собственность");*/ }

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
                                                            Owners.Add(new Owner(FIO, part, Convert.ToString(Math.Round(partOf,2)), document));
                                                            FIO = "";
                                                        }
                                                        else if (person.Name.LocalName == "Governance") 
                                                        {
                                                        foreach (XElement point in person.Elements())
                                                            if (point.Name.LocalName == "Name")
                                                                FIO += point.Value + " ";
                                                            // Записываем государственного собственника 
                                                            Owners.Add(new Owner(FIO, part, Convert.ToString(Math.Round(partOf, 2)), document));
                                                            FIO = "";

                                                        }


                                        }
                                }
                                catch { /*Console.WriteLine("Ошибка получения имени собственника."); */}
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
            flat_temp.addressWithDot = addressWithDot;
            return (flat_temp);
            //Конец XmlProcessing
        }

        //Удаление файлов в директории
        public static void clrDir(string adr)
        {
            //*** Рассмотреть внедрение DirectoryInfo
            DirectoryInfo dir = new DirectoryInfo(adr);
            try
            {
                foreach (var file in dir.GetFiles())
                    file.Delete();
            }
            catch {/* //Console.WriteLine("Проблемы с удалением"); */}
        }

        //Чтение Xml файлов 
        public static void ReedXML(ref List<Item_flat> Flats, IEnumerable<string> FileNames, ref ProgressBar progressBar)
        {
            if (FileNames.Count() > 0)
            {
                //XDocument xDoc = new XDocument();
                foreach (string xml_filename in FileNames)
                {
                    progressBar.Value++; //прогресс бар
                    try { Flats.Add(XmlProcessing(XDocument.Load(xml_filename))); }
                    catch { continue; };
                    Console.WriteLine("Прочитал: " + Flats.Last().address + ", " + Flats.Last().numFlat);
                }
            }
        }

        //Чтение Xml файлов сразу из архивов
        public static void ReedZip(ref List<Item_flat> Flats, IEnumerable<string> FileNames, ref ProgressBar progressBar)
        {
            //Временная директория
            string home = Directory.GetCurrentDirectory() + "\\";
            string temp = home + "__tmp__";
            try { Directory.CreateDirectory(temp); }
            catch {/* Console.WriteLine("Ошибка при создании временной папки.");*/ }

            foreach (string arhive_1lvl in FileNames)
            {
                progressBar.Value++; //прогресс бар
                try
                {
                    clrDir(temp); //чистим временную папку
                    ZipFile.ExtractToDirectory(arhive_1lvl, temp); //распаковываем основной архив

                    //Распаковка вложениых архивов
                    foreach (string arhive_1lv2 in Directory.EnumerateFiles(temp, "*.zip", SearchOption.AllDirectories))
                    {
                        //РАзархивация
                        ZipFile.ExtractToDirectory(arhive_1lv2, temp);
                        //Получили XML
                        foreach (string xml_filename in Directory.EnumerateFiles(temp, "*.xml", SearchOption.AllDirectories))
                        {
                            Flats.Add(XmlProcessing(XDocument.Load(xml_filename)));
                            Console.WriteLine("Прочитал: " + Flats.Last().address + ", " + Flats.Last().numFlat);

                            //Выгрузка
                            //Directory.CreateDirectory(home + "\\Выгрузка\\");
                            //File.Copy(xml_filename, home + "\\Выгрузка\\" + All_flats.Last().address + " кв " + All_flats.Last().numFlat + ".xml");
                        }
                    }
                }
                catch {/* Console.WriteLine("Ошибка при работе с архивами.");*/ }

                //Удаление временной дериктории
                clrDir(temp);
                try { Directory.Delete(temp); }
                catch { /*Console.WriteLine("Ошибка при удалении временной папки.");*/ }
            }
            //Удаление временной дериктории
            clrDir(temp);
            try { Directory.Delete(temp); }
            catch {/* Console.WriteLine("Ошибка при удалении временной папки.");*/ }

        }

        //Переименование Xml файлов 
        public static void RenameXML(List<string> FileNames, string Target, ref ProgressBar progressBar, bool SaveDuplicate = false)
        {
            //Console.WriteLine("\nКопирую с переименованием найденные XML в папку \"Переименованные файлы\".");
            //Временная директориz и папка для сохранения
            string target = Target + "Переименованные файлы\\";

            //Решение проблем с прошлой итерацией
            while (Directory.Exists(target)) target += "Новые Переименованные файлы\\";

            try { Directory.CreateDirectory(target); }
            catch {/* Console.WriteLine("Ошибка при создании папки.");*/ }

            //int numFiles = 1;
            if (FileNames.Count() > 0)
                foreach (string xml_filename in FileNames)
                {
                    progressBar.Value++; //прогресс бар

                    //Получение данных об объекте
                    Item_flat flat = XmlProcessing(XDocument.Load(xml_filename));
                    bool SaveThis = true; // нужно ли ещё сохранять этот файл
                                          //Устранение непригодных символов
                    string newName = "";
                    foreach (char ch in flat.address + ", " + flat.numFlat + ".xml")
                    {
                        if (ch == '/')
                            newName += '-';
                        else
                            newName += ch;
                    }

                    if (File.Exists(target + newName))
                    {
                        try
                        {
                            Item_flat taretFlat = XmlProcessing(XDocument.Load(target + newName));

                            if (!flat.Equals(taretFlat))
                            {
                                File.Copy(xml_filename, target + "Другой " + newName);
                                //Console.WriteLine($"файл {numFiles++}. Сохранён дубликат с отличающимеся данными: " + newName);
                                continue;
                            }
                            else if (SaveDuplicate)
                            {
                                if (File.Exists(target + "Дубликат " + newName))
                                {
                                    //Console.WriteLine("Один дубликат уже есть: Дубликат " + newName);
                                    continue;
                                }

                                File.Copy(xml_filename, target + "Дубликат " + newName);
                                //Console.WriteLine($"файл {numFiles++}. Сохранён дубликат с теми же данными: " + newName);
                                continue;
                            }
                        }
                        catch
                        { /*Console.WriteLine("Ошибка при работе с дубликатом");*/ continue; }

                    }
                    //копирование
                    try
                    {
                        if (SaveThis)
                        {
                            File.Copy(xml_filename, target + newName);
                            //Console.WriteLine($"файл {numFiles++}. Сохранён: " + newName);
                        }
                    }
                    catch { /*Console.WriteLine("Ошибка при копировании XML."); continue;*/ }

                }
        }

        //Переименование Xml файлов 
        public static void RenameFiles(List<string> FileNames, string Target, string Type, ref ProgressBar progressBar, bool SaveDuplicate = false)
        {
            //Console.WriteLine("\nКопирую с переименованием найденные XML в папку \"Переименованные файлы\".");
            //Временная директориz и папка для сохранения
            string target = Target + "Переименованные файлы\\";
            //Решение проблем с прошлой итерацией
            while (Directory.Exists(target)) target += "Новые Переименованные файлы\\";
            try { Directory.CreateDirectory(target); }
            catch {/* Console.WriteLine("Ошибка при создании папки.");*/ }

            string temp = target + "__tmp__";
            if (Type == ".zip")
            {                
                Directory.CreateDirectory(temp);
            }

            //int numFiles = 1;
            if (FileNames.Count() > 0)
                foreach (string filename in FileNames)
                {
                    progressBar.Value++; //прогресс бар

                    //Получение данных об объекте
                    Item_flat flat = new Item_flat();
                    if (Type ==".xml")
                        flat = XmlProcessing(XDocument.Load(filename));
                    if (Type == ".zip")
                    {
                        try
                        {
                            clrDir(temp); //чистим временную папку
                            ZipFile.ExtractToDirectory(filename, temp); //распаковываем основной архив

                            //Распаковка вложениых архивов
                            foreach (string arhive_1lv2 in Directory.EnumerateFiles(temp, "*.zip", SearchOption.AllDirectories))
                            {                                
                                ZipFile.ExtractToDirectory(arhive_1lv2, temp);
                                //Получили XML
                                foreach (string xml_filename in Directory.EnumerateFiles(temp, "*.xml", SearchOption.AllDirectories))                                
                                    flat = XmlProcessing(XDocument.Load(xml_filename));                                
                            }
                        }
                        catch { /*Console.WriteLine("Ошибка при работе с архивами.");*/ }
                    }
                    
                    bool SaveThis = true; // нужно ли ещё сохранять этот файл
                                          //Устранение непригодных символов
                    string newName = "";
                    foreach (char ch in flat.address + ", " + flat.numFlat + Type)
                    {
                        if (ch == '/')
                            newName += '-';
                        else
                            newName += ch;
                    }

                    if (File.Exists(target + newName))
                    {
                        try
                        {
                            Item_flat taretFlat = XmlProcessing(XDocument.Load(target + newName));

                            if (!flat.Equals(taretFlat))
                            {
                                File.Copy(filename, target + "Другой " + newName);
                                //Console.WriteLine($"файл {numFiles++}. Сохранён дубликат с отличающимеся данными: " + newName);
                                continue;
                            }
                            else if (SaveDuplicate)
                            {
                                if (File.Exists(target + "Дубликат " + newName))
                                {
                                    //Console.WriteLine("Один дубликат уже есть: Дубликат " + newName);
                                    continue;
                                }

                                File.Copy(filename, target + "Дубликат " + newName);
                                //Console.WriteLine($"файл {numFiles++}. Сохранён дубликат с теми же данными: " + newName);
                                continue;
                            }
                        }
                        catch
                        { /*Console.WriteLine("Ошибка при работе с дубликатом");*/ continue; }

                    }
                    //копирование
                    try
                    {
                        if (SaveThis)
                        {
                            File.Copy(filename, target + newName);
                            //Console.WriteLine($"файл {numFiles++}. Сохранён: " + newName);
                        }
                    }
                    catch { /*Console.WriteLine("Ошибка при копировании XML."); continue;*/ }

                }
            if (Directory.Exists(temp))
            {
                clrDir(temp);
                try { Directory.Delete(temp); }
                catch {/* Console.WriteLine("Ошибка при удалении временной папки.");*/ }
            }
        }

        //Переименование архивов
        public static void RenameZip(List<string> FileNames, string Target, ref ProgressBar progressBar, bool SaveDuplicate = false)
        {

        }

        //Извлечение XML из архивов
        public static void UnZip(List<string> FileNames, bool NeedRename, ref ProgressBar progressBar)
        {
            //Console.WriteLine("\nРаспаковываю архивы.");
            //Временная директориz и папка для сохранения
            string home = Directory.GetCurrentDirectory() + "\\";
            string temp = home + "__tmp__";
            string target = home + "\\Распакованные файлы\\";
            try
            {
                Directory.CreateDirectory(temp);
                Directory.CreateDirectory(target);
            }
            catch { /*Console.WriteLine("Ошибка при создании временной папки.");*/ }

            foreach (string arhive_1lvl in FileNames)
            {
                progressBar.Value++; //прогресс бар
                try
                {
                    clrDir(temp); //чистим временную папку
                    ZipFile.ExtractToDirectory(arhive_1lvl, temp); //распаковываем основной архив

                    //Распаковка вложениых архивов
                    foreach (string arhive_1lv2 in Directory.EnumerateFiles(temp, "*.zip", SearchOption.AllDirectories))
                    {
                        //РАзархивация
                        ZipFile.ExtractToDirectory(arhive_1lv2, temp);
                        //Получили XML
                        foreach (string xml_filename in Directory.EnumerateFiles(temp, "*.xml", SearchOption.AllDirectories))
                        {
                            //Выгрузка
                            Item_flat flat = XmlProcessing(XDocument.Load(xml_filename));
                            if(NeedRename)
                                File.Copy(xml_filename, target + flat.address + ", " + flat.numFlat + ".xml");
                            else
                                File.Copy(xml_filename, target + Path.GetFileName(xml_filename));
                        }
                    }
                }
                catch { /*Console.WriteLine("Ошибка при работе с архивами.");*/ }
            }

            //Удаление временной дериктории
            if (Directory.Exists(temp))
            {
                clrDir(temp);
                try { Directory.Delete(temp); }
                catch {/* Console.WriteLine("Ошибка при удалении временной папки.");*/ }

            }
        }

        //Распределяем квартиры по адресам. Список зданий
        public static List<Buildding> MakeBuildings(List<Item_flat> Flats)
        {
            List<Buildding> All_buildings = new List<Buildding>();
            string active_address = "Ну это точно не попадётся, палка-копалка";
            int active_building = 0;

            //перебор квартир
            foreach (Item_flat flat in Flats)
            {
                if (active_address == flat.address)
                {
                    All_buildings[active_building].flats.Add(flat);
                }
                else
                {
                    // Проверяем наличие дома
                    bool need_new = true;
                    for (int n = 0; n < All_buildings.Count(); n++)
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
                    if (need_new)
                    {
                        All_buildings.Add(new Buildding(flat.address));
                        active_building = All_buildings.Count() - 1;
                        active_address = flat.address;
                        All_buildings[active_building].addressWithDot = flat.addressWithDot;
                    }
                    All_buildings[active_building].flats.Add(flat);
                }
            }
            return All_buildings;
        }
    }

}
