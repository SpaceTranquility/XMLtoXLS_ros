using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace XmlToXls_3
{
    public partial class MainForm : Form
    {
        string home, sourseDirectory, TargetDirectory;

        List<string> All_XML;
        List<string> All_Zip;

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar.Visible = false;
            label_Progress.Visible = false;
            label_NumericInfo.Visible = false;

            home = Directory.GetCurrentDirectory() + "\\";
            sourseDirectory = TargetDirectory = home;
            textBox_SourceAddress.Text = home;
            textBox_TargetAddress.Text = home;
        }

        private void button_SourceAddress_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.SelectedPath = sourseDirectory;
            FBD.ShowNewFolderButton = true;
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                sourseDirectory = FBD.SelectedPath;
                textBox_SourceAddress.Text = sourseDirectory;
            }
        }

        private void button_TargetAddress_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.SelectedPath = TargetDirectory;
            FBD.ShowNewFolderButton = true;
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                TargetDirectory = FBD.SelectedPath;
                textBox_TargetAddress.Text = TargetDirectory;
            }
        }

        private void button_Start_Click(object sender, EventArgs e)
        {
            //Создаём коллекцию нужных файлов в дериктории
            All_XML = Directory.EnumerateFiles(sourseDirectory, "*.xml", SearchOption.AllDirectories).ToList();
            All_Zip = Directory.EnumerateFiles(sourseDirectory, "*.zip", SearchOption.AllDirectories).ToList();

            label_NumericInfo.Visible = true;
            label_NumericInfo.Text = $"Найдено {All_XML.Count} XML и {All_Zip.Count} архивов.";

            //Активация кнопок
            if (All_Zip.Count != 0)
            {
                button_UnPack.Enabled = true;
                button_UnPack_Rename.Enabled = true;
                button_ReNameZip.Enabled = true;
            }
            if (All_XML.Count != 0 || All_Zip.Count != 0)
                button_Processing.Enabled = true;
        }

        //Обработка всех файлов
        private void button_Processing_Click(object sender, EventArgs e)
        {
            bool choiceVisible = true; //Отображение пустых
            /*
            //Выбор отображения
            bool choiceVisible = false;
            Console.Write("\n\nЗаписывать объекты для которых не удалось найти собственников?\n" +
               "1 - Да. Отоброжать.\n" +
               "2 - Нет. Не отображать. (по умолчанию)\n" +
               "Ввидите номер выбранного варианта и нажмите Enter: ");
            try { if (Convert.ToInt32(Console.ReadLine()) == 1) choiceVisible = true; }
            catch { Console.WriteLine("Некорретное значение."); }
            */

            //Все квартиры из всех файлов
            List<Source.Item_flat> All_flats = new List<Source.Item_flat>();

            progressBar.Visible = true;
            label_Progress.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;

            //Чтение всехх Xml файлов 
            label_Progress.Text = "Читаю файлы:";
            progressBar.Maximum = All_XML.Count + All_Zip.Count;
            Source.ReedXML(ref All_flats, All_XML, ref progressBar);
            //Проход по архивам (если есть)
            Source.ReedZip(ref All_flats, All_Zip, ref progressBar);

            // Получим список строений
            label_Progress.Text = "Распределяю помещения:";
            progressBar.Maximum = All_flats.Count;
            progressBar.Value = 0;
            List<Source.Buildding> All_buildings = Source.MakeBuildings(All_flats);

            // Создаём Файлы реестров
            label_Progress.Text = "Заполняю реестр:";
            progressBar.Maximum = All_flats.Count;
            progressBar.Value = 0;
            Source.MakeXls(All_buildings, choiceVisible, TargetDirectory, ref progressBar);
            label_Progress.Text = "Готово.";
        }

        private void button_ReNameZip_Click(object sender, EventArgs e)
        {
            progressBar.Visible = true;
            label_Progress.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;

            label_Progress.Text = "Переименовываю.";
            progressBar.Maximum = All_Zip.Count;
            Source.RenameFiles(All_Zip, TargetDirectory, ".zip", ref progressBar);
            label_Progress.Text = "Готово.";
        }

        private void button_UnPack_Click(object sender, EventArgs e)
        {
            progressBar.Visible = true;
            label_Progress.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;

            label_Progress.Text = "Распаковываю.";
            progressBar.Maximum = All_Zip.Count;
            Source.UnZip(All_Zip, false, ref progressBar);
            label_Progress.Text = "Готово.";
        }

        private void button_UnPack_Rename_Click(object sender, EventArgs e)
        {
            progressBar.Visible = true;
            label_Progress.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;

            label_Progress.Text = "Переименовываю.";
            progressBar.Maximum = All_Zip.Count;
            Source.UnZip(All_Zip, true, ref progressBar);
            label_Progress.Text = "Готово.";
        }
    }
}
