using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using ComboBox = Autodesk.Revit.UI.ComboBox;
using Form = System.Windows.Forms.Form;
using Newtonsoft.Json;
using Autodesk.Revit.UI.Selection;

namespace bricksCountKursovoy
{
    public partial class bricksCountForm : Form
    {
        private ExternalEvent f_ExEvent;
        private CommandHandler f_Handler;
        List<Brick> materials = new List<Brick>();
        private string[] typesOfBricks = new string[] { "Одинарный", "Полуторный", "Двойной" };
        private string[] typesOfFill = new string[] { "В 1 кирпич", "В 2 кирпича", "В 2.5 кирпича" };
        private int[] amountOfBricksFor1Br = new int[3] { 51, 39, 26 };
        private int[] amountOfBricksFor2Br = new int[3] { 204, 156, 104 };
        private int[] amountOfBricksFor25Br = new int[3] { 255, 195, 130 };
        private Wall wall;
        private UIApplication uiApplication;

        public bricksCountForm(ExternalEvent exEvent, CommandHandler handler, UIApplication uiapp)
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            comboBox1.Items.Add("Нет");
            comboBox1.SelectedIndex = comboBox1.Items.IndexOf("Нет");
            comboBox2.Items.AddRange(typesOfBricks);
            comboBox2.SelectedIndex = comboBox2.Items.IndexOf("Одинарный");
            comboBox3.Items.AddRange(typesOfFill);
            comboBox3.SelectedIndex = comboBox3.Items.IndexOf("В 2 кирпича");
            f_ExEvent = exEvent;
            f_Handler = handler;
            textBox1.ReadOnly = true;
            uiApplication = uiapp;
            String JSONtxt =
                    File.ReadAllText(@"C:\Users\Sergey\AppData\Roaming\Autodesk\Revit\Addins\2020\MaterialDoc.json");
            IEnumerable<Brick> materialsjson =
                Newtonsoft.Json.JsonConvert.DeserializeObject<IEnumerable<Brick>>(JSONtxt);
            foreach (var elm in materialsjson)
            {
                comboBox1.Items.Add(elm.Name + " " + elm.InPack + "м2");
                materials.Add(elm);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            f_ExEvent.Dispose();
            f_ExEvent = null;
            f_Handler = null;
            base.OnFormClosed(e);
        }

        
        // Расчет для разных типов кирпичей

        public double type1bricksCount(Brick mat) // Одинарный
        {
            double amountOfBricks = 0;
            switch (comboBox3.Text)
            {
                case "В 1 кирпич":
                    amountOfBricks = 102;
                    break;
                case "В 2 кирпича":
                    amountOfBricks = 204;
                    break;
                case "В 2.5 кирпича":
                    amountOfBricks = 255;
                    break;
            }
            return amountOfBricks;
        }
        public double type15bricksCount(Brick mat) // Полуторный
        {
            double amountOfBricks = 0;
            switch (comboBox3.Text)
            {
                case "В 1 кирпич":
                    amountOfBricks = 78;
                    break;
                case "В 2 кирпича":
                    amountOfBricks = 156;
                    break;
                case "В 2.5 кирпича":
                    amountOfBricks = 195;
                    break;
            }
            return amountOfBricks;
        }
        public double type2bricksCount(Brick mat) // Двойной
        {
            double amountOfBricks = 0;
            switch (comboBox3.Text)
            {
                case "В 1 кирпич":
                    amountOfBricks = 52;
                    break;
                case "В 2 кирпича":
                    amountOfBricks = 104;
                    break;
                case "В 2.5 кирпича":
                    amountOfBricks = 130;
                    break;
            }
            return amountOfBricks;
        }


        // расчет площади
        public double AreaCount()
        {
            Solid wallSolid = null;
            GeometryElement wallGeometry = wall.get_Geometry(new Options());  // Получение геометрии стены
            foreach (var geomObject in wallGeometry)
            {
                if (geomObject is Solid)
                {
                    Solid solid = geomObject as Solid;
                    if (solid.Volume > 0)
                        wallSolid = solid;
                }
                if (geomObject is GeometryInstance geomInst)
                {
                    // Геометрия экземпляра получается таким образом, что пересечение работает должным образом, не требуя преобразования

                    GeometryElement instElem = geomInst.GetInstanceGeometry();

                    foreach (GeometryObject instObj in instElem)
                    {
                        if (instObj is Solid solid)
                        {
                            if (solid.Volume > 0)
                                wallSolid = solid;
                        }
                    }
                }
            }
            var faces = wallSolid.Faces; // получаем поверхности стены
            double ar = 0;
            foreach (Face face in faces)
            {
                var faceArea = face.Area; // складываем площади
                ar += faceArea; // поверхностей 
            }
            return ar / 32; // 32 - универсальное число для преобразования площади, которая получается с помощью Revit API, в действительную (примерно);
        }

        private bool SelectWall() // выбор стены
        {
            Selection selection = uiApplication.ActiveUIDocument.Selection;
            UIDocument uidoc = uiApplication.ActiveUIDocument;
            Document doc = uidoc.Document;
            try
            {
                Reference pickedRef = null;
                pickedRef = selection.PickObject(ObjectType.Element,
                        "Выберите стену");
                Element elem = doc.GetElement(pickedRef);
                if (elem is Wall) // элемент - стена?
                {
                    textBox1.Text = "";
                    wall = (Wall)elem;
                    return true;
                }
                else 
                {
                    TaskDialog.Show("Предупреждение", "Выберите стену. Вы выбрали элемент с именем  " + elem.Name);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return false;
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

            private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
            {
            textBox1.Text = "";
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (SelectWall())
            {
                label4.Text = "Стена выбрана";
            }
            this.Activate();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void okbtn_Click_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (textBox2.Text.Replace(" ", string.Empty) != "" && textBox3.Text.Replace(" ", string.Empty) != "")
            {
                try
                {
                    Brick newMat = new Brick();
                    newMat.Name = textBox2.Text;
                    newMat.bType = comboBox2.Text;
                    newMat.InPack = double.Parse(textBox3.Text);
                    foreach (var material in materials)
                    {
                        if (newMat.Name == material.Name && newMat.InPack == material.InPack)
                        {
                            MessageBox.Show("Идентичный этому тип кирпича с одинаковым количеством кирпича на м2 уже существует!");
                            return;
                        }
                    }

                    materials.Add(newMat);
                    comboBox1.Items.Add(newMat.Name + " " + newMat.InPack + "м2");

                    using (StreamWriter file =
                        File.CreateText(@"C:\Users\Sergey\AppData\Roaming\Autodesk\Revit\Addins\2020\MaterialDoc.json"))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        serializer.Serialize(file, materials);
                    }

                    MessageBox.Show("Успешное добавление типа кирпича!");
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message + "\n" + exception.StackTrace);
                    throw;
                }
                this.Activate();
            }
            else
            {
                TaskDialog.Show("Предупреждение", "Все поля должны быть заполнены!");
                this.Activate();
            }
        }

        private void cnlbtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void bricksCountForm_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void computebtn_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "Нет" && label4.Text == "Стена выбрана")
            {
                try
                {
                    string bricksName = comboBox1.Text;
                    double amountOfBricks = 0, amountOfPacks = 0;
                    double area = AreaCount();
                    foreach (var material in materials)
                    {
                        if (material.Name + " " + material.InPack + "м2" == bricksName)
                        {
                            amountOfPacks = Math.Ceiling(area / material.InPack);
                            switch (material.bType)
                            {
                                case "Одинарный":
                                    amountOfBricks = type1bricksCount(material);
                                    break;
                                case "Полуторный":
                                    amountOfBricks = type15bricksCount(material);
                                    break;
                                case "Двойной":
                                    amountOfBricks = type2bricksCount(material);
                                    break;
                            }
                            textBox1.Text = "На стену вам понадобится " + amountOfPacks + " упаковок данного кирпича, в каждой по " + (amountOfBricks) + " кирпичей (Всего: " + (amountOfBricks * amountOfPacks).ToString() + " штук).";
                        }
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message + "\n" + exception.StackTrace);
                    throw;
                }
            }
            else if (comboBox1.Text == "Нет")
            {
                TaskDialog.Show("Предупреждение", "Пожалуйста, выберите тип кирпича!");
            }
            else if (label4.Text != "Стена выбрана")
            {
                TaskDialog.Show("Предупреждение", "Пожалуйста, выберите стену!");
            }
                this.Activate();
        }
    }
}
