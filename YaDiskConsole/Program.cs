using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using NPOI.XWPF.UserModel;
using YandexDiskNET;


namespace YaDiskConsole
{

    public static class ToWord
    {
        public static void ImportReport(List<SortUser> spisokSortU) //Импорт отчета в word документ в папку Reports. 
        {
            XWPFDocument document = new XWPFDocument();
            XWPFParagraph p2 = document.CreateParagraph();
            p2.Alignment = ParagraphAlignment.CENTER;
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("Отчет");
            r2.IsBold = true;
            r2.FontFamily = "Times New Roman";
            r2.FontSize = 14;

            XWPFParagraph p3 = document.CreateParagraph();
            p3.Alignment = ParagraphAlignment.CENTER;
            XWPFRun r3 = p3.CreateRun();
            r3.SetText("о загруженных документах в портфолио");
            r3.IsBold = true;
            r3.FontFamily = "Times New Roman";
            r3.FontSize = 14;

            XWPFParagraph p4 = document.CreateParagraph();
            p4.Alignment = ParagraphAlignment.CENTER;
            XWPFRun r4 = p4.CreateRun();
            r4.SetText("за период с " + DateTime.Now.AddMonths(-1).ToShortDateString() +
                " по " + DateTime.Now.Date.ToShortDateString());
            r4.IsBold = true;
            r4.FontFamily = "Times New Roman";
            r4.FontSize = 14;
            r4.TextPosition = 40;

            foreach (var item in spisokSortU)
            {

                XWPFParagraph p = document.CreateParagraph();
                XWPFRun r = p.CreateRun();
                r.SetText(item.fio);
                r.IsBold = true;
                r.FontFamily = "Times New Roman";
                r.FontSize = 14;
                foreach (var i in item.files)
                {
                    XWPFParagraph p1 = document.CreateParagraph();
                    XWPFRun r1 = p1.CreateRun();
                    r1.SetText(i);
                    r1.FontFamily = "Times New Roman";
                    r1.FontSize = 14;

                }
            }

            string basepath = AppDomain.CurrentDomain.BaseDirectory.ToString();
            string nameOfFile = "Reports\\Отчет о портфолио на " + DateTime.Now.Date.ToShortDateString() + ".docx";
            string newpath = Path.Combine(basepath, nameOfFile);


            try
            {
                //Проверяем, не создан ли данный путь в предыдущие запуски программы.
                if (!Directory.Exists(Path.Combine(basepath, "Reports")))
                {
                    //Путь пока не создан... 
                    try
                    {
                        //Пытаемся создать папку:
                        Directory.CreateDirectory(Path.Combine(basepath, "Reports"));
                    }
                    catch (IOException ex)
                    {
                        //В случае ошибок ввода-вывода выдаем сообщение об ошибке
                        Console.WriteLine("Не получается создать папку Reports. Создайте самостоятельно.");
                        Console.ReadLine();
                        //Вновь генерируем обшибку. В случае необходимости реакцию на ошибку ввода-вывода
                        //можно изменить именно тут:
                        throw ex;
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        //В случае ошибки с нехваткой прав вновь выдаем сообщение:
                        Console.WriteLine(ex.Message);
                        //И вновь генерируем ошибку. Если нужно обработаь ошибку более детально, то тут как раз
                        //самое место это сделать. 
                        throw ex;
                    }
                }
                using (FileStream fs = new FileStream(newpath, FileMode.Create))
                {
                    document.Write(fs);
                }

            }
            catch
            {
                Console.WriteLine("Ошибка записи файла");

            }

        }
    }
    public struct User
    {
        public User(string _fio, string _name, string _category)
        {
            fio = _fio;
            name = _name;
            category = _category;
        }
        public string name; // имя файла
        public string fio; // фио 
        public string category; //категория работы

    }
    public struct SortUser
    {
        public SortUser(string _fio, string _category)
        {
            fio = _fio;
            category = _category;
            files = new List<string>();
        }
        public string fio;
        public string category;
        public List<string> files;

        public void Add(string n)
        {
            files.Add(n);
        }

        public void Display()
        {
            Console.WriteLine(fio + " - " + category);
            foreach (var item in files.OrderBy(x => x))
            {
                Console.WriteLine(item);
            }
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string oauth = "bla-bla"; //токен
            YandexDiskRest disk = new YandexDiskRest(oauth);
            if (disk.GetDiskInfo().ErrorResponse.Message != null)
            {
                Console.WriteLine("Проблема с подключением. Проверьте токен!");

            }

            ResInfo filesByDateFields = disk.GetResourceByDate(
               (int)5e6,
               new Media_type[]
               {
                    Media_type.Audio,
                    Media_type.Backup,
                    Media_type.Compressed,
                    Media_type.Book,
                    Media_type.Data,
                    Media_type.Development,
                    Media_type.Diskimage,
                    Media_type.Document,
                    Media_type.Encoded,
                    Media_type.Executable,
                    Media_type.Flash,
                    Media_type.Font,
                    Media_type.Image,
                    Media_type.Settings,
                    Media_type.Spreadsheet,
                    Media_type.Text,
                    Media_type.Unknown,
                    Media_type.Video,
                    Media_type.Web
               },
               new ResFields[] {
                    ResFields.Antivirus_status,
                    ResFields.Created,
                    ResFields.Md5,
                    ResFields.Media_type,
                    ResFields.Mime_type,
                    ResFields.Modified,
                    ResFields.Name,
                    ResFields.Path,
                    ResFields.Preview,
                    ResFields.Public_key,
                    ResFields.Public_url,
                    ResFields.Sha256,
                    ResFields.Size,
                    ResFields.Type,
                    ResFields._Embedded
               },
               true, "120x240");

            List<string> resSpisokPath = new List<string>(); //Список путей к файлам

            if (filesByDateFields.ErrorResponse.Message == null)
            {
                if (filesByDateFields._Embedded.Items.Count != 0)
                {
                    foreach (var item in filesByDateFields._Embedded.Items)
                    {
                        DateTime date1 = new DateTime(); // здесь будет дата создания
                        try
                        {
                            string d = item.Created.ToString();
                            string[] q = d.Split('/');
                            string temp = q[1];
                            q[1] = q[0];
                            q[0] = temp;
                            string s = "";
                            foreach (var i in q) s += i + "/";
                            s = s.TrimEnd('/');
                            date1 = DateTime.Parse(s);
                        }
                        catch
                        {
                            Console.WriteLine("Ошибка обработки даты создания в файле " + item.Name.ToString());
                            Console.WriteLine("Нажмите Enter");
                            Console.ReadLine();
                        }


                        if (date1 > DateTime.Now.AddMonths(-1))
                        {
                            resSpisokPath.Add(item.Path.ToString());
                        }


                    }
                }
            }
            List<User> spisokU = new List<User>(); // список ФИО, категорий и путей
            List<SortUser> spisokSortU = new List<SortUser>(); // список сгруппиррованный по фио

            foreach (var item in resSpisokPath)
            {
                string[] p = item.Split('/');
                string fio = "";
                string kategory = "";
                string name = "";


                if (p[2] == "Сотрудники")
                {
                    kategory = p[3];
                    fio = p[4];
                    for (int i = 5; i < p.Length; i++)
                    {
                        name += " " + p[i];
                    }
                }
                if (p[2] == "Студенты")
                {
                    kategory = p[2];
                    fio = p[4] + " - " + p[5];
                    name = " " + p[p.Length - 1];

                }
                spisokU.Add(new User(fio, name, kategory));
                if (spisokSortU.Where(x => x.fio == fio).Count() == 0)
                {
                    SortUser sortUser = new SortUser(fio, kategory);
                    sortUser.Add(name);
                    spisokSortU.Add(sortUser);
                }
                else
                {
                    spisokSortU.First(x => x.fio == fio).Add(name);
                }


            }
            foreach (var item in spisokSortU) item.Display();
            ToWord.ImportReport(spisokSortU);
            Console.WriteLine("Нажмите Enter");
            Console.ReadLine();



        }
    }
}
