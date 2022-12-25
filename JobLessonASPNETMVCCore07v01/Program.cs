using TemplateEngine.Docx;

namespace JobLessonASPNETMVCCore07v01
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DiscsInformation.DiskReport();
        }
        public static class DiscsInformation
        {
            private static bool _exit = true;
            private static string _getPath = null;
            private static string _pathOld;
            private static string _pathNew;

            public static List<DriveInfoSave> infoSaves { get; set; } = new List<DriveInfoSave>();

            public static DriveInfoSave driveInfoSave { get; set; }
            public static void Run()
            {
                do
                {
                    Console.WriteLine(_allCommand[0]);
                    Console.Write(PathLine.GetPath(_getPath));
                    Console.Write(">>");

                    Input(Console.ReadLine());
                    Console.WriteLine();
                } while (_exit);
            }

            private static async void Input(string str)
            {
                str = str.ToLower();
                switch (str)
                {

                    case "help":
                        foreach (var i in _allCommand)
                        {
                            Console.WriteLine(i);
                        }
                        break;

                    case "q":
                        _exit = false;
                        break;

                    case "d":
                        foreach (var i in PathLine.GetDrives())
                        {
                            Console.WriteLine(i);
                        }
                        break;

                    case "show":
                        Console.WriteLine(@"Введите адрес каталога в формате 'd:\folder\subfolder'");
                        string str2 = Console.ReadLine();
                        PathLine.GetShow(@$"{str2}");
                        break;

                    case "cd":
                        Console.WriteLine("Введите путь к переходу в новую директорию");
                        _getPath = PathLine.GetPath(Console.ReadLine());
                        break;

                    case "null":
                        _getPath = null;
                        break;

                    case "nc":
                        if (_getPath == null)
                        {
                            Console.WriteLine("Не указана директория, где создать новую папку");
                        }
                        else
                        {
                            Console.WriteLine("Введите название папки");
                            _getPath = PathLine.CreateCatalog(_getPath, Console.ReadLine());
                        }
                        break;

                    case "nf":
                        if (_getPath == null)
                        {
                            Console.WriteLine("Не указана директория, где сохранить новый файл");
                        }
                        else
                        {
                            Console.WriteLine("Введите название и расширение файла в формате 'text.txt'");
                            PathLine.CreateFile(_getPath, Console.ReadLine());
                        }

                        break;

                    case "dc":
                        if (_getPath == null)
                        {
                            Console.WriteLine("Не указанна директория, где хранится папка");
                        }
                        else
                        {
                            Console.WriteLine("Введите название папки для удаления");
                            PathLine.DeleteCatalog(_getPath, Console.ReadLine());
                        }
                        break;

                    case "df":
                        if (_getPath == null)
                        {
                            Console.WriteLine("Не удаётся найти директорию, из которой удалить файл");
                        }
                        else
                        {
                            Console.WriteLine("Введите название и расширение файла в формате для удаления 'text.txt'");
                            PathLine.DeleteFile(_getPath, Console.ReadLine());
                        }
                        break;

                    case "mf":
                        Console.WriteLine(@"Введите название файла для перемещения в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название файла, куда переместить в формате 'd:\folder\text.txt'");
                        _pathNew = Console.ReadLine();

                        PathLine.MoveToFile(_pathOld, _pathNew);
                        break;

                    case "mc":
                        Console.WriteLine(@"Введите название папки для перемещения в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название новой папки, куда переместить в формате 'd:\folder\text'");
                        _pathNew = Console.ReadLine();

                        PathLine.MoveToCatalog(_pathOld, _pathNew);
                        break;

                    case "rc":
                        Console.WriteLine(@"Введите название папки, которую надо переименовать в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название новой папки в формате 'd:\folder\text'");
                        _pathNew = Console.ReadLine();

                        PathLine.RenameCatalog(_pathOld, _pathNew);
                        break;

                    case "rf":
                        Console.WriteLine(@"Введите название файла, который надо переименовать в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название файла, в который надо переименовать в формате 'd:\folder\text.txt'");
                        _pathNew = Console.ReadLine();

                        PathLine.RenameFile(_pathOld, _pathNew);
                        break;

                    case "cc":
                        Console.WriteLine(@"Введите название папки, которую надо копировать в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название новой папки, который надо создать в формате 'd:\folder\text'");
                        _pathNew = Console.ReadLine();

                        PathLine.CopyCatalog(_pathOld, _pathNew);
                        break;

                    case "cf":
                        Console.WriteLine(@"Введите название файла, который надо копировать в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Введите название нового файла, который надо создать при копировании в формате 'd:\folder\text.txt'");
                        _pathNew = Console.ReadLine();

                        PathLine.CopyFile(_pathOld, _pathNew);
                        break;

                    case "gc":
                        Console.WriteLine(@"Произвести поиск папок по маске в папке в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Указать часть папки в формате 'c*'");
                        _pathNew = Console.ReadLine();

                        PathLine.GetFullDirectories(_pathOld, _pathNew);
                        break;

                    case "gf":
                        Console.WriteLine(@"Произвести поиск файлов по маске в папке в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Произвести поиск файлов по маске файлов в формате '*.txt'");
                        _pathNew = Console.ReadLine();

                        PathLine.GetFullFiles(_pathOld, _pathNew);
                        break;


                    case "ic":
                        Console.WriteLine(@"Введите адрес папки для вычесления размера в формате 'd:\folder\text'");
                        _pathOld = Console.ReadLine();


                        PathLine.InfoCatalog(_pathOld);
                        break;


                    case "if":
                        Console.WriteLine(@"Введите адрес файла для вычисления данных о файле в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        PathLine.InfoFile(_pathOld);
                        break;



                    case "sf":
                        Console.WriteLine(@"Введите адрес файла в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        PathLine.TextFileStatistics(_pathOld);
                        break;


                    case "af":
                        Console.WriteLine(@"Введите адрес файла для изменения атрибута только для чтения в формате 'd:\folder\text.txt'");
                        _pathOld = Console.ReadLine();

                        Console.WriteLine(@"Передайте параметр  для файла 'f'-не активен только для чтения, 't'- активен только для чтения");
                        _pathNew = Console.ReadLine();

                        PathLine.AttributeFile(_pathOld, _pathNew);
                        break;

                    case "dr":
                        DiskReport();
                        break;

                    default:
                        break;
                }
            }

            private static List<string> _allCommand = new List<string>()
        {
            "help - выводит справочную информацию о командах",
            "q- выход из программы",
            "d- показать диски",
            "show -показать содержимое диска или папки",
            "'null' - выйти из каталога",
            @"'cd' затем 'C:\Windows'- переход в данный каталог",
            "'nc'- создать новую папку в данной директории",
            "'nf'- создать новый файл в данной директории",
            "'dc'- удалить папку в данной директории",
            "'df'- удалить файл в данной директории",
            "'mf'- переместить файл",
            "'mc'- переместить папку",
            "'rf'- переименовать файл",
            "'cc'- копировать папку",
            "'cf'- копировать файл",
            "'gc'- поиск папок по маске",
            "'gf'- поиск файлов по маске",
            "'ic'- вычислить размер папки",
            "'if'- вычислить размер файла",
            "'sf'- статистические данные о файле",
            "'af'- установить/удалить атрибут только для чтения в файле",
            "'dr'- отчёт по использованию дисков",

        };


            /// <summary>
            /// Отчёт по дискам
            /// </summary>
            public static void DiskReport()
            {
                Console.WriteLine("Отчет по дискам:");
                DriveInfo[] drives = DriveInfo.GetDrives();

                foreach (DriveInfo drive in drives)
                {
                    driveInfoSave = new DriveInfoSave();
                    Console.WriteLine($"Название: {drive.Name}");
                    driveInfoSave.Name = drive.Name;

                    Console.WriteLine($"Тип: {drive.DriveType}");
                    driveInfoSave.DriveType = drive.DriveType;
                    if (drive.IsReady)
                    {
                        Console.WriteLine($"Объем диска: {drive.TotalSize}");
                        driveInfoSave.TotalSize = drive.TotalSize;
                        Console.WriteLine($"Свободное пространство: {drive.TotalFreeSpace}");
                        driveInfoSave.TotalFreeSpace = drive.TotalFreeSpace;
                        Console.WriteLine($"Метка диска: {drive.VolumeLabel}");
                        driveInfoSave.VolumeLabel = drive.VolumeLabel;
                    }

                    Console.WriteLine();
                    infoSaves.Add(driveInfoSave);
                }
                TemplateMicrosoftWord(infoSaves);
            }

            private static void TemplateMicrosoftWord(List<DriveInfoSave> infoSave, string output = "")
            {
                if (infoSave is null)
                {
                    return;
                }
                if (string.IsNullOrWhiteSpace(output))
                {
                    output = Path.Combine(Directory.GetCurrentDirectory(),
                    "DriveInfoSave01.docx");
                }
                if (File.Exists(output))
                {
                    File.Delete(output);
                }
                //File.Copy("D:\\Программирование\\GeekBrains\\ASP.NET MVC Core\\MVC\\Report01.docx", output);
                File.Copy("Report01.docx", output);
                List<TableRowContent> rows = new List<TableRowContent>();

                foreach (DriveInfoSave drive in infoSave)
                {
                    rows.Add(new TableRowContent(new List<FieldContent>()
                {
                    new FieldContent("Field Name",drive.Name),
                    new FieldContent("Field DriveType",drive.DriveType.ToString()),
                    new FieldContent("Field TotalSize",drive.TotalSize.ToString()),
                    new FieldContent("Field TotalFreeSpace",drive.TotalFreeSpace.ToString()),
                    new FieldContent("Field VolumeLabel",drive.VolumeLabel.ToString()),
                }));
                }


                var valuesToFill = new Content(

                TableContent.Create("Lines", rows)
                );

                using (var outputDocument = new TemplateProcessor(output)
                .SetRemoveContentControls(true))
                {
                    outputDocument.FillContent(valuesToFill);
                    outputDocument.SaveChanges();
                }

            }

        }

        public class DriveInfoSave
        {
            public string Name { get; set; }
            public DriveType DriveType { get; set; }
            public long TotalSize { get; set; }
            public long TotalFreeSpace { get; set; }
            public string VolumeLabel { get; set; }

        }
    }
    public static class PathLine
    {
        /// <summary>
        /// Показать все диски
        /// </summary>
        /// <returns></returns>
        public static List<DriveInfo> GetDrives()
        {
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            List<DriveInfo> newAllDrives = new List<DriveInfo>();
            foreach (var driveInfo in allDrives)
            {
                if (driveInfo.IsReady)
                {
                    newAllDrives.Add(driveInfo);
                }
            }
            return newAllDrives;
        }

        /// <summary>
        /// Показать все файлы
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<FileInfo> GetFiles(String path)
        {
            if (path == "" || path == null)
            {
                return null;
            }
            if (Path.HasExtension(path))
            {
                path = Path.GetDirectoryName(path);
            }
            return new DirectoryInfo(path).GetFiles().ToList();
        }

        /// <summary>
        /// Показать все папки и файлы в заданной директории
        /// </summary>
        /// <param name="dirName"></param>
        public static void GetShow(string dirName)
        {
            Console.WriteLine();
            if (Directory.Exists(dirName))
            {
                Console.WriteLine("Подкаталоги:");
                string[] dirs = Directory.GetDirectories(dirName);
                foreach (string s in dirs)
                {
                    Console.WriteLine(s);
                }
                Console.WriteLine();
                Console.WriteLine("Файлы:");
                string[] files = Directory.GetFiles(dirName);
                foreach (string s in files)
                {
                    Console.WriteLine(s);
                }
            }
        }

        /// <summary>
        /// Название директории
        /// </summary>
        /// <param name="dirName"></param>
        /// <returns></returns>
        public static string GetPath(string dirName)
        {
            if (Directory.Exists(dirName))
            {
                return dirName;
            }
            else
                return null;
        }

        /// <summary>
        /// Создать папку
        /// </summary>
        /// <param name="oldPath"></param>
        /// <param name="newCatalog"></param>
        /// <returns></returns>
        public static string CreateCatalog(string oldPath, string newCatalog)
        {
            string oldPath2 = @$"{oldPath}\{newCatalog}";
            DirectoryInfo dirInfo2 = new DirectoryInfo(oldPath2);
            DirectoryInfo dirInfo = new DirectoryInfo(oldPath);

            if (dirInfo2.Exists)
            {
                //dirInfo.Create();
                Console.WriteLine("такой каталог уже есть");
            }
            else
            {
                dirInfo.CreateSubdirectory(newCatalog);
                Console.WriteLine("папка успешно создана");
            }
            return dirInfo.FullName;
        }

        /// <summary>
        /// Создать файл
        /// </summary>
        /// <param name="path"></param>
        /// <param name="file"></param>
        public static void CreateFile(string path, string file)
        {
            using (FileStream fstream = new FileStream($@"{path}\{file}", FileMode.OpenOrCreate))
            {
            }
        }

        /// <summary>
        /// Удалить папку
        /// </summary>
        /// <param name="path"></param>
        /// <param name="folderDil"></param>
        public static void DeleteCatalog(string path, string folderDil)
        {
            try
            {
                path = @$"{path}\{folderDil}";
                DirectoryInfo dirInfo = new DirectoryInfo(path);
                dirInfo.Delete(true);
                Console.WriteLine("Каталог удалён");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Удалить файл
        /// </summary>
        /// <param name="path"></param>
        /// <param name="file"></param>
        public static void DeleteFile(string path, string file)
        {
            path = @$"{path}\{file}";
            FileInfo fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }
        }

        /// <summary>
        /// Переместить папку
        /// </summary>
        /// <param name="pathOld"></param>
        /// <param name="pathNew"></param>
        public static void MoveToCatalog(string pathOld, string pathNew)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(pathOld);
            if (dirInfo.Exists && Directory.Exists(pathNew) == false)
            {
                dirInfo.MoveTo(pathNew);
            }
        }

        /// <summary>
        /// Переместить файл
        /// </summary>
        /// <param name="pathOld"></param>
        /// <param name="pathNew"></param>
        public static void MoveToFile(string pathOld, string pathNew)
        {
            FileInfo fileInf = new FileInfo(pathOld);
            if (fileInf.Exists)
            {
                fileInf.MoveTo(pathNew);
            }
        }

        /// <summary>
        /// Переименовать папку
        /// </summary>
        /// <param name="pathOld"></param>
        /// <param name="pathNew"></param>
        public static void RenameCatalog(string pathOld, string pathNew)
        {
            Directory.Move(pathOld, pathNew);
        }

        /// <summary>
        /// Переименовать файл
        /// </summary>
        /// <param name="pathOld"></param>
        /// <param name="pathNew"></param>
        public static void RenameFile(string pathOld, string pathNew)
        {
            File.Move(pathOld, pathNew);
        }

        public static void CopyFile(string toCopy, string newFile)
        {
            File.Copy(toCopy, newFile);
        }

        public static void CopyCatalog(string toCopyCatalog, string newCopyCatalog)
        {
            DirectoryInfo diSource = new DirectoryInfo(toCopyCatalog);
            DirectoryInfo diTarget = new DirectoryInfo(newCopyCatalog);

            if (Directory.Exists(diTarget.FullName) == false)
            {
                Directory.CreateDirectory(diTarget.FullName);
            }

            // Копируем все файлы в новую директорию
            foreach (FileInfo fi in diSource.GetFiles())
            {
                fi.CopyTo(Path.Combine(diTarget.ToString(), fi.Name), true);
            }

            // Копируем рекурсивно все поддиректории
            foreach (DirectoryInfo diSourceSubDir in diSource.GetDirectories())
            {
                diTarget.CreateSubdirectory(diSourceSubDir.Name);
            }
        }

        /// <summary>
        /// Поиск файлов по маске
        /// </summary>
        /// <param name="getCatalog"></param>
        /// <param name="getFiles"></param>
        public static void GetFullFiles(string getCatalog, string getFiles)
        {
            string[] files = Directory.GetFiles(getCatalog, getFiles);

            Console.WriteLine("Всего файлов: {0}.", files.Length);
            foreach (string f in files)
            {
                Console.WriteLine(f);
            }
        }

        /// <summary>
        /// Поиск папок по маске
        /// </summary>
        public static void GetFullDirectories(string getPath, string pathCatalog)
        {
            string[] dirs = Directory.GetDirectories(getPath, pathCatalog);
            Console.WriteLine("Всего каталогов: {0}", dirs.Length);

            foreach (string d in dirs)
            {
                Console.WriteLine(d);
            }
        }

        /// <summary>
        /// Вывод данных о папке
        /// </summary>
        /// <param name="infoCatalog"></param>
        public static void InfoCatalog(string infoCatalog)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(infoCatalog);

            Console.WriteLine($"Название каталога: {dirInfo.Name}");
            Console.WriteLine($"Полное название каталога: {dirInfo.FullName}");
            Console.WriteLine($"Время создания каталога: {dirInfo.CreationTime}");
            Console.WriteLine($"Корневой каталог: {dirInfo.Root}");
        }

        /// <summary>
        /// Вывод данных о файле
        /// </summary>
        /// <param name="infoFile"></param>
        public static void InfoFile(string infoFile)
        {
            FileInfo fileInf = new FileInfo(infoFile);
            if (fileInf.Exists)
            {
                Console.WriteLine("Имя файла: {0}", fileInf.Name);
                Console.WriteLine("Время создания: {0}", fileInf.CreationTime);
                Console.WriteLine("Размер: {0}", fileInf.Length);
            }
        }

        public static void TextFileStatistics(string fileStatict)
        {
            try
            {
                using (StreamReader sr = new StreamReader(fileStatict))
                {
                    Console.WriteLine(sr.ReadToEnd());     //весь текст
                    string word = "";
                    string[] words;

                    while (sr.EndOfStream != true)
                    {
                        word += sr.ReadLine();
                    }
                    words = word.Split(' ');
                    Console.WriteLine("Количество слов: " + words.Length);
                }

                using (StreamReader sr = new StreamReader(fileStatict))
                {
                    string line;
                    int i = 0;
                    while ((line = sr.ReadLine()) != null) //читаем по одной линии(строке) пока не вычитаем все из потока (пока не достигнем конца файла)
                    {
                        i++;
                        //Console.WriteLine(line);
                    }
                    Console.WriteLine("Количество строк: " + i.ToString());
                }

                using (StreamReader sr = new StreamReader(fileStatict))
                {
                    string line;
                    int i = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line[0] == '\t')
                        {
                            i++;
                        }
                    }
                    Console.WriteLine("Количество абзацев: " + i.ToString());
                }

                using (StreamReader sr = new StreamReader(fileStatict))
                {
                    string s = sr.ReadToEnd();
                    char[] ch = s.ToCharArray();

                    Console.WriteLine("Количество символов с пробелами: " + ch.Length.ToString());
                }

                using (StreamReader sr = new StreamReader(fileStatict))
                {
                    string st = sr.ReadToEnd();
                    string[] SMass;
                    int sumStr = 0;

                    SMass = st.Split(' ', '\n');

                    foreach (var i in SMass)
                    {
                        if (i == "" || i == "\r")
                        {
                            continue;
                        }
                        sumStr++;
                    }
                    Console.WriteLine("Количество слов без пробелов: " + sumStr);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void AttributeFile(string pathFile, string attrib)
        {
            FileInfo fileInfo = new FileInfo(pathFile);

            bool at = false;
            if (attrib == "t")
            {
                at = true;
            }

            fileInfo.IsReadOnly = at;

            Console.WriteLine("Атрибут IsReadOnly: " + fileInfo.IsReadOnly);
        }
    }
}
