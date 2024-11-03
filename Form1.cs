using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Aspose.Words;
using IronXL;
using IronWord;
using SixLabors.ImageSharp.Drawing;

namespace WordToExcelConverter
{
    //TODO: Найти причину увеличения времени обработки каждого файла в зависимости от их общего количества
    public partial class Form1 : System.Windows.Forms.Form
    {
        class Words
        {
            public string RusWord { get; set; }
            public string EngId { get; set; }
            public string EngWord { get; set; }
            public string EngTranscription { get; set; }
            public int SW { get; set; }
            public int DLM { get; set; }
            public string HeaderId { get; set; }
            public int HeaderLevel { get; set; }

            public Words(string rusWord, string engId, string engWord, string engTranscription, int sw, int dlm)
            {
                RusWord = rusWord;
                EngId = engId;
                EngWord = engWord;
                EngTranscription = engTranscription;
                SW = sw;
                DLM = dlm;
            }

            public Words(string rusWord, string engId, string engWord, string engTranscription, int sw, int dlm, string headerId, int headerLevel)
            {
                RusWord = rusWord;
                EngId = engId;
                EngWord = engWord;
                EngTranscription = engTranscription;
                SW = sw;
                DLM = dlm;
                HeaderId = headerId;
                HeaderLevel = headerLevel;
            }
        }

        /*Двумерный массив заголовков*/
        static string[,] headersArr = {
                { "<p", "</p>" },
                { "<h1", "</h1>" },
                { "<h2", "</h2>" },
                { "<h3", "</h3>" },
                { "<h4", "</h4>" },
                { "<h5", "</h5>" }
        };

        /*Глобальные переменные связи слов с заголовками*/
        //Массив идентификаторов тем (Topics)
        static string[] TIDs = new string[4];
        static int headerLevel = 0;

        /*Регулярные выражения для поиска в тексте*/
        static Regex regexRus = new Regex(@"[А-Яа-яёáéё́и́óýы́э́ю́я́А́Е́Ё́И́О́У́Ы́Э́Ю́Я́().]");
        static Regex regexEngId = new Regex(@"[A-Za-z0-9]");
        static Regex regexWordFile = new Regex(@"\\+(\w|\w)+.do(c|cx)");
        static Regex regexWordFileFolder = new Regex(@"(\s|\S)+.do(c|cx)");
        static Regex regexFirstWordFileToConvertionFolder = new Regex(@"(\s|\S)+_1.do(c|cx)");
        static Regex regexLevel = new Regex(@"\([0-9]{1}\)");

        /*Сведения о файлах и путях*/
        //Справочные константы
    //const string coreDir = "D:\\Task_3_Files\\";
    //const string defaultDocFileName = "Английский_Итог_067"; 
        //Магические числа
        const int PARAGRAPHSPERPAGE = 30;
        const int PAGES = 4;
        const int PARAGRAPHS = PARAGRAPHSPERPAGE * PAGES;

        /*Тэги*/
        List<int> listTagsSW = new List<int>();
        List<int> listTagsDLM = new List<int>();

        /*Глобальные счётчики*/
        int globalCounter = 0;

        //Инициализация главного окна
        public Form1()
        {
            //Базовая инициализация
            InitializeComponent();
            //Подсчитать и вывести количество подходящих файлов
            CountValidWordFiles();
        }

    /*Кнопки*/

        //Очистка консоли
        private void button4_Click(object sender, EventArgs e)
        {
            ClearConsole();
        }

        //Выбор пути исходного файла Word для разбиения
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.ShowDialog();
                if (!openFileDialog1.FileName.Trim().Equals(""))
                {
                    CheckWordFilePath(openFileDialog1.FileName);
                    textBox7.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        //Выбор пути к файлам Word
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.ShowDialog();
                if (!folderBrowserDialog1.SelectedPath.Equals(""))
                {
                    textBox1.Text = folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        //Выбор пути для генерации Excel
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog2.ShowDialog();
                if (!folderBrowserDialog2.SelectedPath.Equals(""))
                {
                    textBox4.Text = folderBrowserDialog2.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        //Разбиение файла Word
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                ClearConsole();
                GenerateLog("Начинаем разбиение");
                string messageBoxText;
                if (textBox1.Text.Trim().Equals(""))
                {
                    throw new Exception("Путь для итоговых файлов разбитого Word документа пуст");
                }
                else if (Directory.Exists(textBox1.Text.Trim()))
                {
                    messageBoxText = "Данный процесс удалит все файлы из папки:\r\n" + textBox1.Text.Trim() + "\r\nПродолжить?";
                }
                else 
                {
                    messageBoxText = "Данный процесс создаст новую папку и скопирует разбитые doc-файлы по пути:\r\n" + textBox1.Text.Trim() + "\r\nПродолжить?";
                }
                DialogResult result = MessageBox.Show(
                    messageBoxText,
                    "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1);
                switch (result) 
                {
                    case DialogResult.Yes:
                        if (textBox7.Text.Trim().Equals(""))
                        {
                            throw new Exception("Путь к исходному файлу Word для разбиения пуст");
                        }
                        else
                        {
                            DivideWordFileByUserCounter(textBox7.Text);
                            GenerateLog("Разбиение завершено");
                            CountValidWordFiles();
                        }
                        break;
                    case DialogResult.No:
                        throw new Exception("Операция разбиения файла Word отменена");
                    default:
                        throw new Exception("Некорректный ввод операции подтверждения. Операция разбиения файла Word отменена");
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
                GenerateLog("Разбиение прервано");
            }
        }

        //Запуск процесса конвертации
        //TODO: Переписать логику обработку наименований doc-файлов для конвертации
        //TODO: Переписать логику обращений к итоговым путям конвертации
        private void button2_Click(object sender, EventArgs e)
        {
            int textBoxWordFilesStart;
            int textBoxWordFilesCounter;
            int ExcelFileRowsLength;

            string textHtmlTemp;
            string defaultDocFileName;

            WorkBook workBook;
            WorkSheet workSheet;

            /*Переменные хранящие результат для записи в Excel-файл*/
            //RusWord, EngId, EngWord, EngTranscription, SW, DLM
            List<Words> listWords = new List<Words>();
            //0 - текст ассоциации, 1 - текст тэга #$D, 2 - текст тэга (*)
            List<string[]> listAssociations = new List<string[]>();

            try
            {
                ClearConsole();
                /*Создание и проверка папок для файлов*/
                string pathL1_A = textBox4.Text + "\\L1_A";
                CheckOrGenerateFilePath(pathL1_A);
                string pathL1_I = textBox4.Text + "\\L1_I";
                CheckOrGenerateFilePath(pathL1_I);
                string pathL2_T = textBox4.Text + "\\L2_T";
                CheckOrGenerateFilePath(pathL2_T);
                //Проверка пути на наличие doc-файлов
                defaultDocFileName = CheckWordFolderPath(textBox1.Text);
                //TODO: Сделать дополнительную проверку на данное поле
                textBoxWordFilesStart = Convert.ToInt32(textBox6.Text.Trim());
                textBoxWordFilesCounter = Convert.ToInt32(textBox2.Text.Trim());
                if (textBoxWordFilesCounter > 0 && textBoxWordFilesCounter <= Convert.ToInt32(textBox5.Text.Trim()))
                {
                    //TODO: Сделать генерацию через каждые 5 файлов
                    //Генерация стартового шаблона Excel
                    GenerateExcelWords(textBox4.Text, defaultDocFileName);

                    for (int i = textBoxWordFilesStart; i < textBoxWordFilesStart + textBoxWordFilesCounter; i++)
                    {
                        workBook = WorkBook.Load(textBox4.Text + "\\" + defaultDocFileName + "_Words.xlsx");
                        workSheet = workBook.GetWorkSheet("Words");

                        GenerateLog($"Файл №{i}: Конвертация начата!");
                        //Получение текста html-файла, сконвертированного из Word
                        CheckOrGenerateFilePath(textBox4.Text + "\\result\\html");
                        textHtmlTemp = WordToHtmlConversion(textBox4.Text, defaultDocFileName, i);
                        //                        GenerateLog("Файл №" + i + ": Конвертация Word-файла в формат html прошла успешно!");
                        //Получение русского слова, англ.id, англ.транскрипции
                        listWords = GetWordsList(textHtmlTemp);
                        //                        GenerateLog($"Файл №{i}: Конвертация слов и английской транскрипции успешно завершена!");
                        //Получение ассоциаций
                        listAssociations = GetAssociationsList(textHtmlTemp);
                        //Запись в Excel
                        if (listWords.Count == listAssociations.Count)
                        {
                            ExcelFileRowsLength = workSheet.Rows.Length;
                            for (int j = 0; j < listWords.Count; j++)
                            {
                                //LID
                                workSheet[$"B{ExcelFileRowsLength + j + 1}"].Value = listWords[j].EngId;
                                //TID
                                workSheet[$"C{ExcelFileRowsLength + j + 1}"].Value = listWords[j].HeaderId;
                                //DLID
                                workSheet[$"D{ExcelFileRowsLength + j + 1}"].Value = listWords[j].HeaderLevel;
                                ////L1
                                workSheet[$"F{ExcelFileRowsLength + j + 1}"].Value = listWords[j].RusWord;
                                //L1_D
                                workSheet[$"J{ExcelFileRowsLength + j + 1}"].Value = listAssociations[j][1];
                                //L1_I
                                workSheet[$"K{ExcelFileRowsLength + j + 1}"].Value = listAssociations[j][2];
                                File.WriteAllText(pathL1_I + "\\" + listWords[j].EngId + ".txt", listAssociations[j][2]);
                                //L2
                                workSheet[$"O{ExcelFileRowsLength + j + 1}"].Value = listWords[j].EngWord;
                                //L2_T
                                //workSheet[$"R{ExcelFileRowsLength + j + 1}"].Value = listWords[j].EngTranscription;
                                //File.WriteAllText(pathL2_T + "\\" + listWords[j].EngId + ".txt", listWords[j].EngTranscription);
                                //L1_A
                                workSheet[$"X{ExcelFileRowsLength + j + 1}"].Value = listAssociations[j][0];
                                File.WriteAllText(pathL1_A + "\\" + listWords[j].EngId + ".html", listAssociations[j][0]);
                                //SW
                                workSheet[$"AD{ExcelFileRowsLength + j + 1}"].Value = listTagsSW[j];
                                //DLM
                                workSheet[$"AE{ExcelFileRowsLength + j + 1}"].Value = listTagsDLM[j];
                            }
                        }
                        else
                            throw new Exception($"Файл №{i}: Ошибка! Обнаружено несовпадение количества слов и ассоциаций");
                        //Обнуление переменных
                        listTagsSW.Clear();
                        listTagsDLM.Clear();
                        //Сохранение файла Excel
                        workBook.SaveAs(textBox4.Text + "\\" + defaultDocFileName + "_Words.xlsx");
                        workBook.Close();
                    }
                    GenerateLog("Конвертация завершена!");
                }
                else
                    throw new Exception("Указано некорректное количество Word-документов для конвертации");
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

    /*Изменение полей*/    
        
        //Изменение строки, содержащей путь к файлам Word
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            CountValidWordFiles();
        }

        //Изменение счетчика количества потенциальных файлов doc для конвертации
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text.Equals("") || Int32.Parse(textBox5.Text) < 1)
            {
                DisableConversionElements();
            }
            else
            {
                EnableConversionElements();
            }
        }

    /*Файловая система*/

        //Проверка пути на наличие и генерация в случае отсутствия
        private void CheckOrGenerateFilePath(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

    /*Генерация XML-файлов для разбиения*/

        //Разбиение исходного файла Word через заданное число страниц
        private void DivideWordFileByUserCounter(string wordFilePath)
        {
            //Создание директории функции по умолчанию
            string defPath = textBox1.Text;
            Directory.CreateDirectory(defPath);
            //Удаление всех файлов из результирующей директории
            Directory.Delete(defPath, true);
            //Получение наименования исходного Word-файла
            string fileName = GetWordFileNameFromPath(wordFilePath);
            //Создание директорий для вспомогательных файлов
            string zipFilesPath = defPath + "\\ZipFiles";
            string xmlFilesPath = defPath + "\\XmlFiles";
            Directory.CreateDirectory(zipFilesPath);
            Directory.CreateDirectory(xmlFilesPath);
            //Перемещение doc-файла в директорию ZipFiles в формате .zip
            string zipPath = zipFilesPath + "\\" + fileName + ".zip";
            if (System.IO.File.Exists(zipPath))
                System.IO.File.Delete(zipPath);
            System.IO.File.Copy(wordFilePath, zipPath);
            //Перемещение xml-файлов в обозначенные директории
            string xmlFilePath = xmlFilesPath + "\\";
            using (var zipArchive = ZipFile.Open(zipPath, ZipArchiveMode.Read))
            {
                int totalEntries = zipArchive.Entries.Count;
                foreach (var e in zipArchive.Entries)
                {
                    if (e.FullName.Contains("word/document.xml"))
                    {
                        xmlFilePath += e.Name;
                        if (System.IO.File.Exists(xmlFilePath))
                            System.IO.File.Delete(xmlFilePath);
                        e.ExtractToFile(xmlFilePath);
                    }
                }
            }
            //Получение абзацев
            GenerateLog("Получение абзацев");
            List<string> paragraphsList = GetParagraphsFromXml(xmlFilePath);
            EraseLastLog();
            //Подсчет количества файлов
            int filesCounter = paragraphsList.Count / PARAGRAPHS + 1;
            //Создание XML-файлов разбитого документа
            GenerateLog("Генерация XML-файлов разбитого документа");
            GenerateXmlByDividedWordFile(paragraphsList, xmlFilesPath, filesCounter);
            EraseLastLog();
            //Создание zip-файла обрезанного документа
            GenerateLog("Процесс переноса разметки XML в обрезанные doc-файлы");
            string editedZipFileName;
            for (int i = 0; i < filesCounter; i++)
            {
                GenerateLog("Генерация: " + (i + 1) + "-го doc-файла из: " + filesCounter);
                editedZipFileName = zipPath.Replace(fileName, fileName + "_" + (i + 1));
                System.IO.File.Copy(zipPath, editedZipFileName);
                using (var zipArchive = ZipFile.Open(editedZipFileName, ZipArchiveMode.Update))
                {
                    zipArchive.GetEntry("word/document.xml").Delete();
                    zipArchive.CreateEntryFromFile(xmlFilesPath + "\\paragraphs_" + (i + 1) + ".xml", "word\\document.xml");
                }
                System.IO.File.Copy(editedZipFileName, defPath + "\\" + fileName + "_" + (i + 1) + ".docx");
                EraseLastLog();
            }
            EraseLastLog();
            Directory.Delete(defPath + "\\XmlFiles", true);
            Directory.Delete(defPath + "\\ZipFiles", true);
        }

        //Получение абзацев
        public List<string> GetParagraphsFromXml(string path) 
        {
            //Получаем абзацы и добавляем в конце каждого абзаца закрывающий тэг абзаца в XML
            List<string> paragraphs = File.ReadAllText(path)
                .Split(new string[] { "</w:p>" }, StringSplitOptions.None)
                .Select(p => p += "</w:p>").ToList();
            //Убираем тэг абзаца у самого последнего абзаца, т.к. он содержит закрывающий тэг файла в XML
            paragraphs[paragraphs.Count - 1] = paragraphs[paragraphs.Count - 1]
                .Substring(0, paragraphs[paragraphs.Count - 1].Length - "</w:p>".Length);
            //Производим разделение первого абзаца на служебный абзац стилей и абзац текста
            int openBodyTagIndex = paragraphs[0].LastIndexOf("<w:body>") + "<w:body>".Length;
            paragraphs.Insert(0, "");
            paragraphs[0] = paragraphs[1].Substring(0, openBodyTagIndex);
            paragraphs[1] = paragraphs[1].Substring(openBodyTagIndex);

            return paragraphs;
        }

        //Генерация XML-файлов с обрезанными данными из документа
        private void GenerateXmlByDividedWordFile(List<string> paragraphs, string xmlFilesPath, int filesCounter) 
        {
            //Получение открывающего и закрывающего абзаца в качестве шаблона для всех разбитых файлов
            string beginParagraph = paragraphs[0];
            string endParagraph = paragraphs[paragraphs.Count - 1];

            int lastFileParagraphsCounter = paragraphs.Count % PARAGRAPHS - 2;
            string paragraphFileName;

            for (int i = 0; i < filesCounter; i++)
            {
                //Создание одного из файлов, содержащего обрезанную часть большого документа с открывающим и закрывающим абзацем
                paragraphFileName = xmlFilesPath + "\\paragraphs_" + (i + 1) + ".xml";
                if (System.IO.File.Exists(paragraphFileName))
                    System.IO.File.Delete(paragraphFileName);

                GenerateLog("Генерация: " + (i + 1) + "-го XML-файла из: " + filesCounter);
                using (StreamWriter writer = new StreamWriter(paragraphFileName, false))
                {
                    writer.WriteLine(beginParagraph);
                    if (i == filesCounter - 1)
                        WriteInXmlFile(writer, paragraphs, i, lastFileParagraphsCounter);
                    else
                        WriteInXmlFile(writer, paragraphs, i, PARAGRAPHS);
                    writer.WriteLine(endParagraph);
                }
                EraseLastLog();
            }
        }

        //Запись в XML-файл
        private void WriteInXmlFile(StreamWriter writer, List<string> paragraphs, int fileCounter, int paragraphsPerFile)
        {
            for (int i = 1; i <= paragraphsPerFile; i++)
            {
                writer.WriteLine(paragraphs[fileCounter * PARAGRAPHS + i]);
            }
        }

        //Получение имени файла Word из пути
        private string GetWordFileNameFromPath(string path)
        {
            try
            {
                if (path.Contains(".doc"))
                {
                    int startIndex = path.LastIndexOf("\\") + 1;
                    return path.Substring(startIndex, path.LastIndexOf(".") - startIndex);
                }
                else
                {
                    throw new Exception("Ошибка! Некорректный формат файла для разбиения");
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
                return null;
            }
        }

        //Удаление уже использованных doc-файлов

    /*Функции для генерации Words-файла Excel*/

        //Генерация итогового Words-файла Excel
        private void GenerateExcelWords(string path, string defaultDocFileName)
        {
            //Вспомогательные переменные
            int counter = 0;
            string pathTemp = path + "\\" + defaultDocFileName + "_Words.xlsx";
            string[] fieldsExcel = new string[31] {
                "M",
                "LID", //2
                "TID",
                "DLID",
                "L_P",
                "L1", //6
                "L1_V",
                "L1_F",
                "L1_S",
                "L1_D", //10
                "L1_I", //11
                "L1_T",
                "L1_T2",
                "L1_R",
                "L2", //15
                "L2_V",
                "L2_F",
                "L2_T", //18
                "L2_T2",
                "L2_S",
                "L2_R",
                "L2_D",
                "L2_I",
                "L1_A", //24
                "L1_A_P",
                "L1_A_S",
                "L2_A",
                "L2_A_P",
                "L2_A_S",
                "SW", //30
                "DLM", //31
            };
            
            //pathTemp
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            workBook.CreateWorkSheet("Words");
            WorkSheet workSheet = workBook.GetWorkSheet("Words");

            foreach (var cell in workSheet["A1:AE1"])
            {
                cell.Text = fieldsExcel[counter];
                counter++;
            }

            workBook.SaveAs(pathTemp);
        }

        //TODO: Переписать для наименования путей для более прозрачной работы
        //Конвертация Word в HTML и на выходе получим строковую переменную, содержащую текст html-файла
        private string WordToHtmlConversion(string coreDir, string defaultDocFileName, int wordFileIndex)
        {
            string htmlWordsFilePath;
            string htmlWordsFilePathTemp;
            string docFileNameTemp;
            string textHtmlTemp;

            htmlWordsFilePath = coreDir + "\\result\\html\\" + defaultDocFileName + "_" + wordFileIndex + ".html";
                
            //Проверка существующих файлов html
            if (File.Exists(htmlWordsFilePath))
                File.Delete(htmlWordsFilePath);

            /*Генерация стартового шаблона HTML-файла*/
//        Console.WriteLine("\r\nФайл №" + (wordFileIndex));
            //Генерация тегов в начале HTML-файла Words
            GenerateHeadTagsHtml(htmlWordsFilePath);
            //Проверка на чтение первого файла для добавления названия столбцов в HTML - файл
            if (wordFileIndex == 1)
                GenerateDefaultTableHtml(htmlWordsFilePath);
            
            //Получение пути к doc-файлу для генерации
            docFileNameTemp = textBox1.Text + "\\" + defaultDocFileName + "_" + wordFileIndex + ".docx";

            /*Генерация временного HTML-файла на основе DOCX-файла Words*/
            Document doc = new Document(docFileNameTemp);
            htmlWordsFilePathTemp = coreDir + "\\result\\html\\" + defaultDocFileName + "_" + wordFileIndex + ".html";
            doc.Save(htmlWordsFilePathTemp, Aspose.Words.SaveFormat.Html);

            /*Чтение HTML-файла, сконвертированного из doc-файла*/
//        Console.WriteLine("Чтение HTML-файла\r\n");
            textHtmlTemp = GetText(htmlWordsFilePathTemp);

//            GenerateLog("Файл №" + wordFileIndex + ": Конвертация Word-файла в формат html прошла успешно!");
            
            return textHtmlTemp;
        }

        //Получение списка ассоциаций
        private List<string[]> GetAssociationsList(string text)
        {
            int tagsIndexCounter = 0;
            int substrIndex;
            //Вспомогательные строковые переменные
            string textTemp;
            string stringAssoc;
            string stringAssocHtml;
            string[] spanSplitted;
            //Ассоциации
            string L1_I = "";
            string L1_D = "";
            //Итоговая возвращаемая переменная
            List<string[]> assocListResult = new List<string[]>();

            try
            {
                do
                {
                    //Обнуляем служебную переменную, в которую записывается ассоциация
                    stringAssoc = "";
                    stringAssocHtml = "";

                    //Обрезаем весь текст начиная с тега </h5>, чтобы получить ассоциации сразу после слов
                    if (text.Contains(headersArr[5, 1]))
                        text = text.Substring(text.IndexOf(headersArr[5, 1]) + headersArr[5, 1].Length);
                    else
                        break;
                    textTemp = text;

                //Получение html ассоциации
                    if (textTemp.Contains(headersArr[0, 0]) && textTemp.Contains(headersArr[0, 1]))
                    {
                        stringAssocHtml = textTemp.Substring(textTemp.IndexOf(headersArr[0, 0]));
                        stringAssocHtml = textTemp.Substring(0, textTemp.IndexOf(headersArr[0, 1]) + headersArr[0, 1].Length);
                    }
                //Получение html ассоциации

                    //Обрезаем текст между тэгами <p> и </p> после строки слов
                    if (textTemp.Contains(headersArr[0, 0]) && textTemp.Contains(headersArr[0, 1]))
                        textTemp = textTemp.Substring(textTemp.IndexOf(headersArr[0, 0]), textTemp.IndexOf(headersArr[0, 1]) - textTemp.IndexOf(headersArr[0, 0]));
                    else
                        break;

                    //Получаем строку c ассоциацией
                    spanSplitted = textTemp.Split(new string[] { "</span>" }, StringSplitOptions.None);

                    foreach (string span in spanSplitted)
                    {
                        if (span.LastIndexOf(">") != -1)
                        {
                            stringAssoc += span.Substring(span.LastIndexOf(">") + 1);
                        }
                    }

                    /*Удаление посторонних символов Ч.1*/
                    stringAssoc = stringAssoc.Replace("&#xa0;", "");

                    /*Условия для пустой записи, если ассоциация пропущена в файле Word*/
                    if (stringAssoc.Trim().StartsWith("+") || stringAssoc.Trim().StartsWith("-") || stringAssoc.Trim().StartsWith("?") || stringAssoc.Trim().Equals(""))
                    {
                        assocListResult.Add(new string[] { "+", "", "" });
                        continue;
                    }

                    /*Проверка на тэги*/
                    //SW
                    if (textTemp.Contains("#$SW"))
                        listTagsSW[tagsIndexCounter] = 1;
                    //DLM
                    if (textTemp.Contains("#@1"))
                        listTagsDLM[tagsIndexCounter] = 1;
                    //L1_D
                    if (stringAssoc.Contains("#$D"))
                    {
                        L1_D = stringAssoc.Substring(stringAssoc.IndexOf("#$D") + 3);
                        stringAssoc = stringAssoc.Substring(0, stringAssoc.IndexOf("#$D")).Trim();
                    }
                    else
                        L1_D = "";
                    //L1_I
                    if (stringAssoc.Contains("(*)"))
                    {
                        L1_I = stringAssoc.Substring(stringAssoc.IndexOf("(*)") + 3);
                        stringAssoc = stringAssoc.Substring(0, stringAssoc.IndexOf("(*)")).Trim();
                    }
                    else
                        L1_I = "";

                    /*Удаление посторонних символов Ч.2*/
                    stringAssoc = stringAssoc.Replace("*", "").Replace("#$SW", "").Replace("#@1", "").Replace("=","").Trim();

                    //Добавляем ассоциацию в итоговый лист
                    assocListResult.Add(new string[] { stringAssocHtml, L1_D, L1_I});

                    //Обрезание исходного текста
                    substrIndex = text.IndexOf(headersArr[0, 1]);
                    if (substrIndex == -1)
                        break;
                    else
                        text = text.Substring(substrIndex + headersArr[5, 1].Length);

                    //Увеличиваем число идентификатора для вставки в лист тэгов
                    tagsIndexCounter++;

                } while (true);
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }

            return assocListResult;
        }

        //Получение списка слов
        private List<Words> GetWordsList(string text)
        {
        
        /*Общее*/
            bool wordsEndFlag = false;
            bool topicsEndFlag = false;

        /*Words*/
            //Счётчик слов
            int counter = 0;
            //Индекс символа с которого необходимо обрезать исходный текст
            int substrIndex;
            //Вспомогательные строковые переменные
            string textTemp = "";
            string stringWords;
            string[] stringWordsResult = new string[4];
            string[] spanSplitted;
            //Строчная переменная для сбора вывода в консоль
            string outResultString = "";
            //Итоговая возвращаемая переменная
            List<Words> listWordsResult = new List<Words>();

        /*Topics*/
            //Строковые переменные текстов файлов
            string htmlHeaderTextTemp;
            string headerId = "";
            string headerLevelString;
            //Численные переменные текстов файлов
            int countHeader = 1;
            int minHeaderIndex = int.MaxValue;
            int minHeaderId = int.MaxValue;
            //Вспомогательные массивы индексов
            int[] headersIndexCount = new int[6] { int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue };

            try
            {
                do
                {
                /*Чтение и сбор информации Words*/

                    //Обнуляем служебную переменную, в которую записываются искомы строки в формате: рус.слово - анг.id [траскрипция] - (опционально)/рус.транскрипция/
                    stringWords = "";

                    //Добавляем дефолтный элемент в лист тэгов
                    listTagsSW.Add(0);
                    listTagsDLM.Add(2);

                    //Обрезаем весь текст до тега <h5
                    if (text.Contains(headersArr[5, 0]))
                        textTemp = text.Substring(text.IndexOf(headersArr[5, 0]));
                    else
                        wordsEndFlag = true;   

                    //Получаем минимальный индекс заголовка
                    headersIndexCount = GetHeadersFirstIndexes(text);
                    minHeaderIndex = headersIndexCount.Min(); //значение

                    if (wordsEndFlag && topicsEndFlag)
                        break;
                    else if (!wordsEndFlag && (text.IndexOf(textTemp) < minHeaderIndex || topicsEndFlag))
                    {
                        //Обрезаем весь текст до </h5>
                        textTemp = textTemp.Substring(0, textTemp.IndexOf(headersArr[5, 1]));

                        if (textTemp.Contains("#$SW"))
                            listTagsSW[listTagsSW.Count - 1] = 1;

                        if (textTemp.Contains("#@1"))
                            listTagsDLM[listTagsDLM.Count - 1] = 1;

                        //Получаем строку, в формате: рус.слово - анг.id [траскрипция] - (опционально)/рус.транскрипция/
                        spanSplitted = textTemp.Split(new string[] { "</span>" }, StringSplitOptions.None);
                        foreach (string span in spanSplitted)
                        {
                            if (span.LastIndexOf(">") != -1)
                            {
                                stringWords += span.Substring(span.LastIndexOf(">") + 1);
                            }
                        }
                        counter++;
                        globalCounter++;

                        //Важно! Запуск процесса разбиения и получения слов
                        stringWordsResult = GetSplitedStringWords(stringWords);

                        /*Вывод логов в консоль*/
                        outResultString += $"\r\nCтрока со словами №{counter}: {stringWords}\r\n";
                        //GenerateSubLog($"\r\nCтрока со словами №{counter}: {stringWords}\r\n");
                        for (int i = 0; i < stringWordsResult.Length; i++)
                        {
                            stringWordsResult[i] = stringWordsResult[i].Replace("xa0", "").Trim();
                            switch (i)
                            {
                                case 0:
                                    outResultString += $"Русское слово: |{stringWordsResult[0]}|\r\n";
                                    break;
                                case 1:
                                    outResultString += $"Английский идентификатор: |{stringWordsResult[1]}|\r\n";
                                    break;
                                case 2:
                                    outResultString += $"Английская транскрипция: |{stringWordsResult[2]}|\r\n";
                                    break;
                                default:
                                    break;
                            }
                        }

                        /*Контрольная запись в возвращаемую переменную*/
                        listWordsResult.Add(new Words(stringWordsResult[0], stringWordsResult[1], ConvertEngIdToWord(stringWordsResult[1]), stringWordsResult[2], listTagsSW[listTagsSW.Count - 1], listTagsDLM[listTagsDLM.Count - 1], GetHeaderId(TIDs), headerLevel));

                        /*TODO: Добавить логику для обрезания текста после проверки на Words и топики*/
                        /*Обрезание текста*/
                        substrIndex = text.IndexOf(headersArr[5, 1]);
                        if (substrIndex == -1)
                            break;
                        else
                            text = text.Substring(substrIndex + headersArr[5, 1].Length);
                    }
                    else
                    {
                        minHeaderId = GetIdOfMinInt(headersIndexCount); //ключ
                        if (minHeaderIndex == int.MaxValue)
                        {
                            topicsEndFlag = true;
                            continue;
                        }
                        else
                        {
                            htmlHeaderTextTemp = GetHeaderText(text, minHeaderIndex, minHeaderId).Replace("&#xa0;", "");
                            headerId = GetIdFromHeaderText(htmlHeaderTextTemp).Trim();

                            if (htmlHeaderTextTemp == "")
                            {
                                if (text.Contains(headersArr[minHeaderId, 1]))
                                {
                                    if (minHeaderId == 0)
                                        text = text.Substring(minHeaderIndex);
                                    text = text.Substring(text.IndexOf(headersArr[minHeaderId, 1]) + headersArr[minHeaderId, 1].Length);
                                }
                                else
                                {
                                    topicsEndFlag = true;
                                }
                                continue;
                            }
                            
                            htmlHeaderTextTemp = htmlHeaderTextTemp.Substring(0, htmlHeaderTextTemp.IndexOf("#")).Trim();
                            GenerateLog($"{countHeader}. {htmlHeaderTextTemp} | #{headerId}");

                            if (minHeaderId > 0)
                                TIDs[minHeaderId - 1] = headerId;
                            if (minHeaderId == 4)
                            {
                                headerLevelString = regexLevel.Match(htmlHeaderTextTemp).Value;
                                headerLevel = Convert.ToInt32(headerLevelString.Replace("(", "").Replace(")", ""));
                            }

                            // Увеличение счётчика заголовков
                            countHeader++;

                            //Обрезаем исходный текст
                            if (text.Contains(headersArr[minHeaderId, 1]))
                            {
                                if (minHeaderId == 0)
                                    text = text.Substring(minHeaderIndex);
                                text = text.Substring(text.IndexOf(headersArr[minHeaderId, 1]) + headersArr[minHeaderId, 1].Length);
                            }
                            else
                                topicsEndFlag = true;
                        }
                    }

                } while (true);
            }
            catch (IndexOutOfRangeException)
            {
                GenerateLog("Ошибка! Индекс вне границ массива!");
            }
            catch (Exception ex)
            {
                GenerateLog($"Ошибка в слове \"{stringWordsResult[0]}\": {ex.Message}");
            }

            GenerateLog(outResultString);

            return listWordsResult;
        }

        //Производит конвертацию английского слова из идентификатора в слово
        private string ConvertEngIdToWord(string engId)
        {
            string engWord = "";

            Regex regexEng = new Regex(@"[A-Za-z -]");

            foreach (var match in regexEng.Matches(engId))
            {
                engWord += match.ToString();
            }

            return engWord;
        }

        //Производит разбиение и получает слово на русском, англ. id и англ. транскрипцию
        //TODO: Переписать цикл под нормы написания кода
        private string[] GetSplitedStringWords(string text)
        {
            string[] result = new string[3] { "", "", "" };
            string[] textStrings = text.ToCharArray().Select(c => c.ToString()).ToArray();
            
            int state = 0;

            for (int i = 0; i < textStrings.Length; i++)
            {
                if (textStrings[i].Equals(" "))
                    result[state] += textStrings[i];
                else if (textStrings[i].Equals("–") || textStrings[i].Equals("-"))
                {
                    if (state == 0 && regexRus.IsMatch(textStrings[i + 1]))
                        result[0] += textStrings[i];
                    else if (state == 1 && regexEngId.IsMatch(textStrings[i + 1]))
                        result[1] += textStrings[i];
                    else if (state < 2)
                        state++;
                }
                else if (state == 0 && regexRus.IsMatch(textStrings[i]))
                    result[0] += textStrings[i];
                else if (state == 1 && regexEngId.IsMatch(textStrings[i]))
                    result[1] += textStrings[i];
                else if (textStrings[i].Contains("["))
                {
                    state = 2;
                    result[2] += textStrings[i];
                }
                else if (textStrings[i].Contains("]"))
                {
                    result[2] += textStrings[i];
                    break;
                }
                else if (state == 2)
                    result[2] += textStrings[i];
            }

            return result;
        }

        //Генерация стартовых тегов html-файла Words
        private void GenerateHeadTagsHtml(string path)
        {
            string textHTML =
                "<html>" +
                    "<head>" +
                        "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                        "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                        "<meta name=\"generator\" content=\"Aspose.Words for .NET 24.1.0\" />" +
                        "<title></title>" +
                    "</head>" +
                    "<body>" +
                        "<table>";

            File.AppendAllText(path, textHTML);
        }

        //Генерация таблицы HTML по умолчанию для дальнейшего конвертирования в Excel
        static void GenerateDefaultTableHtml(string path)
        {
            string textHTML =
                "<tr>" +
                    "<td>M</td>" +
                    "<td>LID</td>" +
                    "<td>TID</td>" +
                    "<td>DLID</td>" +
                    "<td>DLM</td>" +
                    "<td>L_P</td>" +
                    "<td>L1</td>" +
                    "<td>L1_V</td>" +
                    "<td>L1_F</td>" +
                    "<td>L1_S</td>" +
                    "<td>L1_D</td>" +
                    "<td>L1_I</td>" +
                    "<td>L1_T</td>" +
                    "<td>L1_T2</td>" +
                    "<td>L1_R</td>" +
                    "<td>L2</td>" +
                    "<td>L2_V</td>" +
                    "<td>L2_F</td>" +
                    "<td>L2_T</td>" +
                    "<td>L2_T2</td>" +
                    "<td>L2_S</td>" +
                    "<td>L2_R</td>" +
                    "<td>L2_D</td>" +
                    "<td>L2_I</td>" +
                    "<td>L1_A</td>" +
                    "<td>L1_A_P</td>" +
                    "<td>L1_A_S</td>" +
                    "<td>L2_A</td>" +
                    "<td>L2_A_P</td>" +
                    "<td>L2_A_S</td>" +
                    "<td>SW</td>" +
                "</tr>";

            File.AppendAllText(path, textHTML);
        }

        //Получение текста HTML-файла, сконвертированного из doc-файла
        static string GetText(string path)
        {
            StreamReader streamReader = new StreamReader(path, System.Text.Encoding.UTF8);
            string text = streamReader.ReadToEnd();
            streamReader.Close();

            return text;
        }

        //Проверка выбранный путь на соответствие формату .doc/.docx
        private void CheckWordFilePath(string path)
        {
            if (!File.Exists(path))
                throw new Exception("Ошибка! Указан некорректный путь к файлу Word!");
            else if (!regexWordFile.IsMatch(path))
                throw new Exception("Ошибка! Выбранный файл не соответствует форматам: *.doc, *.docx");
        }

        //Проверка выбранного пути на наличие файлов _###.doc/.docx,
        //используемых для генерации Excel-таблиц и возвращение наименования файлов для конвертации
        private string CheckWordFolderPath(string path)
        {
            string wordFileName = "";

            string[] files = Directory.GetFiles(path);

            if (!Directory.Exists(path))
            {
                throw new Exception("Ошибка! Указан некорректный путь к файлам Word!");
            }
            else if (!files.Select(f => regexWordFileFolder.IsMatch(f)).Any())
            {
                CountValidWordFiles();
                throw new Exception("Ошибка! В папке не найдено файлов форматов: *.doc, *.docx");
            }
            else if (!files.Select(f => regexFirstWordFileToConvertionFolder.IsMatch(f)).First())
            {
                throw new Exception("Ошибка! В папке не содержится первого файла среди файлов, прошедших операцию разделения.");
            }
            else
            {
                wordFileName = GetWordFileNameFromPath(files.FirstOrDefault(f => regexFirstWordFileToConvertionFolder.IsMatch(f)));
                wordFileName = wordFileName.Substring(0, wordFileName.LastIndexOf("_"));
            }

            return wordFileName;
        }

        private void CountValidWordFiles()
        {
            if (Directory.Exists(textBox1.Text))
            {
                textBox5.Text = Convert.ToString(Directory.GetFiles(textBox1.Text).Select(f => regexWordFileFolder.IsMatch(f)).Count());
            }
            else
            {
                textBox5.Text = "";
            }
        }
        
    /*Функции взаимодействия с главным окном*/

        //Генерация логов в консоль
        private void GenerateLog(string message)
        {
            textBox3.Text = "========================================================\r\n\r\n" + message + "\r\n\r\n========================================================\r\n\r\n" + textBox3.Text;
            Application.DoEvents();
        }

        //Убрать последний лог из консоли
        private void EraseLastLog()
        {
            string consoleText = textBox3.Text;
            string logSeparator = "========================================================\r\n\r\n";
            consoleText = consoleText.Substring(consoleText.IndexOf(logSeparator) + logSeparator.Length);
            consoleText = consoleText.Substring(consoleText.IndexOf(logSeparator) + logSeparator.Length);

            textBox3.Text = consoleText;
            Application.DoEvents();
        }

        //Чистка логов в консоли
        private void ClearConsole()
        {
            textBox3.Clear();
            textBox3.ClearUndo();
        }

        //Включает элементы, которые необходимы для запуска процесса конвертации
        private void EnableConversionElements()
        {
            button2.Enabled = true;
            button3.Enabled = true;
            textBox4.Enabled = true;
        }

        //Отключает элементы, которые необходимы для запуска процесса конвертации
        private void DisableConversionElements()
        {
            button2.Enabled = false;
            button3.Enabled = false;
            textBox4.Enabled = false;
        }

    /*Функции для генерации Topic-файла Excel*/

        static int[] GetHeadersFirstIndexes(string text)
        {
            int indexTemp = 0;
            int[] headersIndexCount = new int[5] { int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue };
            
            for (int i = 0; i < headersIndexCount.Length; i++)
            {
                if (i == 0)
                    indexTemp = GetHeaderPTagIndex(text);
                else 
                    indexTemp = text.IndexOf(headersArr[i, 0]);
                
                if (indexTemp != -1)
                    headersIndexCount[i] = indexTemp;
            }

            return headersIndexCount;
        }

        static int GetHeaderPTagIndex(string text)
        {
            string textTemp = text;
            string textHeaderTemp;

            Regex regexHeaderP = new Regex(@"(\s|\S)+ #([A-Z0-9])+");

            do
            {
                //Обрезаем весь текст до тега <p
                if (textTemp.Contains(headersArr[0, 0]) && textTemp.Contains(headersArr[0, 1]))
                    textHeaderTemp = textTemp.Substring(textTemp.IndexOf(headersArr[0, 0]));
                else
                    return -1;

                //Обрезаем весь текст до </p>
                textHeaderTemp = textHeaderTemp.Substring(0, textHeaderTemp.IndexOf(headersArr[0, 1]) + headersArr[0, 1].Length);

                //Получаем текст фразы
                textHeaderTemp = textHeaderTemp.Substring(0, textHeaderTemp.LastIndexOf("</span>"));
                textHeaderTemp = textHeaderTemp.Substring(textHeaderTemp.LastIndexOf(">"));

                //TODO: Исправить проверку на header
                if (regexHeaderP.IsMatch(textHeaderTemp))
                    return text.IndexOf(textHeaderTemp);
                else
                    textTemp = textTemp.Substring(textTemp.IndexOf(headersArr[0, 1]) + headersArr[0, 1].Length);

            } while (true);
        }

        static int GetIdOfMinInt(int[] headersIndexCount)
        {
            int minInt = headersIndexCount.Min();
            for (int i = 0; i < headersIndexCount.Length; i++)
            {
                if (headersIndexCount[i] == minInt)
                    return i;
            }

            return 0;
        }

        static string GetHeaderText(string htmlText, int substrIndex, int headerIndex)
        {
            string htmlTextResult = "";

            string htmlTextTemp = htmlText.Substring(substrIndex);
            htmlTextTemp = htmlTextTemp.Substring(0, htmlTextTemp.IndexOf(headersArr[headerIndex, 1]));
            string[] htmlTextSplitted = htmlTextTemp.Split(new string[] { "</span>" }, StringSplitOptions.None);
            foreach (string hTS in htmlTextSplitted)
            {
                if (hTS.LastIndexOf(">") != -1)
                {
                    htmlTextResult += hTS.Substring(hTS.LastIndexOf(">") + 1);
                }
            }

            return htmlTextResult.Trim();
        }

        static string GetIdFromHeaderText(string headerText)
        {
            return headerText.Substring(headerText.IndexOf("#") + 1);
        }

        static string GetHeaderId(string[] TIDs)
        {
            string headerId = "";

            for (int i = 0; i < TIDs.Length; i++)
            {
                headerId += TIDs[i];
            }

            return headerId;
        }
    }
}
