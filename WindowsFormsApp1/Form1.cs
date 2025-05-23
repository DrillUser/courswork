using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        // Приватное поле для хранения авторов и их дисциплин в виде словаря, где ключ - автор, значение - множество дисциплин
        private readonly Dictionary<string, HashSet<string>> authorDisciplineCount = new Dictionary<string, HashSet<string>>();

        // Приватное поле для хранения количества публикаций по дисциплинам в виде словаря, где ключ - дисциплина, значение - количество
        private readonly Dictionary<string, int> disciplineCount = new Dictionary<string, int>();

        // Конструктор формы Form1
        public Form1()
        {
            // Инициализация компонентов формы
            InitializeComponent();
            // Установка позиции формы при открытии - по центру экрана
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        // Обработчик события клика по кнопке button1
        private void button1_Click(object sender, EventArgs e)
        {
            // Создание диалогового окна для выбора файлов
            using (var openFileDialog = new OpenFileDialog())
            {
                // Установка фильтра для отображения только Word-документов
                openFileDialog.Filter = "Word Documents|*.doc;*.docx";
                // Разрешение выбора нескольких файлов
                openFileDialog.Multiselect = true;

                // Проверка, что пользователь выбрал файлы и нажал OK
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Перебор всех выбранных файлов
                    foreach (string filePath in openFileDialog.FileNames)
                    {
                        // Обработка каждого Word-файла
                        ProcessWordFile(filePath);
                    }
                }
            }
        }

        // Метод для обработки Word-файла
        private void ProcessWordFile(string filePath)
        {
            // Установка модели многопоточности для работы с COM-объектами (Single Thread Apartment)
            System.Threading.Thread.CurrentThread.ApartmentState = ApartmentState.STA;

            // Объявление переменных для приложения Word и документа
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document doc = null;

            try
            {
                // Создание экземпляра Word Application со следующими параметрами:
                wordApp = new Microsoft.Office.Interop.Word.Application()
                {
                    // Скрытый режим (без отображения окна Word)
                    Visible = false,
                    // Отключение системных предупреждений
                    DisplayAlerts = WdAlertLevel.wdAlertsNone,
                    // Уровень безопасности (низкий для автоматизации)
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow
                };

                // Открытие документа Word с параметрами:
                doc = wordApp.Documents.Open(
                    // Путь к файлу
                    FileName: filePath,
                    // Без подтверждения конвертации
                    ConfirmConversions: false,
                    // Только для чтения
                    ReadOnly: true,
                    // Не добавлять в список последних файлов
                    AddToRecentFiles: false
                );

                // Пауза для завершения инициализации документа (300 мс)
                System.Threading.Thread.Sleep(300);

                // Получение имени файла без расширения
                string fileName = Path.GetFileNameWithoutExtension(filePath);

                // Выбор метода обработки в зависимости от содержимого имени файла:
                if (fileName.Contains("РП АИУС")) ProcessRPAIUS(doc);       // Если имя содержит "РП АИУС"
                else if (fileName.Contains("РП ПиОА")) ProcessRPPIOA(doc);  // Если имя содержит "РП ПиОА"
                else if (fileName.Contains("РП ТСАУ")) ProcessRPTSAU(doc);  // Если имя содержит "РП ТСАУ"
                else ProcessDefault(doc);                                   // Для всех остальных случаев

                // Извлечение информации о дисциплинах и часах из файла
                ExtractDisciplineHours(filePath);
            }
            catch (Exception ex)
            {
                // Обработка исключений - вывод ошибки в консоль
                Console.WriteLine($"Ошибка обработки файла {filePath}: {ex.Message}");
            }
            finally
            {

                // Закрытие документа, если он был открыт
                if (doc != null)
                {
                    // Закрытие без сохранения изменений (false)
                    doc.Close(false);
                    // Освобождение COM-объекта
                    Marshal.ReleaseComObject(doc);
                }

                // Завершение работы приложения Word, если оно было запущено
                if (wordApp != null)
                {
                    // Выход из приложения Word
                    wordApp.Quit();
                    // Освобождение COM-объекта
                    Marshal.ReleaseComObject(wordApp);
                }
            }

            // Обновление диаграмм после обработки файла
            UpdateCharts();
        }

        // Обработка документа типа "РП АИУС"
        private void ProcessRPAIUS(Document doc)
        {
            // Получение имени автора из параграфов документа
            string author = GetAuthorFromParagraphs(doc);

            // Получение названия дисциплины из заголовка документа
            string discipline = GetDisciplineFromHeader(doc.Content);

            // Добавление автора и дисциплины в коллекцию
            AddAuthorDiscipline(author, discipline);

            // Увеличение счетчика публикаций по дисциплине
            AddDisciplineCount(discipline);
        }

        // Обработка документа типа "РП ПиОА"
        private void ProcessRPPIOA(Document doc)
        {
            // Проверка, что документ содержит хотя бы одну страницу
            if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) > 0)
            {
                // Получение диапазона текста первой страницы
                var firstPageRange = GetFirstPageRange(doc);

                // Получение и очистка имени автора из первого якорного текста
                string author = CleanAuthorName(GetFirstAnchorText(firstPageRange));

                // Получение и очистка названия дисциплины из второго якорного текста
                string discipline = CleanDisciplineName(GetSecondAnchorText(firstPageRange));

                // Добавление автора и дисциплины в коллекцию
                AddAuthorDiscipline(author, discipline);

                // Увеличение счетчика публикаций по дисциплине
                AddDisciplineCount(discipline);
            }
        }

        // Обработка документа типа "РП ТСАУ"
        private void ProcessRPTSAU(Document doc)
        {
            // Проверка, что документ содержит хотя бы одну страницу
            if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) > 0)
            {
                // Получение диапазона текста первой страницы
                var firstPageRange = GetFirstPageRange(doc);

                // Получение и очистка имени автора из первого якорного текста
                string author = CleanAuthorName(GetFirstAnchorText(firstPageRange));

                // Получение названия дисциплины из заголовка документа и его очистка
                string discipline = CleanDisciplineName(GetDisciplineFromHeader(doc.Content));

                // Добавление автора и дисциплины в коллекцию
                AddAuthorDiscipline(author, discipline);

                // Увеличение счетчика публикаций по дисциплине
                AddDisciplineCount(discipline);
            }
        }

        // Обработка документов по умолчанию (не подходящих под специальные категории)
        private void ProcessDefault(Document doc)
        {
            // Проверка, что документ содержит хотя бы одну страницу
            if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) > 0)
            {
                // Получение диапазона первой страницы документа
                var firstPageRange = GetFirstPageRange(doc);

                // Подсчет количества якорных элементов (ссылок/меток) на первой странице
                int anchorCount = CountAnchors(firstPageRange);

                // Инициализация переменных для автора и дисциплины
                string author = null;
                string discipline = null;

                // Логика извлечения данных в зависимости от количества якорей:
                if (anchorCount == 1)
                {
                    // Если 1 якорь - берем только автора
                    author = CleanAuthorName(GetFirstAnchorText(firstPageRange));
                }
                else if (anchorCount == 2)
                {
                    // Если 2 якоря - первый для автора, второй для дисциплины
                    author = CleanAuthorName(GetFirstAnchorText(firstPageRange));
                    discipline = CleanDisciplineName(GetSecondAnchorText(firstPageRange));
                }
                else if (anchorCount >= 3)
                {
                    // Если 3+ якоря - второй для автора, третий для дисциплины
                    author = CleanAuthorName(GetSecondAnchorText(firstPageRange));
                    discipline = CleanDisciplineName(GetThirdAnchorText(firstPageRange));
                }

                // Если автор не найден через якоря, пробуем извлечь из таблицы
                if (string.IsNullOrEmpty(author))
                {
                    author = GetAuthorFromTable(doc);
                }

                // Если дисциплина не найдена через якоря, пробуем извлечь из заголовка
                if (string.IsNullOrEmpty(discipline))
                {
                    discipline = CleanDisciplineName(GetDisciplineFromHeader(doc.Content));
                }

                // Добавляем найденные данные в коллекции
                AddAuthorDiscipline(author, discipline);
                AddDisciplineCount(discipline);
            }
        }

        // Метод для получения диапазона первой страницы документа
        private Range GetFirstPageRange(Document doc)
        {
            try
            {
                // Начальная позиция (начало документа)
                int start = 0;

                // Конечная позиция - начало второй страницы минус 1 символ
                // (или конец документа, если страница одна)
                int end = doc.GoTo(
                    What: WdGoToItem.wdGoToPage,  // Переход по страницам
                    Which: WdGoToDirection.wdGoToAbsolute,  // Абсолютная позиция
                    Count: 2  // Вторая страница
                ).Start - 1;  // Берем позицию перед началом второй страницы

                // Возвращаем диапазон, защищаясь от отрицательных значений
                return doc.Range(Math.Max(0, start), Math.Max(0, end));
            }
            catch
            {
                // В случае ошибки возвращаем весь документ как fallback
                return doc.Content;
            }
        }

        // Подсчет количества якорных элементов (картинок и текстовых блоков) в указанном диапазоне
        private int CountAnchors(Range range)
        {
            int count = 0;

            // Подсчет inline-изображений в диапазоне
            count += range.InlineShapes.Cast<InlineShape>()
                .Count(shape => shape.Type == WdInlineShapeType.wdInlineShapePicture);

            // Подсчет текстовых блоков (TextBox) в диапазоне
            count += range.ShapeRange.Cast<Microsoft.Office.Interop.Word.Shape>()
                .Count(shape => shape.Type == MsoShapeType.msoTextBox);

            return count;
        }

        // Извлечение имени автора из параграфов документа
        private string GetAuthorFromParagraphs(Document doc)
        {
            // Перебор всех параграфов документа
            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                // Получение и очистка текста параграфа
                string text = paragraph.Range.Text.Trim();

                // Проверка, начинается ли параграф с "Автор:"
                if (text.StartsWith("Автор:", StringComparison.OrdinalIgnoreCase))
                {
                    // Извлечение части после двоеточия
                    string author = text.Substring(text.IndexOf(':') + 1).Trim();

                    // Удаление дополнительной информации после запятой (если есть)
                    int commaIndex = author.IndexOf(',');
                    return commaIndex >= 0 ? author.Substring(0, commaIndex).Trim() : author;
                }
            }
            return null; // Автор не найден
        }

        // Извлечение автора из таблицы документа
        private string GetAuthorFromTable(Document doc)
        {
            // Проверка наличия хотя бы двух таблиц в документе
            if (doc.Tables.Count >= 2)
            {
                Table secondTable = doc.Tables[2]; // Работа со второй таблицей

                // Перебор всех строк таблицы
                for (int i = 1; i <= secondTable.Rows.Count; i++)
                {
                    // Получение текста из первой ячейки
                    string potentialAuthor = secondTable.Rows[i].Cells[1].Range.Text.Trim();

                    // Проверка, содержит ли ячейка слово "автор"
                    if (potentialAuthor.ToLower().Contains("автор"))
                    {
                        // Извлечение имени автора из соседней ячейки
                        string authorName = secondTable.Rows[i].Cells[2].Range.Text.Trim();
                        return CleanAuthorName(authorName);
                    }
                }
            }
            return null; // Автор не найден
        }

        // Очистка и нормализация имени автора
        private string CleanAuthorName(string author)
        {
            if (string.IsNullOrEmpty(author)) return null;

            // Удаление специальных символов (например, \u0007 - символ звонка)
            author = author.Replace("\u0007", "").Trim();

            // Удаление дополнительной информации после запятой
            int commaIndex = author.IndexOf(',');
            return commaIndex >= 0 ? author.Substring(0, commaIndex).Trim() : author;
        }

        // Очистка и нормализация названия дисциплины
        private string CleanDisciplineName(string discipline)
        {
            if (string.IsNullOrEmpty(discipline)) return null;

            // Удаление кавычек и специальных символов
            return discipline.Replace("\"", "")
                .Replace("«", "")
                .Replace("»", "")
                .Replace("\r", "") // Удаление символа возврата каретки
                .Trim(); // Удаление пробелов по краям
        }

        // Добавление связи автор-дисциплина в словарь
        private void AddAuthorDiscipline(string author, string discipline)
        {
            // Проверка, что оба параметра не пустые
            if (!string.IsNullOrEmpty(author) && !string.IsNullOrEmpty(discipline))
            {
                // Если автор еще не добавлен в словарь
                if (!authorDisciplineCount.ContainsKey(author))
                {
                    // Создаем новую запись с пустым HashSet для дисциплин
                    authorDisciplineCount[author] = new HashSet<string>();
                }
                // Добавляем дисциплину в HashSet автора (HashSet автоматически исключает дубликаты)
                authorDisciplineCount[author].Add(discipline);
            }
        }

        // Увеличение счетчика упоминаний дисциплины
        private void AddDisciplineCount(string discipline)
        {
            // Проверка, что дисциплина не пустая
            if (!string.IsNullOrEmpty(discipline))
            {
                // Если дисциплина уже есть в словаре
                if (disciplineCount.ContainsKey(discipline))
                {
                    // Увеличиваем счетчик на 1
                    disciplineCount[discipline]++;
                }
                else
                {
                    // Создаем новую запись с начальным значением 1
                    disciplineCount[discipline] = 1;
                }
            }
        }

        // Обновление всех диаграмм
        private void UpdateCharts()
        {
            UpdateAuthorChart();      // Обновление диаграммы авторов
            UpdateDisciplineChart();  // Обновление диаграммы дисциплин
            UpdateHoursChart();       // Обновление диаграммы часов (метод не показан)
        }

        // Обновление диаграммы авторов
        private void UpdateAuthorChart()
        {
            // Очистка предыдущих данных
            chart1.Series.Clear();

            // Создание новой серии данных
            var series = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Дисциплины по авторам",  // Название серии
                Color = System.Drawing.Color.Blue,  // Цвет столбцов
                IsValueShownAsLabel = true,  // Показывать значения на столбцах
                ChartType = SeriesChartType.Column  // Тип диаграммы - столбчатая
            };

            // Добавление данных для каждого автора
            foreach (var kvp in authorDisciplineCount)
            {
                // X - имя автора, Y - количество дисциплин
                series.Points.AddXY(kvp.Key, kvp.Value.Count);
            }

            // Добавление серии на диаграмму
            chart1.Series.Add(series);

            // Настройка внешнего вида диаграммы
            ConfigureChart(chart1, "Авторы", "Количество дисциплин");
        }

        // Обновление диаграммы дисциплин
        private void UpdateDisciplineChart()
        {
            // Очистка предыдущих данных
            chart2.Series.Clear();

            // Создание новой серии данных
            var series = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Количество упоминаний дисциплин",  // Название серии
                Color = System.Drawing.Color.Blue,  // Цвет столбцов
                IsValueShownAsLabel = true,  // Показывать значения на столбцах
                ChartType = SeriesChartType.Column,  // Тип диаграммы - столбчатая
                LabelFormat = "0"  // Формат отображения значений (целые числа)
            };

            // Добавление данных для каждой дисциплины
            foreach (var kvp in disciplineCount)
            {
                // X - название дисциплины, Y - количество упоминаний
                series.Points.AddXY(kvp.Key, kvp.Value);
            }

            // Добавление серии на диаграмму
            chart2.Series.Add(series);

            // Настройка внешнего вида диаграммы
            ConfigureChart(chart2, "Дисциплины", "Количество упоминаний");

            // Наклон подписей по оси X на 45 градусов для лучшей читаемости
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
        }

        // Настройка внешнего вида диаграммы
        private void ConfigureChart(System.Windows.Forms.DataVisualization.Charting.Chart chart, string xTitle, string yTitle)
        {
            // Получаем первую область диаграммы (ChartArea)
            var area = chart.ChartAreas[0];

            // Устанавливаем заголовки осей
            area.AxisX.Title = xTitle;  // Заголовок оси X
            area.AxisY.Title = yTitle;  // Заголовок оси Y

            // Настройка оси Y:
            area.AxisY.IsStartedFromZero = true;  // Всегда начинать с 0
            area.AxisY.Minimum = 0;              // Минимальное значение 0

            // Автоматическое определение максимального значения оси Y
            if (chart.Series[0].Points.Count > 0)  // Если есть точки данных
            {
                // Находим максимальное значение Y и округляем вверх
                area.AxisY.Maximum = Math.Ceiling(chart.Series[0].Points.FindMaxByValue().YValues[0]);
            }
            else  // Если нет данных
            {
                area.AxisY.Maximum = 10;  // Устанавливаем значение по умолчанию
            }

            // Устанавливаем интервалы для оси Y
            area.AxisY.Interval = 1;           // Интервал основных делений
            area.AxisY.LabelStyle.Interval = 1; // Интервал подписей
        }

        // Извлечение названия дисциплины из заголовка документа
        private string GetDisciplineFromHeader(Range range)
        {
            string discipline = string.Empty;

            // Паттерны для поиска маркеров дисциплины в тексте
            string[] patterns = {
        "(название дисциплины)",
        "(вид практики)",
        "(название)"
    };

            // Паттерны заголовков рабочих программ
            string[] programPatterns = {
        "РАБОЧАЯ ПРОГРАММА УЧЕБНОЙ ДИСЦИПЛИНЫ",
        "РАБОЧАЯ УЧЕБНАЯ ПРОГРАММА ПО ДИСЦИПЛИНЕ",
        "ПРОГРАММА ПРАКТИКИ"
    };

            // Получаем весь текст из переданного диапазона
            string fullText = range.Text;

            // Поиск по каждому паттерну маркера дисциплины
            foreach (string pattern in patterns)
            {
                int index = fullText.IndexOf(pattern);
                if (index != -1)  // Если паттерн найден
                {
                    // Находим начало строки перед найденным паттерном
                    int startOfLine = fullText.LastIndexOf('\n', index);
                    if (startOfLine == -1)
                    {
                        startOfLine = 0;  // Если перевод строки не найден, начинаем с начала
                    }

                    // Извлекаем текст перед найденным паттерном
                    string previousLine = fullText.Substring(startOfLine, index - startOfLine).Trim();

                    // Проверяем, содержит ли эта строка заголовок программы
                    foreach (string programPattern in programPatterns)
                    {
                        int programIndex = previousLine.IndexOf(programPattern);
                        if (programIndex != -1)  // Если заголовок найден
                        {
                            // Извлекаем текст после заголовка программы
                            discipline = previousLine.Substring(programIndex + programPattern.Length).Trim();
                            break;  // Прерываем цикл после первого совпадения
                        }
                    }

                    // Если не нашли заголовок программы, берем всю строку как дисциплину
                    if (string.IsNullOrEmpty(discipline))
                    {
                        discipline = previousLine;
                    }

                    break;  // Прерываем цикл после первого найденного паттерна
                }
            }

            // Очищаем результат от лишних символов и возвращаем
            return discipline.Replace("\r", "").Trim();
        }


        // Получение текста из первого якорного элемента (картинки или текстового блока)
        private string GetFirstAnchorText(Range range)
        {
            // Поиск среди встроенных фигур (InlineShapes)
            foreach (InlineShape shape in range.InlineShapes)
            {
                // Проверка, является ли фигура изображением
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    // Возвращаем текст, связанный с изображением (подпись или альтернативный текст)
                    return shape.Range.Text.Trim();
                }
            }

            // Поиск среди обычных фигур (Shapes)
            foreach (Microsoft.Office.Interop.Word.Shape shape in range.ShapeRange)
            {
                // Проверка, является ли фигура текстовым блоком
                if (shape.Type == MsoShapeType.msoTextBox)
                {
                    // Возвращаем текст из текстового блока
                    return shape.TextFrame.TextRange.Text.Trim();
                }
            }

            // Возвращаем null, если якорные элементы не найдены
            return null;
        }

        // Получение текста из второго якорного элемента
        private string GetSecondAnchorText(Range range)
        {
            int anchorCounter = 0; // Счетчик найденных якорей

            // Поиск среди встроенных фигур
            foreach (InlineShape shape in range.InlineShapes)
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    anchorCounter++;
                    // Если это второй найденный якорь
                    if (anchorCounter == 2)
                    {
                        return shape.Range.Text.Trim();
                    }
                }
            }

            // Поиск среди обычных фигур
            foreach (Microsoft.Office.Interop.Word.Shape shape in range.ShapeRange)
            {
                if (shape.Type == MsoShapeType.msoTextBox)
                {
                    anchorCounter++;
                    // Если это второй найденный якорь (с учетом всех типов)
                    if (anchorCounter == 2)
                    {
                        return shape.TextFrame.TextRange.Text.Trim();
                    }
                }
            }

            return null; // Второй якорь не найден
        }

        // Получение текста из третьего якорного элемента
        private string GetThirdAnchorText(Range range)
        {
            int anchorCounter = 0; // Счетчик найденных якорей

            // Поиск среди встроенных фигур
            foreach (InlineShape shape in range.InlineShapes)
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    anchorCounter++;
                    // Если это третий найденный якорь
                    if (anchorCounter == 3)
                    {
                        return shape.Range.Text.Trim();
                    }
                }
            }

            // Поиск среди обычных фигур
            foreach (Microsoft.Office.Interop.Word.Shape shape in range.ShapeRange)
            {
                if (shape.Type == MsoShapeType.msoTextBox)
                {
                    anchorCounter++;
                    // Если это третий найденный якорь (с учетом всех типов)
                    if (anchorCounter == 3)
                    {
                        return shape.TextFrame.TextRange.Text.Trim();
                    }
                }
            }

            return null; // Третий якорь не найден
        }

        // Словарь для хранения количества часов по дисциплинам
        private Dictionary<string, int> disciplineHours = new Dictionary<string, int>();

        // Метод для извлечения информации о часах из документа
        private void ExtractDisciplineHours(string filePath)
        {
            // Создаем экземпляр Word приложения
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;

            try
            {
                // Открываем документ
                doc = wordApp.Documents.Open(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string discipline = string.Empty;

                // Получаем название дисциплины в зависимости от типа файла
                if (fileName.Contains("РП АИУС") || fileName.Contains("РП ТСАУ"))
                {
                    // Для АИУС и ТСАУ берем дисциплину из заголовка
                    discipline = GetDisciplineFromHeader(doc.Content);
                }
                else if (fileName.Contains("РП ПиОА"))
                {
                    // Для ПиОА берем из второго якорного текста на первой странице
                    if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) > 0)
                    {
                        var firstPageRange = GetFirstPageRange(doc);
                        discipline = CleanDisciplineName(GetSecondAnchorText(firstPageRange));
                    }
                }
                else
                {
                    // Для остальных документов анализируем количество якорей
                    if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) > 0)
                    {
                        Range firstPageRange = GetFirstPageRange(doc);
                        int anchorCount = CountAnchors(firstPageRange);

                        // Выбираем источник дисциплины в зависимости от количества якорей
                        if (anchorCount >= 3)
                        {
                            discipline = GetThirdAnchorText(firstPageRange);
                        }
                        else if (anchorCount == 2)
                        {
                            discipline = GetSecondAnchorText(firstPageRange);
                        }
                        else
                        {
                            discipline = GetDisciplineFromHeader(doc.Content);
                        }
                    }
                }

                // Очищаем название дисциплины
                discipline = CleanDisciplineName(discipline);

                // Если дисциплина найдена
                if (!string.IsNullOrEmpty(discipline))
                {
                    int hours = 0;

                    // Особый случай для РП АИУС - обработка таблицы
                    if (fileName.Contains("РП АИУС") && doc.ComputeStatistics(WdStatistic.wdStatisticPages) >= 2)
                    {
                        // Берем третью таблицу в документе
                        Table table = doc.Tables[3];

                        if (table != null && table.Columns.Count >= 2)
                        {
                            // Получаем ячейку (3 строка, 2 колонка)
                            Cell cell = table.Cell(3, 2);
                            string cellText = cell.Range.Text;

                            // Разбиваем текст ячейки на строки
                            string[] lines = cellText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                            // Обрабатываем каждую строку
                            foreach (string line in lines)
                            {
                                string cleanLine = line.Trim();

                                // Удаляем слова "зачёт", "экз" и подобные с последующими символами
                                cleanLine = Regex.Replace(cleanLine, @"(зачёт|зачет|экз|зкз)[.-]*\s*", "", RegexOptions.IgnoreCase).Trim();

                                // Удаляем оставшиеся дефисы/точки в начале строки
                                cleanLine = Regex.Replace(cleanLine, @"^[.-]\s*", "").Trim();

                                // Пробуем преобразовать в число
                                if (int.TryParse(cleanLine, out int value))
                                {
                                    hours += value;
                                }
                            }
                        }
                    }
                    else if (doc.ComputeStatistics(WdStatistic.wdStatisticPages) >= 2)
                    {
                        // Стандартная обработка для других файлов
                        // Получаем диапазон второй страницы
                        Range secondPageRange = doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 2);
                        Range endRange = doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 3);
                        secondPageRange.End = endRange.Start;

                        // Ищем часы в тексте второй страницы
                        string pageText = secondPageRange.Text;
                        hours = FindAcademicHours(pageText);
                    }

                    // Если часы найдены, добавляем в словарь
                    if (hours > 0)
                    {
                        if (disciplineHours.ContainsKey(discipline))
                        {
                            disciplineHours[discipline] += hours; // Суммируем часы для существующей дисциплины
                        }
                        else
                        {
                            disciplineHours[discipline] = hours; // Добавляем новую запись
                        }
                    }
                }
            }
            finally
            {
                // Гарантированное закрытие документа и Word
                doc?.Close(false);
                wordApp.Quit();
            }
        }
        // Метод для поиска количества академических часов в тексте
        private int FindAcademicHours(string text)
        {
            // Нормализация текста:
            // 1. Удаляем переносы слов (дефис + перевод строки)
            // 2. Заменяем все переводы строк на пробелы
            // 3. Заменяем двойные пробелы на одинарные
            text = text.Replace("-\r\n", "").Replace("\r\n", " ").Replace("  ", " ");

            // Очистка чисел от окружающих символов:
            // Заменяем подчеркивания и пробелы вокруг чисел на единичные пробелы
            // Пример: "_100_" -> " 100 "
            text = Regex.Replace(text, @"[_\s]+(\d+)[_\s]+", " $1 ");

            // Основной паттерн поиска:
            // Ищем числа, после которых идет слово "час" (в любом падеже),
            // возможно с прилагательным "академических" перед ним
            string pattern = @"(\d+)\s*(?=(академических\s*)?час[аов]+\b)";

            // Поиск всех совпадений с регулярным выражением
            var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
            if (matches.Count > 0)
            {
                // Берем последнее совпадение (наиболее вероятно правильное)
                var lastMatch = matches[matches.Count - 1];
                // Пробуем преобразовать найденное значение в число
                if (int.TryParse(lastMatch.Groups[1].Value, out int hours))
                {
                    return hours;
                }
            }

            // Возвращаем 0, если ничего не найдено
            return 0;
        }

        // Метод для обновления диаграммы часов по дисциплинам
        private void UpdateHoursChart()
        {
            // Очищаем предыдущие данные
            chart3.Series.Clear();

            // Создаем новую серию данных
            var series = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Академические часы по дисциплинам",  // Название серии
                Color = System.Drawing.Color.Green,          // Цвет столбцов
                IsValueShownAsLabel = true,                 // Показывать значения на столбцах
                ChartType = SeriesChartType.Column,          // Тип диаграммы - столбчатая
                LabelFormat = "0",                          // Формат чисел (целые)
            };

            // Добавляем данные в серию
            foreach (var kvp in disciplineHours)
            {
                // X - название дисциплины, Y - количество часов
                series.Points.AddXY(kvp.Key, kvp.Value);
            }

            // Добавляем серию на диаграмму
            chart3.Series.Add(series);

            // Настройка оси X:
            chart3.ChartAreas[0].AxisX.Title = "Дисциплины";         // Заголовок оси
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = -45;       // Наклон подписей на 45°
            chart3.ChartAreas[0].AxisX.Interval = 1;                 // Интервал между подписями
            chart3.ChartAreas[0].AxisX.IsMarginVisible = false;      // Убираем отступы по краям

            // Настройка оси Y:
            chart3.ChartAreas[0].AxisY.Title = "Академические часы"; // Заголовок оси
            chart3.ChartAreas[0].AxisY.Minimum = 0;                  // Минимальное значение 0

            // Вычисление максимального значения с округлением:
            // 1. Находим максимальное значение из всех часов
            int maxValue = disciplineHours.Values.DefaultIfEmpty().Max();
            // 2. Округляем до ближайших 50 в большую сторону
            int roundedMax = ((maxValue / 50) + 1) * 50;
            // 3. Устанавливаем максимальное значение оси
            chart3.ChartAreas[0].AxisY.Maximum = roundedMax;

            // Установка интервалов для оси Y:
            chart3.ChartAreas[0].AxisY.Interval = 50;               // Интервал основных делений
            chart3.ChartAreas[0].AxisY.LabelStyle.Interval = 50;     // Интервал подписей
            chart3.ChartAreas[0].AxisY.MajorGrid.Interval = 50;     // Интервал сетки
        }

    }
}
