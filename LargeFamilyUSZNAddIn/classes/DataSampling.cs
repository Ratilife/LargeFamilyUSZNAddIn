using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using Catalyst;
//using Mosaik.Core;

namespace LargeFamilyUSZNAddIn.classes
{
    internal class DataSampling
    {
        #region ВыборкаФИО
        /*
        // поиск ФИО
        public async Task<List<string>> RecognizeNames(List<string> texts)
        {
            //Настройка хранилища моделей:
            //Здесь задается текущее хранилище моделей. В данном случае используется онлайн-репозиторий с локальным кэшем на диске.
            Storage.Current = new OnlineRepositoryStorage(new DiskStorage("catalyst-models"));
            //Инициализация NLP пайплайна:
            // Этот шаг инициализирует NLP пайплайн для русского языка, загружая необходимые модели.
            var nlp = await Pipeline.ForAsync(Language.Russian);
            //Создание списка для результатов:
            //Создается пустой список для хранения распознанных имен.
            var results = new List<string>();
            //Обработка каждого текста:
            foreach (var text in texts)
            {
                var doc = new Document(text);
                nlp.ProcessSingle(doc);

                //Для каждого текста из списка создается объект Document, который затем обрабатывается пайплайном NLP.
                //Для каждого предложения в документе извлекаются сущности. Если метка сущности равна “PER” 
                //(что означает, что это имя), значение сущности добавляется в список результатов.
                foreach (var sentence in doc)
                {
                    foreach (var entity in sentence.Entities)
                    {
                        if (entity.Label == "PER") // "PER" - метка для имен
                        {
                            results.Add(entity.Value);
                        }
                    }
                }
            }

            return results;
        }
   */
        #endregion

        #region ВыборкаЧисла
       
        public List<string> FindNumbersByCustomPattern(List<string> texts, string mask)
        {
            //Преобразование маски в регулярное выражение:
            string pattern = ConvertMaskToRegex(mask);
            // Создание объекта регулярного выражения:
            //Создается объект Regex с преобразованным шаблоном регулярного выражения.
            var regex = new Regex(pattern);
            //Создание списка для результатов:
            //Создается пустой список для хранения найденных строк, соответствующих маске.
            var results = new List<string>();
            //Итерация по текстам:
            foreach (var text in texts)
            {
                //TODO проверить код 
                if (string.IsNullOrEmpty(text)) 
                {
                    results.Add(text);
                }
                    var matches = regex.Matches(text);
                foreach (Match match in matches)
                {
                    results.Add(match.Value);
                }
            }

            return results;
        }

        //Метод для преобразования маски в регулярное выражение:
        public string ConvertMaskToRegex(string mask)
        {

            // Заменяем символы маски на соответствующие регулярные выражения
            string regexPattern = mask
                .Replace("#", @"\d");  // Заменяем # на \d (цифра)
                //.Replace("П", @"\p{L}"); // Заменяем П на \p{L} (буква)
            return regexPattern;
        }

        public List<string> TransformByMask(List<string> dataList, string mask)
        {
            var result = new List<string>();

            // Итерируем по каждому элементу списка
            foreach (var data in dataList)
            {
                // Удаляем все символы, кроме цифр
                string numericData = new string(data.Where(char.IsDigit).ToArray());

                // Если длина числовых данных меньше, чем длина маски, пропускаем этот элемент
                if (numericData.Length < mask.Count(c => c == '#'))
                {
                    result.Add(data); // Добавляем исходное значение без изменений
                    continue;
                }

                // Применяем маску к числовым данным
                string transformedData = ApplyMask(numericData, mask);

                // Добавляем преобразованные данные в результат
                result.Add(transformedData);
            }

            return result;
        }

        // Метод для применения маски к строке с цифрами
        private string ApplyMask(string numericData, string mask)
        {
            var maskedData = new System.Text.StringBuilder();
            int dataIndex = 0;

            // Проходим по символам маски
            foreach (var maskChar in mask)
            {
                if (maskChar == '#') // Если в маске символ '#', заменяем его цифрой
                {
                    maskedData.Append(numericData[dataIndex]);
                    dataIndex++;
                }
                else // Если это любой другой символ (разделитель), добавляем его как есть
                {
                    maskedData.Append(maskChar);
                }
            }

            return maskedData.ToString();
        }


        #endregion

    }
}
