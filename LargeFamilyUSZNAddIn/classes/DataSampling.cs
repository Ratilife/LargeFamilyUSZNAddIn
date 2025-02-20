﻿using DocumentFormat.OpenXml.Drawing.Charts;
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

        /**
         * 
         * <summary>
         * Находит и возвращает строки, соответствующие заданной маске, из списка текстов.
         * </summary>
         * 
         * <param name="texts">Список строк, в которых будет производиться поиск.</param>
         * <param name="mask">Строка, представляющая маску, которая будет преобразована в регулярное выражение для поиска.</param>
         * 
         * <returns>
         * Возвращает список строк, содержащих найденные совпадения, соответствующие маске.
         * Если в тексте встречаются пустые строки, они также добавляются в результат.
         * </returns>
         * 
         * <remarks>
         * Метод выполняет следующие шаги:
         * 1. Преобразует заданную маску в регулярное выражение, используя метод ConvertMaskToRegex.
         * 2. Создает объект регулярного выражения на основе полученного шаблона.
         * 3. Итерирует по каждому тексту в списке.
         * 4. Проверяет, является ли текущий текст пустым или равным null. Если да, добавляет его в список результатов.
         * 5. Ищет все совпадения регулярного выражения в текущем тексте.
         * 6. Добавляет все найденные совпадения в результирующий список.
         * 7. Возвращает список найденных совпадений.
         * 
         * <note>
         * Если текст содержит пустые строки, они будут добавлены в список результатов без изменений.
         * Метод не выполняет дополнительных действий по обработке или фильтрации пустых строк, кроме их добавления в результат.
         * </note>
         */
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

        public (List<string> matchedResults, List<string> nonMatchedResults) FindNumbersByCustomPatternSeparator(List<string> texts, string mask)
        {
            // Преобразование маски в регулярное выражение:
            string pattern = ConvertMaskToRegex(mask);

            // Создание объекта регулярного выражения:
            var regex = new Regex(pattern);

            // Создание списков для результатов:
            var matchedResults = new List<string>();       // Список строк, соответствующих маске
            var nonMatchedResults = new List<string>();    // Список строк, которые не прошли по маске

            // Символы, которые нужно отбросить из непрошедших по маске строк
            char[] separators = new char[] { ',', ';', ':', '/','0','1','2','3','4','5','6','7','8','9' };

            // Итерация по текстам:
            foreach (var text in texts)
            {
                // Пропуск пустых строк
                if (string.IsNullOrEmpty(text))
                {
                    matchedResults.Add(text);
                    nonMatchedResults.Add(text);
                    //continue;
                }

                // Поиск совпадений по регулярному выражению
                var matches = regex.Matches(text);

                // Создаем переменную для хранения текста без совпадений
                string DiscardedText = text;

                //if (matches.Count > 0)
                //{
                    // Если есть совпадения, добавляем их в список совпадений
                    foreach (Match match in matches)
                    {
                        matchedResults.Add(match.Value);
                        // Удаляем совпадающее значение из DiscardedText
                        DiscardedText = DiscardedText.Replace(match.Value, string.Empty).Trim();
                        // Если совпадений нет, очищаем строку от заданных символов и добавляем в nonMatchedResults
                        string cleanedText = new string(DiscardedText.Where(c => !separators.Contains(c)).ToArray());
                        nonMatchedResults.Add(cleanedText.Trim());  // Удаляем возможные пробелы в начале и конце
                    }
                    
                //}
                //else
                //{
                    // Если совпадений нет, очищаем строку от заданных символов и добавляем в nonMatchedResults
                    //string cleanedText = new string(text.Where(c => !separators.Contains(c)).ToArray());
                    //nonMatchedResults.Add(cleanedText);
                //}
            }

            // Возвращаем два списка: совпадения и несоответствующие элементы
            return (matchedResults, nonMatchedResults);
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

        #region СравнениеЧисла

        public Dictionary<string, string> FindNumbersByCustomPatternDictionary(Dictionary<string, string> textsDict, string mask)
        {
            // Преобразование маски в регулярное выражение:
            string pattern = ConvertMaskToRegex(mask);

            // Создание объекта регулярного выражения:
            var regex = new Regex(pattern);

            // Создание словаря для результатов:
            var resultDict = new Dictionary<string, string>();
            int index = 1;
            // Итерация по элементам словаря:
            foreach (var kvp in textsDict)
            {
                string key = kvp.Key;     // Ключ (исходное значение)
                string text = kvp.Value;  // Текст для обработки

                // Список для хранения найденных строк по каждому ключу:
                var results = new List<string>();

                // Пропуск пустых строк
                if (string.IsNullOrEmpty(text))
                {
                    results.Add(text);
                }
                else
                {
                    // Поиск совпадений по регулярному выражению
                    var matches = regex.Matches(text);
                    foreach (Match match in matches)
                    {
                        //results.Add(match.Value); // Добавляем совпадение
                        // Добавляем результаты в словарь с тем же ключом
                        string key_index = $"{key}({index})";
                        resultDict.Add(key_index , match.Value);
                        index++;
                    }
                }
            }

            return resultDict;
        }

        
        public Dictionary<string, string> TransformByMaskDictionary(Dictionary<string, string> dataDict, string mask)
        {
            var result = new Dictionary<string, string>();

            // Итерируем по каждому элементу словаря
            foreach (var kvp in dataDict)
            {
                string key = kvp.Key;      // Исходное значение (ключ)
                string data = kvp.Value;   // Данные для преобразования (значение)

                // Удаляем все символы, кроме цифр
                string numericData = new string(data.Where(char.IsDigit).ToArray());
                // Проверяем, если хотя бы одно из числовых значений удовлетворяет длине маски
                if (numericData.Length >= mask.Count(c => c == '#'))
                {
                    // Применяем маску к числовым данным
                    string transformedData = ApplyMask(numericData, mask);

                    // Добавляем результат в словарь
                    result[key] = transformedData;
                }
            }

            return result;
        }

        /*//совпадение
        public List<string> CompareDictionaries(Dictionary<string, string> dict1, Dictionary<string, string> dict2)
        {
            List<string> result = new List<string>();

            // Создаем обратный словарь, где ключи - значения dict2, а значения - ключи dict2
            Dictionary<string, List<string>> reverseDict2 = new Dictionary<string, List<string>>();

            // Заполняем обратный словарь
            foreach (var kvp2 in dict2)
            {
                if (!reverseDict2.ContainsKey(kvp2.Value))
                {
                    reverseDict2[kvp2.Value] = new List<string>();
                }
                reverseDict2[kvp2.Value].Add(kvp2.Key);
            }

            // Итерируем по dict1 и проверяем совпадения по значению в обратном словаре
            foreach (var kvp1 in dict1)
            {
                if (reverseDict2.TryGetValue(kvp1.Value, out List<string> keys2))
                {
                    // Добавляем все соответствующие ключи из dict2
                    foreach (var key2 in keys2)
                    {
                        result.Add(kvp1.Key);
                        result.Add(key2);
                    }
                }
            }

            return result;
        }
        */

        //на случай если работающий метод не работает правельно заменить
        public List<string> CompareDictionaries(Dictionary<string, string> dict1, Dictionary<string, string> dict2)
        {
            List<string> result = new List<string>();

            foreach (var kvp1 in dict1)
            {
                foreach (var kvp2 in dict2)
                {
                    // Сравниваем значения словарей
                    if (kvp1.Value == kvp2.Value)
                    {
                        // Добавляем ключи из обоих словарей в список
                        result.Add(kvp1.Key);
                        result.Add(kvp2.Key);
                    }
                }
            }

            return result;
        }
        /*
        public List<string> CompareDictionaries(Dictionary<string, string> dict1, Dictionary<string, string> dict2)
        {
            List<string> result = new List<string>();
            HashSet<string> matchedKeysDict1 = new HashSet<string>();
            HashSet<string> matchedKeysDict2 = new HashSet<string>();

            // Итерируем по dict1 и dict2, чтобы найти совпадения
            foreach (var kvp1 in dict1)
            {
                foreach (var kvp2 in dict2)
                {
                    if (kvp1.Value == kvp2.Value)
                    {
                        // Если найдено совпадение, запоминаем ключи
                        matchedKeysDict1.Add(kvp1.Key);
                        matchedKeysDict2.Add(kvp2.Key);
                    }
                }
            }

            // Добавляем в результат ключи из dict1, которые не были найдены в dict2
            foreach (var kvp1 in dict1)
            {
                if (!matchedKeysDict1.Contains(kvp1.Key))
                {
                    result.Add(kvp1.Key);
                }
            }

            // Добавляем в результат ключи из dict2, которые не были найдены в dict1
            foreach (var kvp2 in dict2)
            {
                if (!matchedKeysDict2.Contains(kvp2.Key))
                {
                    result.Add(kvp2.Key);
                }
            }

            return result;
            }

       */

        // Основной метод, который сравнивает два словаря по маске и возвращает список ключей с совпадающими значениями.
        public List<string> CompareDictionaries2(Dictionary<string, string> dict1, Dictionary<string, string> dict2, string mask)
        {
            List<string> result = new List<string>();

            foreach (var kvp1 in dict1)
            {
                // Проверяем, есть ли ключ в dict2 и получаем его значение
                if (dict2.TryGetValue(kvp1.Key, out string value2))
                {
                    // Применяем маску к значениям словарей
                    string value1Processed = ApplyMask(kvp1.Value, mask);
                    string value2Processed = ApplyMask(value2, mask);

                    // Сравниваем обработанные строки
                    if (value1Processed == value2Processed)
                    {
                        result.Add(kvp1.Key);
                    }
                }
            }

            return result;
        }

        public List<string> CompareDictionariesDifference(Dictionary<string, string> dict1, Dictionary<string, string> dict2)
        {
            List<string> result = new List<string>();

            // Создаем обратный словарь, где ключи - значения dict2, а значения - ключи dict2
            Dictionary<string, List<string>> reverseDict2 = new Dictionary<string, List<string>>();

            // Заполняем обратный словарь
            foreach (var kvp2 in dict2)
            {
                if (!reverseDict2.ContainsKey(kvp2.Value))
                {
                    reverseDict2[kvp2.Value] = new List<string>();
                }
                reverseDict2[kvp2.Value].Add(kvp2.Key);
            }

            // Итерируем по dict1 и проверяем, есть ли совпадения в обратном словаре
            foreach (var kvp1 in dict1)
            {
                // Если совпадений по значению нет, добавляем ключ из dict1 в результат
                if (!reverseDict2.ContainsKey(kvp1.Value))
                {
                    result.Add(kvp1.Key);
                }
            }

            // Также проверяем dict2 и добавляем ключи из dict2, которые не имеют совпадений в dict1
            foreach (var kvp2 in dict2)
            {
                if (!dict1.ContainsValue(kvp2.Value))
                {
                    result.Add(kvp2.Key);
                }
            }

            return result;
        }


        /*
        // Метод для сравнения двух списков после применения маски.
        private bool CompareValueLists(List<string> values1, List<string> values2)
        {
            foreach (var val1 in values1)
            {
                if (values2.Any(val2 => val1 == val2))
                {
                    return true; // Совпадение найдено.
                }
            }

            return false; // Нет совпадений.
        }

        //TODO Проверить код
        // Метод для применения маски к каждому значению списка.
        private List<string> ApplyMaskList(List<string> values, string mask)
        {
            return values
                .Select(value => ExtractByMask(value, mask))
                .ToList();
        }

        // Метод для извлечения части строки на основе маски.
        private string ExtractByMask(string value, string mask)
        {
            // Пример использования маски: здесь можно изменить реализацию для более сложной логики.
            // Например, маска может указывать, какие символы оставить или какие позиции учитывать.
            // Сейчас предполагается, что маска просто фильтрует строки на основе содержания.
            // (например, возвращает строку, если она содержит определенные символы из маски).

            return new string(value.Where(c => mask.Contains(c)).ToArray());
        }
        */
        #endregion
    }
}
