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

        #endregion

    }
}
