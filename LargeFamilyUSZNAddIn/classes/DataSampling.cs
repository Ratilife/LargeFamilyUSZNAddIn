using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Catalyst;
//using Mosaik.Core;

namespace LargeFamilyUSZNAddIn.classes
{
    internal class DataSampling
    {
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
    }
}
