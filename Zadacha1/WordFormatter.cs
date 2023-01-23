﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace Zadacha1
{
    internal class WordFormatter
    {
        public WordFormatter(string sourcePath,string distPath)
        {
            this.sourcePath= sourcePath;
            this.distPath = distPath;
        }
        public WordFormatter(){}
        //определение параметров класса
        /// <summary>текущий номер раздела в тексте</summary>
        int _sectionNumber = 0;
        /// <summary>текущий номер  рисунка в тексте<</summary>
        int _pictureNumber = 0;
        /// <summary>текущий номер  таблиц в тексте<</summary>
        int _tableNumber = 0;
        /// <summary>нумерация источников в списке литературы</summary>
        int _sourceNumber = 0;
        /// <summary>текущий номер  раздела в закладках</summary>
        int _sectionBookmark = 0;
        /// <summary>текущий номер  раздела в закладках</summary>
        int _pictureBookMark = 0;
        /// <summary>текущий номер  таблиц в закладках</summary>
        int _tableBookMark = 0;
        /// <summary>текущий номер  вставки кода в закладках</summary>
        int _codeBookMark = 0;
        /// <summary>путь до исходного шаблона</summary>
        string sourcePath = "шаблон.rtf";
        /// <summary>путь до выходного файла</summary>
        string distPath = "result.rtf";
        /// <summary>список шаблонных строк в тексте для форматирования</summary>
        string[] templateStringList =
            {
                "[*номер раздела*]",                    //0
                "[*номер рисунка*]",                    //1
                "[*номер таблицы*]",                    //2
                "[*ссылка на следующий рисунок*]",      //3
                "[*ссылка на предыдущий рисунок*]",     //4
                "[*ссылка на таблицу*]",                //5
                "[*таблица ",                           //6
                 "[*cписок литературы*]",               //7
                 "[*код",                               //8
                };
        /// <summary>префиксы названий закладок, которыми будем пользоваться</summary>
        string[] listBookMarkes =
        {
                    "_numberSection", //0
                    "_numberPicture", //1
                    "_numberTable",   //2
                    "_literature",    //3
                    "_code"           //4
                };

        /// <summary>список литературы</summary>
        List<string> sourceList = new List<string>();
        /// <summary>пустая ссылка для передачи пустого параметра в COM-объект</summary>
        Object missing = System.Type.Missing;
        public void Make()
        {
            //открытие приложения Word
            var application = new Application();
            //делаем приложение видимым пользователю
            application.Visible = true;
            //открываем документ
            Document document = application.Documents.Open(sourcePath, false);

            //CODEPART 1 Сборка ранее определенного списка литературы
            //список уже определенных источников

            //CODEPART 2 Определение параметров уже определенных закладок
            //делаем скрытые закладки видимыми
            document.Bookmarks.ShowHidden = true;
            //и устанавливаем их сортировку по их порядку в документе
            document.Bookmarks.DefaultSorting = WdBookmarkSortBy.wdSortByLocation;

            //обходим все закладки в документе
            foreach (Bookmark bookmark in document.Bookmarks)
            {
                //этот цикл нам нужен, чтобы не перезаписывать уже определенные закладки
            }

            //CODEPART 3 Первый обход абзацев - форматирование, вставка закладок, перекресных ссылок
            //для того, чтобы выравнивать по центру рисунок, следующий перед подрисуночной подписью
            //будем фиксировать ссылку на предыдущий параграф
            Paragraph prevParagraph = null;

            //обходим в документе все параграфы
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                //CODEPART 3.1 Исключение абзацев, которые не нужно трогать

                //флаг, для того, чтобы определять, нужно ли применять форматирование
                //к параграфу, как к основному тексту
                bool isStandartFormat = true;

                //проверяем наличие шаблонных строк в абзаце 
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    //если есть шаблонная строка
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        switch (i)
                        {
                            //CODEPART 3.2 Редактирование абзаца заголовка раздела
                            case 0:// "[*номер раздела*]"
                                {
                                    //выделяем параграф, ищем поиском шаблонную строку, вставляем закладку
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    application.Selection.Range.Bookmarks.Add("_numberSection" + ++_sectionBookmark, application.Selection.Range);

                                    //форматируем абзац
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 15;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.Font.Bold = 1;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    isStandartFormat = false;//дальнейшее форматирование не нужно
                                }
                                break;
                            //CODEPART 3.3 Редактирование подрисуночной подписи
                            case 1://"[*номер рисунка*]"
                                {
                                    //так как у подрисуночной подписи еще нужно вставить подпись Рисунок
                                    //ищем шаблонную строку в выделении
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);

                                    //переводим курсор в начало шаблонной строки
                                    Microsoft.Office.Interop.Word.Range range = application.Selection.Range;
                                    range.Start = application.Selection.Range.Start;
                                    range.End = application.Selection.Range.Start;
                                    range.Select();

                                    //вставляем текст
                                    application.Selection.Text = "Рисунок" + '\u00A0';

                                    //снова ищем шаблонную строку
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);

                                    //переносим курсор в конец шаблонной строки
                                    range = application.Selection.Range;
                                    range.Start = application.Selection.Range.End;
                                    range.End = application.Selection.Range.End;
                                    range.Select();

                                    //вставляем тире
                                    application.Selection.Text = " –";

                                    //снова ищем шаблонную строку, чтобы вставить закладку
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    document.Bookmarks.Add("_numberPicture" + ++_pictureBookMark, application.Selection.Range);

                                    //форматируем абзац
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 12;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    //так как это подрисуночная подпись, рисунок тоже нужно выравнять по ширине 
                                    if (prevParagraph != null)
                                    {
                                        //а рисунок на предыдущем абзаце
                                        //форматируем
                                        prevParagraph.Format.SpaceBefore = 12;
                                        prevParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    }

                                    isStandartFormat = false;//дальнейшее форматирование не нужно
                                }
                                break;
                            //CODEPART 3.4 Редактирование заголовка таблицы
                            case 2://"[*номер таблицы*]"
                                {
                                    //так как у надтабличной подписи еще нужно вставить подпись Таблица
                                    //ищем шаблонную строку в выделении
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);

                                    //переводим курсор в начало шаблонной строки
                                    Microsoft.Office.Interop.Word.Range range = application.Selection.Range;
                                    range.Start = application.Selection.Range.Start;
                                    range.End = application.Selection.Range.Start;
                                    range.Select();

                                    //вставляем текст
                                    application.Selection.Text = "Таблица" + '\u00A0';

                                    //снова ищем шаблонную строку
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);

                                    //переносим курсор в конец шаблонной строки
                                    range = application.Selection.Range;
                                    range.Start = application.Selection.Range.End;
                                    range.End = application.Selection.Range.End;
                                    range.Select();

                                    //вставляем тире
                                    application.Selection.Text = " –";

                                    //снова ищем шаблонную строку, чтобы вставить закладку
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    application.Selection.Range.Bookmarks.Add("_numberTable" + ++_tableBookMark, application.Selection.Range);

                                    //форматируем абзац
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 12;
                                    paragraph.Format.SpaceBefore = 12;
                                    paragraph.Format.SpaceAfter = 0;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    isStandartFormat = false;//дальнейшее форматирование не нужно
                                }
                                break;

                            //CODEPART 3.5 Вставка перекрестной ссылки на предыдущий рисунок
                            case 4://"[*ссылка на предыдущий рисунок*]"
                                {
                                }
                                break;
                            //CODEPART 3.5 Вставка таблицы из файла
                            case 6://"[*таблица "
                                {
                                }
                                break;
                            //CODEPART 3.6 Вставка закладки списка литературы
                            case 7://"[*cписок литературы*]"
                                {
                                    //если есть шаблонная строка для места вставки списка литературы
                                    //заменяем закладкой
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    document.Bookmarks.Add("_literature", application.Selection.Range);
                                }
                                break;
                            //CODEPART 3.7 Вставка кода из файла
                            case 8://"[*код"
                                {
                                    //если есть шаблонная строка для места вставки кода
                                    //заменяем закладкой
                                    paragraph.Range.Select();
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    document.Bookmarks.Add("_code" + ++_codeBookMark, application.Selection.Range);
                                    //TODO (задание на 5) вставить код из файла - CourierNew 8 пт одинарный без отступа в рамке
                                }
                                break;
                        }
                    }
                }

                //CODEPART 3.8 Сбор внутритекстовых ссылок на литературу

                //CODEPART 3.9 Стандартное форматирование абзаца
                //если нужно абзац форматировать как обычный текст
                if (isStandartFormat)
                {
                    //то форматируем основной текст
                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    paragraph.Range.Font.Name = "Times New Roman";
                    paragraph.Range.Font.Size = 14;
                    paragraph.Format.SpaceAfter = 0;
                    paragraph.Format.SpaceBefore = 0;
                    paragraph.Range.Font.Bold = 0;
                    paragraph.Range.HighlightColorIndex = 0;

                    paragraph.Format.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                    paragraph.Format.CharacterUnitFirstLineIndent = 1.5f;
                }

                //фиксируем ссылку на текущий абзац
                prevParagraph = paragraph;

            }

            //CODEPART 4 Формирование списка литературы

            //CODEPART 5 Заполнение закладок номера таблиц, рисунков и разделов
            //осталось переопределить номера таблиц, рисунков и разделов
            //снова обходим все параграфы
            foreach (Paragraph paragraph in document.Paragraphs)
            {
            }


            //обновляем все поля, чтобы перекрестные ссылки забрали текст закладок
            document.Fields.Update();

            //сохраняем документ
            document.SaveAs2(distPath);
            application.Quit();

        }
    }
}
