using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Collections;
using System.IO;
using Microsoft.Win32;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WpfApplication1
{
    //class CreateWordFile
    public partial class MainWindow : System.Windows.Window
    {
        partial void CreateWordFile()
        {
            // Проверка того, что выбрано хотя бы одно Организационное мероприятие
            int orgnums = 0;
            for (int org = 0; org < listOfOrgArrs.Count; org++)
            {
                if(listOfOrgArrs[org].isWork == true)
                { orgnums++; }
            }
            if (orgnums > 0)
            {
                // Проверка. Если работы на ВЛ не производятся, то должно быть выбрано только одно мероприятие (последнее)
                if ((listOfOrgArrs[listOfOrgArrs.Count - 1].isWork == true && orgnums == 1) || (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork == false))
                {

                    // Проверка на незаполненность полей у Персонала, ответственного за переключения
                    bool isEmptyFieldInListOfPresonal = false;
                    for (int p = 0; p < listOfPersonal.Count; p++)
                    {
                        // Если не заполнена организация
                        if (listOfPersonal[p].organisationOfPersonal == "")
                        { isEmptyFieldInListOfPresonal = true; }
                        else
                        {
                            for (int q = 0; q < listOfPersonal[p].Person.Count; q++)
                            {
                                // Если не заполнено имя или должность
                                if (listOfPersonal[p].Person[q].nameOfPerson == "" || listOfPersonal[p].Person[q].role == "")
                                { isEmptyFieldInListOfPresonal = true; }
                                if (listOfPersonal[p].organisationOfPersonal == "SO")
                                {
                                    // Если не заполнено действие у работников СО
                                    if (listOfPersonal[p].Person[q].action == "")
                                    { isEmptyFieldInListOfPresonal = true; }
                                }
                            }
                        }
                    }
                    if (isEmptyFieldInListOfPresonal == false)
                    {
                        //progressBar1.IsIndeterminate = true;
                        try
                        {
                            Microsoft.Office.Interop.Word.Application winword =
                                new Microsoft.Office.Interop.Word.Application();

                            winword.Visible = /*true*/false;

                            object missing = System.Reflection.Missing.Value;

                            //Создание нового документа
                            Microsoft.Office.Interop.Word.Document document =
                                winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                            //winword.Visible = true;
                            document.Content.SetRange(0, 0);

                            //Добавление текста
                            Microsoft.Office.Interop.Word.Paragraph sec0 = document.Content.Paragraphs.Add(ref missing);
                            document.Content.Paragraphs.LeftIndent = -28;
                            sec0.Range.Font.Size = 14;                          // Задаётся формат шрифта в таблице
                            sec0.Range.Font.Name = "Times New Roman";
                            //sec0.Format.SpaceBefore = 24;
                            document.PageSetup.TopMargin = 50;          // смещение документа вверх

                            Microsoft.Office.Interop.Word.Table tab0 = document.Tables.Add(sec0.Range, 2, 1, 2, ref missing);
                            tab0.Range.Font.Size = 14;
                            tab0.Range.Font.Name = "Times New Roman";
                            tab0.Borders.Enable = 0;
                            tab0.Cell(1, 1).Range.Text = "Дата составления программы:                            «___»______________   201__г.";
                            tab0.Cell(2, 1).Range.Text = "Дата производства переключений:                    «___»______________   201__г.";
                            //tab0.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 24;
                            tab0.Cell(2, 1).Range.ParagraphFormat.SpaceAfter = 24;
                            tab0.Cell(1, 1).Range.ParagraphFormat.SpaceBefore = 24;
                            tab0.Cell(2, 1).Range.ParagraphFormat.SpaceBefore = 24;
                            //winword.Selection.MoveDown(WdUnits.wdParagraph, 1, WdMovementType.wdMove);
                            /*winword.Selection.TypeParagraph();
                            winword.Selection.Paragraphs.SpaceAfter = 0;
                            sec0.SpaceAfterAuto = -1;
                            winword.Selection.ParagraphFormat.SpaceAfter = 24;*/
                            /*sec0.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            sec0.LineSpacing = 40;
                            sec0.Format.SpaceBefore = 50;
                            sec0.Format.SpaceAfter = 100;
                            //sec0.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;

                            //document.Content.Paragraphs.SpaceAfter = 24;*/

                            // Заполнение заглавия документа
                            Microsoft.Office.Interop.Word.Paragraph zaglav = document.Content.Paragraphs.Add(ref missing);
                            //zaglav.Format.LeftIndent = -30;
                            //zaglav.LeftIndent = -30;
                            Microsoft.Office.Interop.Word.Table tab01 = document.Tables.Add(zaglav.Range, 3, 1, ref missing, ref missing);
                            tab01.Borders.Enable = 0;
                            tab01.Range.Font.Size = 14;
                            tab01.Range.Font.Name = "Times New Roman";
                            //tab01.Columns.Borders.DistanceFromLeft = 20;
                            tab01.Rows.SetLeftIndent(-53, WdRulerStyle.wdAdjustNone);
                            tab01.Columns.SetWidth(546, WdRulerStyle.wdAdjustNone);
                            //tab01.Range.ParagraphFormat.LineSpacing = 1;
                            tab01.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            tab01.Range.ParagraphFormat.SpaceAfter = 0;
                            tab01.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            tab01.Cell(1, 1).Range.Text = "ПРОГРАММА ПЕРЕКЛЮЧЕНИЙ № 1";
                            tab01.Cell(1, 1).Range.Font.Bold = 1;

                            /*switch (comboBox1.SelectedIndex)
                            {
                                case 0:
                                    tab01.Cell(2, 1).Range.Text = "ПО ВЫВОДУ В РЕМОНТ";
                                    break;
                                case 1:
                                    tab01.Cell(2, 1).Range.Text = "ПО ВЫВОДУ В РЕЗЕРВ";
                                    break;
                                case 2:
                                    tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ ИЗ РЕЗЕРВА";
                                    break;
                                case 3:
                                    tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ В РАБОТУ";
                                    break;
                            }*/

                            switch (mainParamsOfSP.aim)
                            {
                                case "Вывод в ремонт":
                                    tab01.Cell(2, 1).Range.Text = "ПО ВЫВОДУ В РЕМОНТ";
                                    break;
                                case "Вывод в резерв":
                                    tab01.Cell(2, 1).Range.Text = "ПО ВЫВОДУ В РЕЗЕРВ";
                                    break;
                                case "Ввод из резерва":
                                    tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ ИЗ РЕЗЕРВА";
                                    break;
                                case "Ввод в работу":
                                    tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ В РАБОТУ";
                                    break;
                            }

                            tab01.Cell(3, 1).Range.Text = mainParamsOfSP.nameLine /*textBox2.Text*/ + "\n";
                            zaglav.Range.Text = "\n";
                            //winword.Selection.MoveDown(WdUnits.wdParagraph, 1, WdMovementType.wdMove);

                            // Заполнение пунктов 1 и 2
                            Microsoft.Office.Interop.Word.Paragraph para_12 = document.Content.Paragraphs.Add(ref missing);
                            para_12.LeftIndent = 20;
                            Microsoft.Office.Interop.Word.Table table12 = document.Tables.Add(para_12.Range, 4, 1, 2, ref missing);
                            table12.Borders.Enable = 0;
                            table12.Range.Font.Size = 13;
                            table12.Range.Font.Name = "Times New Roman";
                            table12.Cell(1, 1).Range.Select();          // Выделение текущей ячейки

                            Object unit = WdUnits.wdLine;               // Операции по перемещению выделения вверх
                            Object count = 1;
                            Object extend = WdMovementType.wdMove;
                            winword.Selection.MoveUp(ref unit, ref count, ref extend);
                            winword.Selection.Delete(WdUnits.wdCharacter, 1);           // Удалить 1 символ справа от выделения

                            table12.Rows.SetLeftIndent(-35, WdRulerStyle.wdAdjustNone);
                            table12.Columns.SetWidth(510, WdRulerStyle.wdAdjustNone);

                            table12.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            table12.Range.ParagraphFormat.SpaceAfter = 0;
                            table12.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                            table12.Cell(1, 1).Range.Text = "1. ЦЕЛЬ ПРОГРАММЫ:";
                            table12.Cell(1, 1).Range.Font.Bold = 1;
                            table12.Cell(2, 1).Range.Text = mainParamsOfSP.aim /*comboBox1.Text*/ + " " + mainParamsOfSP.nameLine /*textBox2.Text*/ + ".\n";
                            table12.Cell(3, 1).Range.Text = "2. ОБЪЕКТЫ ПЕРЕКЛЮЧЕНИЙ:";
                            table12.Cell(3, 1).Range.Font.Bold = 1;

                            string substs = "";
                            for (int r = 0; r < listOfPowerObjects.Count; r++)
                            {
                                if (listOfPowerObjects[r].isUsed == true)
                                {
                                    if (r != listOfPowerObjects.Count - 1)
                                    {
                                        substs = substs + listOfPowerObjects[r].NamePO + ", ";
                                    }
                                    else
                                    {
                                        substs = substs + listOfPowerObjects[r].NamePO + ".";
                                    }
                                }
                            }
                            table12.Cell(4, 1).Range.Text = substs + "\n";

                            // Заполнение пунктa 3
                            Microsoft.Office.Interop.Word.Paragraph para_03 = document.Content.Paragraphs.Add(ref missing);
                            para_03.LeftIndent = 20;
                            Microsoft.Office.Interop.Word.Table table03 = document.Tables.Add(para_03.Range, 2, 1, ref missing, ref missing);
                            table03.Borders.Enable = 0;
                            table03.Range.Font.Size = 13;
                            table03.Range.Font.Name = "Times New Roman";
                            table03.Rows.SetLeftIndent(-35, WdRulerStyle.wdAdjustNone);
                            table03.Columns.SetWidth(510, WdRulerStyle.wdAdjustNone);

                            table03.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            table03.Range.ParagraphFormat.SpaceAfter = 0;
                            table03.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            table03.Cell(1, 1).Range.Text = "3. УСЛОВИЯ ВЫПОЛНЕНИЯ ПЕРЕКЛЮЧЕНИЙ:";
                            table03.Range.Font.Bold = 1;
                            table03.Cell(2, 1).Range.Text = "3.1. Схемы объектов переключений:\n";
                            para_03.Range.Text = "\n";

                            Microsoft.Office.Interop.Word.Paragraph para_3 = document.Content.Paragraphs.Add(ref missing);
                            Microsoft.Office.Interop.Word.Table table3 = document.Tables.Add(para_3.Range, 1, 2, ref missing, ref missing);
                            table3.Borders.Enable = 1;
                            table3.Range.Font.Size = 13;
                            table3.Range.Font.Name = "Times New Roman";

                            table3.Rows.SetLeftIndent(-29, WdRulerStyle.wdAdjustNone);
                            //table3.Columns.SetWidth(510, WdRulerStyle.wdAdjustNone);
                            table3.Columns[1].SetWidth(88, WdRulerStyle.wdAdjustNone);
                            table3.Columns[2].SetWidth(432, WdRulerStyle.wdAdjustNone);
                            /*sec5.Columns[3].SetWidth(340, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[4].SetWidth(43, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[5].SetWidth(43, WdRulerStyle.wdAdjustNone);*/

                            table3.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                            table3.Cell(1, 1).Range.Select();          // Выделение текущей ячейки

                            winword.Selection.MoveUp(WdUnits.wdLine, 1, WdMovementType.wdMove); // Операции по перемещению выделения вверх
                            winword.Selection.Delete(WdUnits.wdCharacter, 1);           // Удалить 1 символ справа от выделения

                            table3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            table3.Range.ParagraphFormat.SpaceAfter = 0;
                            int line = 1;
                            int firstcell = 1;
                            for (int obj = 0; obj < listOfPowerObjects.Count; obj++)
                            {
                                string charForBold = "";                            // строка для подсчёта символов, выделяемых жирным шрифтом 
                                int forCharForBoldswi = 0;
                                int forCharForBolddis = 0;
                                int forCharForBoldgrd = 0;

                                if (listOfPowerObjects[obj].isUsed == true)
                                {
                                    List<PowerObject.Equipment> forTableOn = new List<PowerObject.Equipment>();
                                    List<PowerObject.Equipment> forTableOff = new List<PowerObject.Equipment>();
                                    for (int t = 0; t < One_listEquipment.Count; t++)
                                    {
                                        if (listOfPowerObjects[obj].NamePO == One_listEquipment[t].NamePO)
                                        {
                                            var eq = new PowerObject.Equipment();
                                            eq.nameEquip = One_listEquipment[t].nameEquip;
                                            eq.typeEquip = One_listEquipment[t].typeEquip;
                                            eq.stateEquip = One_listEquipment[t].stateEquip;

                                            if (One_listEquipment[t].stateEquip == true)
                                            {
                                                forTableOn.Add(eq);
                                            }
                                            else
                                            {
                                                forTableOff.Add(eq);
                                            }

                                        }
                                    }

                                    table3.Cell(firstcell, 1).Range.Text = listOfPowerObjects[obj].NamePO;
                                    if (forTableOn.Count > 0)
                                    {
                                        table3.Cell(firstcell, 2).Range.Text = "Включены:";
                                        table3.Cell(firstcell, 2).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                                        if (forTableOff.Count > 0)
                                        {
                                            table3.Rows.Add();
                                            line++;
                                        }

                                        string swi = "Выключатели ";
                                        string dis = "Разъединители ";
                                        string grd = "Заземляющие ножи ";

                                        for (int s = 0; s < forTableOn.Count; s++)
                                        {
                                            switch (forTableOn[s].typeEquip)
                                            {
                                                case "Switch":
                                                    swi = swi + forTableOn[s].nameEquip + ", ";
                                                    charForBold = swi;
                                                    forCharForBoldswi = charForBold.Length - 12;
                                                    break;
                                                case "Disconnector":
                                                    dis = dis + forTableOn[s].nameEquip + ", ";
                                                    charForBold = dis;
                                                    forCharForBolddis = charForBold.Length - 14;
                                                    break;
                                                case "GroundDisconnector":
                                                    grd = grd + forTableOn[s].nameEquip + ", ";
                                                    charForBold = grd;
                                                    forCharForBoldgrd = charForBold.Length - 17;
                                                    break;
                                            }
                                        }
                                        if (swi != "Выключатели ")
                                        {
                                            //text = text.Substring(0, text.Length - 2);
                                            swi = swi.Substring(0, swi.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = swi + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBoldswi - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1 || forTableOff.Count > 0 || dis != "Разъединители " || grd != "Заземляющие ножи ")
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                        if (dis != "Разъединители ")
                                        {
                                            dis = dis.Substring(0, dis.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = dis + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBolddis - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1 || forTableOff.Count > 0 || grd != "Заземляющие ножи ")
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                        if (grd != "Заземляющие ножи ")
                                        {
                                            grd = grd.Substring(0, grd.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = grd + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBoldgrd - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1 || forTableOff.Count > 0)
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                    }
                                    if (forTableOff.Count > 0)
                                    {
                                        table3.Cell(firstcell + 1, 2).Range.Text = "Отключены:";
                                        table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                                        table3.Rows.Add();
                                        table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                        table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                        line++;


                                        string swi2 = "Выключатели ";
                                        string dis2 = "Разъединители ";
                                        string grd2 = "Заземляющие ножи ";

                                        for (int s = 0; s < forTableOff.Count; s++)
                                        {

                                            switch (forTableOff[s].typeEquip)
                                            {
                                                case "Switch":
                                                    swi2 = swi2 + forTableOff[s].nameEquip + ", ";
                                                    charForBold = swi2;
                                                    forCharForBoldswi = charForBold.Length - 12;
                                                    break;
                                                case "Disconnector":
                                                    dis2 = dis2 + forTableOff[s].nameEquip + ", ";
                                                    charForBold = dis2;
                                                    forCharForBolddis = charForBold.Length - 14;
                                                    break;
                                                case "GroundDisconnector":
                                                    grd2 = grd2 + forTableOff[s].nameEquip + ", ";
                                                    charForBold = grd2;
                                                    forCharForBoldgrd = charForBold.Length - 17;
                                                    break;
                                            }
                                        }
                                        if (swi2 != "Выключатели ")
                                        {
                                            swi2 = swi2.Substring(0, swi2.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = swi2 + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBoldswi - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1 || dis2 != "Разъединители " || grd2 != "Заземляющие ножи ")
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                        if (dis2 != "Разъединители ")
                                        {
                                            dis2 = dis2.Substring(0, dis2.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = dis2 + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBolddis - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1 || grd2 != "Заземляющие ножи ")
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                        if (grd2 != "Заземляющие ножи ")
                                        {
                                            grd2 = grd2.Substring(0, grd2.Length - 2);
                                            table3.Cell(firstcell + 1, 2).Range.Text = grd2 + ".";

                                            // Выделение жирным названия оборудования
                                            table3.Cell(firstcell + 1, 2).Range.Select();
                                            Object unit3 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            Object count3 = 1;
                                            Object extend3 = WdMovementType.wdMove;
                                            winword.Selection.MoveRight(ref unit3, ref count3, ref extend3);
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            count3 = forCharForBoldgrd - 1;
                                            extend3 = WdMovementType.wdExtend;
                                            winword.Selection.MoveLeft(ref unit3, ref count3, ref extend3);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                            table3.Cell(firstcell + 1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                            if (obj != listOfPowerObjects.Count - 1/* || grd2 != "Заземляющие ножи "*/)
                                            {
                                                table3.Rows.Add();
                                                line++;
                                            }
                                            //table3.Rows.Add();
                                            table3.Cell(firstcell, 1).Merge(table3.Cell(firstcell + 1, 1));
                                            table3.Cell(firstcell, 2).Merge(table3.Cell(firstcell + 1, 2));
                                            //line++;

                                        }
                                    }
                                    firstcell++;
                                }
                            }
                            
                            // Проверка наличия текста в последней строке таблицы и удаление её, если там пусто                            
                            string row1 = table3.Cell(firstcell,1).Range.Text;
                            string row2 = table3.Cell(firstcell, 1).Range.Text;                            
                            if (row1 != "" && row1 != "\r\a" && row1 != null &&
                                row2 != "" && row2 != "\r\a" && row2 != null)
                            { }
                            else
                            {                             
                                table3.Cell(firstcell, 1).Range.Select();                                
                                winword.Selection.Cells.Delete(WdDeleteCells.wdDeleteCellsEntireRow); // Удаление лишней строки в таблице                                
                            }                           
                            
                            para_03.Range.Text = "\n";

                            // Заполнение пунктов 3.1 (конец), 3.2, 3.3, 3.4
                            Microsoft.Office.Interop.Word.Paragraph para_321 = document.Content.Paragraphs.Add(ref missing);

                            Microsoft.Office.Interop.Word.Table table32 = document.Tables.Add(para_321.Range, 1, 1, ref missing, ref missing);
                            table32.Borders.Enable = 0;
                            table32.Range.Font.Size = 13;
                            table32.Range.Font.Name = "Times New Roman";
                            table32.Rows.SetLeftIndent(-35, WdRulerStyle.wdAdjustNone);
                            //table32.Columns.SetWidth(500, WdRulerStyle.wdAdjustNone);
                            table32.Columns.SetWidth(537, WdRulerStyle.wdAdjustNone);

                            table32.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            table32.Range.ParagraphFormat.SpaceAfter = 0;
                            table32.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                            table32.Cell(1, 1).Range.Select();
                            Object unit1 = WdUnits.wdLine;               // Операции по перемещению выделения вверх
                            Object count1 = 1;
                            Object extend1 = WdMovementType.wdMove;
                            winword.Selection.MoveUp(ref unit, ref count, ref extend);
                            winword.Selection.Delete(WdUnits.wdCharacter, 1);                   // Удаляется строка сверху


                            table32.Cell(1, 1).Range.Text = "Состав защит " + mainParamsOfSP.nameLine /*textBox2.Text*/ +
                                " соответствует «Инструкции по обслуживанию устройств релейной защиты и сетевой автоматики ЛЭП, находящихся в диспетчерском управлении РДУ».";
                            table32.Cell(1, 1).Range.Font.Bold = 1;

                            table32.Rows.Add();

                            table32.Cell(1, 1).Range.ParagraphFormat.LineUnitAfter = 1;
                            table32.Cell(1, 1).Range.ParagraphFormat.LineUnitBefore = 1;

                            // Пункт 3.2
                            Object unit32 = WdUnits.wdCharacter;               // Операции по перемещению выделения вверх                
                            Object extend32 = WdMovementType.wdMove;
                            Object count32 = 3;
                            if (mainParamsOfSP.inducedVoltage == true/*(bool)checkBox1.IsChecked*/)
                            {
                                table32.Cell(2, 1).Range.Text = "3.2. Наличие наведённого напряжения после отключения и заземления в РУ: Да.";
                                count32 = 3;
                            }
                            else
                            {
                                table32.Cell(2, 1).Range.Text = "3.2. Наличие наведённого напряжения после отключения и заземления в РУ: Нет.";
                                count32 = 4;
                            }
                            table32.Rows.Add();

                            table32.Cell(2, 1).Range.Select();
                            winword.Selection.MoveRight(ref unit32, 1, ref extend1);
                            winword.Selection.MoveLeft(ref unit32, 1, ref extend1);
                            extend32 = WdMovementType.wdExtend;
                            winword.Selection.MoveLeft(ref unit32, ref count32, ref extend32);
                            winword.Selection.Font.Bold = 0;

                            // Пункт 3.3
                            if (mainParamsOfSP.isUsedARM == true/*(bool)checkBox2.IsChecked*/)
                            {
                                table32.Cell(3, 1).Range.Text = "3.3. Выполнение переключений с использованием АРМ: Да.";
                                count32 = 3;
                            }
                            else
                            {
                                table32.Cell(3, 1).Range.Text = "3.3. Выполнение переключений с использованием АРМ: Нет.";
                                count32 = 4;
                            }
                            table32.Rows.Add();

                            table32.Cell(3, 1).Range.Select();
                            winword.Selection.MoveRight(ref unit32, 1, ref extend1);
                            winword.Selection.MoveLeft(ref unit32, 1, ref extend1);
                            extend32 = WdMovementType.wdExtend;
                            winword.Selection.MoveLeft(ref unit32, ref count32, ref extend32);
                            winword.Selection.Font.Bold = 0;

                            // Пункт 3.4
                            if (mainParamsOfSP.ferroresonance == true/*(bool)checkBox3.IsChecked*/)
                            {
                                table32.Cell(4, 1).Range.Text = "3.4. Имеется возможность возникновения феррорезонанса.";
                            }
                            else
                            {
                                table32.Cell(4, 1).Range.Text = "3.4. Отсутствует возможность возникновения феррорезонанса.";
                            }


                            GoToNextPage(winword, document, missing);

                            //table32.Rows.Add();

                            /*table32.Cell(4, 1).Range.Select();
                            winword.Selection.MoveRight(ref unit32, 1, ref extend32);
                            winword.Selection.MoveLeft(ref unit32, 1, ref extend32);
                            extend32 = WdMovementType.wdExtend;
                            winword.Selection.MoveLeft(ref unit32, ref count32, ref extend32);
                            winword.Selection.Font.Bold = 0;*/


                            /*table3.Rows.Add();
                            int iinntt = table3.Rows.Count;
                            table3.Rows[iinntt].HeightRule = WdRowHeightRule.wdRowHeightAuto;                
                            table3.Cell(iinntt,1).Merge(table3.Cell(iinntt, 2));
                            table3.Cell(iinntt, 1).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                            table3.Cell(iinntt, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            table3.Cell(iinntt,1).Range.Text = "Состав защит " + textBox2.Text + 
                                " соответствует «Инструкции по обслуживанию устройств релейной защиты и сетевой автоматики ЛЭП, находящихся в диспетчерском управлении РДУ».";
                            table3.Cell(iinntt, 1).Range.Font.Bold = 1;

                            table3.Rows.Add();
                            table3.Cell(iinntt,1).Range.ParagraphFormat.LineUnitAfter = 1;
                            table3.Cell(iinntt, 1).Range.ParagraphFormat.LineUnitBefore = 1;

                            /*para_03.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                            para_03.Range.Font.Name = "Times New Roman";
                            //para_03.Range.ParagraphFormat.LineUnitAfter = 1;
                            //para_03.Range.ParagraphFormat.LineUnitBefore = 1;
                            //para_03.Range.Font.Bold = 1;

                            para_03.Range.Text = "\n";

                            Object unit2 = WdUnits.wdLine;               // Операции по перемещению выделения вверх
                            Object count2 = 1;
                            Object extend2 = WdMovementType.wdMove;
                            winword.Selection.MoveUp(ref unit, ref count, ref extend);


                            para_03.Range.Text = "\n" + "Состав защит " +
                                textBox2.Text + " соответствует «Инструкции по обслуживанию устройств релейной защиты и сетевой автоматики ЛЭП, находящихся в диспетчерском управлении РДУ».";
                            para_03.Range.Select();          // Выделение текущей ячейки

                            winword.Selection.ParagraphFormat.LineUnitAfter = 1;        // добавляются отсутпы сверху и снизу
                            winword.Selection.ParagraphFormat.LineUnitBefore = 1;
                            winword.Selection.Font.Bold = 1;                            // Шрифт жирный
                            winword.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            winword.Selection.ParagraphFormat.LineSpacing = 11;
                            winword.Selection.ParagraphFormat.LeftIndent = 20;
                            winword.Selection.ParagraphFormat.RightIndent = 20;

                            Object unit1 = WdUnits.wdLine;               // Операции по перемещению выделения вверх
                            Object count1 = 1;
                            Object extend1 = WdMovementType.wdMove;
                            winword.Selection.MoveUp(ref unit, ref count, ref extend);
                            winword.Selection.Delete(WdUnits.wdCharacter, 1);                   // Удаляется строка сверху
                            para_03.Range.Select();
                            winword.Selection.MoveDown(ref unit, ref count, ref extend);               
                            para_03.Range.Text = para_03.Range.Text+"\n";

                            para_03.Range.Font.Bold = 0;                                 // Шрифт обычный



                            */

                            /*Microsoft.Office.Interop.Word.Paragraph para_32 = document.Content.Paragraphs.Add(ref missing);
                            para_32.LeftIndent = 20;
                            para_32.Range.Text = "23456";
                            para_32.Range.Select();
                            //Microsoft.Office.Interop.Word.Paragraph para_32 = document.Content.Paragraphs.Add(ref missing);
                            para_32.LeftIndent = 20;
                            para_32.Range.Text = "23456";*/
                            /*Microsoft.Office.Interop.Word.Table table32 = document.Tables.Add(para_32.Range, 2, 1, ref missing, ref missing);
                            table32.Borders.Enable = 0;
                            table32.Range.Font.Size = 13;
                            table32.Range.Font.Name = "Times New Roman";
                            table32.Rows.SetLeftIndent(-35, WdRulerStyle.wdAdjustNone);
                            table32.Columns.SetWidth(510, WdRulerStyle.wdAdjustNone);

                            table32.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            table32.Range.ParagraphFormat.SpaceAfter = 0;
                            table32.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;


                            table32.Cell(1, 1).Range.Text = "3. УСЛОВИЯ ВЫПОЛНЕНИЯ ПЕРЕКЛЮЧЕНИЙ:";
                            table32.Range.Font.Bold = 1;*/




                            // ТАБЛИЦА ДЛЯ РАЗДЕЛОВ 4-7

                            //Добавление текста
                            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                            para1.Range.Text = " ";


                            para1.Format.LeftIndent = -30;    // Сдвиг всей таблицы влево
                                                              //para1.Range.ParagraphFormat.LeftIndent = -40;
                                                              //para1.Range.InsertParagraphAfter();
                            Microsoft.Office.Interop.Word.Table sec5 = document.Tables.Add(para1.Range, 1, 5, 2, ref missing);
                            // Таблица
                            para1.Format.LeftIndent = -30;

                            sec5.Borders.Enable = 1;

                            sec5.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                            sec5.Range.Font.Name = "Times New Roman";

                            /*para1.Range.Font.Size = 13;
                            para1.Range.Font.Name = "Times New Roman";*/


                            // Хорошие размеры таблицы
                            sec5.Columns[1].SetWidth(63, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[2].SetWidth(42, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[3].SetWidth(340, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[4].SetWidth(43, WdRulerStyle.wdAdjustNone);
                            sec5.Columns[5].SetWidth(43, WdRulerStyle.wdAdjustNone);

                            sec5.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                            // Шапка таблицы для пунктов 4 - 7
                            //sec5.Rows[1].Cells[1].Range.ParagraphFormat.LeftIndent = -5;        // Смещение текста в ячейке влево
                            sec5.Cell(1, 1).Range.ParagraphFormat.LeftIndent = -5;
                            //sec5.Rows[1].Cells[1].Range.Text = "Персонал,\nвыполняющий операцию";
                            sec5.Cell(1, 1).Range.Text = "Персонал,\nвыполняющий операцию";
                            //sec5.Rows[1].Cells[2].Range.Text = "п.п.";
                            sec5.Cell(1, 2).Range.Text = "п.п.";
                            //sec5.Rows[1].Cells[3].Range.Text = "Объект переключений,\nоперация, сообщение";
                            sec5.Cell(1, 3).Range.Text = "Объект переключений,\nоперация, сообщение";
                            //sec5.Rows[1].Cells[4].Range.Text = "Время\nотдачи команды";
                            sec5.Cell(1, 4).Range.Text = "Время\nотдачи команды";
                            //sec5.Rows[1].Cells[5].Range.Text = "Время\nвыполнения\nкоманды";
                            sec5.Cell(1, 5).Range.Text = "Время\nвыполнения\nкоманды";
                            sec5.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            sec5.Rows[1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            sec5.Rows.Add();

                            //sec5.Cell(2, 1).Range.Text = "testure";

                            int nORs = 0;          //numberOfRows
                            nORs = sec5.Rows.Count;

                            // Заполнение 5 пункта Программы переключений
                            //sec5.Rows[nORs].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            sec5.Cell(nORs, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            sec5.Cell(nORs, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            sec5.Cell(nORs, 3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            sec5.Cell(nORs, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            sec5.Cell(nORs, 5).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            //section5.Rows[nORs].Cells[2].Range.Text = "5.";




                            // ПУНКТ 4
                            sec5.Cell(nORs, 2).Range.Text = "4.";
                            sec5.Rows.Add();
                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                            sec5.Cell(nORs, 3).Range.Text = "МЕРОПРИЯТИЯ ПО ПОДГОТОВКЕ К ВЫПОЛНЕНИЮ ПЕРЕКЛЮЧЕНИЙ";
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                            nORs++;

                            int powObjOn = 0;
                            int numnum = 0;

                            // Если операции на ВЛ не будут производиться (выбран последний пункт)
                            if (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork == true)
                            {
                                /*int orgArrsIsWork = 0;
                                for (int i = 0; i < listOfOrgArrs.Count - 1; i++)
                                {
                                    if (listOfOrgArrs[i].isWork == true)
                                    {
                                        orgArrsIsWork++;
                                    }
                                }
                                if (orgArrsIsWork != 0)
                                {

                                }*/
                            }
                            else
                            {
                                sec5.Cell(nORs, 2).Range.Text = "4.1.";
                                sec5.Rows.Add();
                                sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                                sec5.Cell(nORs, 3).Range.Text = "ОРГАНИЗАЦИОННЫЕ";
                                sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                nORs++;

                                sec5.Rows.Add();
                                int numbOfOperation = 0;
                                int rowForDelite = 0;

                                if (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork != true)
                                {
                                    // Если работы на ВЛ

                                    int lld = listOfOrgArrs.Count - 1;

                                    int numeratonOfBullits = 1;
                                    // Если будет производиться работа на ВЛ, то...
                                    if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true)
                                    {
                                        numnum++;
                                        sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.lineOrganisation /*textBox3.Text*/;
                                        sec5.Cell(nORs, 2).Range.Text = "4.1." + numnum + ".";
                                        numeratonOfBullits++;
                                        sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                        sec5.Cell(nORs, 3).Range.Text = "Подтверждение готовности к производству линейных работ на " +
                                            mainParamsOfSP.nameLine /*textBox2.Text*/ + " (при наличии диспетчерской заявки)";
                                        sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                        nORs++;

                                        sec5.Rows.Add();

                                        /*
                                        // Подтверждения выполнения линейных работ на ВЛ до ЛР на энергообъектах
                                        {
                                            rowForDelite = nORs - 1;
                                            int forpmes = 1; // переменная для отображения названия организации, если нет галочки работы на ВЛ

                                            for (int k = 0; k < listOfPowerObjects.Count; k++)   // Проверка выполнения работ на Энергообъектах
                                            {
                                                if (listOfPowerObjects[k].isUsed == true)
                                                {
                                                    numnum++;
                                                    string nameDisconnector = "";
                                                    for (int j = 0; j < One_listEquipment.Count; j++)
                                                    {
                                                        if (One_listEquipment[j].NamePO == listOfPowerObjects[k].NamePO &&
                                                            One_listEquipment[j].typeEquip == "Disconnector")
                                                        { nameDisconnector = One_listEquipment[j].nameEquip; }
                                                    }
                                                    if (nameDisconnector != "")
                                                    {
                                                        if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork != true && forpmes == 1)
                                                        {
                                                            sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.lineOrganisation;
                                                        }
                                                        forpmes++;
                                                        sec5.Cell(nORs, 2).Range.Text = "4.1." + Convert.ToString(numnum) + ".";
                                                        sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                        sec5.Cell(nORs, 3).Range.Text = "Подтверждение готовности к производству работ на участке " +
                                                                        mainParamsOfSP.nameLine + " в пределах " + listOfPowerObjects[k].NamePO +
                                                                        " до линейного разъединителя " + nameDisconnector +
                                                                        " (при наличии диспетчерской заявки)";

                                                        // Выделение названия оборудованияжирным шрифтом
                                                        int charsInLine = mainParamsOfSP.nameLine.Length;
                                                        int charsInPO = listOfPowerObjects[k].NamePO.Length;
                                                        int charsInEquip = nameDisconnector.Length;

                                                        int commonChars = 57 + charsInLine + 12 + charsInPO + 28;

                                                        Object unit41 = WdUnits.wdCharacter;               // Операции по перемещению выделения вверх                
                                                        Object extend41 = WdMovementType.wdMove;
                                                        Object count41 = charsInEquip;

                                                        sec5.Cell(nORs, 3).Range.Select();
                                                        winword.Selection.MoveLeft(ref unit41, 1, ref extend1);
                                                        winword.Selection.MoveRight(ref unit41, commonChars, ref extend1);
                                                        extend41 = WdMovementType.wdExtend;
                                                        winword.Selection.MoveRight(ref unit41, ref count41, ref extend41);
                                                        winword.Selection.Font.Bold = 1;
                                                        winword.Selection.Font.Italic = 1;

                                                        sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                                        if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true || forpmes != 2)
                                                        {
                                                            if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true)
                                                                sec5.Cell(rowForDelite, 1).Merge(sec5.Cell(nORs, 1));
                                                            else
                                                                sec5.Cell(rowForDelite + 1, 1).Merge(sec5.Cell(nORs, 1));
                                                        }

                                                        nORs++;
                                                        sec5.Rows.Add();
                                                    }
                                                }
                                            }
                                        }*/
                                    }

                                    for (int i = 0; i < listOfPowerObjects.Count; i++) // Считается количество включенных ПС
                                    {
                                        if (listOfPowerObjects[i].isUsed == true)
                                        {
                                            powObjOn++;
                                        }
                                    }
                                    if (powObjOn != 0) // Если включенных ПС не равно нулю, то
                                    {
                                        /*if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork != true)
                                        {*/
                                            rowForDelite = nORs - 1;
                                            int forpmes = 1; // переменная для отображения названия организации, если нет галочки работы на ВЛ

                                            for (int k = 0; k < listOfOrgArrs.Count - 2; k++)   // Проверка выполнения работ на Энергообъектах
                                            {
                                                if (listOfOrgArrs[k].isWork == true)
                                                {
                                                    numnum++;
                                                    string nameDisconnector = "";
                                                    for (int j = 0; j < One_listEquipment.Count; j++)
                                                    {
                                                        if (One_listEquipment[j].NamePO == listOfOrgArrs[k].PObject &&
                                                            One_listEquipment[j].typeEquip == "Disconnector")
                                                        { nameDisconnector = One_listEquipment[j].nameEquip; }
                                                    }
                                                    if (nameDisconnector != "")
                                                    {
                                                        if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork != true && forpmes == 1)
                                                        {
                                                            sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.lineOrganisation;
                                                        }
                                                        forpmes++;
                                                        sec5.Cell(nORs, 2).Range.Text = "4.1." + Convert.ToString(numnum) + ".";
                                                        sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                        sec5.Cell(nORs, 3).Range.Text = "Подтверждение готовности к производству работ на участке " +
                                                                        mainParamsOfSP.nameLine + " в пределах " + listOfOrgArrs[k].PObject +
                                                                        " до линейного разъединителя " + nameDisconnector +
                                                                        " (при наличии диспетчерской заявки)";

                                                        // Выделение названия оборудованияжирным шрифтом
                                                        int charsInLine = mainParamsOfSP.nameLine.Length;
                                                        int charsInPO = listOfOrgArrs[k].PObject.Length;
                                                        int charsInEquip = nameDisconnector.Length;

                                                        int commonChars = 57 + charsInLine + 12 + charsInPO + 28;

                                                        Object unit41 = WdUnits.wdCharacter;               // Операции по перемещению выделения вверх                
                                                        Object extend41 = WdMovementType.wdMove;
                                                        Object count41 = charsInEquip;

                                                        sec5.Cell(nORs, 3).Range.Select();
                                                        winword.Selection.MoveLeft(ref unit41, 1, ref extend1);
                                                        winword.Selection.MoveRight(ref unit41, commonChars, ref extend1);
                                                        extend41 = WdMovementType.wdExtend;
                                                        winword.Selection.MoveRight(ref unit41, ref count41, ref extend41);
                                                        winword.Selection.Font.Bold = 1;
                                                        winword.Selection.Font.Italic = 1;

                                                        sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                                        if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true || forpmes != 2)
                                                        {
                                                            if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true)
                                                                sec5.Cell(rowForDelite, 1).Merge(sec5.Cell(nORs, 1));
                                                            else
                                                                sec5.Cell(rowForDelite + 1, 1).Merge(sec5.Cell(nORs, 1));
                                                        }

                                                        nORs++;
                                                        sec5.Rows.Add();
                                                    }
                                                }
                                            }
                                        /*}*/
                                    }
                                }
                            }

                            if (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork == true)
                            {
                                sec5.Rows.Add();
                            }

                            sec5.Cell(nORs, 1).Select();
                            winword.Selection.Cells.Delete(WdDeleteCells.wdDeleteCellsEntireRow); // Удаление лишней строки в таблице

                            // Если работы производятся на ВЛ
                            if (powObjOn != 0) // Если включенных ПС не равно нулю, то
                            {
                                if (listOfOrgArrs[listOfOrgArrs.Count -2].isWork == true)
                                {
                                    for (int i = 0; i < listOfPowerObjects.Count; i++)
                                    {
                                        if (listOfPowerObjects[i].isUsed == true)
                                        {
                                            sec5.Rows.Add();


                                            numnum++;
                                            sec5.Cell(nORs, 1).Range.Text = listOfOrgArrs[i].PObject/*listOfPowerObjects[i].NamePO*/;
                                            sec5.Cell(nORs, 2).Range.Text = "4.1." + Convert.ToString(numnum) + ".";
                                            string operation = "";
                                            switch (mainParamsOfSP.aim/*comboBox1.SelectedIndex*/)
                                            {
                                                case "Вывод в ремонт":
                                                    operation = "выводу в ремонт";
                                                    break;
                                                case "Вывод в резерв":
                                                    operation = "выводу в резерв";
                                                    break;
                                                case "Ввод из резерва":
                                                    operation = "вводу из резерва";
                                                    break;
                                                case "Ввод в работу":
                                                    operation = "вводу в работу";
                                                    break;
                                            }
                                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                            sec5.Cell(nORs, 3).Range.Text = "Подтверждение:\n" +
                                                "- наличия проверенного бланка переключений и возможности применения указанной в нем последовательности переключений по " +
                                                operation + " " + mainParamsOfSP.nameLine /*textBox2.Text*/ + ";\n" + "- готовности персонала к производству переключений";
                                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                                            nORs++;
                                        }
                                    }
                                }
                            }

                            // Информация по участвующим подстанциям
                            if (powObjOn != 0) // Если включенных ПС не равно нулю, то
                            {
                                if (listOfOrgArrs[listOfOrgArrs.Count-2].isWork != true) // Если галочка "Работы на ВЛ" не включена
                                {
                                    for (int i = 0; i < listOfOrgArrs.Count - 2; i++)
                                    {
                                        if (listOfOrgArrs[i].isWork == true)
                                        {
                                            sec5.Rows.Add();


                                            numnum++;
                                            sec5.Cell(nORs, 1).Range.Text = listOfOrgArrs[i].PObject/*listOfPowerObjects[i].NamePO*/;
                                            sec5.Cell(nORs, 2).Range.Text = "4.1." + Convert.ToString(numnum) + ".";
                                            string operation = "";
                                            switch (mainParamsOfSP.aim/*comboBox1.SelectedIndex*/)
                                            {
                                                case "Вывод в ремонт":
                                                    operation = "выводу в ремонт";
                                                    break;
                                                case "Вывод в резерв":
                                                    operation = "выводу в резерв";
                                                    break;
                                                case "Ввод из резерва":
                                                    operation = "вводу из резерва";
                                                    break;
                                                case "Ввод в работу":
                                                    operation = "вводу в работу";
                                                    break;
                                            }
                                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                            sec5.Cell(nORs, 3).Range.Text = "Подтверждение:\n" +
                                                "- наличия проверенного бланка переключений и возможности применения указанной в нем последовательности переключений по " +
                                                operation + " " + mainParamsOfSP.nameLine /*textBox2.Text*/ + ";\n" + "- готовности персонала к производству переключений";
                                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                                            nORs++;
                                        }
                                    }
                                }                                                            
                            }

                            string secondBullit = "1";
                            if (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork == true)
                            {
                                secondBullit = "1";
                                //sec5.Rows.Add();
                            }
                            else { secondBullit = "2"; }

                            sec5.Rows.Add();
                            sec5.Cell(nORs, 2).Range.Text = "4." + secondBullit + ".";
                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            sec5.Cell(nORs, 3).Range.Text = "РЕЖИМНЫЕ:\n(ДОПУСТИМЫЕ ПЕРЕТОКИ ПО КОНТРОЛИРУЕМЫМ СЕЧЕНИЯМ И ЛЭП НА ВРЕМЯ ПЕРЕКЛЮЧЕНИЙ)";
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                            nORs++;
                            sec5.Cell(nORs, 3).Range.Font.Bold = 0;

                            sec5.Rows.Add();

                            sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.dispOffice /*textBox1.Text*/;
                            sec5.Cell(nORs, 2).Range.Text = "4." + secondBullit + ".1.";
                            string operation2 = "";
                            switch (mainParamsOfSP.aim/*comboBox1.SelectedIndex*/)
                            {
                                case "Вывод в ремонт":
                                    operation2 = "выводу в ремонт";
                                    break;
                                case "Вывод в резерв":
                                    operation2 = "выводу в резерв";
                                    break;
                                case "Ввод из резерва":
                                    operation2 = "вводу из резерва";
                                    break;
                                case "Ввод в работу":
                                    operation2 = "вводу в работу";
                                    break;
                            }
                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            sec5.Cell(nORs, 3).Range.Text = "Перед началом производства переключений по " + operation2 +
                                " " + mainParamsOfSP.nameLine/*textBox2.Text*/ + " перетоки по контролируемым сечениям, а также токи по ЛЭП и электросетевому " +
                                "оборудованию должны быть приведены в соответствие с режимными указаниями к диспетчерской заявке. " +
                                "При отсутствии в диспетчерской заявке режимных указаний перетоки по контролируемым сечениям, а " +
                                "также токи по ЛЭП и электросетевому оборудованию должны соответствовать указаниям Положения по " +
                                "управлению режимами работы в операционной зоне Филиала АО «СО ЕЭС» " + progOption.roditPadezh;
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                            nORs++;


                            // ПУНКТ 5
                            sec5.Cell(nORs, 2).Range.Text = "5.";

                            sec5.Rows.Add();

                            //section5.Rows[nORs].Cells[3].Range.Font.Bold = 1;
                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                            //section5.Rows[nORs].Cells[3].Range.Text = "ПОРЯДОК И ПОСЛЕДОВАТЕЛЬНОСТЬ ВЫПОЛНЕНИЯ ОПЕРАЦИЙ:";
                            sec5.Cell(nORs, 3).Range.Text = "ПОРЯДОК И ПОСЛЕДОВАТЕЛЬНОСТЬ ВЫПОЛНЕНИЯ ОПЕРАЦИЙ:";

                            //section5.Rows[nORs].Cells[4].Merge(section5.Rows[nORs].Cells[5]);   // Объединение 4 и 5 ячейки в строке
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));

                            //nORs++;


                            // Заполнение из actionsList
                            for (int i = 0; i < actionsList.Count; i++)
                            {
                                int index = sec5.Rows.Count;
                                /*if (i != 0)
                                {
                                    nORs++;                        
                                }
                                sec5.Rows.Add();
                                    index = sec5.Rows.Count;*/
                                if (actionsList[i].FlName != "Начало")
                                {
                                    nORs++;
                                    sec5.Rows.Add();
                                    index = sec5.Rows.Count;
                                    //section5.Rows[nORs].Cells[1].Range.Text = actionsList[i].FlName;
                                    sec5.Cell(nORs, 1).Range.Text = actionsList[i].FlName;
                                    //section5.Rows.Add();                        

                                    for (int stolb = 0; stolb < 4; stolb++)     // Объединение столбцов встороке
                                    {
                                        sec5.Cell(nORs, 1).Merge(sec5.Cell(nORs, 2));
                                    }
                                }
                                if (actionsList[i].SecondLevelList.Count > 0)
                                {
                                    for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                                    {
                                        nORs++;
                                        sec5.Rows.Add();
                                        index = sec5.Rows.Count;

                                        int firstRowForMerge = nORs;
                                        //section5.Rows[nORs].Cells[1].Range.Text = actionsList[i].SecondLevelList[j].SlName;
                                        sec5.Cell(nORs, 1).Range.Text = actionsList[i].SecondLevelList[j].SlName;
                                        if (actionsList[i].SecondLevelList[j].slCommand == mainParamsOfSP.typeDO/*comboBox3.Text*/)
                                        {
                                            //section5.Rows[nORs].Cells[4].Merge(section5.Rows[nORs].Cells[5]);   // Объединение 4 и 5 ячейки в строке
                                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));       // Объединение 4 и 5 ячейки в строке
                                        }

                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList.Count > 0)
                                        {
                                            for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                                            {
                                                if (k != 0)
                                                {
                                                    nORs++;
                                                    sec5.Rows.Add();
                                                    index = sec5.Rows.Count;
                                                }
                                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated == true)
                                                {
                                                    index = sec5.Rows.Count;
                                                    //section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber;
                                                    sec5.Cell(nORs, 2).Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber;

                                                    // Удаление в стретьем столбце номера действия
                                                    string text3 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                                    int itNumb = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber.Length;
                                                    text3 = text3.Remove(0, itNumb);
                                                    //section5.Rows[nORs].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                    sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                    //section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                                    sec5.Cell(nORs, 3).Range.Text = text3 /*actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName*/;

                                                    // Здесь происходит поиск названия оборудования и выделения его жирным шрифтом
                                                    if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isConsistEquip == true)
                                                    {
                                                        string s1 = text3;
                                                        string s2 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].equipmentName;

                                                        Object unit5 = WdUnits.wdCharacter;
                                                        Object extend5 = WdMovementType.wdMove;
                                                        Object count5 = s1.IndexOf(s2);

                                                        sec5.Cell(nORs, 3).Range.Select();
                                                        winword.Selection.MoveLeft(ref unit32, 1, ref extend1);
                                                        winword.Selection.MoveRight(ref unit5, count5, ref extend5);

                                                        Object extend51 = WdMovementType.wdExtend;
                                                        winword.Selection.MoveRight(ref unit5, s2.Length, ref extend51);
                                                        winword.Selection.Font.Bold = 1;
                                                        winword.Selection.Font.Italic = 1;
                                                    }

                                                    string gng = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                                    int lastNors = nORs;

                                                    //nORs++;
                                                    /*if (k != actionsList[i].SecondLevelList[j].ThirdLevelList.Count - 1)
                                                    {
                                                        sec5.Rows.Add();
                                                        index = sec5.Rows.Count;
                                                    }*/
                                                    //int intovich = sec5.Rows.Count;

                                                    /*if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isConsistEquip == true)
                                                    {
                                                        //section5.Rows[lastNors].Cells[4].Merge(section5.Rows[lastNors].Cells[5]);   // Объединение 4 и 5 ячейки в строке
                                                        sec5.Cell(lastNors, 4).Merge(sec5.Cell(lastNors, 5));       // Объединение 4 и 5 ячейки в строке
                                                    }*/
                                                }
                                                else
                                                {
                                                    //section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                                    sec5.Cell(nORs, 2).Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;

                                                    for (int qw = 0; qw < 3; qw++)
                                                    {
                                                        sec5.Cell(nORs, 2).Merge(sec5.Cell(nORs, 3));
                                                    }

                                                    //section5.Rows[nORs].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                    sec5.Cell(nORs, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                                    if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count > 0)
                                                    {
                                                        //nORs = section5.Rows.Count;
                                                        nORs++;
                                                        sec5.Rows.Add();
                                                        index = sec5.Rows.Count;

                                                        int forMergeRows_4_5 = nORs;   // Для объединение по вертикали 4и пятого столбцов
                                                        for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                                        {
                                                            if (l != 0)
                                                            {
                                                                nORs++;
                                                                sec5.Rows.Add();
                                                                index = sec5.Rows.Count;
                                                            }
                                                            if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated == true)
                                                            {
                                                                string prov = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                                //section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber;
                                                                sec5.Cell(nORs, 2).Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber;
                                                                //section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;

                                                                // Удаление в стретьем столбце номера действия
                                                                string text3 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                                int itNumb = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber.Length;
                                                                text3 = text3.Remove(0, itNumb);

                                                                sec5.Cell(nORs, 3).Range.Text = text3 /*actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName*/;

                                                                // Здесь происходит поиск названия оборудования и выделения его жирным шрифтом
                                                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isConsistEquip == true)
                                                                {
                                                                    string s1 = text3;
                                                                    string s2 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].equipmentName;

                                                                    Object unit5 = WdUnits.wdCharacter;
                                                                    Object extend5 = WdMovementType.wdMove;
                                                                    Object count5 = s1.IndexOf(s2);

                                                                    sec5.Cell(nORs, 3).Range.Select();
                                                                    winword.Selection.MoveLeft(ref unit32, 1, ref extend1);
                                                                    winword.Selection.MoveRight(ref unit5, count5, ref extend5);

                                                                    Object extend51 = WdMovementType.wdExtend;
                                                                    winword.Selection.MoveRight(ref unit5, s2.Length, ref extend51);
                                                                    winword.Selection.Font.Bold = 1;
                                                                    winword.Selection.Font.Italic = 1;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                //section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                                sec5.Cell(nORs, 3).Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                            }

                                                            //section5.Rows[nORs].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                                            /*if (l != actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count -1)
                                                            { */
                                                            /*section5.Rows.Add(); /*}*/
                                                            /*nORs = section5.Rows.Count;*/
                                                        }
                                                        if (nORs > forMergeRows_4_5)
                                                        {
                                                            for (int de = 1; de < nORs - forMergeRows_4_5 + 1; de++)
                                                            {
                                                                sec5.Cell(forMergeRows_4_5, 4).Merge(sec5.Cell(forMergeRows_4_5 + de, 4));
                                                                sec5.Cell(forMergeRows_4_5, 5).Merge(sec5.Cell(forMergeRows_4_5 + de, 5));
                                                            }
                                                        }
                                                        /*if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count > 1)
                                                        {
                                                            for (int d = 1; d < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; d++)
                                                            {
                                                                //section5.Rows[forCombine4_5].Cells[4].Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                                                                sec5.Cell(forCombine4_5,4).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                                                                sec5.Cell(forCombine4_5, 4).Borders[WdBorderType.wdBorderTop].Color = WdColor.wdColorRed;
                                                                //section5.Rows[forCombine4_5].Cells[4].TopPadding = 0;
                                                                //новый обращается через общий массив ячеек
                                                                //section5.Cell(forCombine4_5, 4).Merge(section5.Cell(forCombine4_5 + d, 4));
                                                                //старый код обращался к ячейке через коллекцию строк 
                                                                //section5.Rows[forCombine4_5].Cells[4].Merge(section5.Rows[forCombine4_5 + d].Cells[4]);
                                                            }
                                                        }*/
                                                    }
                                                    //nORs ++;
                                                    /*if (k != actionsList[i].SecondLevelList[j].ThirdLevelList.Count - 1)
                                                    {
                                                        sec5.Rows.Add();
                                                        index = sec5.Rows.Count;
                                                    }*/
                                                }

                                            }
                                        }
                                        if (nORs > firstRowForMerge)
                                        {
                                            for (int fr = 1; fr < nORs - firstRowForMerge + 1; fr++)
                                            {
                                                //int nextRowForMerge = 
                                                sec5.Cell(firstRowForMerge, 1).Merge(sec5.Cell(firstRowForMerge + fr, 1));
                                            }
                                        }
                                    }
                                }
                            }

                            // ПУНКТ 6
                            sec5.Rows.Add();
                            nORs++;

                            sec5.Cell(nORs, 2).Range.Text = "6.";
                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            sec5.Cell(nORs, 3).Range.Text = "КОНТРОЛЬ СООТВЕТСТВИЯ ФАКТИЧЕСКОГО РЕЖИМА В СОЗДАННОЙ СХЕМЕ ИНСТРУКТИВНЫМ УКАЗАНИЯМ";
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                            sec5.Rows.Add();
                            nORs++;

                            sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.dispOffice/*textBox1.Text*/;
                            sec5.Cell(nORs, 2).Range.Text = "6.1.";
                            sec5.Cell(nORs, 3).Range.Font.Bold = 0;
                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                            sec5.Cell(nORs, 3).Range.Text = "На время отключенного состояния " + mainParamsOfSP.nameLine/*textBox2.Text*/ +
                                " перетоки по контролируемым сечениям, а также токи по ЛЭП и электросетевому оборудованию " +
                                "должны быть приведены в соответствие с режимными указаниями к диспетчерской заявке. " +
                                "При отсутствии в диспетчерской заявке режимных указаний перетоки по контролируемым сечениям, " +
                                "а также токи по ЛЭП и электросетевому оборудованию должны соответствовать указаниям Положения " +
                                "по управлению режимами работы в операционной зоне Филиала АО «СО ЕЭС» " + progOption.roditPadezh;
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                            // ПУНКТ 7
                            sec5.Rows.Add();
                            nORs++;

                            sec5.Cell(nORs, 2).Range.Text = "7.";
                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            sec5.Cell(nORs, 3).Range.Text = "ОРГАНИЗАЦИОННЫЕ МЕРОПРИЯТИЯ";
                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));


                            sec5.Rows.Add();
                            nORs++;

                            int numberOfPO = 0;
                            for (int i = 0; i < listOfPowerObjects.Count; i++)
                            {
                                if (listOfPowerObjects[i].isUsed == true)
                                { numberOfPO++; }
                            }

                            if (numberOfPO == 2)
                            {
                                sec5.Cell(nORs, 1).Range.Text = "Операции по п.п. 7.1, 7.2 выполнять одновременно";
                            }
                            else if (numberOfPO > 2)
                            {
                                sec5.Cell(nORs, 1).Range.Text = "Операции по п.п. 7.1 - 7." + numberOfPO + " выполнять одновременно";
                            }
                            sec5.Cell(nORs, 1).Merge(sec5.Cell(nORs, 5));

                            sec5.Rows.Add();
                            nORs++;

                            int numchik = 0;
                            int RowForMerge = nORs;
                            sec5.Cell(nORs, 1).Range.Text = mainParamsOfSP.dispOffice/*textBox1.Text*/;
                            for (int j = 0; j < listOfPowerObjects.Count; j++)
                            {
                                if (listOfPowerObjects[j].isUsed == true)
                                {
                                    if (j == 1)
                                    {
                                        //WdRowHeightRule.wdRowHeightAuto = 0;
                                        sec5.Cell(nORs - 1, 3).HeightRule = WdRowHeightRule.wdRowHeightAuto;
                                    }
                                    numchik++;
                                    for (int k = 0; k < 2; k++)
                                    {
                                        if (k == 0)
                                        {
                                            sec5.Cell(nORs, 2).Range.Text = "7." + numchik;
                                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                            sec5.Cell(nORs, 3).Range.Text = "Сообщить оперативному персоналу " +
                                                listOfPowerObjects[j].NamePO + ":";
                                            sec5.Cell(nORs, 3).Range.Font.Bold = 1;
                                            sec5.Cell(nORs, 3).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                            if (nORs != RowForMerge)
                                            { sec5.Cell(RowForMerge, 1).Merge(sec5.Cell(nORs, 1)); }


                                            sec5.Rows.Add();
                                            nORs++;
                                        }
                                        else
                                        {
                                            string operation3 = "";
                                            int num_oper = 0;
                                            switch (mainParamsOfSP.aim/*comboBox1.SelectedIndex*/)
                                            {
                                                case "Вывод в ремонт":
                                                    operation3 = "выведена в ремонт";
                                                    num_oper = 0;
                                                    break;
                                                case "Вывод в резерв":
                                                    operation3 = "выведена в резерв";
                                                    num_oper = 1;
                                                    break;
                                                case "Ввод из резерва":
                                                    operation3 = "введена из резерва";
                                                    num_oper = 2;
                                                    break;
                                                case "Ввод в работу":
                                                    operation3 = "введена в работу";
                                                    num_oper = 3;
                                                    break;
                                            }
                                            sec5.Cell(nORs, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                            if (num_oper == 1)
                                            {
                                                sec5.Cell(nORs, 3).Range.Text = "«" + mainParamsOfSP.nameLine/*textBox2.Text*/ + " " + operation3 + ".";
                                            }
                                            else
                                            {
                                                sec5.Cell(nORs, 3).Range.Text = "«" + mainParamsOfSP.nameLine/*textBox2.Text*/ + " " + operation3 + ". Время окончания работ на ВЛ______________»";
                                            }
                                            sec5.Cell(nORs, 4).Merge(sec5.Cell(nORs, 5));
                                            sec5.Cell(RowForMerge, 1).Merge(sec5.Cell(nORs, 1));
                                            sec5.Cell(nORs - 1, 2).Merge(sec5.Cell(nORs, 2));
                                            sec5.Cell(nORs - 1, 3).Merge(sec5.Cell(nORs, 3));
                                            sec5.Cell(nORs - 1, 4).Merge(sec5.Cell(nORs, 4));

                                            sec5.Rows.Add();
                                            //nORs++;
                                        }
                                    }
                                }
                            }

                            sec5.Cell(nORs, 1).Select();
                            winword.Selection.Cells.Delete(WdDeleteCells.wdDeleteCellsEntireRow); // Удаление лишней строки в таблице
                            sec5.Cell(nORs, 1).Select();
                            winword.Selection.Cells.Delete(WdDeleteCells.wdDeleteCellsEntireRow); // Удаление лишней строки в таблице





                            // ПУНКТ 8
                            // Если работы не производятся вообще, то пункт 8 не выполняется
                            if (listOfOrgArrs[listOfOrgArrs.Count - 1].isWork != true)
                            {
                                sec5.Cell(nORs - 1, 3).Select();
                                //int lastNotClearPage = winword.Selection.Information[WdInformation.wdActiveEndPageNumber]; // Текущая страница

                                /*WdStatistic stat = WdStatistic.wdStatisticPages;
                                int totalPages = document.ComputeStatistics(stat, ref missing);              // Всего страниц*/


                                /*winword.Selection.MoveRight(WdUnits.wdCharacter, 3, WdMovementType.wdMove);*/

                                //Microsoft.Office.Interop.Word.Paragraph para_08 = document.Content.Paragraphs.Add(ref missing);

                                GoToNextPage(winword, document, missing/*, para_08*/);
                                /*int lastNotClearPage = winword.Selection.Information[WdInformation.wdActiveEndPageNumber]; // Текущая страница

                                WdStatistic stat = WdStatistic.wdStatisticPages;
                                int totalPages = document.ComputeStatistics(stat, ref missing);              // Всего страниц

                                // Пока каретка не перейдёт на следующую страницу добавляем строки
                                while (lastNotClearPage == totalPages)
                                {
                                    para_08.Range.Text = "\n";
                                    winword.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                                    //currentPage = winword.Selection.Information[WdInformation.wdActiveEndPageNumber];
                                    totalPages = document.ComputeStatistics(stat, ref missing);
                                }*/

                                /*// Удаляется лишняя строка
                                winword.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdExtend);
                                winword.Selection.Delete();*/

                                Microsoft.Office.Interop.Word.Paragraph para_8 = document.Content.Paragraphs.Add(ref missing);
                                Microsoft.Office.Interop.Word.Table table8 = document.Tables.Add(para_8.Range, 1, 4, 2, ref missing);

                                para_8.Format.LeftIndent = -30;    // Сдвиг всей таблицы влево

                                //para_8.Range.InsertParagraphAfter();

                                // Таблица
                                para_8.Format.LeftIndent = -30;

                                table8.Borders.Enable = 1;

                                table8.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                                table8.Range.Font.Name = "Times New Roman";

                                /*para1.Range.Font.Size = 13;
                                para1.Range.Font.Name = "Times New Roman";*/


                                // Хорошие размеры таблицы
                                table8.Columns[1].SetWidth(63, WdRulerStyle.wdAdjustNone);
                                table8.Columns[2].SetWidth(42, WdRulerStyle.wdAdjustNone);
                                table8.Columns[3].SetWidth(340, WdRulerStyle.wdAdjustNone);
                                table8.Columns[4].SetWidth(86/*43*/, WdRulerStyle.wdAdjustNone);
                                //table8.Columns[5].SetWidth(43, WdRulerStyle.wdAdjustNone);

                                table8.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                                // Шапка таблицы                                
                                table8.Cell(1, 1).Range.ParagraphFormat.LeftIndent = -5;        // Смещение текста в ячейке влево                        
                                                                                                //table8.Cell(1, 1).Range.Text = "Персонал,\nвыполняющий операцию";                        
                                table8.Cell(1, 2).Range.Text = "8.";
                                table8.Cell(1, 2).Range.Font.Size = 13;
                                table8.Cell(1, 2).Range.Font.Name = "Times New Roman";

                                table8.Cell(1, 3).Range.Text = "МЕРОПРИЯТИЯ ПО ОБЕСПЕЧЕНИЮ БЕЗОПАСНОСТИ ПРОВЕДЕНИЯ РАБОТ";
                                table8.Cell(1, 3).Range.Font.Size = 13;
                                table8.Cell(1, 3).Range.Font.Name = "Times New Roman";

                                table8.Cell(1, 4).Range.Font.Size = 13;
                                table8.Cell(1, 4).Range.Font.Name = "Times New Roman";

                                table8.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                table8.Rows[1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                                table8.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;

                                table8.Rows.Add();

                                table8.Cell(1, 3).Range.Font.Bold = 1;


                                // ПУНКТ 8.1.
                                int currRow = 2;
                                int firstRowForMerge = 2;
                                int itemNumber = 1;
                                int arrangementsSelect = 0;

                                table8.Cell(currRow, 1).Range.Text = mainParamsOfSP.dispOffice;
                                table8.Cell(currRow, 2).Range.Text = "8." + itemNumber + ".";
                                for (int i = 0; i < listOfOrgArrs.Count - 2; i++)
                                {
                                    if (listOfOrgArrs[i].isWork == true)
                                    { arrangementsSelect++; }
                                }
                                string text8 = "При производстве работ на ";
                                string addedText = " участке " + mainParamsOfSP.nameLine + " в пределах ";
                                // Проверка условия выполнения работ на ВЛ
                                if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true)
                                {
                                    text8 = text8 + mainParamsOfSP.nameLine;
                                    if (arrangementsSelect != 0)
                                    { text8 = text8 + " и/или "; }
                                }
                                // Если работа выполняется на одной из ПС, то добавляются эти ПС
                                for (int i = 0; i < listOfPowerObjects.Count; i++)
                                {
                                    if (listOfPowerObjects[i].isUsed == true)
                                    {
                                        addedText = addedText + listOfPowerObjects[i].NamePO + " и/или ";

                                    }
                                }
                                text8 = text8 + addedText;

                                text8 = text8.Remove(text8.Length - 7);
                                /*if (arrangementsSelect != 0)
                                {
                                    text8 = text8.Remove(text8.Length - 7);
                                }*/

                                text8 = text8 + " сообщить диспетчеру " + mainParamsOfSP.lineOrganisation + ":";
                                int boldText = text8.Length;
                                text8 = text8 + "\n«Операции по отключению, заземлению, переключениям во вторичных цепях выполнены. " +
                                    mainParamsOfSP.nameLine + " отключена и заземлена в сторону ВЛ на ";
                                addedText = "";
                                for (int i = 0; i < listOfPowerObjects.Count; i++)
                                {
                                    if (listOfPowerObjects[i].isUsed == true)
                                    {
                                        addedText = addedText + listOfPowerObjects[i].NamePO + " и ";
                                    }
                                }

                                text8 = text8 + addedText;
                                text8 = text8.Remove(text8.Length - 3);
                                text8 = text8 + ". На приводах линейных разъединителей " + mainParamsOfSP.nameLine + " на ";

                                addedText = "";
                                for (int i = 0; i < listOfPowerObjects.Count; i++)
                                {
                                    if (listOfPowerObjects[i].isUsed == true)
                                    {
                                        addedText = addedText + listOfPowerObjects[i].NamePO + " и ";
                                    }
                                }
                                text8 = text8 + addedText;
                                text8 = text8.Remove(text8.Length - 3);
                                text8 = text8 + " вывешены плакаты «Не включать! Работа на линии». На ";

                                addedText = "";
                                for (int i = 0; i < listOfPowerObjects.Count; i++)
                                {
                                    if (listOfPowerObjects[i].isUsed == true)
                                    {
                                        addedText = addedText + listOfPowerObjects[i].NamePO + " и ";
                                    }
                                }
                                text8 = text8 + addedText;
                                text8 = text8.Remove(text8.Length - 3);
                                text8 = text8 + " приняты меры препятствующие подаче напряжения на " +
                                    mainParamsOfSP.nameLine +
                                    " вследствие ошибочного или самопроизвольного включения коммутационных аппаратов»";

                                table8.Cell(currRow, 3).Range.Text = text8;
                                table8.Cell(currRow, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                table8.Cell(currRow, 3).Select();
                                winword.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                                // Выделение жирным шрифтом части текста
                                Object unit8 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                                                                  //Object count8 = 1;
                                Object extend8 = WdMovementType.wdExtend;

                                winword.Selection.MoveRight(ref unit8, boldText, ref extend8);
                                winword.Selection.Font.Bold = 1;


                                //==========Конец заполнения пункта 8.1.================

                                // Если выбран пункт "Работы на ВЛ"
                                if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork == true)
                                {
                                    table8.Rows.Add();
                                    currRow++;
                                    itemNumber++;

                                    table8.Cell(currRow, 3).Range.Font.Bold = 0;

                                    table8.Cell(currRow, 2).Range.Text = "8." + itemNumber + ".";
                                    text8 = "Дать команду диспетчеру " + mainParamsOfSP.lineOrganisation +
                                        " (при производстве линейных работ на " + mainParamsOfSP.nameLine + ":";

                                    boldText = text8.Length;

                                    text8 = text8 + "\n";
                                    text8 = text8 + "«После выполнения иных технических мероприятий, предусмотренных нарядом, организуйте выдачу разрешения на подготовку рабочего места и допуск для линейных работ на " +
                                        mainParamsOfSP.nameLine + ". Работы закончить до ____(ЧЧ:ММ), «___»________20   г. с аварийной готовностью ____ч»";

                                    table8.Cell(currRow, 3).Range.Text = text8;
                                    table8.Cell(currRow, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                    table8.Cell(currRow, 3).Select();
                                    winword.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                                    // Выделение жирным шрифтом части текста
                                    //Object unit8 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                    //Object count8 = 1;
                                    //Object extend8 = WdMovementType.wdExtend;

                                    winword.Selection.MoveRight(ref unit8, boldText, ref extend8);
                                    winword.Selection.Font.Bold = 1;

                                    table8.Cell(firstRowForMerge, 1).Merge(table8.Cell(currRow, 1));

                                    {
                                        for (int i = 0; i < listOfPowerObjects.Count; i++)
                                        {
                                            if (listOfPowerObjects[i].isUsed == true)
                                            {
                                                table8.Rows.Add();
                                                currRow++;
                                                itemNumber++;

                                                string lineDisconnectorName = "";

                                                table8.Cell(currRow, 3).Range.Font.Bold = 0;

                                                table8.Cell(currRow, 2).Range.Text = "8." + itemNumber + ".";
                                                text8 = "Дать команду диспетчеру " + mainParamsOfSP.lineOrganisation +
                                                    " (при производстве работ на участке " + mainParamsOfSP.nameLine +
                                                    " в пределах " + listOfPowerObjects[i].NamePO + " до линейного разъединителя ";

                                                for (int j = 0; j < One_listEquipment.Count; j++)
                                                {
                                                    if (One_listEquipment[j].NamePO == listOfPowerObjects[i].NamePO &&
                                                        One_listEquipment[j].typeEquip == "Disconnector")
                                                    {
                                                        lineDisconnectorName = One_listEquipment[j].nameEquip;
                                                        text8 = text8 + lineDisconnectorName;
                                                    }
                                                }

                                                text8 = text8 + "):";

                                                boldText = text8.Length;

                                                text8 = text8 + "\n";
                                                text8 = text8 + "«После выполнения иных технических мероприятий, предусмотренных нарядом, организуйте выдачу разрешения на подготовку рабочего места и допуск для работ на участке " +
                                                    mainParamsOfSP.nameLine + " в пределах " + listOfPowerObjects[i].NamePO + " до линейного разъединителя " +
                                                    lineDisconnectorName + ". Работы закончить до ____(ЧЧ:ММ), «___»________20   г. с аварийной готовностью ____ч»";

                                                table8.Cell(currRow, 3).Range.Text = text8;
                                                table8.Cell(currRow, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                                table8.Cell(currRow, 3).Select();
                                                winword.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                                                // Выделение жирным шрифтом части текста
                                                //Object unit8 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                                //Object count8 = 1;
                                                //Object extend8 = WdMovementType.wdExtend;

                                                winword.Selection.MoveRight(ref unit8, boldText, ref extend8);
                                                winword.Selection.Font.Bold = 1;

                                                table8.Cell(firstRowForMerge, 1).Merge(table8.Cell(currRow, 1));

                                                string s1 = text8.Remove(0, boldText);
                                                /*string s2 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].equipmentName;*/

                                                Object unit8_1 = WdUnits.wdCharacter;
                                                Object extend8_1 = WdMovementType.wdMove;
                                                Object count8_1 = s1.IndexOf(lineDisconnectorName) + boldText;

                                                table8.Cell(currRow, 3).Range.Select();
                                                winword.Selection.MoveLeft(ref unit8_1, 1, ref extend8_1);
                                                winword.Selection.MoveRight(ref unit8_1, count8_1, ref extend8_1);

                                                Object extend51 = WdMovementType.wdExtend;
                                                winword.Selection.MoveRight(ref unit8_1, lineDisconnectorName.Length, WdMovementType.wdExtend);
                                                winword.Selection.Font.Bold = 1;
                                                winword.Selection.Font.Italic = 1;

                                            }
                                        }
                                    }
                                }


                                // Если выбрана работа на энергообъектах, то выполяется проход по каждой
                                for (int i = 0; i < listOfOrgArrs.Count - 2; i++)
                                {
                                    // Если работа на ВЛ не выполняется
                                    if (listOfOrgArrs[listOfOrgArrs.Count - 2].isWork != true)
                                    {
                                        if (listOfOrgArrs[i].isWork == true)
                                        {
                                            table8.Rows.Add();
                                            currRow++;
                                            itemNumber++;

                                            string lineDisconnectorName = "";

                                            table8.Cell(currRow, 3).Range.Font.Bold = 0;

                                            table8.Cell(currRow, 2).Range.Text = "8." + itemNumber + ".";
                                            text8 = "Дать команду диспетчеру " + mainParamsOfSP.lineOrganisation +
                                                " (при производстве работ на участке " + mainParamsOfSP.nameLine +
                                                " в пределах " + listOfOrgArrs[i].PObject + " до линейного разъединителя ";

                                            for (int j = 0; j < One_listEquipment.Count; j++)
                                            {
                                                if (One_listEquipment[j].NamePO == listOfOrgArrs[i].PObject &&
                                                    One_listEquipment[j].typeEquip == "Disconnector")
                                                {
                                                    lineDisconnectorName = One_listEquipment[j].nameEquip;
                                                    text8 = text8 + lineDisconnectorName;
                                                }
                                            }

                                            text8 = text8 + "):";

                                            boldText = text8.Length;

                                            text8 = text8 + "\n";
                                            text8 = text8 + "«После выполнения иных технических мероприятий, предусмотренных нарядом, организуйте выдачу разрешения на подготовку рабочего места и допуск для работ на участке " +
                                                mainParamsOfSP.nameLine + " в пределах " + listOfOrgArrs[i].PObject + " до линейного разъединителя " +
                                                lineDisconnectorName + ". Работы закончить до ____(ЧЧ:ММ), «___»________20   г. с аварийной готовностью ____ч»";

                                            table8.Cell(currRow, 3).Range.Text = text8;
                                            table8.Cell(currRow, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                            table8.Cell(currRow, 3).Select();
                                            winword.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                                            // Выделение жирным шрифтом части текста
                                            //Object unit8 = WdUnits.wdCharacter;               // Операции по перемещению каретки вправо
                                            //Object count8 = 1;
                                            //Object extend8 = WdMovementType.wdExtend;

                                            winword.Selection.MoveRight(ref unit8, boldText, ref extend8);
                                            winword.Selection.Font.Bold = 1;

                                            table8.Cell(firstRowForMerge, 1).Merge(table8.Cell(currRow, 1));

                                            string s1 = text8.Remove(0, boldText);
                                            /*string s2 = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].equipmentName;*/

                                            Object unit8_1 = WdUnits.wdCharacter;
                                            Object extend8_1 = WdMovementType.wdMove;
                                            Object count8_1 = s1.IndexOf(lineDisconnectorName) + boldText;

                                            table8.Cell(currRow, 3).Range.Select();
                                            winword.Selection.MoveLeft(ref unit8_1, 1, ref extend8_1);
                                            winword.Selection.MoveRight(ref unit8_1, count8_1, ref extend8_1);

                                            Object extend51 = WdMovementType.wdExtend;
                                            winword.Selection.MoveRight(ref unit8_1, lineDisconnectorName.Length, WdMovementType.wdExtend);
                                            winword.Selection.Font.Bold = 1;
                                            winword.Selection.Font.Italic = 1;

                                        }
                                    }
                                }
                            }


                            GoToNextPage(winword, document, missing);


                            Microsoft.Office.Interop.Word.Paragraph para_9 = document.Content.Paragraphs.Add(ref missing);

                            // Заглавие раздела 9
                            string text9 = "ПЕРСОНАЛ, УЧАСТВУЮЩИЙ В ";

                            switch (mainParamsOfSP.aim/*comboBox1.SelectedIndex*/)
                            {
                                case "Вывод времонт":
                                    text9 = text9 + "ВЫВОДЕ В РЕМОНТ";
                                    break;
                                case "Вывод в резерв":
                                    text9 = text9 + "ВЫВОДЕ В РЕЗЕРВ";
                                    break;
                                    /*case 2:
                                        tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ ИЗ РЕЗЕРВА";
                                        break;
                                    case 3:
                                        tab01.Cell(2, 1).Range.Text = "ПО ВВОДУ В РАБОТУ";
                                        break;*/
                            }

                            text9 = text9 + " " + mainParamsOfSP.nameLine + " И ОРГАНИЗАЦИИ БЕЗОПАСНОГО ПРОВЕДЕНИЯ РАБОТ НА ВЛ";

                            para_9.Range.Text = text9;
                            para_9.Range.InsertParagraphAfter();

                            // ПУНКТ 9 (Персонал)
                            Microsoft.Office.Interop.Word.Table table9 = document.Tables.Add(para_9.Range, 1, 4, 2, ref missing);

                            para_9.Format.LeftIndent = -30;    // Сдвиг всей таблицы влево



                            // Таблица
                            para_9.Format.LeftIndent = -30;

                            table9.Borders.Enable = 1;

                            table9.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                            table9.Range.Font.Name = "Times New Roman";


                            table9.Cell(1, 1).Select();

                            // Выделение жирным шрифтом заглавия
                            Object unit9 = WdUnits.wdCharacter;
                            Object count9 = 1;
                            Object extend9 = WdMovementType.wdExtend;

                            winword.Selection.MoveLeft(ref unit9, count9, WdMovementType.wdMove);
                            winword.Selection.MoveLeft(ref unit9, count9, WdMovementType.wdMove);

                            winword.Selection.MoveLeft(ref unit9, text9.Length, extend9);

                            winword.Selection.Font.Bold = 1;

                            winword.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            winword.Selection.Range.Font.Size = 13;
                            winword.Selection.Range.Font.Name = "Times New Roman";
                            /*para1.Range.Font.Size = 13;
                            para1.Range.Font.Name = "Times New Roman";*/


                            // Хорошие размеры таблицы
                            table9.Columns[1].SetWidth(50, WdRulerStyle.wdAdjustNone);
                            table9.Columns[2].SetWidth(180, WdRulerStyle.wdAdjustNone);
                            table9.Columns[3].SetWidth(140, WdRulerStyle.wdAdjustNone);
                            table9.Columns[4].SetWidth(120, WdRulerStyle.wdAdjustNone);


                            table9.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                            // Шапка таблицы 9
                            table9.Cell(1, 1).Range.Text = "№№\nп/п";
                            table9.Cell(1, 1).Range.Font.Size = 13;
                            table9.Cell(1, 1).Range.Font.Name = "Times New Roman";

                            table9.Cell(1, 2).Range.Text = "Организация\n(объект переключений)";
                            table9.Cell(1, 2).Range.Font.Size = 13;
                            table9.Cell(1, 2).Range.Font.Name = "Times New Roman";

                            table9.Cell(1, 3).Range.Text = "Фамилия, инициалы";
                            table9.Cell(1, 3).Range.Font.Size = 13;
                            table9.Cell(1, 3).Range.Font.Name = "Times New Roman";

                            table9.Cell(1, 4).Range.Text = "Должность";
                            table9.Cell(1, 4).Range.Font.Size = 13;
                            table9.Cell(1, 4).Range.Font.Name = "Times New Roman";

                            table9.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            table9.Rows[1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            // Заполнения таблицы 9

                            int currRow9 = 2;
                            int itemNumber9 = 1;
                            for (int i = 0; i < listOfPersonal.Count; i++)
                            {
                                // Если человек не из Системного опереатора
                                if (listOfPersonal[i].organisationOfPersonal != "SO")
                                {
                                    table9.Rows.Add();

                                    table9.Cell(currRow9, 1).Range.Text = itemNumber9 + ".";
                                    table9.Cell(currRow9, 2).Range.Text = listOfPersonal[i].organisationOfPersonal;
                                    table9.Cell(currRow9, 3).Range.Text = listOfPersonal[i].Person[0].nameOfPerson;
                                    table9.Cell(currRow9, 4).Range.Text = listOfPersonal[i].Person[0].role;

                                    currRow9++;
                                    itemNumber9++;
                                }
                            }

                            // Таблица с людьми из СО
                            para_9.Range.InsertParagraphAfter();

                            // ПУНКТ 9 (Персонал)
                            Microsoft.Office.Interop.Word.Table table91 = document.Tables.Add(para_9.Range, 1, 3, 2, ref missing);
                            para_9.Format.LeftIndent = -10;    // Сдвиг всей таблицы влево
                            
                            table91.Borders.Enable = 1;

                            table91.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                            table91.Range.Font.Name = "Times New Roman";

                            // Хорошие размеры таблицы
                            table91.Columns[1].SetWidth(160, WdRulerStyle.wdAdjustNone);
                            table91.Columns[2].SetWidth(125, WdRulerStyle.wdAdjustNone);
                            table91.Columns[3].SetWidth(225, WdRulerStyle.wdAdjustNone);

                            table91.Cell(1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;


                            table91.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                            int lastRow = 1;
                            for (int p = 0; p < listOfPersonal.Count; p++)
                            {
                                if (listOfPersonal[p].organisationOfPersonal == "SO")
                                {
                                    for (int q = 0; q < listOfPersonal[p].Person.Count; q++)
                                    {
                                        table91.Cell(q + 1, 1).Range.Text = listOfPersonal[p].Person[q].action;
                                        table91.Cell(q + 1, 3).Range.Text = listOfPersonal[p].Person[q].role;
                                        table91.Rows.Add();
                                        lastRow++;
                                    }
                                }
                            }

                            // Удаление лишней строки в таблице
                            table91.Cell(lastRow, 1).Select();
                            winword.Selection.Cells.Delete(WdDeleteCells.wdDeleteCellsEntireRow); // Удаление лишней строки в таблице



                            // Создание колонтитулов
                            WdStatistic stat1 = WdStatistic.wdStatisticPages;

                            int totalPages1 = document.ComputeStatistics(stat1, ref missing);              // Всего страниц

                            //winword.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 1, Type.Missing);
                            /*WdBreakType brType = WdBreakType.wdSectionBreakNextPage;

                            winword.Selection.InsertBreak(brType);              // Разрыв разделов

                            winword.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 2, Type.Missing);     // Переход на указанную страницу

                            document.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                            winword.Selection.HeaderFooter.LinkToPrevious = false;                  // Убирается параметр "как в пердыдущем разделе"

                            int secs = winword.Selection.Sections.Count;

                            //winword.Selection.Sections[2].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                            winword.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "2222f"; //текст для нижнего колонтитула текущей страницы

                            winword.Selection.GoTo(WdGoToItem.wdGoToSection, WdGoToDirection.wdGoToAbsolute, 2, Type.Missing);      // переход на слудующий нижний колонтитул
                            winword.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "999kkk"; //текст для нижнего колонтитула текущей страницы



                            winword.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 3, Type.Missing);
                            winword.Selection.InsertBreak(brType);              // Разрыв разделов

                            winword.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 3, Type.Missing);     // Переход на указанную страницу

                            winword.Selection.GoTo(WdGoToItem.wdGoToSection, WdGoToDirection.wdGoToAbsolute, 3, Type.Missing);
                            document.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                            winword.Selection.HeaderFooter.LinkToPrevious = false;                  // Убирается параметр "как в пердыдущем разделе"

                            winword.Selection.GoTo(WdGoToItem.wdGoToSection, WdGoToDirection.wdGoToAbsolute, 3, Type.Missing);      // переход на слудующий нижний колонтитул
                            winword.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "третий3"; //текст для нижнего колонтитула текущей страницы
                            */

                            winword.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToFirst, 1, Type.Missing);

                            //progressBar1.IsIndeterminate = false;
                            //winword.Activate();
                            winword.Visible = true;         // Открытие созданного файла
                            //winword.WindowState = WdWindowState.wdWindowStateNormal;
                        }
                        catch (Exception ex)
                        {
                            //progressBar1.IsIndeterminate = false;
                            MessageBox.Show(ex.Message);
                        }
                }
                    else
                    {
                        MessageBox.Show("В настройках персонала имеются незаполненные поля");
                    }
            }
                else
                {
                    MessageBox.Show("Некорреткно выбраны организационные мероприятия");
                }
            }
            else
            { MessageBox.Show("Не выбрано ни одно организационное мероприятие"); }
        }

        partial void GoToNextPage(Microsoft.Office.Interop.Word.Application app, Microsoft.Office.Interop.Word.Document doc, 
            object missing/*, Microsoft.Office.Interop.Word.Paragraph para*/)
        {
            Microsoft.Office.Interop.Word.Paragraph para_0 = doc.Content.Paragraphs.Add(ref missing); // Параграф, необходимый для выполнения следующих операций

            int lastNotClearPage = /*winword*/app.Selection.Information[WdInformation.wdActiveEndPageNumber]; // Текущая страница

            WdStatistic stat = WdStatistic.wdStatisticPages;
            int totalPages = doc.ComputeStatistics(stat, ref missing);              // Всего страниц

            // Пока каретка не перейдёт на следующую страницу добавляем строки
            while (lastNotClearPage == totalPages)
            {
                para_0.Range.Text = "\n";
                app.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                //currentPage = winword.Selection.Information[WdInformation.wdActiveEndPageNumber];
                totalPages = doc.ComputeStatistics(stat, ref missing);
            }
        }

    }
}
