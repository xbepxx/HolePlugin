using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitApiHole
{
    #region ЗАНЯТИЕ 8. ПЛАГИН "РАССТАНОВКА ОТВЕРСТИЙ". ЧАСТЬ 2.
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    //самым верным методом для определения места пересечения стены и трубопровода/воздуховода является использование фильтра ReferenceIntersector.
    //Принцип действия: в пространстве выбирается точка, из которой "выпускается" воображаемый луч. И фильтр навходит все элементы, которые данный луч пересекает. Мы можем указывать тип искомых объектов,
    //диапазон поиска (к примеру, воздуховод). Из одного из концов испусается луч. Особенность: поиск выполняется на 3D виде. Плюс, луч находит стену 2 раза (2 грани стены). Класс сравнения
    //IEqualityComparer - для выбора одной стен. Далее, при работе со связанными файлами, необходимо учитывать, что в связанном файле не возможно выпонлять транзакцию по целевому объекту, а значит основной файл: АР, ОВК - связь*/
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document arDoc = uidoc.Document;
            Document ovkDoc = arDoc.Application.Documents //список документов в проекте
                .OfType<Document>() //получаем список
                .Where(x => x.Title.Contains("ОВК")) //который должен содержать ОВК
                .FirstOrDefault();
            if (ovkDoc==null) //проверка, если такой файл
            {
                TaskDialog.Show("Ошибка!", "Файл не найден");
                return Result.Cancelled;
            }

            FamilySymbol hole = new FilteredElementCollector(arDoc) //фильтр по отверстиям
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие")) //тут внимательнее, не Name
                .FirstOrDefault();
            
            if (hole==null) //проверка по наличию семейства отверстий
            {
                TaskDialog.Show("Ошибка!", "Не найдено семейство \"Отверстие\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovkDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovkDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc) //получаем 3D вид
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) //убеждаемся, что 3D вид не является шаблоном вида!
                .FirstOrDefault();

            if (view3D==null)
            {
                TaskDialog.Show("Ошибка!", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D); //выбираем перегрузку с 3-мя аргументами: ElementClassFilter(тоже, что и .OfClass)

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Расстановка отверстий"); 
            if (!hole.IsActive)
            {
                hole.Activate();
            }
            transaction0.Commit();
            
            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");      
            foreach (Duct d in ducts)
            {
               Line curve= (d.Location as LocationCurve).Curve as Line; //получаем линию из воздуховода
                XYZ point = curve.GetEndPoint(0); //получаем точку
                XYZ direction = curve.Direction; //получаем направление

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //получаем список пересечений
                    .Where(x => x.Proximity <= curve.Length) //для сужения круга поиска выберем объекты, длина которых не превышает длины воздуховода
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach(ReferenceWithContext refer in intersections) //для каждого полученного пересечения вставим отверстие
                {
                    double proximity = refer.Proximity; //расстояние до объекта
                    Reference reference = refer.GetReference(); //ссылка. Reference - внутренний код в базе
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //получаем уровень
                    XYZ pointhole = point + (direction * proximity); //получаем точку вставки

                    FamilyInstance instHole = arDoc.Create.NewFamilyInstance(pointhole, hole, wall, level, StructuralType.NonStructural); //NonStructural - т.к. не фундамент, не колонна
                    Parameter width = instHole.LookupParameter("Ширина");
                    Parameter height = instHole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                }
            }
            foreach (Pipe p in pipes)
            {
                Line curve = (p.Location as LocationCurve).Curve as Line; //получаем линию из трубы
                XYZ point = curve.GetEndPoint(0); //получаем точку
                XYZ direction = curve.Direction; //получаем направление

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //получаем список пересечений
                    .Where(x => x.Proximity <= curve.Length) //для сужения круга поиска выберем объекты, длина которых не превышает длины трубы
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersections) //для каждого полученного пересечения вставим отверстие
                {
                    double proximity = refer.Proximity; //расстояние до объекта
                    Reference reference = refer.GetReference(); //ссылка. Reference - внутренний код в базе
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //получаем уровень
                    XYZ pointhole = point + (direction * proximity); //получаем точку вставки

                    FamilyInstance instHole = arDoc.Create.NewFamilyInstance(pointhole, hole, wall, level, StructuralType.NonStructural); //NonStructural - т.к. не фундамент, не колонна
                    Parameter width = instHole.LookupParameter("Ширина");
                    Parameter height = instHole.LookupParameter("Высота");
                    width.Set(p.Diameter);
                    height.Set(p.Diameter);
                }
            }

            transaction.Commit();

            return Result.Succeeded;
        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext> //дополнительный класс по фильтрации точек по одной стене
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y) //будут ли два объекта одинаковыми
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId //тоже самое, что ниже, но у связанного файла
                           && xReference.ElementId == yReference.ElementId; //если у обоих элементов совпадает ElementId, то возвращается true
            }

            public int GetHashCode(ReferenceWithContext obj) //возвращает хэшкод объекта
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
    #endregion
}
