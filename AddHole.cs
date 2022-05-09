using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePluginLab8
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document; //подключаем документ 

            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            //связанный файл ovDoc получаем из основного файла обратившись к свойству Application
            //который представляет собой объект класса Application

            if (ovDoc == null)//если файл не будет обнаружен
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled; //возвращаем резельтат "отмена"
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc) //проверяем что наш файл "АР" закружено семейство с отверстиями
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)//это с Revit
                .OfType<FamilySymbol>() //приводим к типу FamilySymbol
                 .Where(x => x.FamilyName.Equals("Отверстия"))//выполняем поиск по названию
                 .FirstOrDefault();
           
            if (familySymbol == null) //выполняем проверку и вслучае когда семейство ненайдено сообщаем и выходим из плагина
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отвестия\"");
                return Result.Cancelled; //возвращаем резельтат "отмена"
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc) //находим все воздуховоды 
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc) // находим все трубы
               .OfClass(typeof(Pipe))
               .OfType<Pipe>()
               .ToList();

            View3D view3D = new FilteredElementCollector(arDoc) //находим 3D вид через фильтр ищем в документе arDoc
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)//важно чтобы 3D вид не являлся шаблоном вида
                .FirstOrDefault();

            if (view3D == null) //выполняем проверку и вслучае когда семейство ненайдено сообщаем и выходим из плагина
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled; //возвращаем резельтат "отмена"
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), 
                FindReferenceTarget.Element, view3D);
            //создаем объект для сосдания применяем конструктор в который передаем фильтр по стенам, FindReferenceTarget.Element, view3D

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");

            foreach (Duct d in ducts) //получаем результат (набор пересикаемых стен) вызываем метод Find
            {

                Line curve = (d.Location as LocationCurve).Curve as Line;//получаем точку которая лежит у воздуховода
                XYZ point = curve.GetEndPoint(0);//вызываем точку из кривой
                XYZ direction = curve.Direction; //вектор направления из кривой
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                //методу Find передаем исходную точку и направление точки
                .Where(x => x.Proximity <= curve.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer()) //метод расширения
                .ToList();// мы получаем набор всех пересечений в виде объекта ReferenceWithContext

                foreach(ReferenceWithContext refer in intersections) //переберем и для каждого определим точку вставки и добавляем
                                                                     //экземпляр семейства отверстия
                {
                    double proximity = refer.Proximity; //расстояние до объекта
                    Reference reference = refer.GetReference(); //ссылку через метод GetReference
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //зная ID можна получить в документе стену на которую идет эта ссылка
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity); //стартовая точка + (направление * расстояние)

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    //метод для вставки экземпляра симейства (изменение в модели )

                    Parameter width = hole.LookupParameter("Ширина"); //ширину и высоту определяем взависимости от диаметра трубы 
                    Parameter heigfh = hole.LookupParameter("Высота");
                    width.Set(d.Diameter); //устанавливаем значение которе берем у воздуховода
                    heigfh.Set(d.Diameter);
                }
            }
            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line; //получаем точку которая лежит у трубопровода
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;//вектор направления из кривой
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //метод расширения
                    .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;//расстояние до объекта
                    Reference reference = refer.GetReference();//ссылку через метод GetReference
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;//зная ID можна получить в документе стену на которую идет эта ссылка
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);//стартовая точка + (направление * расстояние)

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");//ширину и высоту определяем взависимости от диаметра трубы 
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(pipe.Diameter);//устанавливаем значение которе берем у трубопровода
                    height.Set(pipe.Diameter);
                }
            }    

            transaction.Commit();
            return Result.Succeeded;

        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>// класс унаследован от IEqualityComparer
        {
            public bool Equals( ReferenceWithContext x, ReferenceWithContext y)// методEquals
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();
                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                    && xReference.ElementId == yReference.ElementId;// результат мы получаем точки на одной стене
            }

            public int GetHashCode(ReferenceWithContext obj)// метод GetHashCode
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();// возвращает HashCode объекта
                }
            }
        }
    }
}
