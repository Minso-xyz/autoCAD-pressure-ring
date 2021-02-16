using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace TestaUN
{
    public class Global
    {
        internal class Point
        {
            public static Point3d insertPoint;
            public static object getInsertionPoint(string msg, Document doc)
            {
                PromptPointOptions Point = new PromptPointOptions(msg);
                PromptPointResult PointResult;
                PointResult = doc.Editor.GetPoint(Point);
                insertPoint = PointResult.Value;

                // Exit if the user presses ESC or cancels the command
                if (PointResult.Status == PromptStatus.Cancel) return null;
                return insertPoint;
            }
        }
    }
}
