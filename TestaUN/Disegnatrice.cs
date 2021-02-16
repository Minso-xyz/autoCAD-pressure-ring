using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace TestaUN
{
    public class Disegnatrice
    {
        [CommandMethod("Demo")]
        public static void StartFormUN()
        {
            button_insert form = new button_insert();
            form.Show();
        }
    }
}
