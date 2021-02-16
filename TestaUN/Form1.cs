using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace TestaUN
{
    public partial class button_insert : Form
    {
        public button_insert()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Declare the input variables
            double intDia = double.Parse(textBox_intDia.Text);
            double extDia = double.Parse(textBox_extDia.Text);
            double height = double.Parse(textBox_height.Text);
            double crossSection = (extDia - intDia) * 0.5;

            // Boolean : Endless/Split
            Boolean endless = true; 
            Boolean split = false;

            // Input validation
            Boolean isDataValidated = true;

            // Get the document object
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ///// Transferred to the class Global.cs
            //// Prompt for the insertion point
            //PromptPointResult insertPointResult;
            //PromptPointOptions insertPointOption = new PromptPointOptions("");
            //insertPointOption.Message = "\nClick the point where you want to insert the profile UN : ";
            //insertPointResult = doc.Editor.GetPoint(insertPointOption);
            //Point3d insertPoint = insertPointResult.Value;

            //// Exit if the user presses ESC or cancels the command
            //if (insertPointResult.Status == PromptStatus.Cancel) return;

            ////// DATA VALIDATION ///////////////////////////////////////////
            // int Dia. > ext Dia.
            if(intDia >= extDia)
            {
                labelWarning.Text = "int.Diameter must be smaller than ext.Diameter!";
                isDataValidated = false;
            }

            // too low height
            if(height <= 3)
            {
                labelWarning.Text = "Height is too small!";
                isDataValidated = false;
            }

            // too high height
            if(height > 45)
            {
                labelWarning.Text = "Height is too big!";
                isDataValidated = false;
            }

            // Feasible range
            if (crossSection < 5 || crossSection >35)
            {
                labelWarning.Text = "Cross-section is out of range!\n(5 <= F <= 35)";
                isDataValidated = false;
            }

            // Null input
            // Internal diameter (Null)
            if (string.IsNullOrEmpty(this.textBox_intDia.Text))
            {
                labelWarning.Text = "Please insert the value of int.Diameter!";
                isDataValidated = false;
            }

            // External diameter (Null)
            if (string.IsNullOrEmpty(this.textBox_extDia.Text))
            {
                labelWarning.Text = "Please insert the value of ext.Diameter!";
                isDataValidated = false;
            }

            // Height (Null)
            if (string.IsNullOrEmpty(this.textBox_height.Text))
            {
                labelWarning.Text = "Please insert the value of the height!";
                isDataValidated = false;
            }

            // Process futher only if all the condition has been met
            // Prompt for the insertion point
            if(isDataValidated == true)
            {
                this.Hide();
                Global.Point.getInsertionPoint("\nClick the point where you want to insert the profile : ", doc);
            }
            
            // Locking the document
            using (DocumentLock docLock = doc.LockDocument())
            {
                // Transaction
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable bt;
                        bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                        BlockTableRecord btr;
                        btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        double insertPointX = Global.Point.insertPoint.X;
                        double insertPointY = Global.Point.insertPoint.Y;
                        Point3d insertPoint = new Point3d(insertPointX, insertPointY, 0);

                        // Define the dimensions of tacche
                        // Dimensions are according to TAB_27
                        double dimB = 1; 
                        double dimD = 1;

                        if(crossSection >= 5 && crossSection <= 7.5)
                        {
                            dimB = 1.5;
                            dimD = 1.5;
                        }
                        else if(crossSection > 7.5 && crossSection < 12.5)
                        {
                            dimB = 2;
                            dimD = 1.5;
                        }
                        else if (crossSection >= 12.5 && crossSection < 15)
                        {
                            dimB = 2.5;
                            dimD = 1.5;
                        }
                        else if (crossSection >= 15 && crossSection < 17.5)
                        {
                            dimB = 3;
                            dimD = 1.5;
                        }
                        else if (crossSection >= 17.5 && crossSection < 22)
                        {
                            dimB = 3.5;
                            dimD = 2;
                        }
                        else if (crossSection >= 22 && crossSection < 27.6)
                        {
                            dimB = 3.5;
                            dimD = 2.5;
                        }
                        else if (crossSection >= 27.6)
                        {
                            dimB = 3.5;
                            dimD = 3;
                        }
                        else
                        {
                            dimB = 4;
                            dimD = 3.5;
                        }

                        // Number of tacche
                        double numTacche = 2;
                        if (extDia <= 300)
                        {
                            numTacche = 2;
                        }
                        else if (extDia > 300 && extDia <= 500)
                        {
                            numTacche = 4;
                        }
                        else if (extDia > 500 && extDia <= 700)
                        {
                            numTacche = 6;
                        }
                        else if (extDia > 700 && extDia <= 1100)
                        {
                            numTacche = 8;
                        }
                        else if (extDia > 1100 && extDia <= 1500)
                        {
                            numTacche = 12;
                        }
                        else if (extDia > 1500)
                        {
                            numTacche = 16;
                        }
                        else
                        {
                            numTacche = 20;
                        }

                        // Define the dimension of the bottom height
                        // dimE = total height - height imp (altezzaColl)
                        double dimE = 1;

                        if (crossSection <= 7.5)
                        {
                            dimE = 2;
                        }
                        else if (crossSection > 7.5 && crossSection < 10)
                        {
                            dimE = 2.5;
                        }
                        else if (crossSection >= 10 && crossSection < 12.5)
                        {
                            dimE = 3;
                        }
                        else if (crossSection >= 12.5 && crossSection < 15)
                        {
                            dimE = 3.5;
                        }
                        else if (crossSection >= 15 && crossSection < 17.5)
                        {
                            dimE = 4;
                        }
                        else if (crossSection >= 17.5 && crossSection < 20)
                        {
                            dimE = 4.5;
                        }
                        else if (crossSection >= 20 && crossSection < 22.5)
                        {
                            dimE = 5;
                        }
                        else if (crossSection >= 22.5 && crossSection < 25)
                        {
                            dimE = 5.5;
                        }
                        else if (crossSection >= 25 && crossSection < 27.5)
                        {
                            dimE = 6;
                        }
                        else if (crossSection >= 27.5 && crossSection < 30)
                        {
                            dimE = 6.5;
                        }
                        else if (crossSection >= 30 && crossSection < 32.5)
                        {
                            dimE = 7;
                        }
                        else if (crossSection >= 32.5 && crossSection < 35)
                        {
                            dimE = 7.5;
                        }
                        else
                        {
                            dimE = 8;
                        }

                        // Fascia gap
                        /////////////////////// To be defined x = Fnom-Fcoll.
                        double fasciaGap = 1;
                        if(crossSection <= 8)
                        {
                            fasciaGap = 0.4;
                        }
                        else if (crossSection > 8 && crossSection < 10)
                        {
                            fasciaGap = 0.5;
                        }
                        else if (crossSection >= 10 && crossSection < 10.6)
                        {
                            fasciaGap = 0.6;
                        }
                        else if (crossSection >= 10.6 && crossSection < 14)
                        {
                            fasciaGap = 0.8;
                        }
                        else if (crossSection >= 14 && crossSection < 20.1 )
                        {
                            fasciaGap = 1;
                        }
                        else if (crossSection >= 20.1 && crossSection < 23.9)
                        {
                            fasciaGap = 1.2;
                        }
                        else if (crossSection >= 23.9 && crossSection < 26.1)
                        {
                            fasciaGap = 1.3;
                        }
                        else if (crossSection >= 26.1 && crossSection < 27.6)
                        {
                            fasciaGap = 1.4;
                        }
                        else if (crossSection >= 27.6 && crossSection < 30.1)
                        {
                            fasciaGap = 1.5;
                        }
                        else if (crossSection >= 30.1 && crossSection < 35.1)
                        {
                            fasciaGap = 1.6;
                        }
                        else
                        {
                            fasciaGap = 1.8;
                        }


                        // Calculate each point of the profile
                        double point2X = insertPointX;
                        double point2Y = insertPointY - dimB;
                        double point3X = insertPointX - dimD;
                        double point3Y = point2Y;
                        double point4X = point3X;
                        double point4Y = point3Y - (height - dimB);
                        double point5X = point4X;
                        double point5Y = point4Y - (dimE / Math.Sin((45 * (Math.PI / 180))));

                        double point10X = insertPointX + (crossSection - fasciaGap - (2 * dimD));
                        double point10Y = insertPointY;

                        double point9X = point10X;
                        double point9Y = point10Y - dimB;
                        double point8X = point9X + dimD;
                        double point8Y = point9Y;
                        double point7X = point8X;
                        double point7Y = point4Y;

                        double point6X = point7X;
                        double point6Y = point7Y - (dimE / Math.Sin((45 * (Math.PI / 180))));

                        // Create the points with the X,Y point values above
                        Point3d point2 = new Point3d(point2X, point2Y, 0);
                        Point3d point3 = new Point3d(point3X, point3Y, 0);
                        Point3d point4 = new Point3d(point4X, point4Y, 0);
                        Point3d point5 = new Point3d(point5X, point5Y, 0);
                        Point3d point6 = new Point3d(point6X, point6Y, 0);
                        Point3d point7 = new Point3d(point7X, point7Y, 0);
                        Point3d point8 = new Point3d(point8X, point8Y, 0);
                        Point3d point9 = new Point3d(point9X, point9Y, 0);
                        Point3d point10 = new Point3d(point10X, point10Y, 0);

                        //// point5 & point 6 shall be rotated 45°/-45°
                        //Matrix3d currentMatrix = doc.Editor.CurrentUserCoordinateSystem;
                        //CoordinateSystem3d crdSystem = currentMatrix.CoordinateSystem3d;

                        //point5.TransformBy(Matrix3d.Rotation(45 * (Math.PI / 180), crdSystem.Zaxis, point4));
                        //point6.TransformBy(Matrix3d.Rotation(-45 * (Math.PI / 180), crdSystem.Zaxis, point7));

                        //// Draw the profile as polyline
                        //using (Polyline pl = new Polyline())
                        //{
                        //    pl.AddVertexAt(0, new Point2d(point7X, point7Y), 0, 0, 0);
                        //    pl.AddVertexAt(1, new Point2d(point8X, point8Y), 0, 0, 0);
                        //    pl.AddVertexAt(2, new Point2d(point9X, point9Y), 0, 0, 0);
                        //    pl.AddVertexAt(3, new Point2d(point10X, point10Y), 0, 0, 0);
                        //    pl.AddVertexAt(4, new Point2d(insertPointX, insertPointY), 0, 0, 0);
                        //    pl.AddVertexAt(5, new Point2d(point2X, point2Y), 0, 0, 0);
                        //    pl.AddVertexAt(6, new Point2d(point3X, point3Y), 0, 0, 0);
                        //    pl.AddVertexAt(7, new Point2d(point4X, point4Y), 0, 0, 0);
                        //    pl.AddVertexAt(8, new Point2d(point5X, point5Y), 0, 0, 0);
                        //    pl.AddVertexAt(9, new Point2d(point6X, point6Y), 0, 0, 0);

                        //    pl.Closed = true;
                        //    //pl.ColorIndex = 1;

                        //    pl.SetDatabaseDefaults();
                        //    btr.AppendEntity(pl);
                        //    trans.AddNewlyCreatedDBObject(pl, true);
                        //}

                        // Draw the profile section
                        Line ln12 = new Line(insertPoint, point2);
                        Line ln23 = new Line(point2, point3);
                        Line ln34 = new Line(point3, point4);
                        Line ln45 = new Line(point4, point5);
                        //Line ln56 = new Line(point5, point6);
                        Line ln67 = new Line(point6, point7);
                        Line ln78 = new Line(point7, point8);
                        Line ln89 = new Line(point8, point9);
                        Line ln910 = new Line(point9, point10);
                        Line ln101 = new Line(point10, insertPoint);

                        // Point 5 & 6 shall be roated 45° / -45°
                        double pointRotated5X = point5X + dimE;
                        double pointRotated5Y = point5Y + (dimE / Math.Sin((45 * (Math.PI / 180)))) - dimE;
                        double pointRotated6X = point6X - dimE;
                        double pointRotated6Y = point6Y + (dimE / Math.Sin((45 * (Math.PI / 180)))) - dimE;

                        Point3d pointRotated5 = new Point3d(pointRotated5X, pointRotated5Y, 0);
                        Point3d pointRotated6 = new Point3d(pointRotated6X, pointRotated6Y, 0);

                        Line ln56 = new Line(pointRotated5, pointRotated6);

                       
                        // Plot
                        Line[] profileLineArray1 = new Line[] { ln12, ln23, ln34, ln56, ln78, ln89, ln910, ln101 };
                        for (int i = 0; i < 8; i++)
                        {
                            btr.AppendEntity(profileLineArray1[i]);
                            trans.AddNewlyCreatedDBObject(profileLineArray1[i], true);
                        }

                        // ln45 & ln67 shall be rotated 45°/-45°
                        Matrix3d currentMatrix = doc.Editor.CurrentUserCoordinateSystem;
                        CoordinateSystem3d crdSystem = currentMatrix.CoordinateSystem3d;

                        ln45.TransformBy(Matrix3d.Rotation(45 * (Math.PI / 180), crdSystem.Zaxis, point4));
                        ln67.TransformBy(Matrix3d.Rotation(-45 * (Math.PI / 180), crdSystem.Zaxis, point7));

                        // Plot rotated 2 lines : ln45 & ln67
                        ln45.SetDatabaseDefaults();
                        btr.AppendEntity(ln45);
                        trans.AddNewlyCreatedDBObject(ln45, true);

                        ln67.SetDatabaseDefaults();
                        btr.AppendEntity(ln67);
                        trans.AddNewlyCreatedDBObject(ln67, true);

                        // Profile line layer (Disegno)
                        Line[] profileLineArray2 = new Line[] { ln12, ln23, ln34, ln45, ln56, ln67, ln78, ln89, ln910, ln101 };
                        for (int i = 0; i < 10; i++)
                        {
                            profileLineArray2[i].Layer = "DISEGNO";
                        }

                        // Center line (Axis)
                        double axisPoint1X = point8X + 10 + (crossSection * 0.75);  // 10 = length of tacche
                        double axisPoint2X = point8X + 10 + (crossSection * 0.75);  // 10 = length of tacche
                        double axisPoint3X = point8X + 10 + (crossSection * 0.75);  // 10 = length of tacche
                        double axisPoint4X = point8X + 10 + (crossSection * 0.75);  // 10 = length of tacche

                        double axisPoint1Y = point10Y;
                        double axisPoint2Y = point8Y;
                        double axisPoint3Y = point7Y;
                        double axisPoint4Y = point7Y - dimE;

                        Point3d pointCnt1 = new Point3d(axisPoint1X, axisPoint1Y, 0);
                        Point3d pointCnt2 = new Point3d(axisPoint2X, axisPoint2Y, 0);
                        Point3d pointCnt3 = new Point3d(axisPoint3X, axisPoint3Y, 0);
                        Point3d pointCnt4 = new Point3d(axisPoint4X, axisPoint4Y, 0);

                        

                        // Tacche
                        // Calculate each point of tacche
                        // X value
                        double pointTacche1X = point10X + (crossSection * 0.5);
                        double pointTacche2X = pointTacche1X;
                        double pointTacche3X = point10X + (crossSection * 0.5) + 10;
                        double pointTacche4X = pointTacche3X;
                        // Y value
                        double pointTacche1Y = point10Y;
                        double pointTacche2Y = pointTacche1Y - dimB;
                        double pointTacche3Y = pointTacche2Y;
                        double pointTacche4Y = pointTacche1Y;
                        // Tacche points
                        Point3d pointTacche1 = new Point3d(pointTacche1X, pointTacche1Y, 0);
                        Point3d pointTacche2 = new Point3d(pointTacche2X, pointTacche2Y, 0);
                        Point3d pointTacche3 = new Point3d(pointTacche3X, pointTacche3Y, 0);
                        Point3d pointTacche4 = new Point3d(pointTacche4X, pointTacche4Y, 0);


                        // Center lines
                        // Profile lines to the center line
                        Line lnCnt11 = new Line(point10, pointTacche1);
                        Line lnCnt12 = new Line(pointTacche4, pointCnt1);
                        Line lnCnt21 = new Line(point8, pointTacche2);
                        Line lnCnt22 = new Line(pointTacche3, pointCnt2);
                        Line lnCnt3 = new Line(point7, pointCnt3);
                        Line lnCnt4 = new Line(pointRotated6, pointCnt4);

                        // Plot + Layer
                        Line[] centerLineArray = new Line[] { lnCnt11, lnCnt12, lnCnt21, lnCnt22, lnCnt3, lnCnt4 };
                        for (int i = 0; i < 6; i++)
                        {
                            btr.AppendEntity(centerLineArray[i]);
                            trans.AddNewlyCreatedDBObject(centerLineArray[i], true);
                            centerLineArray[i].Layer = "DISEGNO";
                        }

                        // Tacche
                        // Tacche lines
                        Line lnTacche2 = new Line(pointTacche1, pointTacche2);
                        Line lnTacche3 = new Line(pointTacche2, pointTacche3);
                        Line lnTacche4 = new Line(pointTacche3, pointTacche4);

                        // Plot + Layer
                        Line[] taccheLineArray = new Line[] {lnTacche2, lnTacche3, lnTacche4 };
                        for (int i = 0; i < 3; i++)
                        {
                            btr.AppendEntity(taccheLineArray[i]);
                            trans.AddNewlyCreatedDBObject(taccheLineArray[i], true);
                            taccheLineArray[i].Layer = "DISEGNO";
                        }
                        

                        // Axis line in drawing
                        // Calculate 2 points
                        double pointAxisDwg1X = axisPoint1X;
                        double pointAxisDwg2X = axisPoint1X;
                        double pointAxisDwg1Y = axisPoint1Y + dimB;
                        double pointAxisDwg2Y = axisPoint4Y - dimB;

                        // Axis line in drawing point3d
                        Point3d pointAxisDwg1 = new Point3d(pointAxisDwg1X, pointAxisDwg1Y, 0);
                        Point3d pointAxisDwg2 = new Point3d(pointAxisDwg2X, pointAxisDwg2Y, 0);

                        // Draw the line
                        Line axisLineDwg = new Line(pointAxisDwg1, pointAxisDwg2);
                        btr.AppendEntity(axisLineDwg);
                        trans.AddNewlyCreatedDBObject(axisLineDwg, true);

                        // Line layer (Assi)
                        axisLineDwg.Layer = "ASSI";

                        // Actual axis line
                        // ENDLESS/SPLIT selection
                        double endlessSplit = 0;
                        string endlessText = "";
                        string splitText = "";
                        if(radioButton_endless.Checked == true)
                        {
                            endlessSplit = 0;
                            endless = true;
                            endlessText = "Endless (Intera)";
                        }
                        else if (radioButton_split.Checked == true)
                        {
                            endlessSplit = 0.5;
                            split = true;
                            splitText = "Double splits";
                        }
                        // Calculate 2 points
                        double pointAxisReal1X = ((point3X + point8X) * 0.5) + ((intDia + extDia) * 0.25) + endlessSplit;
                        double pointAxisReal1Y = pointAxisDwg1Y;
                        double pointAxisReal2X = pointAxisReal1X;
                        double pointAxisReal2Y = pointAxisDwg2Y;

                        // point
                        Point3d pointAxisReal1 = new Point3d(pointAxisReal1X, pointAxisReal1Y, 0);
                        Point3d pointAxisReal2 = new Point3d(pointAxisReal2X, pointAxisReal2Y, 0);
                        Line lnAxisReal = new Line(pointAxisReal1, pointAxisReal2);
                        btr.AppendEntity(lnAxisReal);
                        trans.AddNewlyCreatedDBObject(lnAxisReal, true);

                        // Line layer (Assi)
                        lnAxisReal.Layer = "DEFPOINTS";


                        // Create the aligned dimension //////////////////////////////////////////////////////
                        // Tolerance variables
                        double tollFascia = 0.1;
                        double tollAltezzaColl = 0.1;
                        double tollAltezzaTotal = 0.1;

                        // Cross-section tolerance range (F = 15)
                        if ((point8X - point3X) < 15)
                        {
                            tollFascia = 0.1;
                        }
                        else if((point8X - point3X) >= 15)
                        {
                            tollFascia = 0.15;
                        }

                        // Altezza Coll. (Height imp)
                        if ((insertPointY-point4Y) < 15)
                        {
                            tollAltezzaColl = 0.1;
                        }
                        else if ((insertPointY - point4Y) >= 15)
                        {
                            tollAltezzaColl = 0.15;
                        }

                        // Altezza total (Total height)
                        if ((insertPointY - pointRotated5Y) < 15)
                        {
                            tollAltezzaTotal = 0.1;
                        }
                        else if ((insertPointY - pointRotated5Y) >= 15)
                        {
                            tollAltezzaTotal = 0.15;
                        }

                        // Open the dimension style and set it to "P"
                        DimStyleTableRecord dstr = new DimStyleTableRecord();
                        DimStyleTable dst = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                        ObjectId dimStyleP = dst["P"];
                        doc.Database.Dimstyle = dimStyleP;

                        // Fascia Coll
                        RotatedDimension fasciaColl = new RotatedDimension();
                        fasciaColl.SetDatabaseDefaults();
                        fasciaColl.XLine1Point = point8;
                        fasciaColl.XLine2Point = point3;
                        fasciaColl.DimLinePoint = new Point3d(0, insertPointY + (dimB * 2.8), 0);
                        fasciaColl.Layer = "QUOTE";
                        fasciaColl.DimensionStyle = doc.Database.Dimstyle;

                        // Tolerance
                        fasciaColl.Dimtol = true;
                        fasciaColl.Dimtp = tollFascia;
                        fasciaColl.Dimtm = tollFascia;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(fasciaColl);
                        trans.AddNewlyCreatedDBObject(fasciaColl, true);


                        // Dimension D (TAB_27)
                        RotatedDimension dimensionD = new RotatedDimension();
                        dimensionD.SetDatabaseDefaults();
                        dimensionD.XLine1Point = point3;
                        dimensionD.XLine2Point = insertPoint;
                        dimensionD.DimLinePoint = new Point3d(0, insertPointY + (dimB * 1.25), 0);
                        dimensionD.Layer = "QUOTE";
                        dimensionD.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(dimensionD);
                        trans.AddNewlyCreatedDBObject(dimensionD, true);


                        // Dimension B (TAB_27)
                        RotatedDimension dimensionB = new RotatedDimension();
                        dimensionB.SetDatabaseDefaults();
                        dimensionB.XLine1Point = insertPoint;
                        dimensionB.XLine2Point = point3;
                        dimensionB.Rotation = Math.PI / 2;
                        dimensionB.DimLinePoint = new Point3d(point3X - (dimD * 1.5), 0, 0);
                        dimensionB.Layer = "QUOTE";
                        dimensionB.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(dimensionB);
                        trans.AddNewlyCreatedDBObject(dimensionB, true);


                        // Altezza coll (height)
                        RotatedDimension altezzaColl = new RotatedDimension();
                        altezzaColl.SetDatabaseDefaults();
                        altezzaColl.XLine1Point = insertPoint;
                        altezzaColl.XLine2Point = point4;
                        altezzaColl.Rotation = Math.PI / 2;
                        altezzaColl.DimLinePoint = new Point3d(point3X - (dimD * 2.8), (point3Y + point4Y) * 0.5, 0);
                        altezzaColl.Layer = "QUOTE";
                        altezzaColl.DimensionStyle = doc.Database.Dimstyle;

                        // Tolerance
                        altezzaColl.Dimtol = true;
                        altezzaColl.Dimtp = tollAltezzaColl;
                        altezzaColl.Dimtm = tollAltezzaColl;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(altezzaColl);
                        trans.AddNewlyCreatedDBObject(altezzaColl, true);


                        // Altezza total (total height)
                        RotatedDimension altezzaTotal = new RotatedDimension();
                        altezzaTotal.SetDatabaseDefaults();
                        altezzaTotal.XLine1Point = insertPoint;
                        altezzaTotal.XLine2Point = pointRotated5;
                        altezzaTotal.Rotation = Math.PI / 2;
                        altezzaTotal.DimLinePoint = new Point3d(point3X - (dimD * 4.15), (insertPointY + pointRotated5Y) * 0.5, 0);
                        altezzaTotal.Layer = "QUOTE";
                        altezzaTotal.DimensionStyle = doc.Database.Dimstyle;

                        // Tolerance
                        altezzaTotal.Dimtol = true;
                        altezzaTotal.Dimtp = tollAltezzaTotal;
                        altezzaTotal.Dimtm = tollAltezzaTotal;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(altezzaTotal);
                        trans.AddNewlyCreatedDBObject(altezzaTotal, true);


                        // Tacche (Dimension A = 10, TAB_27)
                        RotatedDimension dimensionA = new RotatedDimension();
                        dimensionA.SetDatabaseDefaults();
                        dimensionA.XLine1Point = pointTacche1;
                        dimensionA.XLine2Point = pointTacche4;
                        dimensionA.DimLinePoint = new Point3d(0, insertPointY + (dimB * 1.25), 0);
                        dimensionA.Layer = "QUOTE";
                        dimensionA.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(dimensionA);
                        trans.AddNewlyCreatedDBObject(dimensionA, true);


                        // Create an angular dimension (90°)
                        LineAngularDimension2 acLinAngDim = new LineAngularDimension2();
                        acLinAngDim.SetDatabaseDefaults();
                        acLinAngDim.XLine1Start = point4;
                        acLinAngDim.XLine1End = pointRotated5;
                        acLinAngDim.XLine2Start = pointRotated6;
                        acLinAngDim.XLine2End = point7;
                        acLinAngDim.ArcPoint = new Point3d((point4X + point7X) * 0.5, point4Y + 2, 0);
                        acLinAngDim.Layer = "QUOTE";
                        acLinAngDim.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(acLinAngDim);
                        trans.AddNewlyCreatedDBObject(acLinAngDim, true);


                        // External diameter coll.
                        // Diameter dimensions (P-D2)
                        ObjectId dimStylePD2 = dst["P-D2"];

                        RotatedDimension extDiaColl = new RotatedDimension();
                        extDiaColl.SetDatabaseDefaults();
                        extDiaColl.XLine1Point = new Point3d(pointAxisReal1X, point8Y, 0);
                        extDiaColl.XLine2Point = point3;
                        extDiaColl.DimLinePoint = new Point3d(point8X, insertPointY + (dimB * 4.05), 0);
                        extDiaColl.Layer = "QUOTE";
                        doc.Database.Dimstyle = dimStylePD2;
                        extDiaColl.DimensionStyle = doc.Database.Dimstyle;

                        // Tolerance
                        // External diameter
                        extDiaColl.Dimtol = true;
                        double tollExtDia = 1;

                        if(split == true)
                        {
                            if (((pointAxisReal1X - point3X) * 2) <= 1000)
                            {
                                extDiaColl.Dimtp = 1;
                                extDiaColl.Dimtm = 0;
                            }
                            else if (((pointAxisReal1X - point3X) * 2) > 1000)
                            {
                                tollExtDia = Math.Round(extDia * 0.001,2);
                                extDiaColl.Dimtp = tollExtDia;
                                extDiaColl.Dimtm = 0;
                            }
                        }

                        else if (endless == true)
                        {
                            if (((pointAxisReal1X - point3X) * 2) > 600)
                            {
                                tollExtDia = extDia * 0.001;
                                extDiaColl.Dimtp = tollExtDia;
                                extDiaColl.Dimtm = 0;
                            }
                            else if (((pointAxisReal1X - point3X) * 2) <= 200)
                            {
                                tollExtDia = 0.15;
                                extDiaColl.Dimtp = tollExtDia;
                                extDiaColl.Dimtm = tollExtDia;
                            }
                            else if (((pointAxisReal1X - point3X) * 2) > 200 && ((pointAxisReal1X - point3X) * 2) <= 400)
                            {
                                tollExtDia = 0.2;
                                extDiaColl.Dimtp = tollExtDia;
                                extDiaColl.Dimtm = tollExtDia;
                            }
                            else if (((pointAxisReal1X - point3X) * 2) > 400 && ((pointAxisReal1X - point3X) * 2) <= 600)
                            {
                                tollExtDia = 0.25;
                                extDiaColl.Dimtp = tollExtDia;
                                extDiaColl.Dimtm = tollExtDia;
                            }
                        }
                        
                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(extDiaColl);
                        trans.AddNewlyCreatedDBObject(extDiaColl, true);


                        // Diameter at point5
                        RotatedDimension diaPoint5 = new RotatedDimension();
                        diaPoint5.SetDatabaseDefaults();
                        diaPoint5.XLine1Point = new Point3d(pointAxisReal1X, pointRotated5Y, 0);
                        diaPoint5.XLine2Point = pointRotated5;
                        diaPoint5.DimLinePoint = new Point3d(pointRotated5X, pointRotated5Y - (dimB * 2.2), 0);
                        diaPoint5.Layer = "QUOTE";
                        doc.Database.Dimstyle = dimStylePD2;
                        diaPoint5.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(diaPoint5);
                        trans.AddNewlyCreatedDBObject(diaPoint5, true);


                        // Diameter at point6
                        RotatedDimension diaPoint6 = new RotatedDimension();
                        diaPoint6.SetDatabaseDefaults();
                        diaPoint6.XLine1Point = new Point3d(pointAxisReal1X, pointRotated6Y, 0);
                        diaPoint6.XLine2Point = pointRotated6;
                        diaPoint6.DimLinePoint = new Point3d(pointRotated6X, pointRotated6Y - (dimB * 1.2), 0);
                        diaPoint6.Layer = "QUOTE";
                        doc.Database.Dimstyle = dimStylePD2;
                        diaPoint6.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(diaPoint6);
                        trans.AddNewlyCreatedDBObject(diaPoint6, true);


                        // Internal diameter coll.
                        RotatedDimension intDiaColl = new RotatedDimension();
                        intDiaColl.SetDatabaseDefaults();
                        intDiaColl.XLine1Point = new Point3d(pointAxisReal1X, point8Y, 0);
                        intDiaColl.XLine2Point = point8;
                        intDiaColl.DimLinePoint = new Point3d(point8X, insertPointY + (dimB * 2.8), 0);

                        // Override the dimension (ref.)
                        intDiaColl.DimensionText = "(<>)";

                        //acRotDim.DimensionStyle = acCurDb.Dimstyle;
                        intDiaColl.Layer = "QUOTE";

                        doc.Database.Dimstyle = dimStylePD2;
                        intDiaColl.DimensionStyle = doc.Database.Dimstyle;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(intDiaColl);
                        trans.AddNewlyCreatedDBObject(intDiaColl, true);


                        // Get the dimension style from P-D2 to P
                        doc.Database.Dimstyle = dimStyleP;


                        ///// Create the leader ///////////////////////////////////////////////////////
                        Leader leader = new Leader();
                        leader.SetDatabaseDefaults();
                        leader.AppendVertex(new Point3d(pointTacche2X + 2, pointTacche2Y, 0));

                        // Verify the space for leader text
                        if((point8Y - point7Y) >= 6.5)
                        {
                            leader.AppendVertex(new Point3d(pointTacche2X + 2, pointTacche2Y - dimB, 0));
                            leader.AppendVertex(new Point3d(pointTacche2X + 3, pointTacche2Y - dimB, 0));
                        }
                        else if((point8Y - point7Y) <6.5)
                        {
                            leader.AppendVertex(new Point3d(pointTacche2X + 2, pointRotated6Y - dimB, 0));
                            leader.AppendVertex(new Point3d(pointTacche2X + 3, pointRotated6Y - dimB, 0));
                        }
                        leader.HasArrowHead = true;
                        leader.Layer = "QUOTE";

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(leader);
                        trans.AddNewlyCreatedDBObject(leader, true);

                        // Create the leader text (TACCHE COME DA ...)
                        MText leaderText = new MText();
                        leaderText.SetDatabaseDefaults();

                        // Verify the space for leader text
                        if ((point8Y - point7Y) >= 6.5)
                        {
                            leaderText.Location = new Point3d(pointTacche2X + 3, pointTacche2Y - dimB + 0.5, 0);  // Connected to the third point of the leader
                        }
                        else if ((point8Y - point7Y) < 6.5)
                        {
                            leaderText.Location = new Point3d(pointTacche2X + 3, pointRotated6Y - dimB + 0.5, 0);  // Connected to the third point of the leader
                        }
                        leaderText.Width = 15;

                        // Adjust text height
                        double textHeight = 1;
                        if(crossSection < 11)
                        {
                            textHeight = 0.8;
                        }
                        leaderText.TextHeight = textHeight;

                        leaderText.Layer = "0";
                        leaderText.Contents = "TACCHE COME DA\nTAB_27\nN°" + numTacche + "x" + dimB + "x10";

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(leaderText);
                        trans.AddNewlyCreatedDBObject(leaderText, true);


                        ////// HATCH //////////////////////////////////////////////////////////////////////
                        // Create the hatches
                        // Adds the circle to an object id array
                        ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                        for (int i = 0; i < 10; i++)
                        {
                            acObjIdColl.Add(profileLineArray2[i].ObjectId);
                        }

                        // Create the hatch object and append it to the block table record
                        Hatch acHatch = new Hatch();

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(acHatch);
                        trans.AddNewlyCreatedDBObject(acHatch, true);

                        // Set the properties of the hatch object
                        // Associative must be set after the hatch object is appended to the
                        // block table record and before AppendLoop
                        acHatch.SetDatabaseDefaults();
                        acHatch.Layer = "FINE";

                        // Adjust the pattern scale
                        double patternScale = 1;
                        if (crossSection < 10)
                        {
                            patternScale = 0.15;
                        }
                        else if (crossSection >= 10 && crossSection < 19)
                        {
                            patternScale = 0.25;
                        }
                        else if (crossSection >= 19 && crossSection < 34)
                        {
                            patternScale = 0.5;
                        }
                        acHatch.PatternScale = patternScale;   // Pattern scale to be defined according to the size of crossSection

                        // Material type (Nylon / PTFE)
                        string matNylon = "";
                        string matPTFE = "";
                        if(radioButton_nylon.Checked == true)
                        {
                            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI32");
                            matNylon = "Nylon (PA6)";
                        }
                        else if(radioButton_ptfe.Checked == true)
                        {
                            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI34");
                            matPTFE = "PTFE";
                        }
                        acHatch.Associative = true;

                        // Add the new object to Model space and the transaction
                        acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                        acHatch.EvaluateHatch(true);


                        // General information of the profile
                        MText infoText = new MText();
                        infoText.SetDatabaseDefaults();
                        infoText.Location = new Point3d(axisPoint1X + (crossSection * 0.5), insertPointY, 0);  // The right side of the drawing
                        infoText.Width = 20;
                        infoText.TextHeight = textHeight;
                        infoText.Layer = "DEFPOINTS";
                        infoText.Contents = "AUTOGENERATED MALE RING\n" + intDia + "/" + extDia + "x" + height + "\n" + endlessText + splitText + "\nMaterial : " + matNylon + matPTFE;

                        // Add the new object to Model space and the transaction
                        btr.AppendEntity(infoText);
                        trans.AddNewlyCreatedDBObject(infoText, true);

                        // Commit the transaction
                        if(isDataValidated == true)
                        {
                            trans.Commit();
                            // Complete message
                            doc.Editor.WriteMessage("\nThe profile has been successfully created!\n");
                        }
                    }

                    // Exception handling
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage("Error encountered : " + ex.Message);
                        trans.Abort();
                    }
                }
            }

        }
    }
}
