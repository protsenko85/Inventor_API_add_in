using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;

namespace InventorAddIn
{
    public partial class Form1 : Form
    {
        //змінна для підключення до Інвентору
        Inventor.Application mApp;
        //змінна одиниць виміру Інвентору
        UnitsOfMeasure mUOM;

        public Form1()
        {
            //підключення до Інверноту
            Inventor.Application oApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            InitializeComponent();
            mApp = oApp;

            //Підключення до деталі
            PartDocument oPartDoc = AddinGlobal.InventorApp.ActiveDocument as PartDocument;
            PartComponentDefinition oDef = oPartDoc.ComponentDefinition;
            TransientGeometry oTG = AddinGlobal.InventorApp.TransientGeometry;

            //Визначення скетчу і профілю
            PlanarSketch oSketch;
            Profile oProfile;

            //визначення площини для виконання креслень
            WorkPlane oWorkPlane = oDef.WorkPlanes["XY Plane"];

            //додавання площини до скетчу
            oSketch = oDef.Sketches.Add(oWorkPlane, false);

            //додавання точок до екскізу, за якими буватиметься модель
            SketchPoint point1 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(3.3, 0));
            SketchPoint point2 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(3.3, -4.12));
            SketchPoint point3 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.3, -4.12));
            SketchPoint point4 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.3, -7));
            SketchPoint point5 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(7, -7));
            SketchPoint point6 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(7, -5));
            SketchPoint point7 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.6, -5));
            SketchPoint point8 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.6, -4.8));
            SketchPoint point9 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.8, -4.8));
            SketchPoint point10 = oSketch.SketchPoints.Add(oTG.CreatePoint2d(4.8, 0));

            //з'єднання точок
            oSketch.SketchLines.AddByTwoPoints(point1, point2);
            oSketch.SketchLines.AddByTwoPoints(point2, point3);
            oSketch.SketchLines.AddByTwoPoints(point3, point4);
            oSketch.SketchLines.AddByTwoPoints(point4, point5);
            oSketch.SketchLines.AddByTwoPoints(point5, point6);
            oSketch.SketchLines.AddByTwoPoints(point6, point7);
            oSketch.SketchLines.AddByTwoPoints(point7, point8);
            oSketch.SketchLines.AddByTwoPoints(point8, point9);
            oSketch.SketchLines.AddByTwoPoints(point9, point10);
            oSketch.SketchLines.AddByTwoPoints(point10, point1);

            //додавання екскізу в профіль
            oProfile = oSketch.Profiles.AddForSolid();

            //Створення моделі шляхом обертання навколо осі
            WorkAxis AxisEntity = oPartDoc.ComponentDefinition.WorkAxes[2];
            RevolveFeature r = oDef.Features.RevolveFeatures.AddFull(oProfile, AxisEntity, PartFeatureOperationEnum.kJoinOperation);

            //Створення вирізу
            PlanarSketch oSketch2 = oDef.Sketches.Add(r.Faces[3], false);
            SketchPoint CenterPoint = oSketch2.SketchPoints.Add(oTG.CreatePoint2d(0, 5.9));
            WorkPoint wp = oPartDoc.ComponentDefinition.WorkPoints.AddByPoint(CenterPoint);
            PointHolePlacementDefinition phpd = oDef.Features.HoleFeatures.CreatePointPlacementDefinition(wp, oDef.WorkAxes[2]);
            HoleFeature oNewHole = oDef.Features.HoleFeatures.AddCBoreByThroughAllExtent(phpd, 1.2, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, 1.6, 0.6);

            //Створення масиву вирізів
            ObjectCollection oColl = default(ObjectCollection);
            oColl = AddinGlobal.InventorApp.TransientObjects.CreateObjectCollection();
            //parameters for circle pattern
            oColl.Add(oNewHole);
            //circular array
            CircularPatternFeature SPF = oDef.Features.CircularPatternFeatures.Add(oColl, oPartDoc.ComponentDefinition.WorkAxes[2], true, 4, "360", true, PatternComputeTypeEnum.kAdjustToModelCompute);
            
            //встановлення позиції для камери
            Camera cam = AddinGlobal.InventorApp.ActiveView.Camera;
            cam.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
            cam.Apply();
            AddinGlobal.InventorApp.ActiveView.Fit(true);

            PartDocument oDoc = mApp.ActiveDocument as PartDocument;
            PartComponentDefinition oDef1 = oDoc.ComponentDefinition;

            //створення фасок та скруглень
            Edge oEdge = r.Faces[2].Edges[1] as Edge;
            //create object collection
            EdgeCollection oColl2 = default(EdgeCollection);
            oColl2 = AddinGlobal.InventorApp.TransientObjects.CreateEdgeCollection();
            oColl2.Add(oEdge);
            ChamferFeature Ff = default(ChamferFeature);
            Ff = oDef1.Features.ChamferFeatures.AddUsingTwoDistances(oColl2, r.Faces[2], 0.32, 0.25, false, false);


            oEdge = r.Faces[9].Edges[1] as Edge;
            //create object collection
            oColl2 = default(EdgeCollection);
            oColl2 = AddinGlobal.InventorApp.TransientObjects.CreateEdgeCollection();
            oColl2.Add(oEdge);
            ChamferFeature Ff1 = default(ChamferFeature);
            Ff1 = oDef1.Features.ChamferFeatures.AddUsingTwoDistances(oColl2, r.Faces[9], 0.5, 1.5, true, false);

            oEdge = r.Faces[10].Edges[1] as Edge;
            //create object collection
            oColl2 = default(EdgeCollection);
            oColl2 = AddinGlobal.InventorApp.TransientObjects.CreateEdgeCollection();
            oColl2.Add(oEdge);
            Ff1 = default(ChamferFeature);
            Ff1 = oDef1.Features.ChamferFeatures.AddUsingTwoDistances(oColl2, r.Faces[10], 0.2, 0.2, false, false);

            Edge oEdge4 = r.Faces[5].Edges[1] as Edge;
            //create object collection
            EdgeCollection oColl4 = default(EdgeCollection);
            oColl4 = AddinGlobal.InventorApp.TransientObjects.CreateEdgeCollection();
            oColl4.Add(oEdge4);
            FilletFeature Ff4 = default(FilletFeature);
            Ff4 = oDef1.Features.FilletFeatures.AddSimple(oColl4, 0.1);


            Edge oEdge3 = r.Faces[2].Edges[1] as Edge;
            //create object collection
            EdgeCollection oColl3 = default(EdgeCollection);
            oColl3 = AddinGlobal.InventorApp.TransientObjects.CreateEdgeCollection();
            oColl3.Add(oEdge3);
            FilletFeature Ff3 = default(FilletFeature);
            Ff3 = oDef1.Features.FilletFeatures.AddSimple(oColl3, 0.1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            mUOM = mApp.ActiveDocument.UnitsOfMeasure;

            //перевірка на присутність даних для оновлення
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "")
            {
                System.Windows.Forms.MessageBox.Show("Error: Введіть параметр для оновлення...");
            }
            else
            {
                PartDocument oPartDoc = mApp.ActiveDocument as PartDocument;
                Parameter oParam;
                double value;

                //перевірки на правильність данних
                if (textBox1.Text != "")
                {
                    if ((Convert.ToDouble(textBox1.Text) * Convert.ToDouble(oPartDoc.ComponentDefinition.Parameters["d5"].Value*10)) > 363.0)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: зі вказаним розміром отвіри виходять за межі деталі!");
                    }
                    else
                    {
                        //зміна параметрів детали в середині Інверноту
                        oParam = oPartDoc.ComponentDefinition.Parameters["d10"];
                        value = (double)mUOM.GetValueFromExpression(textBox1.Text, UnitsTypeEnum.kDefaultDisplayLengthUnits);
                        oParam.Value = (value * 10);
                    }
                    textBox1.Text = "";
                }
                    
                
                if (textBox2.Text != "")
                {
                    if (Convert.ToDouble(textBox2.Text) >= 22 || (Convert.ToDouble(textBox2.Text) * (oPartDoc.ComponentDefinition.Parameters["d10"].Value*10)) > 363)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: зі вказаним розміром отвіри виходять за межі деталі!");
                    }
                    else if (Convert.ToDouble(textBox2.Text) < oPartDoc.ComponentDefinition.Parameters["d3"].Value)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: великий діаметр отвора не може бути меншим нім малий!");
                    }
                    else
                    {
                        oParam = oPartDoc.ComponentDefinition.Parameters["d5"];
                        value = (double)mUOM.GetValueFromExpression(textBox2.Text, UnitsTypeEnum.kDefaultDisplayLengthUnits);
                        oParam.Value = value;
                    }
                    textBox2.Text = "";
                }

                if(textBox3.Text != "")
                {
                    if (Convert.ToDouble(textBox3.Text) < oPartDoc.ComponentDefinition.Parameters["d6"].Value)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: великий діаметр отвора не може бути меншим нім малий!");
                    }
                    else
                    {
                        oParam = oPartDoc.ComponentDefinition.Parameters["d3"];
                        value = (double)mUOM.GetValueFromExpression(textBox3.Text, UnitsTypeEnum.kDefaultDisplayLengthUnits);
                        oParam.Value = value;
                    }
                    textBox3.Text = "";
                }

                if (textBox4.Text != "")
                {
                    if (Convert.ToDouble(textBox4.Text) >= 20)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: глубина занадто велика!");
                    }
                    else if (Convert.ToDouble(textBox4.Text) <= 0)
                    {
                        System.Windows.Forms.MessageBox.Show("Error: глубина повинна існувати!");
                    }
                    else
                    {
                        oParam = oPartDoc.ComponentDefinition.Parameters["d6"];
                        value = (double)mUOM.GetValueFromExpression(textBox4.Text, UnitsTypeEnum.kDefaultDisplayLengthUnits);
                        oParam.Value = value;
                    }
                    textBox4.Text = "";
                }

                //Прийняття змін
                oPartDoc.Update();
                mApp.ActiveView.Fit(true);
            }


        }
    }
}
