using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CefSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace TeslaRevitTools.FindAffectedSheets
{
    public class CheckedView
    {
        public View View { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }

        public CheckedView(View view)
        {
            View = view;
            BoundingBox = View.get_BoundingBox(null);
        }

        public bool CropBoxContains(Element selectedElement, Document doc)
        {
            if (!View.CropBoxActive) return true;

            BoundingBoxXYZ elBbox = selectedElement.get_BoundingBox(null);

            BoundingBoxXYZ bbox = View.CropBox;
            XYZ bboxmin1 = bbox.Transform.OfPoint(bbox.Min);
            XYZ bboxmax1 = bbox.Transform.OfPoint(bbox.Max);

            List<XYZ> minmax = GetMinMax(bboxmin1, bboxmax1);
            XYZ bboxmin = minmax[0];
            XYZ bboxmax = minmax[1];

            bool outsideX = elBbox.Min.X >= bboxmax.X || elBbox.Max.X <= bboxmin.X;
            bool outsideY = elBbox.Min.Y >= bboxmax.Y || elBbox.Max.Y <= bboxmin.Y;
            bool outsideZ = elBbox.Min.Z >= bboxmax.Z || elBbox.Max.Z <= bboxmin.Z;


            if (View.ViewType == ViewType.Section)
            {
                return !outsideX && !outsideY && !outsideZ;
            }
            else if (View.ViewType == ViewType.FloorPlan)
            {
                try
                {
                    ViewPlan viewPlan = View as ViewPlan;
                    PlanViewRange viewRange = viewPlan.GetViewRange();
                    ElementId topId = viewRange.GetLevelId(PlanViewPlane.TopClipPlane);
                    double topOffset = viewRange.GetOffset(PlanViewPlane.TopClipPlane);

                    ElementId depthId = viewRange.GetLevelId(PlanViewPlane.ViewDepthPlane);
                    double depthOffset = viewRange.GetOffset(PlanViewPlane.ViewDepthPlane);

                    Level topLevel = (doc.GetElement(topId) as Level);
                    if(topLevel == null) return true;
                    double topLevelElevation = topLevel.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();

                    Level depthLevel = (doc.GetElement(depthId) as Level);
                    if (depthLevel == null) return true;

                    double depthLevelElevation = depthLevel.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();

                    double zMin = depthLevelElevation + depthOffset + depthLevel.ProjectElevation;
                    double zMax = topLevelElevation + topOffset + topLevel.ProjectElevation;

                    bool outsideZplanView = elBbox.Min.Z >= zMax || elBbox.Max.Z <= zMin;

                    return !outsideX && !outsideY && !outsideZplanView;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", "Seems like the view has some undefined height or depth:" + Environment.NewLine + ex.ToString());
                    return false;
                }

            }
            else if (View.ViewType == ViewType.CeilingPlan)
            {
                return !outsideX && !outsideY;
            }
            else if (View.ViewType == ViewType.ThreeD) 
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private List<XYZ> GetMinMax(XYZ boxMin, XYZ boxMax)
        {
            List<double> min = new List<double>() { boxMin.X, boxMin.Y, boxMin.Z };
            List<double> max = new List<double>() { boxMax.X, boxMax.Y, boxMax.Z };


            if (boxMin.X > boxMax.X)
            {
                min[0] = boxMax.X;
                max[0] = boxMin.X;
            }

            if (boxMin.Y > boxMax.Y)
            {
                min[1] = boxMax.Y;
                max[1] = boxMin.Y;
            }

            if (boxMin.Z > boxMax.Z)
            {
                min[2] = boxMax.Z;
                max[2] = boxMin.Z;
            }

            XYZ boxMinRes = new XYZ(min[0], min[1], min[2]);
            XYZ boxMaxRes = new XYZ(max[0], max[1], max[2]);

            return new List<XYZ>() { boxMinRes, boxMaxRes };
        }
    }
}
