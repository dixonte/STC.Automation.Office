using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the chart type.
    /// </summary>
    public enum ChartType
    {
        /// <summary>
        /// 3D Area Chart
        /// </summary>
        xl3DArea = -4098,
        /// <summary>
        /// 3D Area Stacked Chart
        /// </summary>
        xl3DAreaStacked = 78,
        /// <summary>
        /// 3D Area Stacked 100  Chart
        /// </summary>
        xl3DAreaStacked100 = 79,
        /// <summary>
        /// 3D Bar Clusters Chart
        /// </summary>
        xl3DBarClustered = 60,
        /// <summary>
        /// 3D Bar Stacked Chart
        /// </summary>
        xl3DBarStacked = 61,
        /// <summary>
        /// 3D Bar Stacked 100 Chart
        /// </summary>
        xl3DBarStacked100 = 62,
        /// <summary>
        /// 3D Column Chart
        /// </summary>
        xl3DColumn = -4100,
        /// <summary>
        /// 3D Column Clustered Chart
        /// </summary>
        xl3DColumnClustered = 54,
        /// <summary>
        /// 3D Column Stacked Chart
        /// </summary>
        xl3DColumnStacked = 55,
        /// <summary>
        /// 3D Column Stacked 100
        /// </summary>
        xl3DColumnStacked100 = 56,
        /// <summary>
        /// 3D Line
        /// </summary>
        xl3DLine = -4101,
        /// <summary>
        /// 3d Pie
        /// </summary>
        xl3DPie = -4102,
        /// <summary>
        /// 3d pie exploded
        /// </summary>
        xl3DPieExploded = 70,
        /// <summary>
        /// Area
        /// </summary>
        xlArea = 1,
        /// <summary>
        /// Area stacked
        /// </summary>
        xlAreaStacked = 76,
        /// <summary>
        /// Area stacked 100
        /// </summary>
        xlAreaStacked100 = 77,
        /// <summary>
        /// bar clustered
        /// </summary>
        xlBarClustered = 57,
        /// <summary>
        /// bar of pie
        /// </summary>
        xlBarOfPie = 71,
        /// <summary>
        /// bar stacked
        /// </summary>
        xlBarStacked = 58,
        /// <summary>
        /// bar stacked 100
        /// </summary>
        xlBarStacked100 = 59,
        /// <summary>
        /// buggle
        /// </summary>
        xlBubble = 15,
        /// <summary>
        /// bubble 3d effect
        /// </summary>
        xlBubble3DEffect = 87,
        /// <summary>
        /// column clustered
        /// </summary>
        xlColumnClustered = 51,
        /// <summary>
        /// column stacked
        /// </summary>
        xlColumnStacked = 52,
        /// <summary>
        /// columnstacked 100
        /// </summary>
        xlColumnStacked100 = 53,
        /// <summary>
        /// bar clustered
        /// </summary>
        xlConeBarClustered = 102,
        /// <summary>
        /// cone bar stacked
        /// </summary>
        xlConeBarStacked = 103,
        /// <summary>
        /// cone bar stacked 100
        /// </summary>
        xlConeBarStacked100 = 104,
        /// <summary>
        /// cone col
        /// </summary>
        xlConeCol = 105,
        /// <summary>
        /// cone col clustered
        /// </summary>
        xlConeColClustered = 99,
        /// <summary>
        /// con col stacked
        /// </summary>
        xlConeColStacked = 100,
        /// <summary>
        /// cone col stacked 100
        /// </summary>
        xlConeColStacked100 = 101,
        /// <summary>
        /// Cylinder bar clustered
        /// </summary>
        xlCylinderBarClustered = 95,
        /// <summary>
        /// Cylinder bar stacked
        /// </summary>
        xlCylinderBarStacked = 96,
        /// <summary>
        /// cylinder bar stacked 100
        /// </summary>
        xlCylinderBarStacked100 = 97,
        /// <summary>
        /// cylinder col
        /// </summary>
        xlCylinderCol = 98,
        /// <summary>
        /// cylinder col clustered
        /// </summary>
        xlCylinderColClustered = 92,
        /// <summary>
        /// cylinder col stacked
        /// </summary>
        xlCylinderColStacked = 93,
        /// <summary>
        /// cylinder col stacked 100
        /// </summary>
        xlCylinderColStacked100 = 94,
        /// <summary>
        /// doughnut
        /// </summary>
        xlDoughnut = -4120,
        /// <summary>
        /// doughnut exploded
        /// </summary>
        xlDoughnutExploded = 80,
        /// <summary>
        /// line
        /// </summary>
        xlLine = 4,
        /// <summary>
        /// line markers
        /// </summary>
        xlLineMarkers = 65,
        /// <summary>
        /// line marker stacked
        /// </summary>
        xlLineMarkersStacked = 66,
        /// <summary>
        /// line marker stacked 100
        /// </summary>
        xlLineMarkersStacked100 = 67,
        /// <summary>
        /// line stacked
        /// </summary>
        xlLineStacked = 63,
        /// <summary>
        /// line stacked 100
        /// </summary>
        xlLineStacked100 = 64,
        /// <summary>
        /// pie
        /// </summary>
        xlPie = 5,
        /// <summary>
        /// pie exploded
        /// </summary>
        xlPieExploded = 69,
        /// <summary>
        /// pie of pie
        /// </summary>
        xlPieOfPie = 68,
        /// <summary>
        /// pyramid bar clustered
        /// </summary>
        xlPyramidBarClustered = 109,
        /// <summary>
        /// pyramid bar stacked
        /// </summary>
        xlPyramidBarStacked = 110,
        /// <summary>
        /// pyramid bar stacked 100
        /// </summary>
        xlPyramidBarStacked100 = 111,
        /// <summary>
        /// pyramid col
        /// </summary>
        xlPyramidCol = 112,
        /// <summary>
        /// pyramid col clustered
        /// </summary>
        xlPyramidColClustered = 106,
        /// <summary>
        /// pyramid col stacked
        /// </summary>
        xlPyramidColStacked = 107,
        /// <summary>
        /// pyramid col stacked 100
        /// </summary>
        xlPyramidColStacked100 = 108,
        /// <summary>
        /// radar
        /// </summary>
        xlRadar = -4151,
        /// <summary>
        /// radar filled
        /// </summary>
        xlRadarFilled = 82,
        /// <summary>
        /// radar markers
        /// </summary>
        xlRadarMarkers = 81,
        /// <summary>
        /// stock HLC
        /// </summary>
        xlStockHLC = 88,
        /// <summary>
        /// Stock OHLC
        /// </summary>
        xlStockOHLC = 89,
        /// <summary>
        /// Stock VHLC
        /// </summary>
        xlStockVHLC = 90,
        /// <summary>
        /// Stock VOHLC
        /// </summary>
        xlStockVOHLC = 91,
        /// <summary>
        /// Surface
        /// </summary>
        xlSurface = 83,
        /// <summary>
        /// SurfceTopView
        /// </summary>
        xlSurfaceTopView = 85,
        /// <summary>
        /// surface tyope view wireframe
        /// </summary>
        xlSurfaceTopViewWireframe = 86,
        /// <summary>
        /// surface wireframe
        /// </summary>
        xlSurfaceWireframe = 84,
        /// <summary>
        /// XY scatter
        /// </summary>
        xlXYScatter = -4169,
        /// <summary>
        /// XY Scatter Lines
        /// </summary>
        xlXYScatterLines = 74,
        /// <summary>
        /// XYScatter Lines No Markers
        /// </summary>
        xlXYScatterLinesNoMarkers = 75,
        /// <summary>
        /// XY Scatter smooth
        /// </summary>
        xlXYScatterSmooth = 72,
        /// <summary>
        /// XY Scatter smooth no markers
        /// </summary>
        xlXYScatterSmoothNoMarkers = 73
    }
}
