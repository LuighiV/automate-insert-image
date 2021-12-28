using System.Collections.Generic;

namespace editdocuments{

    public static class Globals
    {
        // Reference for factors https://stackoverflow.com/a/139712
        static double INCH_FACTOR = 1;
        static double PIXEL_FACTOR = 96;
        static double POINT_FACTOR = 72;
        static double CENTIMETER_FACTOR = 2.54;

        public static List<UnitStruct> AvailableUnits = new List<UnitStruct>();


        static Globals()
        {
            AvailableUnits.Add(new UnitStruct(GUnits.inch, "inch", INCH_FACTOR));
            AvailableUnits.Add(new UnitStruct(GUnits.cm, "cm", CENTIMETER_FACTOR));
            AvailableUnits.Add(new UnitStruct(GUnits.pixel, "pixel", PIXEL_FACTOR));
            AvailableUnits.Add(new UnitStruct(GUnits.point, "point", POINT_FACTOR));
        }
    }
}