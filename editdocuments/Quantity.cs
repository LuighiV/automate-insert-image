using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace editdocuments
{

    public class Quantity
    {

        public UnitStruct Unit = Globals.AvailableUnits[0];
        public double Value = 0;

        public Quantity(double value, GUnits unit=GUnits.inch)
        {
            this.Value = value;
            this.Unit = getUnitStruct(unit);
        }

        public void ToUnit(GUnits unit)
        {

            double newValue = ConvertValue(unit);

            // Apply changes
            this.Value = newValue;
            this.Unit = getUnitStruct(unit);
        }

        public double ConvertValue(GUnits unit)
        {
            UnitStruct currentUnit = this.Unit;
            UnitStruct destinationUnit = getUnitStruct(unit);
            double newValue = this.Value * destinationUnit.Factor / currentUnit.Factor;
            return newValue;
        }

        public void Scale(double scale)
        {
            this.Value = this.Value * scale;
        }

        public void ToCentimeters()
        {
            ToUnit(GUnits.cm);
        }

        public void ToInches()
        {
            ToUnit(GUnits.inch);
        }

        public void ToPixels()
        {
            ToUnit(GUnits.pixel);
        }

        public void ToPoints()
        {
            ToUnit(GUnits.point);
        }

        public UnitStruct getUnitStruct(GUnits unit)
        {
            return Globals.AvailableUnits.Find(element => element.Unit == unit);
        }


        public static Quantity FromInches(double inches)
        {
            var quantity = new Quantity(inches);
            return quantity;
        }

        public static Quantity FromCentimetres(double centimeters)
        {
            var quantity = new Quantity(centimeters, GUnits.cm);
            return quantity;
        }

        public static Quantity FromPixels(double pixels)
        {
            var quantity = new Quantity(pixels, GUnits.pixel);
            return quantity;
        }

        public static Quantity FromPoints(double points)
        {
            var quantity = new Quantity(points, GUnits.point);
            return quantity;
        }
    }
}
