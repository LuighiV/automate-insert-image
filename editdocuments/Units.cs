
namespace editdocuments
{

    public enum GUnits
    {
        inch = 0,
        cm,
        pixel,
        point
    }

    public struct UnitStruct
    {
        public UnitStruct(GUnits unit, string literal, double factor)
        {
            Unit = unit;
            Literal = literal;
            Factor = factor;
        }

        public string Literal { get; }
        public double Factor { get; }

        public GUnits Unit;
    }

}