
namespace editdocuments
{

    public enum GUnits
    {
        inch = 0,
        cm,
        pixel,
        point
    }

    /// <summary>
    /// Structure to save info about the unit
    /// </summary>
    public struct UnitStruct
    {
        /// <summary>
        /// Create the unit structure
        /// </summary>
        /// <param name="unit"> Enter the Unit from Gunit enum</param>
        /// <param name="literal"> Enter the literal representation</param>
        /// <param name="factor"> Enter the conversion factor relative to inches</param>
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