namespace ZipTabel.Model
{
    public static class CellManager
    {
        public static string _value = "Default Value";

        public static ref string ThisValue()
        {
            return ref _value;
        }

        public static  void SetValue(string newValue)
        {
            _value = newValue;
        }
        public static ref string SelectValue(ref string value)
        {
            return ref value;
        }

    }

}
