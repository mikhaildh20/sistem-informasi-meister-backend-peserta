namespace sistem_informasi_produksi_backend.Helper
{
    public class Utilities
    {
        public static string Separator(string input)
        {
            if (input == null) return "";
            try
            {
                string stringNumber = input;
                string result = "";
                int count = 0;
                for (int i = stringNumber.Length - 1; i >= 0; i--)
                {
                    if (count != 0 && count % 3 == 0)
                    {
                        result = '.' + result;
                    }
                    result = stringNumber[i] + result;
                    count++;
                }
                return result;
            }
            catch { return ""; }
        }
    }
}
