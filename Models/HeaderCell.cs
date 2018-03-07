using System.Collections.Generic;

namespace oDataToXls.Models
{
    public class HeaderCell
    {
        public int? Position { get; set; }
        public string Title { get; set; }
        public string Key { get; set; }
        public int? Offset { get; set; }

        public List<HeaderCell> subCells { get; set; }

        public int getLenght()
        {
            int response = 0;
            if(subCells != null)
            {
                foreach(var x in subCells)
                {
                    response += x.getLenght();
                }
            }
            if (response == 0)
                return 1;
            return response;
        }

        public int? getColumn(string key)
        {
            if (this.Key == key)
                return Offset;

            if(subCells != null)
                foreach (var x in subCells)
                {
                    var xo = x.getColumn(key);
                    if (xo != null)
                        return xo;
                }
            return null;
        }
    }
}