using Newtonsoft.Json.Linq;
using oDataToXls.Models;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace oDataToXls.Utils
{

    public class oDataXlsBuilder
    {
        private DataProperties DataProperties;

        private List<DataPropertiesValue> Dimensions;
        private Dictionary<string, oData> DimensionValues = new Dictionary<string, oData>();


        private DataPropertiesValue TimeDimension;
        private oData TimeDimensionValues;
        
        private List<HeaderCell> TopicGroups;
        private int topicCount;


        public async Task Build(string url, string fileName)
        {

            // get base data
            var baseData = await HttpHelper.GetAsync<oDataBase>(url);

            // get headers
            var responseHeaders = await GetHeaders(baseData.value.First(i => i.name == "DataProperties").url);

            // get DimensionValues
            foreach(var dimension in Dimensions)
            {
                DimensionValues.Add(dimension.Key, await HttpHelper.GetAsync<oData>(baseData.value.First(i => i.name == dimension.Key).url));
            }

            // get TimeDimensionValues
            TimeDimensionValues = await HttpHelper.GetAsync<oData>(baseData.value.First(i => i.name == TimeDimension.Key).url);

            // Create XLS
            await CreateXLS(baseData.value.First(i => i.name == "UntypedDataSet").url, fileName);
            
        }

        /// <summary>
        /// Get
        /// </summary>
        /// <param name="url">Url of the UntypedDataSet</param>
        /// <returns></returns>
        private async Task CreateXLS(string url, string outputFile)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            SetHeaderValues(xlWorkSheet, 0, 0);

            // Filling timeDimensions column
            SetTimeDimensions(xlWorkSheet);

            await SetData(url, xlWorkSheet);


            xlWorkBook.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            
        }

       private int SetHeaderValues(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, int position, int offset)
        {
            int dimensionOffsetMultiply = topicCount;
            int row = position;

            var dimension = Dimensions.FirstOrDefault(item => item.Position == position);
            if (dimension != null)
            {
                var key = dimension.Key;
                int index = 0;
                foreach (var dimensionValue in DimensionValues[key].value)
                {
                    xlWorkSheet.Cells[row + 1, offset + index + 2] = dimensionValue.Title;

                    index += SetHeaderValues(xlWorkSheet, position + 1, offset + index);
                }
                return index;
            }else
            {
                return SetTopicHeaders(xlWorkSheet, TopicGroups, offset, row);
            }
        }

        private int SetTopicHeaders(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, List<HeaderCell> topicGroups, int offset, int row)
        {
            int index = 0;
            foreach (var topic in topicGroups)
            {
                xlWorkSheet.Cells[row + 1, offset + index + 2] = topic.Title;

                if (topic.subCells != null)
                {
                    index += SetTopicHeaders(xlWorkSheet, topic.subCells, offset + index, row + 1);
                }else
                {
                    index += 1;
                }
            }
            return index;
        }

        /// <summary>
        /// Put all value into the xls worksheet
        /// </summary>
        /// <param name="xlWorkSheet"></param>
        /// <param name="values"></param>
        private async Task SetData(string url, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet)
        {
            var jsonString = await HttpHelper.GetAsync(url);
            var obj = JObject.Parse(jsonString);
            var values = obj["value"];
            var baseRowOffset = Dimensions.Count() + 2;

            foreach (var dataRow in values)
            {
                int row = baseRowOffset + TimeDimensionValues.value.IndexOf(TimeDimensionValues.value.First(item => item.Key == dataRow[TimeDimension.Key].ToString()));
                int dimensionOffset = 0;
                int dimensionOffsetMultiply = topicCount;

                foreach (var dimension in Dimensions.OrderByDescending(item => item.Position))
                {
                    var key = dimension.Key;
                    int index = DimensionValues[key].value.IndexOf(DimensionValues[key].value.First(item => item.Key == dataRow[key].ToString()));
                    dimensionOffset += index * dimensionOffsetMultiply;
                    dimensionOffsetMultiply = dimensionOffsetMultiply * DimensionValues[key].value.Count();
                }
                /// TODO: fix dimension offset
                foreach (JProperty dataCell in dataRow)
                {
                    var col = GetColumnNumber(dataCell.Name);
                    if (col != -1)
                    {
                        xlWorkSheet.Cells[row + 1, dimensionOffset + col + 2] = dataCell.Value.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Fill the first column with time dimensions
        /// </summary>
        /// <param name="xlWorkSheet"></param>
        private void SetTimeDimensions(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet)
        {
            int tdRow = Dimensions.Count()+2;
            foreach (var timeDimension in TimeDimensionValues.value)
            {
                xlWorkSheet.Cells[tdRow + 1, 1] = timeDimension.Title;
                tdRow++;
            }
        }
        
        /// <summary>
        /// Get the column number for key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private int GetColumnNumber(string key)
        {
            foreach (var x in TopicGroups)
            {
                var xo = x.getColumn(key);
                if (xo != null)
                    return xo.Value;
            }
            return -1;
        }


        /// <summary>
        /// get headers and store them in properties
        /// </summary>
        /// <param name="url">Url of the DataProperties</param>
        /// <returns></returns>
        private async Task<string> GetHeaders(string url)
        {
            DataProperties = await HttpHelper.GetAsync<DataProperties>(url);

            Dimensions = DataProperties.value.Where(item => item.Type == "Dimension").ToList();
            TimeDimension = DataProperties.value.First(item => item.Type == "TimeDimension");
            GetTopicGroups();

            string response = "";
            int offset = 0;
            foreach (var x in TopicGroups.OrderBy(item => item.Position))
            {
                // temp create response string
                response += x.Title + "<br>";

                x.Offset = offset;

                int subOffset = 0;
                foreach (var y in x.subCells.OrderBy(item => item.Position))
                {
                    y.Offset = offset + subOffset;
                    subOffset++;

                    // temp create response string
                    response += " - " + y.Title + "-" + y.Offset + "<br>";
                }

                offset += x.getLenght();
            }
            topicCount = offset;
            return response;
        }

        private void GetTopicGroups()
        {
            TopicGroups = DataProperties.value.Where(item => item.Type == "TopicGroup")
                            .Select(item => new HeaderCell()
                            {
                                Position = item.Position,
                                Title = item.Title,
                                Key = item.Key,
                                subCells = DataProperties.value
                                    .Where(subItem => subItem.Type == "Topic" && subItem.ParentID == item.ID)
                                    .Select(subItem => new HeaderCell()
                                    {
                                        Position = subItem.Position,
                                        Title = subItem.Title,
                                        Key = subItem.Key
                                    }).ToList()
                            })
                            .ToList();
        }
    }
}