using System;
using System.IO;
using Newtonsoft.Json.Linq;
using ClosedXML.Excel;
using Microsoft.Xrm.Sdk;
using System.Linq;
using System.Data;

namespace PP_UTILITIES
{
    public class Plg_ExportToExcel : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                if (context.InputParameters.Contains("meaf_DataJson") && context.InputParameters["meaf_DataJson"] is string dataJson &&
                    context.InputParameters.Contains("meaf_MappingJson") && context.InputParameters["meaf_MappingJson"] is string mappingJson)
                {
                    if (string.IsNullOrWhiteSpace(dataJson))
                    {
                        throw new InvalidPluginExecutionException("Input 'DataJson' cannot be null or empty.");
                    }

                    if (string.IsNullOrWhiteSpace(mappingJson))
                    {
                        throw new InvalidPluginExecutionException("Input 'MappingJson' cannot be null or empty.");
                    }

                    JArray jsonData;
                    JObject mapping;
                    try
                    {
                        jsonData = JArray.Parse(dataJson);
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace($"Error parsing DataJson: {ex.Message}");
                        throw new InvalidPluginExecutionException("The provided DataJson is not valid JSON.");
                    }

                    try
                    {
                        mapping = JObject.Parse(mappingJson);
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace($"Error parsing MappingJson: {ex.Message}");
                        throw new InvalidPluginExecutionException("The provided MappingJson is not valid JSON.");
                    }
                    string generatedExcelBase64 = GenerateExcel(jsonData, mapping, tracingService);
                    context.OutputParameters["meaf_GeneratedExcelBase64"] = generatedExcelBase64;
                }
                else
                {
                    throw new InvalidPluginExecutionException("Required input parameters 'DataJson' and 'MappingJson' are missing or invalid.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error in GenerateExcel Plugin: {ex.Message}");
                throw new InvalidPluginExecutionException($"Error in GenerateExcel Plugin: {ex.Message}", ex);
            }
        }

        private JArray TransformData(JArray data, JObject mapping, ITracingService tracingService)
        {
            var transformedData = data.Select(item =>
            {
                var transformedItem = new JObject();
                foreach (var map in mapping.Properties())
                {
                    string sourceField = map.Name;             
                    string targetField = map.Value.ToString();
                    var value = item[sourceField];
                    if (value != null && value.Type != JTokenType.Null)
                    {
                        transformedItem[targetField] = value;
                    }
                }
                return transformedItem;
            });
            return new JArray(transformedData);
        }

        private string GenerateExcel(JArray data, JObject mapping, ITracingService tracingService)
        {
            try
            {
                var transformedData = TransformData(data, mapping, tracingService);
                DataTable dataTable = new DataTable();
                foreach (var map in mapping.Properties())
                {
                    string targetField = map.Value.ToString();
                    dataTable.Columns.Add(targetField);
                }
                var rows = transformedData.Select(item =>
                {
                    var row = dataTable.NewRow();
                    foreach (var column in dataTable.Columns.Cast<DataColumn>())
                    {
                        row[column.ColumnName] = item[column.ColumnName]?.ToString() ?? string.Empty;
                    }
                    return row;
                }).ToArray(); 

                foreach (var row in rows)
                {
                    dataTable.Rows.Add(row);
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Exported Data");
                    worksheet.Cell(1, 1).InsertTable(dataTable);
                    using (var outputStream = new MemoryStream())
                    {
                        workbook.SaveAs(outputStream);
                        return Convert.ToBase64String(outputStream.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error generating Excel file: {ex.Message}");
                throw new InvalidPluginExecutionException($"Error generating Excel file: {ex.Message}");
            }
        }

    }
}
