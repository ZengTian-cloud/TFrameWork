using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using Excel;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using System.Text;
using System.Reflection;
using System.Reflection.Emit;
using System;
using System.Text.RegularExpressions;

public class ExcelUtility
{
    /// <summary>
    /// 表格数据集合
    /// </summary>
    private DataSet mResultSet;

    /// <summary>
    /// 读取表数据
    /// </summary>
    /// <param name="excelFile">Excel file.</param>
    public ExcelUtility(string excelFile)
    {
        FileStream mStream = File.Open(excelFile, FileMode.Open, FileAccess.Read);
        IExcelDataReader mExcelReader = ExcelReaderFactory.CreateOpenXmlReader(mStream);
        mResultSet = mExcelReader.AsDataSet();
    }

    /// <summary>
    /// 转换为Json格式文件
    /// </summary>
    /// <param name="JsonPath">Json文件路径</param>
    /// <param name="Header">表头行数</param>
    public void ConvertToJson(string JsonPath, Encoding encoding)
    {
        //判断Excel文件中是否存在数据表
        if (mResultSet.Tables.Count < 1)
            return;

        //默认读取第一个数据表
        DataTable mSheet = mResultSet.Tables[0];

        //判断数据表内是否存在数据
        if (mSheet.Rows.Count < 1)
            return;

        //读取数据表行数和列数
        int rowCount = mSheet.Rows.Count;
        int colCount = mSheet.Columns.Count;

        //准备一个列表存储整个表的数据
        List<Dictionary<string, object>> table = new List<Dictionary<string, object>>();

        //读取数据
        for (int i = ExcelTools.DataInitRow-1; i < rowCount; i++)
        {
            //准备一个字典存储每一行的数据
            Dictionary<string, object> row = new Dictionary<string, object>();
            for (int j = 0; j < colCount; j++)
            {
                //读取第1行数据作为表头字段
                string fieldName = mSheet.Rows[ExcelTools.FiledNameRow-1][j].ToString();
                string filedType = mSheet.Rows[ExcelTools.FiledTypeRow - 1][j].ToString();
                if (string.IsNullOrEmpty(fieldName)) continue;
                //Key-Value对应
                row[fieldName] = ConVertDataTypeFormat(mSheet.Rows[i][j], filedType);
               
               // row[field] = mSheet.Rows[i][j];
               //row[fieldName] =Convert.ChangeType(mSheet.Rows[i][j],typeof(int));
            }

            //添加到表数据中
            table.Add(row);
        }

        //生成Json字符串
        string json = JsonConvert.SerializeObject(table, Newtonsoft.Json.Formatting.Indented);
        string directory = Path.GetDirectoryName(JsonPath);
        if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
        //写入文件
        using (FileStream fileStream = new FileStream(JsonPath, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, encoding))
            {
                textWriter.Write(json);
            }
        }
    }

    /// <summary>
    /// 转换为CSV格式文件
    /// </summary>
    public void ConvertToCSV(string CSVPath, Encoding encoding)
    {
        //判断Excel文件中是否存在数据表
        if (mResultSet.Tables.Count < 1)
            return;

        //默认读取第一个数据表
        DataTable mSheet = mResultSet.Tables[0];

        //判断数据表内是否存在数据
        if (mSheet.Rows.Count < 1)
            return;

        //读取数据表行数和列数
        int rowCount = mSheet.Rows.Count;
        int colCount = mSheet.Columns.Count;

        //创建一个StringBuilder存储数据
        StringBuilder stringBuilder = new StringBuilder();

        //读取数据
        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < colCount; j++)
            {
                //使用","分割每一个数值
                stringBuilder.Append(mSheet.Rows[i][j] + ",");
            }
            //使用换行符分割每一行
            stringBuilder.Append("\r\n");
        }
        string directory = Path.GetDirectoryName(CSVPath);
        if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
        //写入文件
        using (FileStream fileStream = new FileStream(CSVPath, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, encoding))
            {
                textWriter.Write(stringBuilder.ToString());
            }
        }

    }

    /// <summary>
    /// 转换为Xml格式文件
    /// </summary>
    public void ConvertToXml(string XmlFile)
    {
        //判断Excel文件中是否存在数据表
        if (mResultSet.Tables.Count < 1)
            return;

        //默认读取第一个数据表
        DataTable mSheet = mResultSet.Tables[0];

        //判断数据表内是否存在数据
        if (mSheet.Rows.Count < 1)
            return;

        //读取数据表行数和列数
        int rowCount = mSheet.Rows.Count;
        int colCount = mSheet.Columns.Count;

        //创建一个StringBuilder存储数据
        StringBuilder stringBuilder = new StringBuilder();
        //创建Xml文件头
        stringBuilder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        stringBuilder.Append("\r\n");
        //创建根节点
        stringBuilder.Append("<Table>");
        stringBuilder.Append("\r\n");
        //读取数据
        for (int i = 1; i < rowCount; i++)
        {
            //创建子节点
            stringBuilder.Append("  <Row>");
            stringBuilder.Append("\r\n");
            for (int j = 0; j < colCount; j++)
            {
                stringBuilder.Append("   <" + mSheet.Rows[0][j].ToString() + ">");
                stringBuilder.Append(mSheet.Rows[i][j].ToString());
                stringBuilder.Append("</" + mSheet.Rows[0][j].ToString() + ">");
                stringBuilder.Append("\r\n");
            }
            //使用换行符分割每一行
            stringBuilder.Append("  </Row>");
            stringBuilder.Append("\r\n");
        }
        //闭合标签
        stringBuilder.Append("</Table>");

        string directory = Path.GetDirectoryName(XmlFile);
        if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
        //写入文件
        using (FileStream fileStream = new FileStream(XmlFile, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.GetEncoding("utf-8")))
            {
                textWriter.Write(stringBuilder.ToString());
            }
        }
    }

    /// <summary>
    /// 将objct数据转换为指定类型后返回object
    /// </summary>
    /// <param name="value"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    private static object ConVertDataTypeFormat(object value, string type)
    {
        //intstring：自定义数据类型 将数据转换为int后再转换为string  防止出现：1 -> 1.0
        object result=null;
        type = type.ToLower();
        switch (type)
        {
            case "int":
                if (Convert.IsDBNull(value)) return 0;
                result = Convert.ChangeType(value, typeof(int));
                break;
            case "string":
                if (Convert.IsDBNull(value)) return "";
                result = Convert.ChangeType(value, typeof(string));
                break;
            case "float":
                if (Convert.IsDBNull(value)) return 0;
                result = Convert.ChangeType(value, typeof(float));
                break;
            case "double":
                if (Convert.IsDBNull(value)) return 0;
                result = Convert.ChangeType(value, typeof(double));
                break;
            case "intstring":
                if (Convert.IsDBNull(value)) return 0;
                var res = Convert.ChangeType(value, typeof(int));
                result = Convert.ChangeType(res, typeof(string));
                break;
            default:
                result = Convert.ChangeType(value, typeof(string));
                break;
        }
        return result;
    }

}
