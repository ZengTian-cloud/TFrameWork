using UnityEngine;
using UnityEditor;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;

public class ExcelTools : EditorWindow
{
	/// <summary>
	/// 当前编辑器窗口实例
	/// </summary>
	private static ExcelTools instance;

	/// <summary>
	/// Excel文件列表
	/// </summary>
	private static List<string> excelList;

	/// <summary>
	/// 项目根路径	
	/// </summary>
	private static string pathRoot;

	/// <summary>
	/// 滚动窗口初始位置
	/// </summary>
	private static Vector2 scrollPos;

	/// <summary>
	/// 输出格式索引
	/// </summary>
	private static int indexOfFormat = 0;

	/// <summary>
	/// 输出格式
	/// </summary>
	private static string[] formatOption = new string[] { "JSON", "CSV", "XML" };

	/// <summary>
	/// 编码索引
	/// </summary>
	private static int indexOfEncoding = 0;

	/// <summary>
	/// 编码选项
	/// </summary>
	private static string[] encodingOption = new string[] { "UTF-8", "GB2312" };

	/// <summary>
	/// 是否保留原始文件
	/// </summary>
	private static bool keepSource = true;

	/// <summary>
    /// 转换后是否关闭窗口
    /// </summary>
	private static bool completeCloseWindow = true;

	private static string OUT_PATH = "Assets/Res/Config";

	/// <summary>
	/// 数据起始行
	/// </summary>
	public static int DataInitRow = 4;
	/// <summary>
	/// 变量名字段行
	/// </summary>
	public static int FiledNameRow = 2;

	/// <summary>
    /// 变量类型行
    /// </summary>
	public static int FiledTypeRow = 3;

	/// <summary>
	/// 默认的输入框宽度
	/// </summary>
	private int defaultFiledWidth = 150;

	/// <summary>
	/// 显示当前窗口	
	/// </summary>
	[MenuItem("Tools/ExcelTools")]
	static void ShowExcelTools()
	{
		Init();
		//加载Excel文件
		LoadExcel();
		instance.Show();
	}

	void OnGUI()
	{
		DrawOptions();
		DrawExport();
	}

	/// <summary>
	/// 绘制插件界面配置项
	/// </summary>
	private void DrawOptions()
	{
		GUILayout.BeginHorizontal();
		EditorGUILayout.LabelField("请选择格式类型:", GUILayout.Width(85));
		indexOfFormat = EditorGUILayout.Popup(indexOfFormat, formatOption, GUILayout.Width(defaultFiledWidth));
		GUILayout.EndHorizontal();

		GUILayout.BeginHorizontal();
		EditorGUILayout.LabelField("请选择编码类型:", GUILayout.Width(85));
		indexOfEncoding = EditorGUILayout.Popup(indexOfEncoding, encodingOption, GUILayout.Width(defaultFiledWidth));
		GUILayout.EndHorizontal();

		if (indexOfFormat == 0)
		{
			GUILayout.BeginHorizontal();
			EditorGUILayout.LabelField("变量名字行", GUILayout.Width(85));
			FiledNameRow = EditorGUILayout.IntField(FiledNameRow, GUILayout.Width(defaultFiledWidth));
			FiledNameRow = Mathf.Clamp(FiledNameRow, 1, int.MaxValue);
			GUILayout.EndHorizontal();

			GUILayout.BeginHorizontal();
			EditorGUILayout.LabelField("变量类型行", GUILayout.Width(85));
			FiledTypeRow = EditorGUILayout.IntField(FiledTypeRow, GUILayout.Width(defaultFiledWidth));
			FiledTypeRow = Mathf.Clamp(FiledTypeRow, 1, int.MaxValue);
			GUILayout.EndHorizontal();

			GUILayout.BeginHorizontal();
			EditorGUILayout.LabelField("数据起始行", GUILayout.Width(85));
			DataInitRow = EditorGUILayout.IntField(DataInitRow, GUILayout.Width(defaultFiledWidth));
			DataInitRow = Mathf.Clamp(DataInitRow, 1, int.MaxValue);
			GUILayout.EndHorizontal();
		}

		GUILayout.BeginHorizontal();
		EditorGUILayout.LabelField("输出路径", GUILayout.Width(85));
		OUT_PATH = EditorGUILayout.DelayedTextField(OUT_PATH, GUILayout.Width(defaultFiledWidth));
		GUILayout.EndHorizontal();

		keepSource = GUILayout.Toggle(keepSource, "保留Excel源文件",GUILayout.Width(defaultFiledWidth));
		completeCloseWindow = GUILayout.Toggle(completeCloseWindow, "转换完成关闭窗口", GUILayout.Width(defaultFiledWidth));
		GUILayout.Space(20);
	}

	Object source;
	/// <summary>
	/// 绘制插件界面输出项
	/// </summary>
	private void DrawExport()
	{
		if (excelList == null) return;
		if (excelList.Count < 1)
		{
			Color color = GUI.contentColor;
			GUI.contentColor = Color.green;
			EditorGUILayout.LabelField("没有Excel文件被选中!");
			GUI.contentColor = color;
		}
		else if (string.IsNullOrEmpty(OUT_PATH))
		{
			Color color = GUI.contentColor;
			GUI.contentColor = Color.green;
			EditorGUILayout.LabelField("输出路径不能为空!");
			GUI.contentColor = color;
		}
		else
		{
			EditorGUILayout.LabelField("下列Excel将被转换为" + formatOption[indexOfFormat] + ":");
			GUILayout.BeginVertical();
			scrollPos = GUILayout.BeginScrollView(scrollPos, false, true, GUILayout.Height(150));
			foreach (string s in excelList)
			{
				GUILayout.BeginHorizontal();
				Color color2 = GUI.contentColor;
				GUI.contentColor = Color.green;
				GUILayout.Toggle(true, s);
				GUI.contentColor = color2;
				GUILayout.EndHorizontal();
			}
			GUILayout.EndScrollView();
			GUILayout.EndVertical();

			Color color = GUI.backgroundColor;
			GUI.backgroundColor = Color.red;
			//输出
			if (GUILayout.Button("开始转换", GUILayout.Height(30)))
			{
				Convert();
			}
			GUI.backgroundColor = color;
		}
	}

	/// <summary>
	/// 转换Excel文件
	/// </summary>
	private static void Convert()
	{
		foreach (string assetsPath in excelList)
		{
			string name= Path.GetFileName(assetsPath);
			//获取Excel文件的绝对路径
			string excelPath = pathRoot + "/" + assetsPath;
			//构造Excel工具类
			ExcelUtility excel = new ExcelUtility(excelPath);

			//判断编码类型
			Encoding encoding = null;
			if (indexOfEncoding == 0)
			{
				encoding = Encoding.GetEncoding("utf-8");
			}
			else if (indexOfEncoding == 1)
			{
				encoding = Encoding.GetEncoding("gb2312");
			}

			//判断输出类型
			string output = pathRoot + "/" + OUT_PATH + "/" + name;
			if (indexOfFormat == 0)
			{
				output = output.Replace(".xlsx", ".json");
				excel.ConvertToJson(output, encoding);
			}
			else if (indexOfFormat == 1)
			{
				output = output.Replace(".xlsx", ".csv");
				excel.ConvertToCSV(output, encoding);
			}
			else if (indexOfFormat == 2)
			{
				output = output.Replace(".xlsx", ".xml");
				excel.ConvertToXml(output);
			}

			//判断是否保留源文件
			if (!keepSource)
			{
				FileUtil.DeleteFileOrDirectory(excelPath);
			}

			//刷新本地资源
			AssetDatabase.Refresh();
		}

		//转换完后关闭插件
		//这样做是为了解决窗口
		//再次点击时路径错误的Bug
		if(completeCloseWindow)
			instance.Close();
		Debug.Log(string.Format("<color=#26CD3F>{0}</color>", "转换完成"), null);
	}

	/// <summary>
	/// 加载Excel
	/// </summary>
	private static void LoadExcel()
	{
		if (excelList == null) excelList = new List<string>();
		excelList.Clear();
		//获取选中的对象
		object[] selection = (object[])Selection.objects;
		//判断是否有对象被选中
		if (selection.Length == 0)
			return;
		//遍历每一个对象判断不是Excel文件
		foreach (Object obj in selection)
		{
			string objPath = AssetDatabase.GetAssetPath(obj);
			if (objPath.EndsWith(".xlsx"))
			{
				excelList.Add(objPath);
			}
		}
	}

	private static void Init()
	{
		//获取当前实例
		instance = EditorWindow.GetWindow<ExcelTools>(false,"ExcelTools",true);
		//初始化
		pathRoot = Application.dataPath;
		//对路径进行处理
		pathRoot = pathRoot.Substring(0, pathRoot.LastIndexOf("/"));
		excelList = new List<string>();
		scrollPos = new Vector2(instance.position.x, instance.position.y + 75);
	}

	void OnSelectionChange()
	{
		//当选择发生变化时重绘窗体
		Show();
		LoadExcel();
		Repaint();
	}
}
