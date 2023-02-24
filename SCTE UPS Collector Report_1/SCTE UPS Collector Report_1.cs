/*
****************************************************************************
*  Copyright (c) 2023,  Skyline Communications NV  All Rights Reserved.    *
****************************************************************************

By using this script, you expressly agree with the usage terms and
conditions set out below.
This script and all related materials are protected by copyrights and
other intellectual property rights that exclusively belong
to Skyline Communications.

A user license granted for this script is strictly for personal use only.
This script may not be used in any way by anyone without the prior
written consent of Skyline Communications. Any sublicensing of this
script is forbidden.

Any modifications to this script by the user are only allowed for
personal use and within the intended purpose of the script,
and will remain the sole responsibility of the user.
Skyline Communications will not be responsible for any damages or
malfunctions whatsoever of the script resulting from a modification
or adaptation by the user.

The content of this script is confidential information.
The user hereby agrees to keep this confidential information strictly
secret and confidential and not to disclose or reveal it, in whole
or in part, directly or indirectly to any person, entity, organization
or administration without the prior written consent of
Skyline Communications.

Any inquiries can be addressed to:

	Skyline Communications NV
	Ambachtenstraat 33
	B-8870 Izegem
	Belgium
	Tel.	: +32 51 31 35 69
	Fax.	: +32 51 31 01 29
	E-mail	: info@skyline.be
	Web		: www.skyline.be
	Contact	: Ben Vandenberghe

****************************************************************************
Revision History:

DATE		VERSION		AUTHOR			COMMENTS

30/09/2019	1.0.0.1		DPR, Skyline	Initial Version
10/05/2020  1.0.0.2		DPR,Skyline		Changed output directory
										Add the parameters: Total Battery voltage, Output Current 1, Output Current 2
										Change column position
05/05/2022  1.0.0.1[transition to GIT]	DPR,Skyline		Changed output directory
										Add the parameters: Total Battery voltage, Output Current 1, Output Current 2
										Change column position
05/01/2023  1.0.0.2		DPR,Skyline		Fixed issue with UPS not showing in the report and shifted columns DCP 200044
****************************************************************************
*/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

using Interop.SLDms;
using Skyline.DataMiner.Automation;

/// <summary>
/// DataMiner Script Class.
/// </summary>
public class Script
{
	public static readonly string Header = "MAC;Name;IP;PS Model;Vendor;Uptime;DOCSISFW;Input Voltage;Output Voltage;Output Power;Output Current;Total Battery Voltage;Output Current 1;" +
		"Output Current 2;Number Batteries;Temperature;Predicted Remaining Time;" +
		"Battery 1;Battery 2;Battery 3;Battery 4;Battery 5;Battery 6;Battery 7;Battery 8;Battery 9;" +
		"Custom Group;Predicted Remaining Time Theoretical;Full Drain Real Backup Time;Absolute Difference;Relative Difference;Element ID Collector;PS ID in the Collector;PS Element ID";

	internal enum LogOutput
	{
		Information,
		Log,
		Both,
	}

	/// <summary>
	/// The Script entry point.
	/// </summary>
	/// <param name="engine">Link with SLScripting process.</param>
	public static void Run(Engine engine)
	{
		Element[] scteElements = engine.FindElementsByProtocol("SCTE UPS Collector", "Production");
		StringBuilder csvData = new StringBuilder(Header + Environment.NewLine);

		int read = 0;
		int errors = 0;
		int inactive = 0;

		foreach (Element element in scteElements)
		{
			try
			{
				if (element.IsActive)
				{
					ProcessElement(engine, element, csvData);
					read++;
				}
				else
				{
					inactive++;
				}
			}
			catch (Exception e)
			{
				errors++;
				engine.Log("Error processing element " + element.Name + "|" + e);
			}
		}

		DateTime dtNow = DateTime.Now;
		string name = String.Format(@"SCTE UPS Collectors Report-{0:yyyy-MM-dd_HH-mm}.csv", dtNow);

		try
		{
			FileInfo fileUps = new FileInfo(String.Format("{0}SCTE UPS Collectors Report-{1:yyyy-MM-dd_HH-mm}.csv", @"C:\Skyline DataMiner\Documents\SCTE UPS Report\", dtNow));

			fileUps.Directory.Create();

			File.WriteAllText(fileUps.FullName, csvData.ToString());

			string message = string.Format(
				"Report:{0} - Processed {1} Collectors. {2} Succeeded, {3} Inactive and {4} Failed",
				name,
				scteElements.Length,
				read,
				inactive,
				errors);

			GenerateMessage(engine, message, LogOutput.Both, String.Empty);
		}
		catch (Exception e)
		{
			engine.Log("Error creating Report: " + name + "|" + e);
		}
	}

	private static void ProcessElement(Engine engine, Element element, StringBuilder csvData)
	{
		Dictionary<string, Dictionary<string, double>> dicBatteries = BuildBatteriesDictionary(engine, element);
		Dictionary<string, Dictionary<string, double>> dicOutput = BuildOutputDictionary(engine, element);

		int[] columnsDeviceIdx = new[] { 0, 2, 12, 15, 16, 28, 29, 30, 33, 34, 35, 36, 45, 47, 51, 59, 123, 124, 125, 130, 131, 22 };

		object[] columnDeviceResult = GetColumns(engine, element, 1000, columnsDeviceIdx);

		if (columnDeviceResult == null || columnDeviceResult.Length != 22)
		{
			return;
		}

		object[] columnIndex = (object[])columnDeviceResult[0];
		object[] columnMac = (object[])columnDeviceResult[1];
		object[] columnPsModule = (object[])columnDeviceResult[2];
		object[] columnName = (object[])columnDeviceResult[3];
		object[] columnIP = (object[])columnDeviceResult[4];
		object[] columnVendor = (object[])columnDeviceResult[5];
		object[] columnDocsSisFw = (object[])columnDeviceResult[6];
		object[] columnUptime = (object[])columnDeviceResult[7];
		object[] columnNumBatteries = (object[])columnDeviceResult[8];
		object[] columnInputVoltage = (object[])columnDeviceResult[9];
		object[] columnOutputVoltage = (object[])columnDeviceResult[10];
		object[] columnTotalOutputVoltage = (object[])columnDeviceResult[11];
		object[] columnOutputCurrent = (object[])columnDeviceResult[12];
		object[] columnOutputPower = (object[])columnDeviceResult[13];
		object[] columnTemperature = (object[])columnDeviceResult[14];
		object[] columnPredRemainingTime = (object[])columnDeviceResult[15];
		object[] columnFullDrRealBckUpTime = (object[])columnDeviceResult[16];
		object[] columnAbsDifference = (object[])columnDeviceResult[17];
		object[] columnRelDifference = (object[])columnDeviceResult[18];
		object[] columnPredRemainingTimeTh = (object[])columnDeviceResult[19];
		object[] columnCustomGroup = (object[])columnDeviceResult[20];
		object[] columnElementId = (object[])columnDeviceResult[21];

		for (int i = 0; i < columnIndex.Length; i++)
		{
			csvData.Append(String.Format(
				"{0};{1};{2};{3};{4};",
				Convert.ToString(columnMac[i]),
				Convert.ToString(columnName[i]).Replace(",", " ").Replace(";", " "),
				Convert.ToString(columnIP[i]),
				Convert.ToString(columnPsModule[i]).Replace(",", " ").Replace(";", " "),
				Convert.ToString(columnVendor[i]).Replace(",", " ").Replace(";", " ")));

			double uptime = Convert.ToDouble(columnUptime[i]);
			string time = ParseFromTimeToString(columnUptime[i]);

			csvData.Append(String.Format("{0}-{1};", uptime, time));

			csvData.Append(String.Format(
				"{0};{1};{2};{3};{4};{5};",
				Convert.ToString(columnDocsSisFw[i]).Replace(",", " ").Replace(";", " "),
				Convert.ToString(columnInputVoltage[i]),
				Convert.ToString(columnOutputVoltage[i]),
				Convert.ToString(columnOutputPower[i]),
				Convert.ToString(columnOutputCurrent[i]),
				Convert.ToString(columnTotalOutputVoltage[i])));

			FillOutputsData(csvData, dicOutput, columnIndex, i);

			csvData.Append(String.Format(
				"{0};{1};",
				Convert.ToString(columnNumBatteries[i]),
				Convert.ToString(columnTemperature[i]).Equals("-1000") ? "N/A" : Convert.ToString(columnTemperature[i])));

			string timepred = ParseFromTimeToString(columnPredRemainingTime[i]);

			csvData.Append(String.Format("{0};", timepred));
			FillBatteriesData(csvData, dicBatteries, columnIndex, i);
			string sCustomGroup = SetCustomGroup(columnCustomGroup, i);

			string sPredRemTimeTh = ParseFromTimeToString(columnPredRemainingTimeTh[i]);
			string sFullDrRealBkUp = ParseFromTimeToString(columnFullDrRealBckUpTime[i]);
			string sAbsDifference = ParseFromTimeToString(columnAbsDifference[i]);
			string sRelDifference = String.IsNullOrEmpty(Convert.ToString(columnRelDifference[i])) ? "N/A" : Convert.ToString(columnRelDifference[i]);
			string sPsIdCollector = Convert.ToString(columnIndex[i]);

			csvData.Append(String.Format(
				"{0};{1};{2};{3};{4};{5};{6};{7}",
				sCustomGroup,
				sPredRemTimeTh,
				sFullDrRealBkUp,
				sAbsDifference,
				sRelDifference,
				Convert.ToString(String.Format("{0}/{1}", element.DmaId, element.ElementId)),
				sPsIdCollector,
				columnElementId[i]));

			csvData.Append(Environment.NewLine);
		}
	}

	private static string SetCustomGroup(object[] columnCustomGroup, int i)
	{
		string cg = Convert.ToString(columnCustomGroup[i]);
		string sCustomGroup;
		if (String.IsNullOrEmpty(cg))
		{
			sCustomGroup = "N/A";
		}
		else if (cg.Equals("-1"))
		{
			sCustomGroup = "Not Configured";
		}
		else
		{
			sCustomGroup = cg;
		}

		return sCustomGroup;
	}

	private static void FillBatteriesData(StringBuilder csvData, Dictionary<string, Dictionary<string, double>> dicBatteries, object[] columnIndex, int i)
	{
		if (dicBatteries.ContainsKey(Convert.ToString(columnIndex[i])))
		{
			Dictionary<string, double> bat = dicBatteries[Convert.ToString(columnIndex[i])];

			AddBattery(csvData, bat, "1");
			AddBattery(csvData, bat, "2");
			AddBattery(csvData, bat, "3");
			AddBattery(csvData, bat, "4");
			AddBattery(csvData, bat, "5");
			AddBattery(csvData, bat, "6");
			AddBattery(csvData, bat, "7");
			AddBattery(csvData, bat, "8");
			AddBattery(csvData, bat, "9");
		}
		else
		{
			csvData.Append(";;;;;;;;;");
		}
	}

	private static void AddBattery(StringBuilder csvData, Dictionary<string, double> bat, string key)
	{
		csvData.Append(bat.ContainsKey(key) ? String.Format("{0};", bat[key]) : ";");
	}

	private static void FillOutputsData(StringBuilder csvData, Dictionary<string, Dictionary<string, double>> dicOutput, object[] columnIndex, int i)
	{
		if (dicOutput.ContainsKey(Convert.ToString(columnIndex[i])))
		{
			if (dicOutput[Convert.ToString(columnIndex[i])].ContainsKey("1"))
			{
				Dictionary<string, double> output = dicOutput[Convert.ToString(columnIndex[i])];

				csvData.Append(String.Format("{0};", Convert.ToString(output["1"])));
			}
			else
			{
				csvData.Append(String.Format("{0};", String.Empty));
			}

			if (dicOutput[Convert.ToString(columnIndex[i])].ContainsKey("2"))
			{
				Dictionary<string, double> output = dicOutput[Convert.ToString(columnIndex[i])];

				csvData.Append(String.Format("{0};", Convert.ToString(output["2"])));
			}
			else
			{
				csvData.Append(String.Format("{0};", String.Empty));
			}
		}
		else
		{
			csvData.Append(";;");
		}
	}

	private static Dictionary<string, Dictionary<string, double>> BuildOutputDictionary(Engine engine, Element element)
	{
		int[] columnsOutputs = new[] { 1, 2, 3 };
		Dictionary<string, Dictionary<string, double>> dicOutput = new Dictionary<string, Dictionary<string, double>>();
		object[] columnOutputResult = GetColumns(engine, element, 5000, columnsOutputs);
		if (columnOutputResult == null || columnOutputResult.Length != 3)
		{
			return dicOutput;
		}

		object[] devices = (object[])columnOutputResult[0];
		object[] columnOutputIndex = (object[])columnOutputResult[1];
		object[] columnOutputVoltage = (object[])columnOutputResult[2];

		for (int i = 0; i < columnOutputIndex.Length; i++)
		{
			if (!dicOutput.ContainsKey(Convert.ToString(devices[i])))
			{
				Dictionary<string, double> dicTemp = new Dictionary<string, double>
				{
					{ Convert.ToString(columnOutputIndex[i]), Convert.ToDouble(columnOutputVoltage[i]) },
				};

				dicOutput.Add(Convert.ToString(devices[i]), dicTemp);
			}
			else
			{
				if (!dicOutput[Convert.ToString(devices[i])].ContainsKey(Convert.ToString(columnOutputIndex[i])))
				{
					dicOutput[Convert.ToString(devices[i])].Add(Convert.ToString(columnOutputIndex[i]), Convert.ToDouble(columnOutputVoltage[i]));
				}
				else
				{
					engine.Log("ERROR:" + element.ElementName + "Duplicate outputs" + Convert.ToString(columnOutputIndex[i]));
				}
			}
		}

		return dicOutput;
	}

	private static Dictionary<string, Dictionary<string, double>> BuildBatteriesDictionary(Engine engine, Element element)
	{
		int[] columnsBatteriesIxd = new[] { 1, 3, 4 };
		Dictionary<string, Dictionary<string, double>> dicBatteries = new Dictionary<string, Dictionary<string, double>>();

		object[] columnBatteriesResult = GetColumns(engine, element, 3000, columnsBatteriesIxd);
		if (columnBatteriesResult == null || columnBatteriesResult.Length != 3)
		{
			return dicBatteries;
		}

		object[] columnDevice = (object[])columnBatteriesResult[0];
		object[] columnIndex = (object[])columnBatteriesResult[1];
		object[] columnVoltage = (object[])columnBatteriesResult[2];

		for (int i = 0; i < columnIndex.Length; i++)
		{
			if (!dicBatteries.ContainsKey(Convert.ToString(columnDevice[i])))
			{
				Dictionary<string, double> dicTemp = new Dictionary<string, double>
				{
					{ Convert.ToString(columnIndex[i]), Convert.ToDouble(columnVoltage[i]) },
				};

				dicBatteries.Add(Convert.ToString(columnDevice[i]), dicTemp);
			}
			else
			{
				if (!dicBatteries[Convert.ToString(columnDevice[i])].ContainsKey(Convert.ToString(columnIndex[i])))
				{
					dicBatteries[Convert.ToString(columnDevice[i])].Add(Convert.ToString(columnIndex[i]), Convert.ToDouble(columnVoltage[i]));
				}
				else
				{
					engine.Log("ERROR:" + element.ElementName + "Duplicate Battery" + Convert.ToString(columnIndex[i]));
				}
			}
		}

		return dicBatteries;
	}

	private static string ParseFromTimeToString(object columnToParse)
	{
		if (int.TryParse(Convert.ToString(columnToParse), out int dTime))
		{
			TimeSpan tsTime = TimeSpan.FromSeconds(dTime);
			string sTime = String.Format(
				"{0}Days {1:D2}h:{2:D2}m:{3:D2}s:{4:D3}ms",
				tsTime.Days,
				tsTime.Hours,
				tsTime.Minutes,
				tsTime.Seconds,
				tsTime.Milliseconds);

			return sTime;
		}

		return "N/A";
	}

	private static void GenerateMessage(Engine engine, string message, LogOutput logOutput, string exception)
	{
		if (logOutput == LogOutput.Both || logOutput == LogOutput.Log)
		{
			string fullMessage = String.IsNullOrEmpty(exception) ? message : message + "|" + exception;
			engine.Log(fullMessage);
		}

		if (logOutput == LogOutput.Both || logOutput == LogOutput.Information)
		{
			engine.GenerateInformation(message);
		}
	}

	private static object[] GetColumns(Engine engine, Element element, int tablePid, int[] columnIdxs)
	{
		object[] resultColumns = null;
		try
		{
			var dms = new DMS();
			var ids = new[] { (uint)element.DmaId, (uint)element.ElementId };

			object returnValue;
			dms.Notify(87, 0, ids, tablePid, out returnValue);

			var table = (object[])returnValue;
			if (table == null || table.Length <= 4)
				return resultColumns;

			var columns = (object[])table[4];
			if (columns == null || columns.Length <= columnIdxs.Max())
				return resultColumns;

			var rowCount = ((object[])columns[0]).Length;
			resultColumns = new object[columnIdxs.Length];
			for (int i = 0; i < columnIdxs.Length; i++)
			{
				resultColumns[i] = new object[rowCount];
			}

			AssignColumnsValues(columnIdxs, resultColumns, columns, rowCount);

			return resultColumns;
		}
		catch (Exception e)
		{
			engine.Log("GetColumns|Error getting columns from table: " + tablePid + ", from element ID: " + element.Id + Environment.NewLine + e);
			return resultColumns;
		}
	}

	private static void AssignColumnsValues(int[] columnIdxs, object[] resultColumns, object[] columns, int rowCount)
	{
		for (int i = 0; i < columnIdxs.Length; i++)
		{
			var column = (object[])columns[columnIdxs[i]];
			for (int j = 0; j < rowCount; j++)
			{
				var cell = (object[])column[j];
				if (cell != null && cell.Length > 0)
				{
					((object[])resultColumns[i])[j] = cell[0];
				}
			}
		}
	}
}