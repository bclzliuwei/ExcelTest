// ExcelTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "ExcelTest.h"
#include <cmath>

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\MSO.DLL" \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("RGB", "RBGXL")

#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB"

#import "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" \
	rename("DialogBox", "DialogBoxXL") \
	rename("RGB", "RBGXL") \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("ReplaceText", "ReplaceTextXL") \
	rename("CopyFile", "CopyFileXL") \
	exclude("IFont", "IPicture") no_dual_interfaces

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// The one and only application object

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL)
	{
		// initialize MFC and print and error on failure
		if (!AfxWinInit(hModule, NULL, ::GetCommandLine(), 0))
		{
			// TODO: change error code to suit your needs
			_tprintf(_T("Fatal Error: MFC initialization failed\n"));
			nRetCode = 1;
		}
		else
		{
			// TODO: code your application's behavior here.
			Excel::_ApplicationPtr xl;
			//A try block is used to trap any errors in communication
//			try
			{
				//Initialise COM interface
				CoInitialize(NULL);
				//Start the Excel Application
				xl.CreateInstance(L"Excel.Application");
				//Make the Excel Application visible, so that we can see it!
				xl->Visible = true;
				//Add a (new) workbook
				xl->Workbooks->Add(Excel::xlWorksheet);
				//Get a pointer to the active worksheet
				Excel::_WorksheetPtr pSheet = xl->ActiveSheet;
				//Set the name of the sheet
				pSheet->Name = "Chart Data";
				//Get a pointer to the cells on the active worksheet
				Excel::RangePtr pRange = pSheet->Cells;
				//Define the number of plot points
				unsigned Nplot = 100;
				//Set the lower and upper limits for x
				double x_low = 0.0, x_high = 20.0;
				//Calculate the size of the (uniform) x interval
				//Note a cast to an double here
				double h = (x_high - x_low) / (double)Nplot;
				//Create two columns of data in the worksheet
				//We put labels at the top of each column to say what it contains
				pRange->Item[1][1] = "x"; 
				pRange->Item[1][2] = "f(x)";
				//Now we fill in the rest of the actual data by
				//using a single for loop
				for (unsigned i = 0; i<Nplot; i++)
				{
					//Calculate the value of x (equally-spaced over the range)
					double x = x_low + i*h;
					//The first column is our equally-spaced x values
					pRange->Item[i + 2][1] = x;
					//The second column is f(x)
					pRange->Item[i + 2][2] = sin(x);
				}
				//The sheet "Chart Data" now contains all the data
				//required to generate the chart
				//In order to use the Excel Chart Wizard,
				//we must convert the data into Range Objects
				//Set a pointer to the first cell containing our data
				Excel::RangePtr pBeginRange = pRange->Item[1][1];
				//Set a pointer to the last cell containing our data
				Excel::RangePtr pEndRange = pRange->Item[Nplot + 1][2];
				//Make a "composite" range of the pointers to the start
				//and end of our data
				//Note the casts to pointers to Excel Ranges
				Excel::RangePtr pTotalRange =
					pSheet->Range[(Excel::Range*)pBeginRange][(Excel::Range*)pEndRange];
				// Create the chart as a separate chart item in the workbook
				Excel::_ChartPtr pChart = xl->ActiveWorkbook->Charts->Add();
				//Use the ChartWizard to draw the chart.
				//The arguments to the chart wizard are
				//Source: the data range,
				//Gallery: the chart type,
				//Format: a chart format (number 1-10),
				//PlotBy: whether the data is stored in columns or rows,
				//CategoryLabels: an index for the number of columns
				// containing category (x) labels
				// (because our first column of data represents
				// the x values, we must set this value to 1)
				//SeriesLabels: an index for the number of rows containing
				// series (y) labels
				// (our first row contains y labels,
				// so we set this to 1)
				//HasLegend: boolean set to true to include a legend
				//Title: the title of the chart
				//CategoryTitle: the x-axis title
				//ValueTitle: the y-axis title
				pChart->ChartWizard((Excel::Range*)pTotalRange,
					(long)Excel::xlXYScatter,
					6L, (long)Excel::xlColumns, 1L, 1L, true,
					"My Graph", "x", "f(x)");
				//Give the chart sheet a name
				pChart->Name = "My Data Plot";
				//Finally Uninitialise the COM interface
				CoUninitialize();
			}
			//If a communication error is thrown, catch it and complain
/*
			catch (_com_error &error)
			{
				cout << "COM error " << endl;
			}
*/
			_tprintf(_T("Hello, World!\n"));
			_gettchar();
		}
	}
	else
	{
		// TODO: change error code to suit your needs
		_tprintf(_T("Fatal Error: GetModuleHandle failed\n"));
		nRetCode = 1;
	}

	return nRetCode;
}
