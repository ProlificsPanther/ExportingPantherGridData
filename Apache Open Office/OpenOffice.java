/*============================================================================

Copyright (c) 2022 Prolifics, Inc
 
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

The Software shall be used for Good, not Evil.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
============================================================================*/

package com.prolifics.java;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.uno.XComponentContext;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XCellRangeData;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.XCellRange;
import com.sun.star.uno.UnoRuntime;

import java.io.PrintWriter;
import java.io.StringWriter;
import com.prolifics.jni.Application;
import com.prolifics.jni.ApplicationInterface;
import com.prolifics.jni.CFunctionsInterface;
import com.prolifics.jni.Constants;
import com.prolifics.jni.ScreenHandlerAdapter;

public class OpenOffice {

	ApplicationInterface ai = Application.getInstance();
	CFunctionsInterface cf = ai.getCFunctions();
	
	public OpenOffice() {
	}
	
	public String addChar(String str, char ch, int position) {
	    StringBuilder sb = new StringBuilder(str);
	    sb.insert(position, ch);
	    return sb.toString();
	}
	public void getArray(String gridName) {
		
		String str = "@screen('@current')!@widget('" + gridName + "')";
		try {
			int propid = cf.sm_prop_id(str);
			int handle = cf.sm_list_objects_start(propid);
			int memcnt = cf.sm_list_objects_count(handle);
			int numrows = cf.sm_prop_get_int(propid, Constants.PR_NUM_OCCURRENCES);
			Object[][] arr = new Object[numrows+1][memcnt];
			
			int member_id;
			String mname;
			String columntitle;
			String editMask;
			String memval;
			StringBuffer sbuf;
			int iData;
			
			for (int i = 0; i < memcnt; i++) {
				member_id = cf.sm_list_objects_next(handle);
				mname = cf.sm_prop_get_str(member_id, Constants.PR_NAME);
				columntitle = cf.sm_prop_get_str(member_id, Constants.PR_COLUMN_TITLE);
				editMask = cf.sm_prop_get_str(member_id, Constants.PR_EDIT_MASK);

				int j = 0 ;
				arr[j][i] = columntitle;

				for (j = 1; j <= numrows; j++) {
					memval = cf.sm_i_fptr(mname, j);
					sbuf = new StringBuffer();
					iData = 0;
					char c; 
					if (editMask == null) {
						arr[j][i] = new String(memval);
						continue;
					}

					for (int k = 0; k < editMask.length(); k++)
					{
						try {
							if (editMask.charAt(k) == '\\') {
								k++;
								c = ' ';
								try {
									c = memval.charAt(iData++);
								} catch (Exception e) {
									; // ignore
								}
								sbuf.append(c);
							} else {
								sbuf.append(editMask.charAt(k));
							}
						} catch (Exception e) {
							; // ignore
						}
					}
					arr[j][i] = sbuf.toString();
				}
			}
			XComponentContext xRemoteContext = Bootstrap.bootstrap();

			if (xRemoteContext == null) {
				System.err.println("ERROR: Could not bootstrap default Office.");
			}

			XMultiComponentFactory xRemoteServiceManager = xRemoteContext.getServiceManager();

			Object desktop = xRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", xRemoteContext);
			XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);

			PropertyValue[] loadProps = new PropertyValue[0];
			XComponent xSpreadsheetComponent = xComponentLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, loadProps);

			XSpreadsheetDocument xSpreadsheetDocument = (XSpreadsheetDocument) UnoRuntime
					.queryInterface(XSpreadsheetDocument.class, xSpreadsheetComponent);
			XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets();
			String[] sheet = xSpreadsheets.getElementNames();
			Object defaultSheet = xSpreadsheets.getByName(sheet[0]);
			XSpreadsheet xSpreadsheet = (XSpreadsheet) UnoRuntime.queryInterface(XSpreadsheet.class, defaultSheet);			
			XCellRange xcellRange = xSpreadsheet.getCellRangeByPosition(0, 0, memcnt - 1, numrows);
			XCellRangeData xcellrangedata = (XCellRangeData) UnoRuntime.queryInterface(XCellRangeData.class, xcellRange);
			xcellrangedata.setDataArray(arr);

		} catch (java.lang.Exception e) {
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			cf.sm_message_box(sw.toString(), "Java Exception", Constants.SM_MB_OK, "");		
		}
	}

}

