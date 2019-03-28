import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.ws.rs.Consumes;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

@Path("/")
public class TrackerRESTService {

	@POST
	@Path("/generateTracker")
	@Consumes(MediaType.APPLICATION_JSON)
	@Produces("application/vnd.ms-excel.sheet.macroEnabled.12")
	public Response generateTracker(InputStream incomingData) {
		Response result = null;
		String fileName = "HCSC_ACE_Status.xlsm";
		System.out.println(incomingData);
		List<DataWrapper> data = convertJSONToWrapper(incomingData);	
		System.out.println(1);
		try {
			if(data.size() > 0) {
				XSSFWorkbook workbook = generateTracker(data);
				System.out.println(1);

				FileOutputStream out = new FileOutputStream(new File(fileName));
				workbook.write(out);
				System.out.println(1);
				out.close();
				System.out.println("xlsm created successfully..");
				File fileDownload = new File(fileName);
		        ResponseBuilder response = Response.ok((Object) fileDownload);
		        response.header("Content-Disposition", "attachment; filename=" + fileName);
		        response.header("Content-Type", "application/vnd.ms-excel.sheet.macroEnabled.12");
		        System.out.println(1);
		        result = response.build();
		        System.out.println(1);
			}
			

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        return result;
	}
	
	private static List<DataWrapper> convertJSONToWrapper(InputStream incomingData) {
		StringBuilder crunchifyBuilder = new StringBuilder();
		DataWrapper record;
		List<DataWrapper> records = new ArrayList<DataWrapper>();
		try {
			BufferedReader in = new BufferedReader(new InputStreamReader(incomingData));
			String line = null;
			while ((line = in.readLine()) != null) {
				crunchifyBuilder.append(line);
			}
		} catch (Exception e) {
			System.out.println("Error Parsing: - ");
		}
		String jsonData = crunchifyBuilder.toString();
		if(jsonData != null) {
			JSONArray arr = new JSONArray(jsonData);
			
			//Now we create the list of records.
			if(arr != null && arr.length() > 0) {
				for (int i = 0; i < arr.length(); i++)
				{
					record = new DataWrapper();
					record.A = getString(arr.getJSONObject(i), "Name");
					record.B = getString(arr.getJSONObject(i), "Title__c");
					record.C = replaceWithDots(getString(arr.getJSONObject(i), "GroomingSessions_ACE__c"));
					record.D = replaceWithDots(getString(arr.getJSONObject(i), "Wireframe_ACE__c"));
					record.E = replaceWithDots(getString(arr.getJSONObject(i), "VD_ACE__c"));
					record.F = replaceWithDots(getString(arr.getJSONObject(i), "BusinessApproval_ACE__c"));
					record.G = replaceWithDots(getString(arr.getJSONObject(i), "PreReview_ACE__c"));
					record.H = replaceWithDots(getString(arr.getJSONObject(i), "Analysis_ACE__c"));
					record.I = replaceWithDots(getString(arr.getJSONObject(i), "InitialTechnicalDraft_ACE__c"));
					record.J = replaceWithDots(getString(arr.getJSONObject(i), "InitialEstimation_ACE__c"));
					record.K = replaceWithDots(getString(arr.getJSONObject(i), "ReadyForDevTeamReview_ACE__c"));
					record.L = replaceWithDots(getString(arr.getJSONObject(i), "TechnicalDesignDocument_ACE__c"));
					record.M = replaceWithDots(getString(arr.getJSONObject(i), "FinalEstimation_ACE__c"));
					record.N = replaceWithDots(getString(arr.getJSONObject(i), "ReadyForDevelopment_ACE__c"));
					record.O = replaceWithDots(getString(arr.getJSONObject(i), "Development_ACE__c"));
					record.P = replaceWithDots(getString(arr.getJSONObject(i), "TestData_ACE__c"));
					record.Q = replaceWithDots(getString(arr.getJSONObject(i), "UnitTesting_ACE__c"));
					record.R = replaceWithDots(getString(arr.getJSONObject(i), "PeerReview_ACE__c"));
					record.S = replaceWithDots(getString(arr.getJSONObject(i), "MockDeployment_ACE__c"));
					record.T = replaceWithDots(getString(arr.getJSONObject(i), "SonarQube_ACE__c"));
					record.U = replaceWithDots(getString(arr.getJSONObject(i), "PullRequest_ACE__c"));
					record.V = replaceWithDots(getString(arr.getJSONObject(i), "UpdateTDDDataDictionaryAndJira_ACE__c"));
					record.W = replaceWithDots(getString(arr.getJSONObject(i), "ACECI_ACE__c"));
					record.X = replaceWithDots(getString(arr.getJSONObject(i), "ACESIT_ACE__c"));
					record.Y = replaceWithDots(getString(arr.getJSONObject(i), "ACEUAT_ACE__c"));
					record.Z = replaceWithDots(getString(arr.getJSONObject(i), "ACEPerform_ACE__c"));
					record.AA = replaceWithDots(getString(arr.getJSONObject(i), "ACEMock_ACE__c"));
					record.AB = replaceWithDots(getString(arr.getJSONObject(i), "Production_ACE__c"));
					record.AC = getString(arr.getJSONObject(i), "Status_FAT__c");
					record.AD = getString(arr.getJSONObject(i), "EstimatedStoryPoints_ACE__c");
					record.AE = getString(arr.getJSONObject(i), "ActualStoryPoints_ACE__c");
					if(doesParentExist(arr.getJSONObject(i), "RecordType")) {
						record.AF = getString(arr.getJSONObject(i).getJSONObject("RecordType"), "Name");
					}
					if(doesParentExist(arr.getJSONObject(i), "Feature_FAT__r")) {
						record.AG = getString(arr.getJSONObject(i).getJSONObject("Feature_FAT__r"), "Title_ACE__c");
						record.AH = getString(arr.getJSONObject(i).getJSONObject("Feature_FAT__r"), "Name");
					}
					record.AJ = getString(arr.getJSONObject(i), "Owner_ACE__c");
					if(doesParentExist(arr.getJSONObject(i), "Sprint_FAT__r")) {
						record.AL = getString(arr.getJSONObject(i).getJSONObject("Sprint_FAT__r"), "Name");
					}
					if(doesParentExist(arr.getJSONObject(i), "Feeds") && doesParentExist(arr.getJSONObject(i).getJSONObject("Feeds") ,"records") &&
							!arr.getJSONObject(i).getJSONObject("Feeds").getJSONArray("records").isEmpty()) {
						record.AK = getString(arr.getJSONObject(i).getJSONObject("Feeds").getJSONArray("records").getJSONObject(0), "Body");
					}
					record.AI = getString(arr.getJSONObject(i), "DevTeam_ACE__c");
					record.AM = getString(arr.getJSONObject(i), "TargetCIDeployment_ACE__c");
					record.AN = replaceWithYesNo(getBoolean(arr.getJSONObject(i), "IntegrationDependency_ACE__c"));
					record.AO = getString(arr.getJSONObject(i), "Type_FAT__c");
					record.AP = replaceWithYesNo(getBoolean(arr.getJSONObject(i), "MuleSoftStory_ACE__c"));
					records.add(record);
				}
			}
		}		
		return records;
	}
	
	private static Boolean doesParentExist(JSONObject parent, String key) {
		Boolean value = false;
		if(parent != null && key != null && parent.has(key)) {
			value = true;
		}
		return value;
	}
	
	private static Boolean getBoolean(JSONObject parent, String key) {
		Boolean value = null;
		if(parent != null && key != null && parent.has(key)) {
			value = parent.getBoolean(key);
		}
		return value;
	}
	
	private static String getString(JSONObject parent, String key) {
		String value = null;
		if(parent != null && key != null && parent.has(key)) {
			value = parent.getString(key);
		}
		return value;
	}
	
	private static String replaceWithYesNo(Boolean value) {
		String result = null;
		if(value != null) {
			if(value) {
				result = "Yes";
			} else {
				result = "No";
			}
		}
		return result;
	}
	
	private static String replaceWithDots(String value) {
		if(value != null) {
			if(value.equals("Not Started")) {
				value = "0";
			} else if(value.equals("In Progress")) {
				value = "1";
			} else if(value.equals("Completed")) {
				value = "2";
			}
		}
		return value;
	}

	private static XSSFWorkbook generateTracker(List<DataWrapper> data) throws IllegalArgumentException, IllegalAccessException, ParseException {
		XSSFWorkbook workbook = null;
		Integer size = data.size() - 1;
		try {
			InputStream is = TrackerRESTService.class.getResourceAsStream("../Template.xlsm");
			workbook = new XSSFWorkbook(is);
			//workbook = new XSSFWorkbook(OPCPackage.open("resources/Template.xlsm"));
			XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Status");
			for (XSSFTable table : sheet.getTables()) {
				for (int i = 0; i < size; i++) {
					addRowToTable(workbook, table);			
					
					// Now we copy all the formatting from the previous row.
					copyRow(workbook, sheet, 2 + i, 3 + i);
				}				
			}
			
			// Now we extend the Conditional formatting.
			Map<Integer, CellRangeAddress[]> formattingsRange = new HashMap<Integer, CellRangeAddress[]>();
			Map<Integer, ConditionalFormattingRule> formattingsRule = new HashMap<Integer, ConditionalFormattingRule>();
			SheetConditionalFormatting sheetFormatting = sheet.getSheetConditionalFormatting();
			int numberOfConditions = sheetFormatting.getNumConditionalFormattings();
			CellRangeAddress[] regions = new CellRangeAddress[1];
			
			//First we collect all the modified rules.
			for (int i = 0; i < numberOfConditions; i++) {
				for (CellRangeAddress region : sheetFormatting.getConditionalFormattingAt(i).getFormattingRanges()) {
					regions = new CellRangeAddress[1];
					if (region.getLastRow() < 1048570) {
						region.setLastRow(region.getLastRow() + size);
					}
					regions[0] = region;
					for (int y = 0; y < sheetFormatting.getConditionalFormattingAt(i).getNumberOfRules(); y++) {
						formattingsRange.put(i, regions);
						formattingsRule.put(i, sheetFormatting.getConditionalFormattingAt(i).getRule(y));
					}
				}
			}
			
			//Now, we add the right ones.
			for(Integer rule : formattingsRange.keySet()) {
				sheetFormatting.addConditionalFormatting(formattingsRange.get(rule), formattingsRule.get(rule));
			}
			
			//Now we delete the old values.
			for (int i = 0; i < numberOfConditions; i++) {
				sheetFormatting.removeConditionalFormatting(i);
			}		

			// Finally, we clone the data validation values.
			CellRangeAddressList regions2;
			DataValidationHelper validationHelper = sheet.getDataValidationHelper();
			for (DataValidation dataValidation : sheet.getDataValidations()) {
				for (CellRangeAddress region : dataValidation.getRegions().getCellRangeAddresses()) {
					if (region.getLastRow() < 1048570) {
						region.setLastRow(region.getLastRow() + size);
						regions2 = new CellRangeAddressList();
						regions2.addCellRangeAddress(CellRangeAddress.valueOf(region.formatAsString()));
						DataValidation validation = validationHelper.createValidation(dataValidation.getValidationConstraint() , regions2);
						sheet.addValidationData(validation);
					}
				}
			}
			
			//Now we insert the data.
			CellReference cr;
			XSSFCell cell;
			Object value;
			List<Integer> text = new ArrayList<Integer>();
			text.add(0);
			text.add(1);
			text.add(28);
			text.add(31);
			text.add(32);
			text.add(33);
			text.add(34);
			text.add(35);
			text.add(36);
			text.add(37);
			text.add(39);
			text.add(40);
			text.add(41);			
			List<Integer> numeric = new ArrayList<Integer>();
			numeric.add(2);
			numeric.add(3);
			numeric.add(4);
			numeric.add(5);
			numeric.add(6);
			numeric.add(7);
			numeric.add(8);
			numeric.add(9);
			numeric.add(10);
			numeric.add(11);
			numeric.add(12);
			numeric.add(13);
			numeric.add(14);
			numeric.add(15);
			numeric.add(16);
			numeric.add(17);
			numeric.add(18);
			numeric.add(19);
			numeric.add(20);
			numeric.add(21);
			numeric.add(22);
			numeric.add(23);
			numeric.add(24);
			numeric.add(25);
			numeric.add(26);
			numeric.add(27);
			numeric.add(29);
			numeric.add(30);
			List<Integer> dateField = new ArrayList<Integer>();
			dateField.add(38);
			for (int i = 0; i < data.size(); i++) {
				for(Field field : data.get(i).getClass().getDeclaredFields()) {
					cr = new CellReference(field.getName() + "" + (i + 3));
					value = field.get(data.get(i));
					cell = sheet.getRow(cr.getRow()).getCell(cr.getCol());
					if(cell == null) {
						cell = sheet.getRow(cr.getRow()).createCell(cr.getCol());
					}
					if(field.get(data.get(i)) != null) {
						Integer column = (int) cr.getCol();
						if(text.contains(column)) {
							cell.setCellValue((String) value);
						} else if(numeric.contains(column)) {
							cell.setCellValue(Double.valueOf((String) value));
						} else if(dateField.contains(column)) {
							cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").parse((String) value));
						}
					} else {
						sheet.getRow(cr.getRow()).removeCell(cell);
					}
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
		return workbook;
	}

	private static void addRowToTable(XSSFWorkbook workbook, XSSFTable table) {

		int lastTableRow = table.getEndCellReference().getRow();
		int totalsRowCount = table.getTotalsRowCount();
		int lastTableDataRow = lastTableRow - totalsRowCount;

		// we will add one row in table data
		lastTableRow++;
		lastTableDataRow++;

		// new table area plus one row
		AreaReference newTableArea = new AreaReference(table.getStartCellReference(),
				new CellReference(lastTableRow, table.getEndCellReference().getCol()), SpreadsheetVersion.EXCEL2007);

		// new table data area plus one row
		AreaReference newTableDataArea = new AreaReference(table.getStartCellReference(),
				new CellReference(lastTableDataRow, table.getEndCellReference().getCol()),
				SpreadsheetVersion.EXCEL2007);

		XSSFSheet sheet = table.getXSSFSheet();
		if (totalsRowCount > 0) {
			// if we have totals rows, shift totals rows down
			sheet.shiftRows(lastTableDataRow, lastTableRow, 1);

			// correcting bug that shiftRows does not adjusting references of the cells
			// if row 3 is shifted down, then reference in the cells remain r="A3", r="B3",
			// ...
			// they must be adjusted to the new row thoug: r="A4", r="B4", ...
			// apache poi 3.17 has done this properly but had have other bugs in shiftRows.
			for (int r = lastTableDataRow; r < lastTableRow + 1; r++) {
				XSSFRow row = sheet.getRow(r);
				if (row != null) {
					long rRef = row.getCTRow().getR();
					for (Cell cell : row) {
						String cRef = ((XSSFCell) cell).getCTCell().getR();
						((XSSFCell) cell).getCTCell().setR(cRef.replaceAll("[0-9]", "") + rRef);
					}
				}
			}
			// end correcting bug

		}

		// if there are CalculatedColumnFormulas do filling them to the new row
		XSSFRow row = sheet.getRow(lastTableDataRow);
		if (row == null)
			row = sheet.createRow(lastTableDataRow);
		for (CTTableColumn tableCol : table.getCTTable().getTableColumns().getTableColumnList()) {
			if (tableCol.getCalculatedColumnFormula() != null) {
				int id = (int) tableCol.getId();
				if (id == 12) {
					id = 44;
				}
				String formula = tableCol.getCalculatedColumnFormula().getStringValue();
				XSSFCell cell = row.getCell(id + 3);
				if (cell == null)
					cell = row.createCell(id + 3);
				cell.setCellFormula(formula);
			}
		}

		table.setArea(newTableArea);

		// correcting bug that Autofilter includes possible TotalsRows after setArea new
		// Autofilter must only contain data area
		table.getCTTable().getAutoFilter().setRef(newTableDataArea.formatAsString());
		// end correcting bug

		table.updateReferences();
	}

	private static void copyRow(XSSFWorkbook workbook, XSSFSheet worksheet, int sourceRowNum, int destinationRowNum) {
		// Get the source / new row
		XSSFRow newRow = worksheet.getRow(destinationRowNum);
		XSSFRow sourceRow = worksheet.getRow(sourceRowNum);

		// If the row exist in destination, push down all rows by 1 else create a new
		// row
		/*
		 * if (newRow != null) { worksheet.shiftRows(destinationRowNum,
		 * worksheet.getLastRowNum(), 1); } else { newRow =
		 * worksheet.createRow(destinationRowNum); }
		 */

		// Loop through source columns to add to new row
		if (sourceRow != null) {
			for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
				// Grab a copy of the old/new cell
				XSSFCell oldCell = sourceRow.getCell(i);
				XSSFCell newCell = newRow.createCell(i);

				// If the old cell is null jump to next cell
				if (oldCell == null) {
					newCell = null;
					continue;
				}

				// Copy style from old cell and apply to new cell
				XSSFCellStyle newCellStyle = workbook.createCellStyle();
				newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
				newCell.setCellStyle(newCellStyle);

				// If there is a cell comment, copy
				if (oldCell.getCellComment() != null) {
					newCell.setCellComment(oldCell.getCellComment());
				}

				// If there is a cell hyperlink, copy
				if (oldCell.getHyperlink() != null) {
					newCell.setHyperlink(oldCell.getHyperlink());
				}

				// Set the cell data type
				newCell.setCellType(oldCell.getCellType());

				// Set the cell data value
				switch (oldCell.getCellType()) {
				case BLANK:
					newCell.setCellValue(oldCell.getStringCellValue());
					break;
				case BOOLEAN:
					newCell.setCellValue(oldCell.getBooleanCellValue());
					break;
				case ERROR:
					newCell.setCellErrorValue(oldCell.getErrorCellValue());
					break;
				case FORMULA:
					newCell.setCellFormula(oldCell.getCellFormula());
					break;
				case NUMERIC:
					newCell.setCellValue(oldCell.getNumericCellValue());
					break;
				case STRING:
					newCell.setCellValue(oldCell.getRichStringCellValue());
					break;
				default:
					break;
				}
			}

			// If there are are any merged regions in the source row, copy to new row
			for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
				CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
				if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
					CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
							(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
							cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
					worksheet.addMergedRegion(newCellRangeAddress);
				}
			}
		}
	}

}