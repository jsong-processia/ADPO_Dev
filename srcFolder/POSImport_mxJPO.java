import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.StringReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;
import java.util.StringTokenizer;
import java.util.logging.SimpleFormatter;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import matrix.db.AttributeType;
import matrix.db.Context;
import matrix.util.StringList;

import com.matrixone.apps.domain.DomainConstants;
import com.matrixone.apps.domain.util.MqlUtil;

public class POSImport_mxJPO {
	
	// Class Variables Declaration
	Workbook wbPnOImportTemplate;
	BufferedWriter bwImportInput;
	BufferedWriter bwPassportImportInput;
	BufferedWriter bwLogger;
	static final String FOLDER_SEPARATOR	=	"\\";
	static final String PRODUCT_NAME	=	"3DEXPERIENCE";
	static String EXP_VERSION	=	"R2018x";
	static String POS_SEPARATOR	=	";";
	static String POS_NULL_CHARACTER	=	"$";
	static final String TYPE_COMPANY	=	"COMPANY";
	static final String CHARACTER_SPACE	=	" ";
	static final String OPERATOR_CREATE_UPDATE	=	"*";
	static final String OPERATOR_ADD	=	"+";
	static final String OPERATOR_REMOVE	=	"-";
	static final String COMMAND_VERSION	=	"VERSION";
	static final String COMMAND_ATTRIBUTE	=	"ATTRIBUTE";
	static final String COMMAND_PROJECT	=	"PROJECT";
	static final String COMMAND_PERSON	=	"PERSON";
	static final String COMMAND_MEMBER	=	"MEMBER";
	static final String COMMAND_CONTEXT	=	"CONTEXT";
	static final String COMMAND_INACTIVE	=	"INACTIVE";
	static final String COMMAND_VISIBILITY	=	"VISIBILITY";
	static final String MAPKEY_HEADERS	=	"Headers";
	static final String COLUMN_PARENT_ORGANIZATION	=	"Parent Organization";
	static final String COLUMN_PARENT_CS	=	"Parent CS";
	static final String COLUMNTYPE_BASIC	=	"Basic";
	static final String COLUMNTYPE_ATTRIBUTE	=	"Attribute";
	static final String COLUMNTYPE_CUSTOMATTRIBUTE	=	"CustomAttribute";
	static final String COLUMNTYPE_RELATIONSHIP	=	"Relationship";
	static final String COLUMNTYPE_PRODUCT	=	"Product";
	static final String COLUMNTYPE_CONTEXT	=	"Context";
	static final String COLUMNTYPE_STATE	=	"State";
	static final String SHEET_ORG_STRUCTURE	=	"OrganizationStructure";
	static final String SHEET_USER_LIST	=	"UserList";
	static final String SHEET_CONFIG_RULES	=	"ConfigRules";
	static final String SHEET_ROLE_ASSIGNMENT	=	"RoleAssignment";
	static final String SHEET_COLLABORATIVE_SPACE	=	"CollaborativeSpaces";
	static final String CONFIG_KEY_MANDATORY_ORGSTRUCTURE	=	"OrgStructure_MandatoryColumns";
	static final String CONFIG_KEY_MANDATORY_USERLIST	=	"UserList_MandatoryColumns";
	static final String CONFIG_KEY_MANDATORY_CS	=	"CS_MandatoryColumns";
	static final String CONFIG_KEY_MANDATORY_CONTEXT	=	"Context_MandatoryColumns";
	static final String CONFIG_KEY_UNIQUE_ORGSTRUCTURE	=	"OrgStructure_UniqueColumns";
	static final String CONFIG_KEY_UNIQUE_USERLIST	=	"UserList_UniqueColumns";
	static final String CONFIG_KEY_UNIQUE_CS	=	"CS_UniqueColumns";
	static final String CONFIG_KEY_UNIQUE_CONTEXT	=	"Context_UniqueColumns";
	static final String CONFIG_KEY_INVALIDCHARS_COMPANYNAME	=	"CompanyName_InvalidChars";
	static final String CONFIG_KEY_INVALIDCHARS_ORGANIZATION	=	"Organization_InvalidChars";
	static final String CONFIG_KEY_INVALIDCHARS_CSNAME	=	"CSName_InvalidChars";
	static final String CONFIG_KEY_INVALIDCHARS_USERNAME	=	"UserName_InvalidChars";
	static final String CONFIG_KEY_PASSPORT_REGISTRATION	=	"PassportRegistration";
	static final String CONFIG_KEY_PASSPORT_COLUMNS	=	"Passport_Columns";
	static final String CONFIG_KEY_LICENSE_SEPARATOR	=	"LicenseSeparator";
	static final String DEFAULT_LICENSE_SEPARATOR	=	",";
	static final String DEFAULT_MANDATORYCOL_SEPARATOR	=	"|";
	static final String DEFAULT_INVALIDCHAR_SEPARATOR	=	" ";
	static final String DEFAULT_CONFIG_VALUE_SEPARATOR	=	"|";
	
	HashMap<String,String> hmConfigRules	=	new HashMap<String,String>();
	StringList slOrgNames	=	new StringList();
	StringList slCSNames	=	new StringList();

	/** Method to generate Input files for OOTB VPLM PnO structure and Passport Users import batch files, from provided
	 * PnO Structure data in pre-defined Excel template format.  
	 * 
	 * @param context, eMatrix object;
	 * @param args, String[] object; holding full path of input PnO Structure file.
	 * @throws Exception, if any error occured during input file generation. 
	 */
	public void generateImportInput(Context context, String[] args) throws Exception{
		if(args == null || args.length < 1){
			printInvalidArgumentsError();
		}else{
			try{
				Properties propPOSToolConfigs = LoadConfigProperties(context);
				generateFiles(propPOSToolConfigs, args[0]);
				
				writeLine(bwLogger, "INFO:: Import Input and Log Files generated.");
				hmConfigRules	=	getConfigRules();
				writeLine(bwLogger, "INFO:: Validating and Updating Config Rules.");
				validateConfigRules(propPOSToolConfigs);
				
				if("Yes".equals(hmConfigRules.get(CONFIG_KEY_PASSPORT_REGISTRATION))){
					writeLine(bwLogger, "INFO:: Writing Global variables to Passport Import Input file.");
					writeInformationBlock(bwPassportImportInput);
				}
				writeLine(bwLogger, "INFO:: Writing Global variables to POS Import Input file.");
				writeInformationBlock(bwImportInput);
				writeLine(bwLogger, "INFO:: Reading data from Input PnO template file.");
				readInputData(context);
			}catch(IOException exLogFile){
				System.out.println("Following Error Occured in IO Operations: "+ exLogFile.getMessage());
				exLogFile.printStackTrace();
			}catch(Exception ex){
				System.out.println("Following Error occurred in current execution: "+ex.getMessage());
				writeLine(bwLogger, "Following Error occurred in current execution: " + ex.getMessage());
				ex.printStackTrace();
			}finally{
				if(bwImportInput != null && bwPassportImportInput != null && bwLogger != null){
					bwImportInput.flush();
					bwImportInput.close();
					bwPassportImportInput.flush();
					bwPassportImportInput.close();
					bwLogger.flush();
					bwLogger.close();
				}
			}
		}
	}
	
	/** Method to load Admin Configuration Settings for POS Automation Tool
	 * 
	 * @param context, Enovia Context
	 * @returns Properties object with all configuration settings.
	 */
	private Properties LoadConfigProperties(Context context) throws Exception{
		Properties propConfigs	=	new Properties();
		try{
			String strConfigProperties = MqlUtil.mqlCommand(context, "print page $1 select content dump", "PRSAdminSettings.properties");
			StringReader srConfigProps = new StringReader(strConfigProperties);
			propConfigs.load(srConfigProps);
		}catch(Exception ex){
			throw ex;
		}
		return propConfigs;
	}
	
	/** Method to create ImportInpuut and Log files.
	 * 
	 * @param strTemplatePath, String object; holds PnO Data Template File path
	 * @throws Exception, if the operation fails
	 */
	private void generateFiles(Properties propConfigs, String strTemplatePath) throws Exception{
		try{
			wbPnOImportTemplate	=	WorkbookFactory.create(new FileInputStream(strTemplatePath));
			String strToolInstallationPath	=	DomainConstants.EMPTY_STRING;
			String strImportInputDirPath	=	DomainConstants.EMPTY_STRING;
			String strLogsDirPath	=	DomainConstants.EMPTY_STRING;
			
			if(strTemplatePath != null && !DomainConstants.EMPTY_STRING.equals(strTemplatePath)){
				if(strTemplatePath.indexOf(".") != -1){	// Template file path
					String strSubFolder	=	strTemplatePath.substring(0,strTemplatePath.lastIndexOf(FOLDER_SEPARATOR));
					strToolInstallationPath	=	strSubFolder.substring(0, strSubFolder.lastIndexOf(FOLDER_SEPARATOR));
				}
			}
		
			Date dtToday = Calendar.getInstance().getTime();
			DateFormat dtTodaysDate = new SimpleDateFormat("dd-MM-yyyy");
			String strTodaysDate = dtTodaysDate.format(dtToday);

			String strImportInputFile	=	propConfigs.getProperty("PRS.POSAutomationTool.PnOImportInputFileName");
			String strPassportImportFile	=	propConfigs.getProperty("PRS.POSAutomationTool.PassportInputFileName");
			String strPOSImportToolLogFile	=	propConfigs.getProperty("PRS.POSAutomationTool.LogFileName");
			String strImportInputDir	=	propConfigs.getProperty("PRS.POSAutomationTool.ImportInputDir");
			String strLogsDir	=	propConfigs.getProperty("PRS.POSAutomationTool.LogsDir");
			
			strImportInputDirPath	=	strToolInstallationPath.concat(FOLDER_SEPARATOR).concat(strImportInputDir);
			strLogsDirPath	=	strToolInstallationPath.concat(FOLDER_SEPARATOR).concat(strLogsDir);

			String strPOSImportToolFileLogPath	=	strLogsDirPath.concat(FOLDER_SEPARATOR).concat(strTodaysDate).concat(FOLDER_SEPARATOR).concat(strPOSImportToolLogFile);
			String strImportInputFilePath	=	strImportInputDirPath.concat(FOLDER_SEPARATOR).concat(strTodaysDate).concat(FOLDER_SEPARATOR).concat(strImportInputFile);
			String strPassportInputFilePath	=	strImportInputDirPath.concat(FOLDER_SEPARATOR).concat(strTodaysDate).concat(FOLDER_SEPARATOR).concat(strPassportImportFile);
			bwImportInput	=	new BufferedWriter(new FileWriter(strImportInputFilePath));
			bwPassportImportInput	=	new BufferedWriter(new FileWriter(strPassportInputFilePath));
			bwLogger	=	new BufferedWriter(new FileWriter(strPOSImportToolFileLogPath));
		}catch(Exception ex){
			System.out.println("Following Error ocurred in creating Files. " + System.lineSeparator() + ex.getMessage());
		}			
	} 
	
	/** Method to validate Config Rules entries and update mandatory entries to default value, if not provided in Input file. 
	 * 
	 * @throws Exception, if any error
	 */
	private void validateConfigRules(Properties propConfigs) throws Exception {
		if(hmConfigRules != null && !hmConfigRules.isEmpty()){
			String strConfigVal	=	DomainConstants.EMPTY_STRING;
			String strDefaultManCols	=	DomainConstants.EMPTY_STRING;
			
			String strDefaultInvalidChars = propConfigs.getProperty("PRS.POSAutomationTool.Default_InvalidChars");
			String strNullChar = propConfigs.getProperty("PRS.POSAutomationTool.Default_NullCharacter");
			String strDefaultInvalidChars_CompanyName	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_CompanyName_InvalidChars");
			String strDefaultSeparator	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_Separator");
			String strDefaultNullChar	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_NullCharacter");
			String strDefaultPassCols	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_PassportCols");
			
			if(hmConfigRules.containsKey(CONFIG_KEY_MANDATORY_ORGSTRUCTURE)){
				strDefaultManCols	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_MandatoryCols_OrgStructure");
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_MANDATORY_ORGSTRUCTURE);
				if(!strConfigVal.equals(strDefaultManCols)){
					hmConfigRules.put(CONFIG_KEY_MANDATORY_ORGSTRUCTURE, strDefaultManCols.concat(DEFAULT_MANDATORYCOL_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_MANDATORY_ORGSTRUCTURE, strDefaultManCols);
			}
			
			if(hmConfigRules.containsKey(CONFIG_KEY_MANDATORY_USERLIST)){
				strDefaultManCols	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_MandatoryCols_Userlist");
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_MANDATORY_USERLIST);
				if(!strConfigVal.equals(strDefaultManCols)){
					hmConfigRules.put(CONFIG_KEY_MANDATORY_USERLIST, strDefaultManCols.concat(DEFAULT_MANDATORYCOL_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_MANDATORY_USERLIST, strDefaultManCols);
			}
			
			if(hmConfigRules.containsKey(CONFIG_KEY_MANDATORY_CS)){
				strDefaultManCols	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_MandatoryCols_CS");
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_MANDATORY_CS);
				if(!strConfigVal.equals(strDefaultManCols)){
					hmConfigRules.put(CONFIG_KEY_MANDATORY_CS, strDefaultManCols.concat(DEFAULT_MANDATORYCOL_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_MANDATORY_CS, strDefaultManCols);
			}
			
			if(hmConfigRules.containsKey(CONFIG_KEY_MANDATORY_CONTEXT)){
				strDefaultManCols	=	propConfigs.getProperty("PRS.POSAutomationTool.Default_MandatoryCols_Context");
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_MANDATORY_CONTEXT);
				if(!strConfigVal.equals(strDefaultManCols)){
					hmConfigRules.put(CONFIG_KEY_MANDATORY_CONTEXT, strDefaultManCols.concat(DEFAULT_MANDATORYCOL_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_MANDATORY_CONTEXT, strDefaultManCols);
			}
			
			if(hmConfigRules.containsKey(CONFIG_KEY_INVALIDCHARS_COMPANYNAME)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_COMPANYNAME);
				if(!strConfigVal.equals(strDefaultInvalidChars_CompanyName)){
					hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_COMPANYNAME, strDefaultInvalidChars_CompanyName.concat(DEFAULT_INVALIDCHAR_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_COMPANYNAME, strDefaultInvalidChars_CompanyName);
			}
			
			if(hmConfigRules.containsKey(CONFIG_KEY_INVALIDCHARS_ORGANIZATION)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_ORGANIZATION);
				if(!strConfigVal.equals(strDefaultInvalidChars)){
					hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_ORGANIZATION, strDefaultInvalidChars.concat(DEFAULT_INVALIDCHAR_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_ORGANIZATION, strDefaultInvalidChars);
			}
			if(hmConfigRules.containsKey(CONFIG_KEY_INVALIDCHARS_USERNAME)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_USERNAME);
				if(!strConfigVal.equals(strDefaultInvalidChars)){
					hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_USERNAME, strDefaultInvalidChars.concat(DEFAULT_INVALIDCHAR_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_USERNAME, strDefaultInvalidChars);
			}
			if(hmConfigRules.containsKey(CONFIG_KEY_INVALIDCHARS_CSNAME)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_CSNAME);
				if(!strConfigVal.equals(strDefaultInvalidChars)){
					hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_CSNAME, strDefaultInvalidChars.concat(DEFAULT_INVALIDCHAR_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_INVALIDCHARS_CSNAME, strDefaultInvalidChars);
			}
			
			if(!hmConfigRules.containsKey(CONFIG_KEY_PASSPORT_REGISTRATION))
				hmConfigRules.put(CONFIG_KEY_PASSPORT_REGISTRATION, "Yes");
			
			if(hmConfigRules.containsKey(CONFIG_KEY_PASSPORT_COLUMNS)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_PASSPORT_COLUMNS);
				if(!strConfigVal.equals(strDefaultPassCols)){
					hmConfigRules.put(CONFIG_KEY_PASSPORT_COLUMNS, strDefaultPassCols.concat(DEFAULT_CONFIG_VALUE_SEPARATOR).concat(strConfigVal));	
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_PASSPORT_COLUMNS, strDefaultPassCols);
			}

			if(hmConfigRules.containsKey(CONFIG_KEY_LICENSE_SEPARATOR)){
				strConfigVal	=	hmConfigRules.get(CONFIG_KEY_LICENSE_SEPARATOR);
				if(!(strDefaultInvalidChars.contains(strConfigVal)) && !(strConfigVal.equals(DEFAULT_LICENSE_SEPARATOR)) && !(strConfigVal.equals(POS_SEPARATOR)) && !(strConfigVal.equals(POS_NULL_CHARACTER))){
					hmConfigRules.put(CONFIG_KEY_LICENSE_SEPARATOR, strConfigVal);	
				}else{
					hmConfigRules.put(CONFIG_KEY_LICENSE_SEPARATOR, DEFAULT_LICENSE_SEPARATOR);
				}
			}else{
				hmConfigRules.put(CONFIG_KEY_LICENSE_SEPARATOR, DEFAULT_LICENSE_SEPARATOR);
			}
			
			if(hmConfigRules.containsKey("Version"))
				EXP_VERSION	=	(hmConfigRules.get("Version") == null || DomainConstants.EMPTY_STRING.equals(hmConfigRules.get("Version")))?EXP_VERSION:hmConfigRules.get("Version");
			else if(hmConfigRules.containsKey("version"))
				EXP_VERSION	=	(hmConfigRules.get("version") == null || DomainConstants.EMPTY_STRING.equals(hmConfigRules.get("version")))?EXP_VERSION:hmConfigRules.get("version");
			if(hmConfigRules.containsKey("Separator") )
				POS_SEPARATOR	=	(hmConfigRules.get("Separator") == null || DomainConstants.EMPTY_STRING.equals(hmConfigRules.get("Separator")))?strDefaultSeparator:hmConfigRules.get("Separator");
			else if(hmConfigRules.containsKey("separator"))
				POS_SEPARATOR	=	(hmConfigRules.get("separator") == null || DomainConstants.EMPTY_STRING.equals(hmConfigRules.get("separator")))?strDefaultSeparator:hmConfigRules.get("separator");
			if(hmConfigRules.containsKey("Null_Character"))
				POS_NULL_CHARACTER	=	(hmConfigRules.get("Null_Character") == null || DomainConstants.EMPTY_STRING.equals(hmConfigRules.get("Null_Character")))?strDefaultNullChar:hmConfigRules.get("Null_Character");
			if(POS_SEPARATOR.equals(POS_NULL_CHARACTER) || strDefaultSeparator.equals(POS_NULL_CHARACTER) || strDefaultNullChar.equals(POS_SEPARATOR)){
				POS_SEPARATOR	=	strDefaultSeparator;
				POS_NULL_CHARACTER	=	strDefaultNullChar;
				try{
					writeLine(bwLogger, "WARNING:: \"SEPARATOR\" and \"NULL CHARACTER\" values can not be same or same as each other's default values. Changing both values to default.");
				}catch(IOException exLogFile){
					throw exLogFile;
				}
			}
		}
		writeLine(bwLogger, "INFO:: Updated Config Rules: "+hmConfigRules);
	}
	
	/** Method to write given String value to Import Input or Log file in specific format
	 * 
	 * @param bwWriter, BufferedWriter object; pointing to output file to write information to.
	 * @param strLine, String object; holding the data to write to input file pointed by bwWriter.
	 * @throws Exception, if any error
	 */
	private void writeLine(BufferedWriter bwWriter, String strLine) throws Exception{
		if(bwWriter != null && !DomainConstants.EMPTY_STRING.equals(strLine)){
			try{
				bwWriter.write(strLine);
				bwWriter.newLine();
			}catch(Exception ex){
				throw ex;
			}
		}
	}
	
	/** Method to declare 3DExperience Version and other global variables in POS Input file. 
	 * 
	 * @param bwInput, BufferedWriter object; pointing to output file to write information to.
	 * @throws Exception, if any error
	 */
	private void writeInformationBlock(BufferedWriter bwInput)throws Exception{
		String strCommandLine	=	OPERATOR_CREATE_UPDATE.concat(COMMAND_VERSION).concat(CHARACTER_SPACE).concat(PRODUCT_NAME).concat(EXP_VERSION);
		
		try{
			writeLine(bwInput, strCommandLine);
			strCommandLine	=	OPERATOR_CREATE_UPDATE.concat("NULL").concat(CHARACTER_SPACE).concat(POS_NULL_CHARACTER);
			writeLine(bwInput, strCommandLine);
			strCommandLine	=	OPERATOR_CREATE_UPDATE.concat("SEPARATOR").concat(CHARACTER_SPACE).concat(POS_SEPARATOR);
			writeLine(bwInput, strCommandLine);
			bwInput.newLine();
		}catch(Exception ex){
			throw ex;
		}
	}

	/** Method to get Column Headers and Type information 
	 * 
	 * @param stDataSheet, Sheet object; referring to current sheet
	 * @return HashMap<String,StringList> object; containing entries in following format {column type=[column headers of that type]} 
	 * @throws Exception, if any error
	 */
	private HashMap<String,StringList> getColumnAndTypeInfo(Sheet stDataSheet) throws Exception{
		int iTotalColumns;
		HashMap<String,StringList> hmColumnsAndTypeInfo	=	new HashMap<String,StringList>();
		
		Row rwColumnTypesRow	=	null;
		Row rwColumnNamesRow	=	null;
		Cell clTypeCell		=	null;
		Cell clHeaderCell	=	null;
		StringList slHeaders	=	new StringList();
		
		rwColumnTypesRow	=	stDataSheet.getRow(0);
		rwColumnNamesRow	=	stDataSheet.getRow(1);
		iTotalColumns		=	rwColumnNamesRow.getLastCellNum();
		
		StringList slBasicColumns			=	new StringList();
		StringList slAttributeColumns		=	new StringList();
		StringList slCustomAttributes		=	new StringList();
		StringList slRelationshipColumns	=	new StringList();
		StringList slProductColumns			=	new StringList();
		StringList slContext				=	new StringList();
		
		try{
			for(int iColIndex = 0; iColIndex < iTotalColumns; iColIndex++){
				if(null != rwColumnTypesRow)
					clTypeCell	=	rwColumnTypesRow.getCell(iColIndex);
				clHeaderCell	=	rwColumnNamesRow.getCell(iColIndex);
				if(clHeaderCell != null){
					slHeaders.addElement(getCellValue(clHeaderCell));
					if(null != clTypeCell ){
						switch(getCellValue(clTypeCell)){
						case COLUMNTYPE_BASIC:
							slBasicColumns.addElement(clHeaderCell.getStringCellValue());
							break;
						case COLUMNTYPE_ATTRIBUTE:
							slAttributeColumns.addElement(clHeaderCell.getStringCellValue());
							break;
						case COLUMNTYPE_CUSTOMATTRIBUTE:
							slCustomAttributes.addElement(clHeaderCell.getStringCellValue());
							break;
						case COLUMNTYPE_RELATIONSHIP:
							slRelationshipColumns.addElement(clHeaderCell.getStringCellValue());
							break;
						case COLUMNTYPE_PRODUCT:
							slProductColumns.addElement(clHeaderCell.getStringCellValue());
						case COLUMNTYPE_CONTEXT:
							slContext.addElement(clHeaderCell.getStringCellValue());
						}	
					}
				}
			}
		}catch(Exception ex){
			writeLine(bwLogger, "ERROR:: \"Column Header\" and \"Column Type\" names should be String values. Please correct Template format and execute tool again.");
			throw new IllegalArgumentException("Invalid Input for Header Type and Header rows.");
		}
		hmColumnsAndTypeInfo.put(MAPKEY_HEADERS, slHeaders);
		if(!slBasicColumns.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_BASIC, slBasicColumns);	
		}
		if(!slAttributeColumns.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_ATTRIBUTE, slAttributeColumns);			
		}
		if(!slCustomAttributes.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_CUSTOMATTRIBUTE, slCustomAttributes);
		}
		if(!slRelationshipColumns.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_RELATIONSHIP, slRelationshipColumns);
		}
		if(!slProductColumns.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_PRODUCT, slProductColumns);
		}
		if(!slContext.isEmpty()){
			hmColumnsAndTypeInfo.put(COLUMNTYPE_CONTEXT, slContext);
		}
		try{
			writeLine(bwLogger, "INFO:: \""+stDataSheet.getSheetName()+"\" Column Map:: "+hmColumnsAndTypeInfo);
		}catch(IOException exLogFile){
			throw exLogFile;
		}
		return hmColumnsAndTypeInfo;
	}
	
	/** Method to get variables depending on input Sheet
	 * 
	 * @param strCurrentSheetName, String object; holding Name of active sheet 
	 * @return HashMap object; containing variable values specific to input sheet
	 */
	private HashMap getSheetVariablesMap(String strCurrentSheetName) throws Exception{
		String strMandatoryColumns	=	DomainConstants.EMPTY_STRING;
		boolean blnPassport	=	false;
		HashMap hmSheetValueMap	=	new HashMap();
		
		switch(strCurrentSheetName){
		case SHEET_ORG_STRUCTURE:
			strMandatoryColumns	=	(String)hmConfigRules.get(CONFIG_KEY_MANDATORY_ORGSTRUCTURE);
			break;
		case SHEET_USER_LIST:
			strMandatoryColumns	=	(String)hmConfigRules.get(CONFIG_KEY_MANDATORY_USERLIST);
			if("Yes".equals(hmConfigRules.get(CONFIG_KEY_PASSPORT_REGISTRATION)))
				blnPassport	=	true;
			break;
		case SHEET_COLLABORATIVE_SPACE:
			strMandatoryColumns	=	(String)hmConfigRules.get(CONFIG_KEY_MANDATORY_CS);
			break;
		case SHEET_ROLE_ASSIGNMENT:
			strMandatoryColumns	=	(String)hmConfigRules.get(CONFIG_KEY_MANDATORY_CONTEXT);
			break;
		}
		hmSheetValueMap.put("MandatoryColumns", strMandatoryColumns);
		hmSheetValueMap.put(CONFIG_KEY_PASSPORT_REGISTRATION, blnPassport);
		try{
			writeLine(bwLogger, "INFO:: \""+strCurrentSheetName+"\" Sheet Variables:: "+hmSheetValueMap);	
		}catch(IOException exLogFile){
			throw exLogFile;
		}
		
		return hmSheetValueMap;
	}
	
	/** Method to read Input PnO data from excel file with predefined format, 
	 * validate and write data to text files which will be provided as input to OOTB POS import and Passport Import batch. 
	 * 
	 * @param context, the eMatrix Context object;
	 * @returns nothing
	 * @throws Exception, if operation failed
	 */
	private void readInputData(Context context) throws Exception{
		
		try{
			if(wbPnOImportTemplate != null){
				HashMap<String, StringList> hmColumnsAndTypeInfo;
				int iTotalSheets	=	wbPnOImportTemplate.getNumberOfSheets();
				
				if(iTotalSheets > 0){
					Sheet stCurrentSheet;
					int iLastRowIndex	=	0;
					int iTotalRows		=	0;
					int iDataRowMinLimit	=	2;
					int iRowCounter		=	0;
					String strInputLine	=	DomainConstants.EMPTY_STRING;
					String strPassInputLine	=	DomainConstants.EMPTY_STRING;
					String strAction	=	DomainConstants.EMPTY_STRING;
					Row rwDataRow;
					Cell clOrgData;
					int iCellCounter	=	0;
					String strCellValue	=	DomainConstants.EMPTY_STRING;
					String strCurrColumn	=	DomainConstants.EMPTY_STRING;
					String strCurrentSheetName	=	DomainConstants.EMPTY_STRING;
					String strMandatoryColumns	=	DomainConstants.EMPTY_STRING;
					boolean blnInvalidName	=	false;
					boolean blnInvalidData	=	false;
					boolean blnPassport		=	false;
					StringList slHeaders	=	null;
					StringList slBasics		=	null;
					StringList slAttributes	=	null;
					StringList slCustomAttributes	=	null;
					StringList slRelationships	=	null;
					StringList slProducts	=	null;
					HashMap hmSheetValueMap	=	null;
					String strUserProducts	=	DomainConstants.EMPTY_STRING;
					String strParent		=	DomainConstants.EMPTY_STRING;
					
					for (int iSheetCounter=0; iSheetCounter<iTotalSheets; iSheetCounter++){	// START:: for loop for all sheets
						stCurrentSheet		=	wbPnOImportTemplate.getSheetAt(iSheetCounter);	
						strCurrentSheetName	=	stCurrentSheet.getSheetName();

						// config rules sheet not needed to be read again.						
						if(stCurrentSheet != null && !SHEET_CONFIG_RULES.equals(strCurrentSheetName)){
							iLastRowIndex	=	stCurrentSheet.getLastRowNum();
							iTotalRows		=	stCurrentSheet.getPhysicalNumberOfRows();
							
							//Check existence of Data row.
							if(iLastRowIndex >= (iDataRowMinLimit-1) && iTotalRows > iDataRowMinLimit){	// START:: Data Row presence check
								writeLine(bwLogger, "INFO:: Processing Sheet \""+strCurrentSheetName+"\".");
								iCellCounter	=	0;
								strCellValue	=	DomainConstants.EMPTY_STRING;
								strCurrColumn	=	DomainConstants.EMPTY_STRING;
								blnInvalidName	=	false;
								blnInvalidData	=	false;
								iRowCounter		=	0;
								strInputLine	=	DomainConstants.EMPTY_STRING;
								strUserProducts	=	DomainConstants.EMPTY_STRING;
								strPassInputLine	=	DomainConstants.EMPTY_STRING;
								
								// Get Sheetwise variables
								hmSheetValueMap		=	getSheetVariablesMap(strCurrentSheetName);
								strMandatoryColumns	=	(String)hmSheetValueMap.get("MandatoryColumns");
								blnPassport			=	(Boolean)hmSheetValueMap.get(CONFIG_KEY_PASSPORT_REGISTRATION);
								
								// Get column Header and column Type information. 
								hmColumnsAndTypeInfo	=	getColumnAndTypeInfo(stCurrentSheet);
								slHeaders	=	hmColumnsAndTypeInfo.get(MAPKEY_HEADERS);
								slBasics	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_BASIC);
								if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_ATTRIBUTE))
									slAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_ATTRIBUTE);
								if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_CUSTOMATTRIBUTE))
									slCustomAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_CUSTOMATTRIBUTE);
								if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_RELATIONSHIP)){
									slRelationships	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_RELATIONSHIP);
								}
								if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_PRODUCT)){
									slProducts	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_PRODUCT);
								}
								
								// Iterate data rows. 
								for(iRowCounter=iDataRowMinLimit; iRowCounter <= iLastRowIndex; iRowCounter++){
									rwDataRow		=	stCurrentSheet.getRow(iRowCounter);
									blnInvalidData	=	false;
									strInputLine	=	OPERATOR_CREATE_UPDATE;
									strUserProducts	=	DomainConstants.EMPTY_STRING;
									strPassInputLine	=	DomainConstants.EMPTY_STRING;
								
									blnInvalidData	=	checkMandatoryColumnContents(strMandatoryColumns, rwDataRow, slHeaders);
									if(blnInvalidData){
										continue;
									}else{
										// Iterate columns
										for(iCellCounter=0; iCellCounter < slHeaders.size(); iCellCounter++){
											strCellValue	=	DomainConstants.EMPTY_STRING;
											strCurrColumn	=	DomainConstants.EMPTY_STRING;
											blnInvalidName	=	false;
											clOrgData		=	rwDataRow.getCell(iCellCounter);
											strCurrColumn	=	(String)slHeaders.get(iCellCounter);
											
											// START: Read and Validate Column Values
											strCellValue	=	getCellValue(clOrgData);
											if(strCellValue == null || DomainConstants.EMPTY_STRING.equals(getCellValue(clOrgData))){
												// Get default value for empty cell depending on column
												strCellValue	=	checkAndReplaceNullCellValue(context, strCurrentSheetName, strCellValue, strCurrColumn, hmColumnsAndTypeInfo, hmSheetValueMap, iRowCounter);
												if(null == strCellValue){
													continue;
												}else if("InvalidData".equals(strCellValue)){
													blnInvalidData	=	true;
													break;
												}
											}else{
												if("Name".equals(strCurrColumn) || "Username".equals(strCurrColumn)){
													// Validate "Name" value for invalid characters.
													blnInvalidName	=	validateName(strCellValue,strCurrentSheetName,getCellValue(rwDataRow.getCell(0)),iRowCounter);
													if(blnInvalidName == true){
														blnInvalidData	=	true;
														break;
													}
												}
												
												// Update User's Product assignment
												if(SHEET_USER_LIST.equals(strCurrentSheetName) && slProducts!= null && slProducts.contains(strCurrColumn)){
													if("Yes".equals(strCellValue)){
														if(slProducts.indexOf(strCurrColumn)>0){
															strUserProducts	=	strUserProducts.concat(hmConfigRules.get(CONFIG_KEY_LICENSE_SEPARATOR));	
														}
														strUserProducts	=	strUserProducts.concat(strCurrColumn);
													}
												}else{
													strCellValue	=	formInputLineFromCellValue(strCurrentSheetName, strCurrColumn, strCellValue, iRowCounter);	
												}
													
												if("InvalidData".equals(strCellValue)){
													blnInvalidData	=	true;
													break;
												}else if("Continue".equals(strCellValue)){
													continue;
												}
											}
											// END: Read and Validate Column Values
											
											// START: Update Organization List
											if(SHEET_ORG_STRUCTURE.equals(strCurrentSheetName) && !blnInvalidData && COLUMN_PARENT_ORGANIZATION.equals(strCurrColumn) ){
												strParent	=	getCellValue(rwDataRow.getCell(1));
												if(!DomainConstants.EMPTY_STRING.equals(strParent) && !slOrgNames.contains(strParent)){
													slOrgNames.addElement(strParent);
												}
											}else if(SHEET_COLLABORATIVE_SPACE.equals(strCurrentSheetName) && !blnInvalidData && COLUMN_PARENT_CS.equals(strCurrColumn) ){
												strParent	=	getCellValue(rwDataRow.getCell(0));
												if(!DomainConstants.EMPTY_STRING.equals(strParent) && !slCSNames.contains(strParent)){
													slCSNames.addElement(strParent);
												}
											}
											
											// START: Form information to write into Passport input file
											if(SHEET_USER_LIST.equals(strCurrentSheetName) && blnPassport && (hmConfigRules.get(CONFIG_KEY_PASSPORT_COLUMNS)).contains(strCurrColumn)){
												if(slBasics != null && slBasics.contains(strCurrColumn)) {
													strPassInputLine	=	strPassInputLine.concat(strCellValue);
													if((iCellCounter > 0) && (iCellCounter < slBasics.size()-1)){
														strPassInputLine	=	strPassInputLine.concat(POS_SEPARATOR);
													}
												}
											}
											
											// START: Form information to write into ImportInput file.
											if(SHEET_USER_LIST.equals(strCurrentSheetName) && !DomainConstants.EMPTY_STRING.equals(strCellValue) && COLUMNTYPE_CONTEXT.equals(strCurrColumn)){
												strAction	=	getCellValue(rwDataRow.getCell(slHeaders.indexOf("Action")));
												strInputLine	=	formContextAssignmentLine(strAction, strCellValue, strUserProducts);
												
											}else if(!DomainConstants.EMPTY_STRING.equals(strCellValue)){
												strInputLine	=	updateInputLineBasedOnColumn(strInputLine, strCellValue, strCurrColumn, hmColumnsAndTypeInfo, rwDataRow, iCellCounter);											
											}
											// END: Form information to write into ImportInput file.
											
											// START: Write lines to ImportInput file.
											if(slBasics != null && slBasics.contains(strCurrColumn)){	
												if(iCellCounter == (slBasics.size()-1)){
													writeLine(bwImportInput, strInputLine);
													strInputLine	=	DomainConstants.EMPTY_STRING;	
												}
												if(SHEET_USER_LIST.equals(strCurrentSheetName) && blnPassport){
													if(iCellCounter == (slBasics.size()-1)){
														writeLine(bwPassportImportInput, strPassInputLine);
														strPassInputLine	=	DomainConstants.EMPTY_STRING;
													}
												}
											}else if(!DomainConstants.EMPTY_STRING.equals(strInputLine) ){
												writeLine(bwImportInput, strInputLine);
												
												if(SHEET_USER_LIST.equals(strCurrentSheetName) && blnPassport){
													if(hmConfigRules.get(CONFIG_KEY_PASSPORT_COLUMNS).contains(strCurrColumn)){
														writeLine(bwPassportImportInput, strInputLine);
													}
												}
												strInputLine	=	DomainConstants.EMPTY_STRING;
											} 
										
											if(iCellCounter == slHeaders.size() && !blnInvalidData){
												bwImportInput.newLine();
											}
											// END: Write lines to ImportInput file.
											
										}	// End: for loop for all column cells.
									}
										
									if(!blnInvalidData){
										bwImportInput.newLine();
										if(SHEET_USER_LIST.equals(strCurrentSheetName) && blnPassport)
											bwPassportImportInput.newLine();
									}
								}	// END: for loop for Excel Rows
							}	// END:: Data Row presence check				
						}	// END:: Sheet null check
					}	// END:: for loop for all sheets
				}	// END:: Total sheet number check
			}	// END:: workbook null check
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	
	/** Method to update input line for Person's context assignment.  
	 * 
	 * @param strAction
	 * @param strCellValue
	 * @param strUserProducts
	 * @return
	 */
	private String formContextAssignmentLine(String strAction, String strCellValue, String strUserProducts){
		String strInputLine	=	DomainConstants.EMPTY_STRING;
		StringList slContexts	=	new StringList();
		String strContext	=	DomainConstants.EMPTY_STRING;
		
		if(!DomainConstants.EMPTY_STRING.equals(strCellValue)){
			if(strCellValue.contains(DEFAULT_CONFIG_VALUE_SEPARATOR)){
				StringTokenizer stContexts	=	new StringTokenizer(strCellValue, DEFAULT_CONFIG_VALUE_SEPARATOR);
				while(stContexts.hasMoreTokens()){
					slContexts.addElement(stContexts.nextToken());
				}
			}else{
				slContexts.addElement(strCellValue);
			}
			
			for(int index=0;index<slContexts.size();index++){
				strContext	=	(String)slContexts.get(index);
				if(!DomainConstants.EMPTY_STRING.equals(strAction) && "Add".equals(strAction)){
					strInputLine	=	strInputLine.concat(OPERATOR_ADD).concat(COMMAND_CONTEXT).concat(CHARACTER_SPACE).concat(strContext);
					if(!DomainConstants.EMPTY_STRING.equals(strUserProducts)){
						strInputLine	=	strInputLine.concat(POS_SEPARATOR).concat(hmConfigRules.get(CONFIG_KEY_LICENSE_SEPARATOR)).concat(POS_SEPARATOR).concat(strUserProducts);
					}
				}else if("Remove".equals(strAction)){
					strInputLine	=	strInputLine.concat(OPERATOR_REMOVE).concat(COMMAND_CONTEXT).concat(CHARACTER_SPACE).concat(strCellValue);
				}
				if(index<(slContexts.size()-1) && !DomainConstants.EMPTY_STRING.equals(strInputLine)){
					strInputLine	=	strInputLine.concat(System.lineSeparator());
				}
			}
		}
		return strInputLine;
	}
	
	/** Method to check if all mandatory columns have value in a data row. 
	 * 
	 * @param strMandatoryColumns, String object; holding "|" separated list of mandatory columns for current sheet.
	 * @param rwDataRow, Row object; holding current data row in a sheet.
	 * @param slHeaders, StringList object; holding list of Column header names of current sheet. 
	 * @returns boolean value; true if any of the mandatory cell is empty for data row, else false.
	 */
	private boolean checkMandatoryColumnContents(String strMandatoryColumns, Row rwDataRow, StringList slHeaders){
		boolean blnSkipRow	=	false;
		if(!DomainConstants.EMPTY_STRING.equals(strMandatoryColumns)){
			String strColumnVal	=	DomainConstants.EMPTY_STRING;
			StringList slColumns	=	new StringList();
			Cell clData	=	null;
			Iterator itrCells	=	null;
			int iCellCntr	=	0;
			
			if(strMandatoryColumns.contains(DEFAULT_MANDATORYCOL_SEPARATOR)){
				StringTokenizer stManColumns	=	new StringTokenizer(strMandatoryColumns, DEFAULT_MANDATORYCOL_SEPARATOR);
				while(stManColumns.hasMoreTokens()){
					slColumns.addElement(stManColumns.nextToken());
				}
			}else{
				slColumns.addElement(strMandatoryColumns);
			}
			for (String strColumn : slColumns) {
				if(slHeaders.contains(strColumn)){
					clData	=	rwDataRow.getCell(slHeaders.indexOf(strColumn));
					if(clData == null || (DomainConstants.EMPTY_STRING.equals(getCellValue(clData)))){
						blnSkipRow	=	true;
						itrCells	=	rwDataRow.iterator();
						while(itrCells.hasNext()){
							strColumnVal	=	getCellValue((Cell)itrCells.next());
							if(!DomainConstants.EMPTY_STRING.equals(strColumnVal)){
								iCellCntr++;
							}
						}
						if(iCellCntr > 0){
							try{
								writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(rwDataRow.getRowNum()+1)+". \""+strColumn +"\" column value can not be empty.");
							}catch(Exception exLogFile){
								System.out.println("Following Error Occured in current execution: "+ exLogFile.getMessage());
								System.out.println("Error Stack Trace: " + exLogFile.getStackTrace());
							}
						}
					}
				}
			}
		}
		return blnSkipRow;
	}
	
	/** Method to handle Null or Empty cells; depending on the corresponding column.
	 *  For Mandatory columns, tool will log error message and move to next data row;
	 *  For other columns, null value will be replaced either by NULL_CHARACTER or default value.
	 * 
	 * @param context, <code>eMatrix Context</code> object
	 * @param strCurrentSheetName, String object; holding name of active Sheet of input excel template file
	 * @param strCurrColumn, String object; holding name of corresponding column in active sheet
	 * @param hmColumnsAndTypeInfo, HashMap<String, StringList> object; containing column headers and type information
	 * @param iRowCounter, integer value; holding current row number in active sheet
	 * @return String object; holding replaced value depending on column
	 * @throws Exception, if any error
	 */
	private String checkAndReplaceNullCellValue(Context context, String strCurrentSheetName, String strCellValue, String strCurrColumn, HashMap<String, StringList> hmColumnsAndTypeInfo, HashMap hmSheetValueMap, int iRowCounter) throws Exception{
		StringList slBasics		=	null;
		StringList slAttributes	=	null;
		StringList slCustomAttributes	=	null;
		String strMandatoryColumns	=	DomainConstants.EMPTY_STRING;
		
		strMandatoryColumns	=	(String)hmSheetValueMap.get("MandatoryColumns");
		int iLastRowIndex	=	wbPnOImportTemplate.getSheet(strCurrentSheetName).getLastRowNum();
		
		slBasics	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_BASIC);
		if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_ATTRIBUTE))
			slAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_ATTRIBUTE);
		if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_CUSTOMATTRIBUTE))
			slCustomAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_CUSTOMATTRIBUTE);
		
		if((strMandatoryColumns).contains(strCurrColumn)){
			// For "OrganizationStructure" sheet, last row may contain pre-defined template data input entries for some columns.
			// Check is to avoid false identification of last row as data row.
			if(SHEET_ORG_STRUCTURE.equals(strCurrentSheetName) && iRowCounter == iLastRowIndex){
				strCellValue	=	"InvalidData";
			}else{
				strCellValue	=	"InvalidData";
				writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". \""+strCurrColumn +"\" column value can not be empty.");
			}
		}else{
			if(slBasics != null && slBasics.contains(strCurrColumn)){
				switch(strCurrColumn){
				case "LicenseType":
					//To-Do: Get Default LicenseType from ConfigRules and set that value. If key not present, use Full licenses as default.
					strCellValue	=	"0";
					break;
				case COLUMN_PARENT_ORGANIZATION:
					if(SHEET_ORG_STRUCTURE.equals(strCurrentSheetName)){
						String strType	=	getCellValue(wbPnOImportTemplate.getSheet(strCurrentSheetName).getRow(iRowCounter).getCell(0));
						if(!TYPE_COMPANY.equals(strType)){
							strCellValue	=	"InvalidData";
							writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". \"Parent Organization\" value can not be empty for \"Business Unit\" or \"Department\".");	
						}else{
							strCellValue	=	POS_NULL_CHARACTER;
						}	
					}else{
						strCellValue	=	POS_NULL_CHARACTER;
					}
					break;
				default:
						strCellValue	=	POS_NULL_CHARACTER;
				}
			}else if((slAttributes != null && slAttributes.contains(strCurrColumn)) || (slCustomAttributes != null && slCustomAttributes.contains(strCurrColumn))){
				try{
					strCellValue	=	(new AttributeType(strCurrColumn)).getDefaultValue(context);
				}catch(Exception ex){
					if(ex.getMessage().contains("does not exist")){
						writeLine(bwLogger, "WARNING:: Ignoring Input at row# "+(iRowCounter+1)+". Attribute \""+strCurrColumn+"\" is not present in target environment.");
					}
				}
			}
		}
		return strCellValue;
	}
	
	/** Method to form part of Input line to be written to Import Input files, from provided cell value, depending on column.
	 * 
	 * @param strCurrentSheetName, String object; holding name of active Sheet of input excel template file
	 * @param strCurrColumn, String object; holding name of corresponding column in active sheet
	 * @param strCellValue, String object; holding value of current cell of active sheet
	 * @param iRowCounter, integer value; holding current row number in active sheet
	 * @return String object; holding part of input line to be written to input files
	 */
	private String formInputLineFromCellValue(String strCurrentSheetName, String strCurrColumn, String strCellValue, int iRowCounter){
		try{
			if(SHEET_ORG_STRUCTURE.equals(strCurrentSheetName)){
				if("Type".equals(strCurrColumn)){
					strCellValue	=	strCellValue.concat(CHARACTER_SPACE);
				}else if(COLUMN_PARENT_ORGANIZATION.equals(strCurrColumn)){
					if(!slOrgNames.contains(strCellValue)){
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". "+strCurrColumn
								+ " \""+strCellValue+"\" is not present in Input Data provided till now.");
						strCellValue	=	"InvalidData";
					}
				}
			}else if(SHEET_COLLABORATIVE_SPACE.equals(strCurrentSheetName)){
				if("Name".equals(strCurrColumn)){
					strCellValue	=	COMMAND_PROJECT.concat(CHARACTER_SPACE).concat(strCellValue).concat(POS_SEPARATOR);
				}else if(COLUMN_PARENT_CS.equals(strCurrColumn)){
					if(!slCSNames.contains(strCellValue)){
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". "+strCurrColumn
								+ " \""+strCellValue+"\" is not present in Input Data provided till now.");
						strCellValue	=	"InvalidData";
					}
				}
			}else if(SHEET_USER_LIST.equals(strCurrentSheetName)){
				if("Name".equals(strCurrColumn)){
					strCellValue	=	COMMAND_PERSON.concat(CHARACTER_SPACE).concat(strCellValue).concat(POS_SEPARATOR);
				}else if("Company".equals(strCurrColumn)){
					if(slOrgNames != null && !slOrgNames.contains(strCellValue)){
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+", Column \""+strCurrColumn +"\". Organization named \""+strCellValue+"\" is not present in Input Data provided till now.");
						strCellValue	=	"InvalidData";
					}
				}else if("LicenseType".equals(strCurrColumn)){
					strCellValue	=	("Full".equals(strCellValue))?"0":"40";
				}
			}else if(SHEET_ROLE_ASSIGNMENT.equals(strCurrentSheetName)){
				if("Role".equals(strCurrColumn)){
					strCellValue	=	COMMAND_CONTEXT.concat(CHARACTER_SPACE).concat(strCellValue).concat(POS_SEPARATOR);
				}
			}
		}catch(Exception exLogFile){
			System.out.println("Following Error Occured in current execution: "+ exLogFile.getMessage());
			System.out.println("Error Stack Trace: " + exLogFile.getStackTrace());
		}
		return strCellValue;
	}
	
	/** Method to update Input Line with expected syntax as per corresponding Column
	 * 
	 * @param strInputLine, String object; holding part of input line to be written to input files
	 * @param strCellValue, String object; holding value of current cell of active sheet
	 * @param strCurrColumn, String object; holding name of corresponding column in active sheet
	 * @param hmColumnsAndTypeInfo, HashMap<String, StringList> object; containing column headers and type information
	 * @param rwDataRow, Row object; holding current row in active sheet
	 * @param iCellCounter, integer value; holding index of current cell
	 * @return String object; holding updated part of input line to be written to input files
	 */
	private String updateInputLineBasedOnColumn(String strInputLine, String strCellValue, String strCurrColumn, HashMap<String, StringList> hmColumnsAndTypeInfo, Row rwDataRow, int iCellCounter){
		String strAction	=	DomainConstants.EMPTY_STRING;
		StringList slHeaders	=	null;
		StringList slBasics		=	null;
		StringList slAttributes	=	null;
		StringList slCustomAttributes	=	null;
		
		slHeaders	=	hmColumnsAndTypeInfo.get(MAPKEY_HEADERS);
		slBasics	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_BASIC);
		if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_ATTRIBUTE))
			slAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_ATTRIBUTE);
		if(hmColumnsAndTypeInfo.containsKey(COLUMNTYPE_CUSTOMATTRIBUTE))
			slCustomAttributes	=	hmColumnsAndTypeInfo.get(COLUMNTYPE_CUSTOMATTRIBUTE);
		
		if(slBasics != null && slBasics.contains(strCurrColumn)){
			strInputLine	=	strInputLine.concat(strCellValue);
			if((iCellCounter > 0) && (iCellCounter < slBasics.size()-1)){
				strInputLine	=	strInputLine.concat(POS_SEPARATOR);
			}
		}else if((slAttributes != null && slAttributes.contains(strCurrColumn)) 
				|| (slCustomAttributes != null && slCustomAttributes.contains(strCurrColumn))){
			if(iCellCounter >= slBasics.size() && !DomainConstants.EMPTY_STRING.equals(strCellValue)){
				strInputLine	=	OPERATOR_ADD.concat(COMMAND_ATTRIBUTE).concat(CHARACTER_SPACE);
				strInputLine	=	strInputLine.concat(strCurrColumn).concat(POS_SEPARATOR).concat(strCellValue);
			}
		}else if(!DomainConstants.EMPTY_STRING.equals(strCellValue) && DomainConstants.RELATIONSHIP_MEMBER.equals(strCurrColumn)){
			strInputLine	=	formMemberRelationshipLine(strCellValue, strCurrColumn, rwDataRow.getRowNum());
		}else if(!DomainConstants.EMPTY_STRING.equals(strCellValue) && COLUMNTYPE_STATE.equals(strCurrColumn)){
			if("Inactive".equals(strCellValue)){
				strInputLine	=	OPERATOR_ADD.concat(COMMAND_INACTIVE);
			}
		}else if(!DomainConstants.EMPTY_STRING.equals(strCellValue) && "Visibility".equals(strCurrColumn)){
			strInputLine	=	OPERATOR_ADD.concat(COMMAND_VISIBILITY).concat(CHARACTER_SPACE).concat(strCellValue);
		}
		return strInputLine;
	}
	
	/** Method to update Import file input with Member relationship information for Person
	 * 
	 * @param strCellValue, String object; holding the value of Member connection of Person. "|" separated list in case multiple connections.
	 * @param strCurrColumn, String object; holding the name of current column.
	 * @param iRowCounter, integer value; holding number of currnet data row. 
	 * @returns String object; holding Member relationship part of current data line in PnOImportInput file. 
	 */
	private String formMemberRelationshipLine(String strCellValue, String strCurrColumn, int iRowCounter){
		String strInputLine	=	DomainConstants.EMPTY_STRING;
		StringList slMembers	=	new StringList();
		String strMember	=	DomainConstants.EMPTY_STRING;
		try{
			if(!DomainConstants.EMPTY_STRING.equals(strCellValue)){
				if(strCellValue.contains(DEFAULT_CONFIG_VALUE_SEPARATOR)){
					StringTokenizer stMembers	=	new StringTokenizer(strCellValue, DEFAULT_CONFIG_VALUE_SEPARATOR);
					while(stMembers.hasMoreTokens()){
						slMembers.addElement(stMembers.nextToken());
					}
				}else{
					slMembers.addElement(strCellValue);
				}
				for(int index=0;index<slMembers.size();index++){
					strMember	=	(String)slMembers.get(index);
					if(!slOrgNames.contains(strMember)){
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+", Column \""+strCurrColumn +"\". Organization named \""+strMember+"\" is not present in Input Data provided till now.");
					}else{
						strInputLine	=	strInputLine.concat(OPERATOR_ADD).concat(COMMAND_MEMBER).concat(CHARACTER_SPACE).concat((String)slMembers.get(index));	
					}
					if(index<(slMembers.size()-1) && !DomainConstants.EMPTY_STRING.equals(strInputLine)){
						strInputLine	=	strInputLine.concat(System.lineSeparator());
					}
				}
			}
		}catch(Exception ex){
			System.out.println("Following Error Occured in current execution: "+ ex.getMessage());
			System.out.println("Error Stack Trace: " + ex.getStackTrace());
		}
		return strInputLine;
	}
	
	/** Method to validate "Name" value for invalid characters.
	 * 
	 * @param strCellValue, String object; holding Name value to be verified.
	 * @param strSheetName, String object; holding name of active Sheet
	 * @param strCellType, String object; holding Input Object's Type value for "Organization" object. It can be empty for other type objects.
	 * @return boolean value; true if name contains invalid characters, else false.
	 */
	private boolean validateName(String strCellValue, String strSheetName, String strCellType, int iRowCounter){
		boolean isInvalid	=	false;
		
		if(!DomainConstants.EMPTY_STRING.equals(strSheetName)){
			String strInvalidChars	=	DomainConstants.EMPTY_STRING;
			String strOrgInvalidChars	=	DomainConstants.EMPTY_STRING;

			if(SHEET_ORG_STRUCTURE.equals(strSheetName)){
				strOrgInvalidChars	=	(String)hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_COMPANYNAME);
				strInvalidChars	=	(String)hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_ORGANIZATION);
				String[] strarrOrgInvalidChars	=	strOrgInvalidChars.split(DEFAULT_INVALIDCHAR_SEPARATOR);
				if(TYPE_COMPANY.equals(strCellType)){
					for (String strInvalid : strarrOrgInvalidChars) {
						if(strCellValue.contains(strInvalid)){
							isInvalid	=	true;
							break;
						}
					}
				}
			}else if(SHEET_USER_LIST.equals(strSheetName) || SHEET_ROLE_ASSIGNMENT.equals(strSheetName)){
				strInvalidChars	=	(String)hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_USERNAME);
			}else if(SHEET_COLLABORATIVE_SPACE.equals(strSheetName)){
				strInvalidChars	=	hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_CSNAME);
			}
			if(!DomainConstants.EMPTY_STRING.equals(strInvalidChars) && strInvalidChars.contains(DEFAULT_INVALIDCHAR_SEPARATOR)){
				String[] strarrInvalidChars	=	strInvalidChars.split(DEFAULT_INVALIDCHAR_SEPARATOR);
				for (String strInvalid : strarrInvalidChars) {
					if(strCellValue.contains(strInvalid)){
						isInvalid	=	true;
						break;
					}
				}	
			}
			try{
				if(isInvalid){
					switch(strSheetName){
					case SHEET_ORG_STRUCTURE:
						if(TYPE_COMPANY.equals(strCellType)){
							writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". Company Name value can not contain following characters: \""+hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_COMPANYNAME)+ hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_ORGANIZATION) +"\".");
						}else{
							writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". Name value can not contain following characters: \""+hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_ORGANIZATION) +"\".");
						}
						break;
					case SHEET_USER_LIST:
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". Name value can not contain following characters: \""+hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_USERNAME) +"\".");
						break;
					case SHEET_COLLABORATIVE_SPACE:
						writeLine(bwLogger, "ERROR:: Invalid Input at row# "+(iRowCounter+1)+". Name value can not contain following characters: \""+hmConfigRules.get(CONFIG_KEY_INVALIDCHARS_CSNAME) +"\".");
						break;
					}
				}
			}catch(Exception ex){
				System.out.println("Following Error Occured in current execution: "+ ex.getMessage());
				System.out.println("Error Stack Trace: " + ex.getStackTrace());
			}
		}
		return isInvalid;
	}
	
	/** Method to get value of given Cell into String/text format.
	 * 
	 * @param clInputCell, Cell object; holding active cell whose value is to be obtained
	 * @return String object; holding value of input Cell object
	 */
	private String getCellValue(Cell clInputCell){
		String strCellValue	=	DomainConstants.EMPTY_STRING;
		FormulaEvaluator fe	=	wbPnOImportTemplate.getCreationHelper().createFormulaEvaluator();
		if(clInputCell != null){
			if(clInputCell.getCellType() == clInputCell.CELL_TYPE_STRING){
				strCellValue	=	clInputCell.getStringCellValue();
			}else if(clInputCell.getCellType() == clInputCell.CELL_TYPE_NUMERIC){
				clInputCell.setCellType(clInputCell.CELL_TYPE_STRING);
				strCellValue	=	clInputCell.getStringCellValue();
			}else if(clInputCell.getCellType() == clInputCell.CELL_TYPE_FORMULA){
				if(clInputCell.getCachedFormulaResultType() == clInputCell.CELL_TYPE_STRING){
					strCellValue	=	clInputCell.getStringCellValue();
				}else if(clInputCell.getCachedFormulaResultType() == clInputCell.CELL_TYPE_NUMERIC){
					strCellValue	=	Double.toString(clInputCell.getNumericCellValue());	
				}else
					strCellValue	=	DomainConstants.EMPTY_STRING;
			}
		}
		return strCellValue.trim();
	}
	
	/** Method to print Error in case of invalid arguments passed to program. Also to print information on required arguments. 
	 * 
	 * @throws Exception, if faced any error while logging invalid argument error to log file.
	 */
	private void printInvalidArgumentsError() throws Exception{
		try{
			System.out.println("ERROR >> Invalid Arguments. Please provide Excel Template path as argument and run this tool again.");
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	
	/** Method to get Config Rules from Input PnO Structure Template file.
	 * 
	 * @return HashMap<String,String> object; Map containing ConfigRules entry in key=value format
	 * @throws Exception, if any error
	 */
	private HashMap<String,String> getConfigRules() throws Exception{
		HashMap<String,String> hmConfigRules	=	new HashMap<String,String>();
		try{
			writeLine(bwLogger, "INFO:: Reading Config Rules from template.");
			Row rwConfigEntry;
			Cell clConfigKey, clConfigValue;
			
			if(null != wbPnOImportTemplate){
				Sheet stConfigRulesSheet	=	wbPnOImportTemplate.getSheet(SHEET_CONFIG_RULES);
				if(stConfigRulesSheet.getLastRowNum() > 0){
					Iterator<Row> itrConfigRows	=	stConfigRulesSheet.rowIterator();
					String strConfigKey=DomainConstants.EMPTY_STRING;
					String strConfigValue=DomainConstants.EMPTY_STRING;
					while (itrConfigRows.hasNext()) {
						strConfigKey	=	DomainConstants.EMPTY_STRING;
						strConfigValue	=	DomainConstants.EMPTY_STRING;
						rwConfigEntry	=	itrConfigRows.next();
						clConfigKey		=	rwConfigEntry.getCell(0);
						clConfigValue	=	rwConfigEntry.getCell(1);
						
						try{
							strConfigKey	=	getCellValue(clConfigKey);
							strConfigValue	=	getCellValue(clConfigValue);
						}catch(NullPointerException exNullValue){
							writeLine(bwLogger, "WARNING:: Null/Empty value present at Row#"+(rwConfigEntry.getRowNum()+1)+" in ConfigRules sheet.");
						}
						hmConfigRules.put(strConfigKey,strConfigValue);
					}
				}
			}
			writeLine(bwLogger, "INFO:: Config Rules Map: "+hmConfigRules);
		}catch(Exception ex){
			ex.printStackTrace();
		}
		return hmConfigRules;
	}
}
