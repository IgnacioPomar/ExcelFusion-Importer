
package es.ipb.excelfusion.config;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import es.ipb.excelfusion.ui.wizard.Step3StructureValidationPage.SheetValidationResult;
import es.ipb.excelfusion.ui.wizard.Step4TypeInferencePage.ColumnDefinition;
import es.ipb.excelfusion.ui.wizard.Step4TypeInferencePage.ColumnType;
import es.ipb.excelfusion.ui.wizard.Step5DatabaseConfigPage.DbType;


/**
 * Contains the full import configuration gathered across all wizard steps.
 * This object is passed to Step6 in order to execute the actual import.
 */
public class ImportConfiguration
{

	// === Step 1 ===
	private File						 dataDirectory;
	private List <File>					 selectedFiles	= new ArrayList <> ();

	// === Step 2 ===
	private Integer						 headerRow;							  // null or >= 1
	private Integer						 dataStartRow;						  // always >= 1
	private boolean						 autoIncrement;
	private List <Boolean>				 fillEmptyColumns;
	private List <String>				 sheetNames		= new ArrayList <> ();

	// === Step 3 ===
	private List <SheetValidationResult> sheetsToImport	= new ArrayList <> ();

	// === Step 4 ===
	private List <ColumnDefinition>		 columns		= new ArrayList <> ();
	private Map <Integer, ColumnType>	 columnTypesByIndex;

	// === Step 5 ===
	private DbType						 dbType;
	private String						 dbHost;
	private int							 dbPort;
	private String						 dbName;
	private String						 dbUser;
	private String						 dbPassword;
	private boolean						 createDbIfMissing;
	private String						 tableName;

	// === Getters / Setters ===

	public File getDataDirectory ()
	{
		return dataDirectory;
	}

	public void setDataDirectory (File dataDirectory)
	{
		this.dataDirectory = dataDirectory;
	}

	public List <File> getSelectedFiles ()
	{
		return selectedFiles;
	}

	public void setSelectedFiles (List <File> selectedFiles)
	{
		this.selectedFiles = selectedFiles;
	}

	public Integer getHeaderRow ()
	{
		return headerRow;
	}

	public void setHeaderRow (Integer headerRow)
	{
		this.headerRow = headerRow;
	}

	public Integer getDataStartRow ()
	{
		return dataStartRow;
	}

	public void setDataStartRow (Integer dataStartRow)
	{
		this.dataStartRow = dataStartRow;
	}

	public boolean isAutoIncrement ()
	{
		return autoIncrement;
	}

	public void setAutoIncrement (boolean autoIncrement)
	{
		this.autoIncrement = autoIncrement;
	}

	public List <String> getSheetNames ()
	{
		return sheetNames;
	}

	public void setSheetNames (List <String> sheetNames)
	{
		this.sheetNames = sheetNames;
	}

	public List <SheetValidationResult> getSheetsToImport ()
	{
		return sheetsToImport;
	}

	public void setSheetsToImport (List <SheetValidationResult> sheetsToImport)
	{
		this.sheetsToImport = sheetsToImport;
	}

	public List <ColumnDefinition> getColumns ()
	{
		return columns;
	}

	public void setColumns (List <ColumnDefinition> columns)
	{
		this.columns = columns;
	}

	public Map <Integer, ColumnType> getColumnTypesByIndex ()
	{
		return columnTypesByIndex;
	}

	public void setColumnTypesByIndex (Map <Integer, ColumnType> columnTypesByIndex)
	{
		this.columnTypesByIndex = columnTypesByIndex;
	}

	public DbType getDbType ()
	{
		return dbType;
	}

	public void setDbType (DbType dbType)
	{
		this.dbType = dbType;
	}

	public String getDbHost ()
	{
		return dbHost;
	}

	public void setDbHost (String dbHost)
	{
		this.dbHost = dbHost;
	}

	public int getDbPort ()
	{
		return dbPort;
	}

	public void setDbPort (int dbPort)
	{
		this.dbPort = dbPort;
	}

	public String getDbName ()
	{
		return dbName;
	}

	public void setDbName (String dbName)
	{
		this.dbName = dbName;
	}

	public String getDbUser ()
	{
		return dbUser;
	}

	public void setDbUser (String dbUser)
	{
		this.dbUser = dbUser;
	}

	public String getDbPassword ()
	{
		return dbPassword;
	}

	public void setDbPassword (String dbPassword)
	{
		this.dbPassword = dbPassword;
	}

	public boolean isCreateDbIfMissing ()
	{
		return createDbIfMissing;
	}

	public void setCreateDbIfMissing (boolean createDbIfMissing)
	{
		this.createDbIfMissing = createDbIfMissing;
	}

	public void setTableName (String tableName)
	{
		this.tableName = tableName;
	}

	public String getTableName ()
	{
		return tableName;
	}

	public void setFillEmptyColumns (List <Boolean> fillColumnList)
	{
		this.fillEmptyColumns = fillColumnList;

	}
}
