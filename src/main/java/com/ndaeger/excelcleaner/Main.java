package com.ndaeger.excelcleaner;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class Main {
    
    private static final DateFormat dformat = new SimpleDateFormat("MM/dd/yyyy hh:mm");

    public static void main(String[] args) throws IOException, ClassNotFoundException, SQLException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        /*

        Before we can do any conversions, we need to convert the access database to a usable excel workbook.

        Keeping it in an excel format is nice for a few different reasons:
            1. I don't have to rewrite the entire program again (it was written to take in an excel file and spit new ones out)
            2. It has to export to excel anyway, and working between 2 excel files is much easier than SQL and excel

        With that being said, first, we need to find the data file for the Access database. Once we get the instance of
        that file, we need to make the database connection, which we can do with the default SQL API in Java.

         */
        File file = new File("CMI_DATA.accdb");
        if (!file.exists()) throw new RuntimeException("CMI_DATA.accdb was not found in the project directory. Please put it in there.");

        //Creating the URL to the access file. jdbc:ucanaccess:// needs to be in front of the path to the file to ensure
        //it makes the connection with the correct driver.
        String url = "jdbc:ucanaccess://" + file.getPath() + ";";
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

        //Making the connection. This closes automatically
        try (Connection connection = DriverManager.getConnection(url)) {

            /*

            The first leg of this conversion from Access to Excel is porting the tbl_PC_Info.
            We just need to run a SELECT query on the table and iterate over all the data, and then copy it over to the excel workbook I have open.

             */
            System.out.println("Converting tbl_PC_Info...");
            Statement statement = connection.createStatement();
            String query = "SELECT PC_Num, Type, PC_Model, Asset_Num, Serial_Num, PO_Num, Username, Location, Prev_User, Net_Port, Comments, Special_Software, Warranty_Expiration FROM tbl_PC_Info";
            ResultSet rs = statement.executeQuery(query);

            //Creating a "PC_Info" sheet in the workbook.
            XSSFSheet sheet = workbook.createSheet("PC_Info");
            XSSFRow row = sheet.createRow(0);

            //Creating the column titles
            row.createCell(0).setCellValue("PC_Num");
            row.createCell(1).setCellValue("PC_Model");
            row.createCell(2).setCellValue("Asset_Num");
            row.createCell(3).setCellValue("Serial_Num");
            row.createCell(4).setCellValue("PO_Num");
            row.createCell(5).setCellValue("Username");
            row.createCell(6).setCellValue("Type");
            row.createCell(7).setCellValue("Location");
            row.createCell(8).setCellValue("Prev_User");
            row.createCell(9).setCellValue("Net_Port");
            row.createCell(10).setCellValue("Comments");
            row.createCell(11).setCellValue("Special_Software");
            row.createCell(12).setCellValue("Warranty_Expiration");

            int currentRow = 1;

            while (rs.next()) {
                //Adding everything to the worksheet after the title row
                row = sheet.createRow(currentRow);
                row.createCell(0).setCellValue(rs.getString("PC_Num"));
                row.createCell(1).setCellValue(rs.getString("PC_Model"));
                row.createCell(2).setCellValue(rs.getString("Asset_Num"));
                row.createCell(3).setCellValue(rs.getString("Serial_Num"));
                row.createCell(4).setCellValue(rs.getString("PO_Num"));
                row.createCell(5).setCellValue(rs.getString("Username"));
                row.createCell(6).setCellValue(rs.getString("Type"));
                row.createCell(7).setCellValue(rs.getString("Location"));
                row.createCell(8).setCellValue(rs.getString("Prev_User"));
                row.createCell(9).setCellValue(rs.getString("Net_Port"));
                row.createCell(10).setCellValue(rs.getString("Comments"));
                row.createCell(11).setCellValue(rs.getString("Special_Software"));
                row.createCell(12).setCellValue(rs.getString("Warranty_Expiration"));
                currentRow++;
            }
            System.out.println("Done!");




            /*
            After the PC database, we need to get the table "tbl_CMIC_Hardware" and put that into the workbook.
             */
            System.out.println("Converting tbl_CMIC_Hardware...");
            statement = connection.createStatement();
            query = "SELECT Netbios_Name0, OSType, OSVersion, Memory, Manufacturer0, Model0, CPU, DriveName, DriveSize, DriveFree, SerialNumber0, LastHWScan, LastHeartbeat, LastLoggedOnUser FROM tbl_CMIC_Hardware";
            rs = statement.executeQuery(query);

            //Creating a "Hardware" sheet in the workbook.
            sheet = workbook.createSheet("Hardware");
            row = sheet.createRow(0);

            //Creating the column titles
            row.createCell(0).setCellValue("Netbios_Name0");
            row.createCell(1).setCellValue("OSType");
            row.createCell(2).setCellValue("OSVersion");
            row.createCell(3).setCellValue("Memory");
            row.createCell(4).setCellValue("Manufacturer0");
            row.createCell(5).setCellValue("Model0");
            row.createCell(6).setCellValue("CPU");
            row.createCell(7).setCellValue("DriveName");
            row.createCell(8).setCellValue("DriveSize");
            row.createCell(9).setCellValue("DriveFree");
            row.createCell(10).setCellValue("SerialNumber0");
            row.createCell(11).setCellValue("LastHWScan");
            row.createCell(12).setCellValue("LastHeartbeat");
            row.createCell(13).setCellValue("LastLoggedOnUser");

            currentRow = 1;

            while (rs.next()) {
                //Adding everything to the worksheet after the title row
                row = sheet.createRow(currentRow);
                row.createCell(0).setCellValue(rs.getString("Netbios_Name0"));
                row.createCell(1).setCellValue(rs.getString("OSType"));
                row.createCell(2).setCellValue(rs.getString("OSVersion"));
                row.createCell(3).setCellValue(rs.getString("Memory"));
                row.createCell(4).setCellValue(rs.getString("Manufacturer0"));
                row.createCell(5).setCellValue(rs.getString("Model0"));
                row.createCell(6).setCellValue(rs.getString("CPU"));
                row.createCell(7).setCellValue(rs.getString("DriveName"));
                row.createCell(8).setCellValue(rs.getString("DriveSize"));
                row.createCell(9).setCellValue(rs.getString("DriveFree"));
                row.createCell(10).setCellValue(rs.getString("SerialNumber0"));
                row.createCell(11).setCellValue(rs.getString("LastHWScan"));
                row.createCell(12).setCellValue(rs.getString("LastHeartbeat"));
                row.createCell(12).setCellValue(rs.getString("LastLoggedOnUser"));
                currentRow++;
            }
            System.out.println("Done!");




            /*
            Now we do the "tbl_DISTINCT_CMIC_Hardware" conversion
             */
            System.out.println("Converting tbl_DISTINCT_CMIC_Hardware...");
            statement = connection.createStatement();
            query = "SELECT Netbios_Name0, OSType, OSVersion, Memory, Manufacturer0, Model0, CPU, LastHWScan, LastHeartbeat, LastLoggedOnUser FROM tbl_DISTINCT_CMIC_Hardware";
            rs = statement.executeQuery(query);

            //Creating a "Distinct_Hardware" sheet in the workbook.
            sheet = workbook.createSheet("Distinct_Hardware");
            row = sheet.createRow(0);

            row.createCell(0).setCellValue("Netbios_Name0");
            row.createCell(1).setCellValue("OSType");
            row.createCell(2).setCellValue("OSVersion");
            row.createCell(3).setCellValue("Memory");
            row.createCell(4).setCellValue("Manufacturer0");
            row.createCell(5).setCellValue("Model0");
            row.createCell(6).setCellValue("CPU");
            row.createCell(7).setCellValue("LastHWScan");
            row.createCell(8).setCellValue("LastHeartbeat");
            row.createCell(9).setCellValue("LastLoggedOnUser");

            currentRow = 1;

            while (rs.next()) {
                //Adding everything after the title row
                row = sheet.createRow(currentRow);
                row.createCell(0).setCellValue(rs.getString("Netbios_Name0"));
                row.createCell(1).setCellValue(rs.getString("OSType"));
                row.createCell(2).setCellValue(rs.getString("OSVersion"));
                row.createCell(3).setCellValue(rs.getString("Memory"));
                row.createCell(4).setCellValue(rs.getString("Manufacturer0"));
                row.createCell(5).setCellValue(rs.getString("Model0"));
                row.createCell(6).setCellValue(rs.getString("CPU"));
                row.createCell(7).setCellValue(rs.getString("LastHWScan"));
                row.createCell(8).setCellValue(rs.getString("LastHeartbeat"));
                row.createCell(9).setCellValue(rs.getString("LastLoggedOnUser"));
                currentRow++;
            }
            System.out.println("Done!");




            /*
            Now we do the "tbl_AD_Info" conversion
             */
            System.out.println("Converting tbl_AD_Info...");
            statement = connection.createStatement();
            query = "SELECT Username, GivenName, FirstName, LastName, DisplayName, Region, PhoneNumber, Email, ProfilePath, Title, Department, Manager, IPPhone, Info FROM tbl_AD_Info";
            rs = statement.executeQuery(query);

            //Creating a "Users" sheet in the workbook.
            sheet = workbook.createSheet("Users");
            row = sheet.createRow(0);

            row.createCell(0).setCellValue("Username");
            row.createCell(1).setCellValue("GivenName");
            row.createCell(2).setCellValue("FirstName");
            row.createCell(3).setCellValue("LastName");
            row.createCell(4).setCellValue("DisplayName");
            row.createCell(5).setCellValue("Region");
            row.createCell(6).setCellValue("PhoneNumber");
            row.createCell(7).setCellValue("Email");
            row.createCell(8).setCellValue("ProfilePath");
            row.createCell(9).setCellValue("Title");
            row.createCell(10).setCellValue("Department");
            row.createCell(11).setCellValue("Manager");
            row.createCell(12).setCellValue("IPPhone");
            row.createCell(13).setCellValue("Info");

            currentRow = 1;

            while (rs.next()) {
                //Adding everything after the title row
                row = sheet.createRow(currentRow);
                row.createCell(0).setCellValue(rs.getString("Username"));
                row.createCell(1).setCellValue(rs.getString("GivenName"));
                row.createCell(2).setCellValue(rs.getString("FirstName"));
                row.createCell(3).setCellValue(rs.getString("LastName"));
                row.createCell(4).setCellValue(rs.getString("DisplayName"));
                row.createCell(5).setCellValue(rs.getString("Region"));
                row.createCell(6).setCellValue(rs.getString("PhoneNumber"));
                row.createCell(7).setCellValue(rs.getString("Email"));
                row.createCell(8).setCellValue(rs.getString("ProfilePath"));
                row.createCell(9).setCellValue(rs.getString("Title"));
                row.createCell(10).setCellValue(rs.getString("Department"));
                row.createCell(11).setCellValue(rs.getString("Manager"));
                row.createCell(12).setCellValue(rs.getString("IPPhone"));
                row.createCell(13).setCellValue(rs.getString("Info"));
                currentRow++;
            }
            System.out.println("Done!");




            /*
            Now we do the "tbl_Additional_Hardware" conversion
             */
            System.out.println("Converting tbl_Additional_Hardware...");
            statement = connection.createStatement();
            query = "SELECT PC_Num, Model, Serial_Num, Asset_Num, FAS_Date_Added, Comments, CellPhoneInfo FROM tbl_Additional_Hardware";
            rs = statement.executeQuery(query);

            //Creating an "Additional_Hardware" sheet in the workbook.
            sheet = workbook.createSheet("Additional_Hardware");
            row = sheet.createRow(0);

            row.createCell(0).setCellValue("PC_Num");
            row.createCell(1).setCellValue("Model");
            row.createCell(2).setCellValue("Serial_Num");
            row.createCell(3).setCellValue("Asset_Num");
            row.createCell(4).setCellValue("FAS_Date_Added");
            row.createCell(5).setCellValue("Comments");
            row.createCell(6).setCellValue("CellPhoneInfo");

            currentRow = 1;

            while (rs.next()) {
                //Adding everything after the title row
                row = sheet.createRow(currentRow);
                row.createCell(0).setCellValue(rs.getString("PC_Num"));
                row.createCell(1).setCellValue(rs.getString("Model"));
                row.createCell(2).setCellValue(rs.getString("Serial_Num"));
                row.createCell(3).setCellValue(rs.getString("Asset_Num"));
                row.createCell(4).setCellValue(rs.getString("FAS_Date_Added"));
                row.createCell(5).setCellValue(rs.getString("Comments"));
                row.createCell(6).setCellValue(rs.getString("CellPhoneInfo"));
                currentRow++;
            }
            System.out.println("Done!");
        }

        System.out.println("Cleaning data...");
        formatAdditionalHardware(workbook);
        formatPCInfo(workbook);
        formatRoomList(workbook);
        System.out.println("Done!");
    }

    /**
     * Pulls the room information from the Access data and saves it to a file.
     * @param parent The parent excel workbook with the raw Access data
     */
    private static void formatRoomList(XSSFWorkbook parent) throws IOException {
        XSSFSheet users = parent.getSheet("Users");
        XSSFWorkbook roomWorkbook = new XSSFWorkbook();
        XSSFSheet rooms = roomWorkbook.createSheet("Rooms");
        XSSFRow titleRow = rooms.createRow(0);
        titleRow.createCell(0).setCellValue("Region");
        titleRow.createCell(1).setCellValue("Room");
        users.forEach(row -> {
            if (row.getCell(2).getStringCellValue().toLowerCase().contains("room")) {
                Row rw = rooms.createRow(rooms.getLastRowNum() + 1);
                rw.createCell(0).setCellValue(row.getCell(5).getStringCellValue());
                rw.createCell(1).setCellValue(row.getCell(3).getStringCellValue());
            }
        });
    
        FileOutputStream roomFile = new FileOutputStream("Room_Database.xlsx");
        roomWorkbook.write(roomFile);
        roomFile.close();
        System.out.println("Saved Rooms!");
        
    }

    /**
     * Pulls the PC information from the Access data and saves it to a file.
     * @param parent The parent excel workbook with the raw Access data.
     */
    private static void formatPCInfo(XSSFWorkbook parent) throws IOException {
        XSSFSheet pcInfo = parent.getSheet("PC_Info");
        XSSFWorkbook pcWorkbook = new XSSFWorkbook();
        XSSFSheet pcSheet = pcWorkbook.createSheet("PC_Sheet");
        XSSFRow titleRow = pcSheet.createRow(0);
        titleRow.createCell(0).setCellValue("PC_Num");      //PC_Num
        titleRow.createCell(1).setCellValue("Owner_Type");  //Determine from the username
        titleRow.createCell(2).setCellValue("Owned_By");    //Should just be Username
        titleRow.createCell(3).setCellValue("Location");    //Location
        titleRow.createCell(4).setCellValue("Model");       //PC_Model
        titleRow.createCell(5).setCellValue("CPU");         //Unable
        titleRow.createCell(6).setCellValue("RAM");         //Check comments or Prev_User
        titleRow.createCell(7).setCellValue("OS");          //Unable
        titleRow.createCell(8).setCellValue("Asset_Num");   //Asset_Num
        titleRow.createCell(9).setCellValue("Serial_Num");  //Serial_Num
        titleRow.createCell(10).setCellValue("PO_Num");     //PO_Num
        titleRow.createCell(11).setCellValue("Net_Port");   //Net_Port
        titleRow.createCell(12).setCellValue("Previous_User");//Prev_User
        titleRow.createCell(13).setCellValue("Special_Software");//Special_Software
        titleRow.createCell(14).setCellValue("Comment");    //Comments
        titleRow.createCell(15).setCellValue("Warranty_Expiration");//Warranty_Expiration
        titleRow.createCell(16).setCellValue("Work_From_Home");//Warranty_Expiration
    
        int rows = pcInfo.getLastRowNum();
        
        pcInfo.forEach(row -> {
            
            String pcNum = "";
            String ownerType = "";
            String ownedBy = "";
            String location = "";
            String model = "";
            String cpu = ""; //Note, we are no longer using access to grab the os/cpu/and ram
            String ram = "";
            String os = "";
            String assetNum = "";
            String serialNum = "";
            String orderNum = "";
            String networkPort = "";
            String previousUser = "";
            String specialSoftware = "";
            String comment = "";
            String warrantyExpiration = "";
            String workFromHome = "false";

            XSSFRow newRow = pcSheet.createRow(pcSheet.getLastRowNum() + 1);
        
            pcNum = row.getCell(0).getStringCellValue();
        
            String typeCell = row.getCell(6).getStringCellValue();
            if (typeCell.toLowerCase().contains("stock")) {
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
                ownedBy = "Stock";
                ownerType = "Stock";
            } else if (typeCell.toLowerCase().contains("reserve")) {
                ownerType = "Reserved";
                ownedBy = "";
                location = "";
            } else if (typeCell.toLowerCase().contains("room")) {
                ownedBy = normalizeUsernameString(row.getCell(5).getStringCellValue());
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                ownerType = "Room";
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            } else if (typeCell.toLowerCase().contains("office")) {
                ownerType = "User";
                ownedBy = row.getCell(5).getStringCellValue().toLowerCase();
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            } else if (typeCell.toLowerCase().contains("pool")) {
                ownerType = "Pool PC";
                ownedBy = row.getCell(5).getStringCellValue();
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            } else if (typeCell.toLowerCase().contains("field")) {
                ownerType = "Field";
                ownedBy = row.getCell(5).getStringCellValue();
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            } else if (typeCell.toLowerCase().contains("teleco")) {
                ownerType = "Telecommuter";
                ownedBy = row.getCell(5).getStringCellValue();
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            } else {
                ownerType = "Other";
                ownedBy = row.getCell(5).getStringCellValue();
                if (ownedBy.toLowerCase().contains("wfh")) {
                    ownedBy = ownedBy.replace("wfh", "");
                    workFromHome = "true";
                }
                location = getLocationFromString(row.getCell(7).getStringCellValue().toLowerCase());
            }
            
            model = row.getCell(1).getStringCellValue();

            //cpu
            //ram = getRam(row.getCell(8).getStringCellValue());
            //os
            assetNum = row.getCell(2).getStringCellValue();
            serialNum = row.getCell(3).getStringCellValue();
            orderNum = row.getCell(4).getStringCellValue();
            networkPort = row.getCell(9).getStringCellValue();
            previousUser = row.getCell(8).getStringCellValue();
            specialSoftware = row.getCell(11).getStringCellValue();
            comment = row.getCell(10).getStringCellValue();
            //warrantyExpiration = row.getCell(12).getStringCellValue();
            if (row.getCell(12).getCellType() == CellType.NUMERIC) {
                Date value = row.getCell(12).getDateCellValue();
                if (value != null) {
                    warrantyExpiration = dformat.format(value);
                }
            } else warrantyExpiration = row.getCell(12).getStringCellValue();
            
            newRow.createCell(0).setCellValue(pcNum);
            newRow.createCell(1).setCellValue(ownerType);
            newRow.createCell(2).setCellValue(ownedBy);
            newRow.createCell(3).setCellValue(location);
            newRow.createCell(4).setCellValue(model);
            newRow.createCell(5).setCellValue(cpu);
            newRow.createCell(6).setCellValue(ram);
            newRow.createCell(7).setCellValue(os);
            newRow.createCell(8).setCellValue(assetNum);
            newRow.createCell(9).setCellValue(serialNum);
            newRow.createCell(10).setCellValue(orderNum);
            newRow.createCell(11).setCellValue(networkPort);
            newRow.createCell(12).setCellValue(previousUser);
            newRow.createCell(13).setCellValue(specialSoftware);
            newRow.createCell(14).setCellValue(comment);
            newRow.createCell(15).setCellValue(warrantyExpiration);
            newRow.createCell(16).setCellValue(workFromHome);
    
            if (row.getRowNum() % 8 == 0 || row.getRowNum() == rows) System.out.printf("%.2f%% complete.\n", (row.getRowNum() / (double)rows) * 100);
            
        });
    
        FileOutputStream pcFile = new FileOutputStream("PC_Database.xlsx");
        pcWorkbook.write(pcFile);
        pcFile.close();
        System.out.println("Saved PCs!");
        
    }

    /**
     * This tries to get CPU, OS, and RAM of a specific PC number.
     * @param hardwareSheet The base hardware sheet
     * @param distinctHardwareSheet The base distinct hardware sheet
     * @param pcNumber The PC number to get the hardware of
     * @return An array of 3 strings, first being the OS, second being the RAM, and the last being the CPU
     */
    private static String[] getHardware(XSSFSheet hardwareSheet, XSSFSheet distinctHardwareSheet, String pcNumber) {
        String[] hardware = new String[3];
        hardware[0] = "";
        hardware[1] = "";
        hardware[2] = "";
        if (pcNumber.isEmpty()) return hardware;
        Optional<Row> optionalRow = StreamSupport.stream(hardwareSheet.spliterator(), true).filter(row -> row.getCell(0).getStringCellValue().equalsIgnoreCase(pcNumber)).findFirst();
        Row row;
        if (!optionalRow.isPresent()) {
            optionalRow = StreamSupport.stream(distinctHardwareSheet.spliterator(), true).filter(row1 -> row1.getCell(0).getStringCellValue().equalsIgnoreCase(pcNumber)).findFirst();
            if (!optionalRow.isPresent()) return hardware;
        }
        row = optionalRow.get();
        hardware[0] = row.getCell(1).getStringCellValue();
        hardware[1] = row.getCell(3).getStringCellValue();
        hardware[2] = row.getCell(6).getStringCellValue();
        return hardware;
    }

    /**
     * Normalizes a persons username. The reason this exists is because many usernames also contained the location of
     * where they were, even though that should have been in the Location column. This pulls the unneeded information out.
     * @param input The username to normalize
     * @return The normalized username.
     */
    private static String normalizeUsernameString(String input) {
        input = input.replaceAll("COTO\\W+|COTO\\W", "");
        input = input.replaceAll("SERO\\W+|SERO\\W", "");
        input = input.replaceAll("NERO\\W+|NERO\\W", "");
        input = input.replaceAll("SWRO\\W+|SWRO\\W", "");
        input = input.replaceAll("CONF\\W+|CONF\\W", "Conference ");
        return input.trim();
    }

    /**
     * Cleans and formats the additional hardware table. This generates 3 different spreadsheets, the first being the
     * Hardware Database which contains the hardware users own. The next spreadsheet created by this is the phone database,
     * that contains all phones that central has. The last spreadsheet created is the Model Database, which just maps hardware
     * types to models of hardware.
     *
     * @param parent The raw access data in spreadsheet format.
     */
    private static void formatAdditionalHardware(XSSFWorkbook parent) throws IOException {
        List<String> models = new ArrayList<>();
        List<String> assetNums = new ArrayList<>();
        XSSFSheet additionalHardware = parent.getSheet("Additional_Hardware");
        XSSFSheet pcInfo = parent.getSheet("PC_Info");
        XSSFSheet userInfo = parent.getSheet("Users");
        
        XSSFWorkbook hardwareModels = new XSSFWorkbook();
        XSSFSheet modelSheet = hardwareModels.createSheet("Models");
        XSSFRow modelTitle = modelSheet.createRow(0);
        modelTitle.createCell(0).setCellValue("Hardware_Type");
        modelTitle.createCell(1).setCellValue("Model");
        
        XSSFWorkbook phoneWorkbook = new XSSFWorkbook();
        XSSFSheet phoneSheet = phoneWorkbook.createSheet("Phones");
        XSSFRow phoneTitle = phoneSheet.createRow(0);
        phoneTitle.createCell(0).setCellValue("Owner_Type");
        phoneTitle.createCell(1).setCellValue("Owned_By");
        phoneTitle.createCell(2).setCellValue("Location");
        phoneTitle.createCell(3).setCellValue("Carrier_Info");
        phoneTitle.createCell(4).setCellValue("Cell_Num");
        phoneTitle.createCell(5).setCellValue("Asset_Num");
        phoneTitle.createCell(6).setCellValue("FAS_Date_Added");
        phoneTitle.createCell(7).setCellValue("Model");
        phoneTitle.createCell(8).setCellValue("Comment");
        phoneTitle.createCell(9).setCellValue("Work_From_Home");

        XSSFWorkbook hardwareWorkbook = new XSSFWorkbook();
        XSSFSheet hardwareSheet = hardwareWorkbook.createSheet("Hardware");
        XSSFRow hardwareTitle = hardwareSheet.createRow(0);
        hardwareTitle.createCell(0).setCellValue("Owner_Type");
        hardwareTitle.createCell(1).setCellValue("Owned_By");
        hardwareTitle.createCell(2).setCellValue("Location");
        hardwareTitle.createCell(3).setCellValue("Hardware_Type");
        hardwareTitle.createCell(4).setCellValue("Serial_Num");
        hardwareTitle.createCell(5).setCellValue("Asset_Num");
        hardwareTitle.createCell(6).setCellValue("FAS_Date_Added");
        hardwareTitle.createCell(7).setCellValue("Model");
        hardwareTitle.createCell(8).setCellValue("Comment");
        hardwareTitle.createCell(8).setCellValue("Work_From_Home");
        
        additionalHardware.forEach(row -> {
            //The "CellPhoneInfo" column is always full when an item is a cell phone. (Note: few mismatches)
            String cellPhoneInfoCheck = row.getCell(6).getStringCellValue();
            //the model will always have the word "Phone" in it when it is a cell phone
            String modelCheck = row.getCell(1).getStringCellValue();
    
            //0 = stock
            //1 = department
            //2 = person
            //3 = room
            String ownerType;
    
            //empty if ownerType is 0
            //department if ownerType is 1
            //person if ownerType is 2
            String ownedBy;
    
            //Always available. Will need to be manually added for "ACTIVE" types (department)
            String location;
    
            //Asset Number only available on about a third of things
            String assetNum;
    
            //Same as asset number
            String dateAdded = "";
    
            //Model should be available on everything.
            String model;
    
            //Comment will help determine the location and ownedBy characteristics
            String comment;

            String workFromHome = "false";

            //PCNum is the same for both phones and hardware
            //covers owner_type, owned_by, and location.
            String pcNumCell = row.getCell(0).getStringCellValue();
            if (pcNumCell.toLowerCase().contains("stock")) {
                location = getLocationFromString(pcNumCell.toLowerCase());
                ownedBy = "Stock";
                ownerType = "Stock";
            } else if (pcNumCell.toLowerCase().contains("active")) {
                ownerType = "Department";
                location = "DEPARTMENT_REGION";
                ownedBy = "DEPARTMENT";
            } else {
                ownedBy = lookupUsername(pcInfo, pcNumCell).toLowerCase();
                if (ownedBy.toLowerCase().contains("wfh")) workFromHome = "true";
                if (ownedBy.toLowerCase().contains("room")) ownerType = "Room";
                else ownerType = "User";
                location = lookupLocation(userInfo, ownedBy).toUpperCase();
            }
            if ((ownerType.equalsIgnoreCase("Room") || ownerType.equalsIgnoreCase("User")) && ownedBy.isEmpty() && location.isEmpty()) ownerType = "";
    
            //Comments are the same for phones and other hardware
            comment = row.getCell(5).getStringCellValue();
            //Asset Number is the same for both
            assetNum = row.getCell(3).getStringCellValue();
            //Date added is the same for both
            if (row.getCell(4).getCellType().equals(CellType.NUMERIC)) {
                Date value = row.getCell(4).getDateCellValue();
                if (value != null) {
                    dateAdded = dformat.format(value);
                }
            }
            if (row.getRowNum() == 0) return;
            if (!cellPhoneInfoCheck.isEmpty() && modelCheck.toLowerCase().contains("phone")) {
                
                //Create a new row after the most recent one
                XSSFRow newRow = phoneSheet.createRow(phoneSheet.getLastRowNum() + 1);
    
                String carrierInfo = row.getCell(1).getStringCellValue();
                String cellNum = row.getCell(2).getStringCellValue();
                model = row.getCell(6).getStringCellValue();
                
                newRow.createCell(0).setCellValue(ownerType);
                newRow.createCell(1).setCellValue(ownedBy);
                newRow.createCell(2).setCellValue(location);
                newRow.createCell(3).setCellValue(carrierInfo);
                newRow.createCell(4).setCellValue(cellNum);
                newRow.createCell(5).setCellValue(assetNum);
                newRow.createCell(6).setCellValue(dateAdded);
                newRow.createCell(7).setCellValue(model);
                newRow.createCell(8).setCellValue(comment);
                newRow.createCell(9).setCellValue(workFromHome);
            } else {
                //Create a new row after the most recent one
                if (assetNum.toLowerCase().contains("nofas") || assetNum.toLowerCase().contains("n/a") || assetNum.toLowerCase().contains("noasset") || assetNum.toLowerCase().contains("not on fas")) assetNum = "";
                if (assetNums.contains(assetNum) && !assetNum.isEmpty()) return;
                else assetNums.add(assetNum);
                XSSFRow newRow = hardwareSheet.createRow(hardwareSheet.getLastRowNum() + 1);
                
                model = row.getCell(1).getStringCellValue();
                String serialNum = row.getCell(2).getStringCellValue();
                String hardwareType = getHardwareTypeFromModel(model);
                
                newRow.createCell(0).setCellValue(ownerType);
                newRow.createCell(1).setCellValue(ownedBy);
                newRow.createCell(2).setCellValue(location);
                newRow.createCell(3).setCellValue(hardwareType);
                newRow.createCell(4).setCellValue(serialNum);
                newRow.createCell(5).setCellValue(assetNum);
                newRow.createCell(6).setCellValue(dateAdded);
                newRow.createCell(7).setCellValue(model);
                newRow.createCell(8).setCellValue(comment);
                newRow.createCell(9).setCellValue(workFromHome);
                
                if (!models.contains(model.toLowerCase())) {
                    models.add(model.toLowerCase());
                    XSSFRow modelRow = modelSheet.createRow(modelSheet.getLastRowNum() + 1);
                    modelRow.createCell(0).setCellValue(hardwareType);
                    modelRow.createCell(1).setCellValue(model);
                }
            }
        });
        
        FileOutputStream hardwareFile = new FileOutputStream("Hardware_Database.xlsx");
        FileOutputStream phoneFile = new FileOutputStream("Phone_Database.xlsx");
        FileOutputStream modelFile = new FileOutputStream("Model_Database.xlsx");
        hardwareWorkbook.write(hardwareFile);
        phoneWorkbook.write(phoneFile);
        hardwareModels.write(modelFile);
        hardwareFile.close();
        hardwareWorkbook.close();
        System.out.println("Saved Hardware!");
        phoneFile.close();
        phoneWorkbook.close();
        System.out.println("Saved Phones!");
        modelFile.close();
        hardwareModels.close();
        System.out.println("Saved Models!");
    }

    /**
     * Utility method to get a hardware type from a model of some type of hardware.
     * @param model The model of hardware to get the hardware type from.
     * @return The hardware type of the model
     */
    private static String getHardwareTypeFromModel(String model) {
        model = model.toLowerCase();
        if (model.contains("vizio") || model.contains("tv")) return "Television";
        else if (model.contains("dell") || model.contains("monitor")) return "Monitor";
        else if (model.contains("headset")) return "Headset";
        else if (model.contains("camera")) return "Camera";
        else if (model.contains("ipad")) return "Tablet";
        else return "Other";
    }

    /**
     * Gets a location from a string.
     * @param input The input to get a location from. NOTE: This needs to be lowercase before putting it into this method
     * @return The location from the input string. Defaulting to "VW" if there was no valid location found.
     */
    private static String getLocationFromString(String input) {
        if (input.contains("swro")) return "SWRO";
        if (input.contains("coto") || input.contains("columbus")) return "COTO";
        if (input.contains("sero")) return "SERO";
        if (input.contains("nero")) return "NERO";
        if (input.contains("field")) return "FIELD";
        else return "VW";
    }

    /**
     * This will look up a username from a PC number and return the username.
     * @param pcInfoSheet The PCInfo table from the access database
     * @param pcNumber The PC number of the PC to get the username from.
     * @return Either an empty string if no username was found, or the username of the user who owns the PC.
     */
    private static String lookupUsername(XSSFSheet pcInfoSheet, String pcNumber) {
        List<Row> results = StreamSupport.stream(pcInfoSheet.spliterator(), true).filter(row ->
            !row.getCell(0).getStringCellValue().isEmpty() && row.getCell(0).getStringCellValue().equalsIgnoreCase(pcNumber)).collect(Collectors.toList());
        if (results.isEmpty()) return "";
        if (results.size() > 1) {
            results.forEach(row -> System.out.print(row.getRowNum() + ", "));
            System.out.println();
            System.out.println("More than one result in table.");
        }
        return results.get(0).getCell(5).getStringCellValue();
    }

    /**
     * Looks up the location of a user.
     * @param userSheet The active directory table from the access database.
     * @param username The username of the user to get the location of.
     * @return The location of the user or an empty string if no valid location was found.
     */
    private static String lookupLocation(XSSFSheet userSheet, String username) {
        List<Row> results = StreamSupport.stream(userSheet.spliterator(), true).filter(row ->
            !row.getCell(0).getStringCellValue().isEmpty() && row.getCell(0).getStringCellValue().equalsIgnoreCase(username.replaceAll(" ", ""))).collect(Collectors.toList());
        if (results.isEmpty()) return "";
        if (results.size() > 1) {
            results.forEach(row -> System.out.print(row.getRowNum() + ", "));
            System.out.println();
            System.out.println("More than one result in table.");
        }
        return results.get(0).getCell(5).getStringCellValue();
    }
    
}
