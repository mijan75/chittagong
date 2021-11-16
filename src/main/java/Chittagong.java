import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import java.util.function.Supplier;


public class Chittagong {

    public Chittagong() {
        readDoc();
//        writeExcel();
    }

    private void writeExcel() {
        String pathname = "C:\\Users\\User\\Desktop\\Data Templet and Coding_Update.xlsx";
        File f = new File(pathname);

//        XSSFWorkbook xssfWorkbook = null;
//        try {
//            xssfWorkbook = new XSSFWorkbook(f);
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
//        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
//        for(int i=1; i<xssfSheet.getPhysicalNumberOfRows(); i++){
//            XSSFRow row = xssfSheet.getRow(i);
//            System.out.println(row.getRowNum());
//        }

        try {
            FileInputStream inputStream = new FileInputStream(f);
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(1);

            Object[][] bookData = {
                    {"The Passionate Programmer", "Chad Fowler", 16},
                    {"Software Craftmanship", "Pete McBreen", 26},
                    {"The Art of Agile Development", "James Shore", 32},
                    {"Continuous Delivery", "Jez Humble", 41},
            };

            int rowCount = sheet.getLastRowNum();

            for (Object[] aBook : bookData) {
                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);

                for (Object field : aBook) {
                    cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }

            }

            inputStream.close();

            FileOutputStream outputStream = new FileOutputStream(pathname);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Chittagong chittagong = new Chittagong();
    }

    public void readDoc() {
        File inputFile = null;
        try {
            inputFile = new File("C:\\Users\\User\\Desktop\\Bakta.Fulbariya.Layla.docx");
            String pathname = "C:\\Users\\User\\Desktop\\Data Templet and Coding_Update.xlsx";
            File f = new File(pathname);
            FileInputStream fis = new FileInputStream(inputFile.getAbsolutePath());

            XWPFDocument document = new XWPFDocument(fis);

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            List<XWPFTable> tables = document.getTables();

            // Basic information
            XWPFTable basicInformation = tables.get(1);
            List<XWPFTableRow> basicInformationRows = basicInformation.getRows();

            String name = basicInformationRows.get(0).getCell(1).getText();
            String contactNumber = basicInformationRows.get(1).getCell(1).getText();
            String gender = basicInformationRows.get(2).getCell(1).getText();
            String age = basicInformationRows.get(3).getCell(1).getText();
            String maritialStatus = basicInformationRows.get(4).getCell(1).getText();
            String village = basicInformationRows.get(5).getCell(1).getText();
            String subDistrict = basicInformationRows.get(6).getCell(1).getText();
            String district = basicInformationRows.get(7).getCell(1).getText();
            String division = basicInformationRows.get(8).getCell(1).getText();

            // Family basic information
            XWPFTable familyInformation = tables.get(2);

            List<XWPFTableRow> familyInformationRows = familyInformation.getRows();
            String familyName1 = familyInformationRows.get(1).getCell(1).getText();
            String familygender1 = familyInformationRows.get(1).getCell(2).getText();
            String familyage1 = familyInformationRows.get(1).getCell(3).getText();
            String familyEducation1 = familyInformationRows.get(1).getCell(4).getText();
            String familyOccupation1 = familyInformationRows.get(1).getCell(5).getText();

            String familyName2 = familyInformationRows.get(2).getCell(1).getText();
            String familygender2 = familyInformationRows.get(2).getCell(2).getText();
            String familyage2 = familyInformationRows.get(2).getCell(3).getText();
            String familyEducation2 = familyInformationRows.get(2).getCell(4).getText();
            String familyOccupation2 = familyInformationRows.get(2).getCell(5).getText();

            String familyName3 = familyInformationRows.get(3).getCell(1).getText();
            String familygender3 = familyInformationRows.get(3).getCell(2).getText();
            String familyage3 = familyInformationRows.get(3).getCell(3).getText();
            String familyEducation3 = familyInformationRows.get(3).getCell(4).getText();
            String familyOccupation3 = familyInformationRows.get(3).getCell(5).getText();

            String familyName4 = familyInformationRows.get(4).getCell(1).getText();
            String familygender4 = familyInformationRows.get(4).getCell(2).getText();
            String familyage4 = familyInformationRows.get(4).getCell(3).getText();
            String familyEducation4 = familyInformationRows.get(4).getCell(4).getText();
            String familyOccupation4 = familyInformationRows.get(4).getCell(5).getText();

            String familyName5 = familyInformationRows.get(5).getCell(1).getText();
            String familygender5 = familyInformationRows.get(5).getCell(2).getText();
            String familyage5 = familyInformationRows.get(5).getCell(3).getText();
            String familyEducation5 = familyInformationRows.get(5).getCell(4).getText();
            String familyOccupation5 = familyInformationRows.get(5).getCell(5).getText();

            Object[][] basicAndFamilyRowData = {
                    {
                            name, division, district, subDistrict, village, gender, age, maritialStatus, contactNumber, "",
                            familyName1, familygender1, familyage1, familyEducation1, familyOccupation1,
                            familyName2, familygender2, familyage2, familyEducation2, familyOccupation2,
                            familyName3, familygender3, familyage3, familyEducation3, familyOccupation3,
                            familyName4, familygender4, familyage4, familyEducation4, familyOccupation4,
                            familyName5, familygender5, familyage5, familyEducation5, familyOccupation5,
                    }
            };
            updateInformation(basicAndFamilyRowData, 0, f, pathname);

            //Activity
            XWPFTable activityTable3 = tables.get(3);
            XWPFTable activityTable4 = tables.get(4);

            Object[][] activityRowData = {
                    {
                            name, activityTable3.getRow(0).getCell(0).getText(), "", "", "", "", "", "",
                            activityTable4.getRow(0).getCell(0).getText()
                    }
            };
            updateInformation(activityRowData, 1, f, pathname);

            //Livelihood
            XWPFTable activityTable5 = tables.get(5);
            XWPFTable activityTable6 = tables.get(6);


            XWPFTable cultivationTable = tables.get(7);
            String agricultureAmount = cultivationTable.getRow(1).getCell(1).getText();
            String agriculturePercent = cultivationTable.getRow(1).getCell(2).getText();

            String nonAgricultureAmount = cultivationTable.getRow(2).getCell(1).getText();
            String nonAgriculturePercent = cultivationTable.getRow(2).getCell(2).getText();

            String salaryFromJobAmount = cultivationTable.getRow(3).getCell(1).getText();
            String salaryFromJobPercent = cultivationTable.getRow(3).getCell(2).getText();


            String remiteneAmount = cultivationTable.getRow(4).getCell(1).getText();
            String remitenePercent = cultivationTable.getRow(4).getCell(2).getText();

            String othersAmount = cultivationTable.getRow(5).getCell(1).getText();
            String othersPercent = cultivationTable.getRow(5).getCell(2).getText();

            XWPFTable activityTable8 = tables.get(8);

            Object[][] livelihoodRowData = {
                    {
                            name, activityTable5.getRow(0).getCell(0).getText(), "", "", "", "", "", "",
                            activityTable6.getRow(0).getCell(0).getText(), "", "", "", "", "", "", "",
                            agricultureAmount, agriculturePercent, nonAgricultureAmount, nonAgriculturePercent,
                            salaryFromJobAmount, salaryFromJobPercent, remiteneAmount, remitenePercent,
                            othersAmount, othersPercent, "", activityTable8.getRow(0).getCell(0).getText()
                    }
            };

            updateInformation(livelihoodRowData, 2, f, pathname);


            //Cultivation
            XWPFTable cultivateTable = tables.get(9);

            String cropName1 = cultivateTable.getRow(1).getCell(0).getText();
            String plantTime1 = cultivateTable.getRow(1).getCell(1).getText();
            String production1 = cultivateTable.getRow(1).getCell(2).getText();
            String yeild1 = cultivateTable.getRow(1).getCell(3).getText();
            String quality1 = cultivateTable.getRow(1).getCell(4).getText();
            String timeOfPlanting1 = cultivateTable.getRow(1).getCell(5).getText();
            String harvest1 = cultivateTable.getRow(1).getCell(6).getText();

            String cropName2 = cultivateTable.getRow(2).getCell(0).getText();
            String plantTime2 = cultivateTable.getRow(2).getCell(1).getText();
            String production2 = cultivateTable.getRow(2).getCell(2).getText();
            String yeild2 = cultivateTable.getRow(2).getCell(3).getText();
            String quality2 = cultivateTable.getRow(2).getCell(4).getText();
            String timeOfPlanting2 = cultivateTable.getRow(2).getCell(5).getText();
            String harvest2 = cultivateTable.getRow(2).getCell(6).getText();

            String cropName3 = cultivateTable.getRow(3).getCell(0).getText();
            String plantTime3 = cultivateTable.getRow(3).getCell(1).getText();
            String production3 = cultivateTable.getRow(3).getCell(2).getText();
            String yeild3 = cultivateTable.getRow(3).getCell(3).getText();
            String quality3 = cultivateTable.getRow(3).getCell(4).getText();
            String timeOfPlanting3 = cultivateTable.getRow(3).getCell(5).getText();
            String harvest3 = cultivateTable.getRow(3).getCell(6).getText();

            String cropName4 = cultivateTable.getRow(4).getCell(0).getText();
            String plantTime4 = cultivateTable.getRow(4).getCell(1).getText();
            String production4 = cultivateTable.getRow(4).getCell(2).getText();
            String yeild4 = cultivateTable.getRow(4).getCell(3).getText();
            String quality4 = cultivateTable.getRow(4).getCell(4).getText();
            String timeOfPlanting4 = cultivateTable.getRow(4).getCell(5).getText();
            String harvest4 = cultivateTable.getRow(4).getCell(6).getText();

            String cropName5 = cultivateTable.getRow(5).getCell(0).getText();
            String plantTime5 = cultivateTable.getRow(5).getCell(1).getText();
            String production5 = cultivateTable.getRow(5).getCell(2).getText();
            String yeild5 = cultivateTable.getRow(5).getCell(3).getText();
            String quality5 = cultivateTable.getRow(5).getCell(4).getText();
            String timeOfPlanting5 = cultivateTable.getRow(5).getCell(5).getText();
            String harvest5 = cultivateTable.getRow(5).getCell(6).getText();


            XWPFTable cultivateTable10 = tables.get(10);

            String crName1 = cultivateTable10.getRow(1).getCell(1).getText();
            String salesDestination1 = cultivateTable10.getRow(1).getCell(2).getText();
            String salesVolume1 = cultivateTable10.getRow(1).getCell(3).getText();
            String totalSales1 = cultivateTable10.getRow(1).getCell(4).getText();
            String unitPrice1 = cultivateTable10.getRow(1).getCell(5).getText();

            String crName2 = cultivateTable10.getRow(2).getCell(1).getText();
            String salesDestination2 = cultivateTable10.getRow(2).getCell(2).getText();
            String salesVolume2 = cultivateTable10.getRow(2).getCell(3).getText();
            String totalSales2 = cultivateTable10.getRow(2).getCell(4).getText();
            String unitPrice2 = cultivateTable10.getRow(2).getCell(5).getText();

            String crName3 = cultivateTable10.getRow(3).getCell(1).getText();
            String salesDestination3 = cultivateTable10.getRow(3).getCell(2).getText();
            String salesVolume3 = cultivateTable10.getRow(3).getCell(3).getText();
            String totalSales3 = cultivateTable10.getRow(3).getCell(4).getText();
            String unitPrice3 = cultivateTable10.getRow(3).getCell(5).getText();

            String crName4 = cultivateTable10.getRow(4).getCell(1).getText();
            String salesDestination4 = cultivateTable10.getRow(4).getCell(2).getText();
            String salesVolume4 = cultivateTable10.getRow(4).getCell(3).getText();
            String totalSales4 = cultivateTable10.getRow(4).getCell(4).getText();
            String unitPrice4 = cultivateTable10.getRow(4).getCell(5).getText();

            String crName5 = cultivateTable10.getRow(5).getCell(1).getText();
            String salesDestination5 = cultivateTable10.getRow(5).getCell(2).getText();
            String salesVolume5 = cultivateTable10.getRow(5).getCell(3).getText();
            String totalSales5 = cultivateTable10.getRow(5).getCell(4).getText();
            String unitPrice5 = cultivateTable10.getRow(5).getCell(5).getText();

            String crName6 = resolve(() -> cultivateTable10.getRow(6).getCell(1).getText()).orElse("");
            String salesDestination6 = resolve(() -> cultivateTable10.getRow(6).getCell(2).getText()).orElse("");
            String salesVolume6 = resolve(() -> cultivateTable10.getRow(6).getCell(3).getText()).orElse("");
            String totalSales6 = resolve(() -> cultivateTable10.getRow(6).getCell(4).getText()).orElse("");
            String unitPrice6 = resolve(() -> cultivateTable10.getRow(6).getCell(5).getText()).orElse("");

            XWPFTable cultivateTable12 = tables.get(12);

            String cName1 = resolve(() -> cultivateTable12.getRow(3).getCell(1).getText()).orElse("");
            String cMarket1 = resolve(() -> cultivateTable12.getRow(3).getCell(2).getText()).orElse("");
            String cOther1 = resolve(() -> cultivateTable12.getRow(3).getCell(3).getText()).orElse("");
            String cTotal1 = resolve(() -> cultivateTable12.getRow(3).getCell(4).getText()).orElse("");

            String cName2 = resolve(() -> cultivateTable12.getRow(4).getCell(1).getText()).orElse("");
            String cMarket2 = resolve(() -> cultivateTable12.getRow(4).getCell(2).getText()).orElse("");
            String cOther2 = resolve(() -> cultivateTable12.getRow(4).getCell(3).getText()).orElse("");
            String cTotal2 = resolve(() -> cultivateTable12.getRow(4).getCell(4).getText()).orElse("");

            String cName3 = resolve(() -> cultivateTable12.getRow(5).getCell(1).getText()).orElse("");
            String cMarket3 = resolve(() -> cultivateTable12.getRow(5).getCell(2).getText()).orElse("");
            String cOther3 = resolve(() -> cultivateTable12.getRow(5).getCell(3).getText()).orElse("");
            String cTotal3 = resolve(() -> cultivateTable12.getRow(5).getCell(4).getText()).orElse("");

            String cName4 = resolve(() -> cultivateTable12.getRow(6).getCell(1).getText()).orElse("");
            String cMarket4 = resolve(() -> cultivateTable12.getRow(6).getCell(2).getText()).orElse("");
            String cOther4 = resolve(() -> cultivateTable12.getRow(6).getCell(3).getText()).orElse("");
            String cTotal4 = resolve(() -> cultivateTable12.getRow(6).getCell(4).getText()).orElse("");

            String cName5 = resolve(() -> cultivateTable12.getRow(7).getCell(1).getText()).orElse("");
            String cMarket5 = resolve(() -> cultivateTable12.getRow(7).getCell(2).getText()).orElse("");
            String cOther5 = resolve(() -> cultivateTable12.getRow(7).getCell(3).getText()).orElse("");
            String cTotal5 = resolve(() -> cultivateTable12.getRow(7).getCell(4).getText()).orElse("");

            XWPFTable cultivateTable13 = tables.get(13);

            String season1_10years = resolve(() -> cultivateTable13.getRow(1).getCell(1).getText()).orElse("");
            String season1_05years = resolve(() -> cultivateTable13.getRow(1).getCell(2).getText()).orElse("");
            String season1CurrentYears = resolve(() -> cultivateTable13.getRow(1).getCell(3).getText()).orElse("");

            String season2_10years = resolve(() -> cultivateTable13.getRow(2).getCell(1).getText()).orElse("");
            String season2_05years = resolve(() -> cultivateTable13.getRow(2).getCell(2).getText()).orElse("");
            String season2CurrentYears = resolve(() -> cultivateTable13.getRow(2).getCell(3).getText()).orElse("");

            String season3_10years = resolve(() -> cultivateTable13.getRow(3).getCell(1).getText()).orElse("");
            String season3_05years = resolve(() -> cultivateTable13.getRow(3).getCell(2).getText()).orElse("");
            String season3CurrentYears = resolve(() -> cultivateTable13.getRow(3).getCell(3).getText()).orElse("");

            String season4_10years = resolve(() -> cultivateTable13.getRow(4).getCell(1).getText()).orElse("");
            String season4_05years = resolve(() -> cultivateTable13.getRow(4).getCell(2).getText()).orElse("");
            String season4CurrentYears = resolve(() -> cultivateTable13.getRow(4).getCell(3).getText()).orElse("");

            XWPFTable cultivateTable14 = tables.get(14);

            String typeOfProblem1 = resolve(() -> cultivateTable14.getRow(1).getCell(1).getText()).orElse("");
            String specifyProblem1 = resolve(() -> cultivateTable14.getRow(1).getCell(2).getText()).orElse("");
            String increaseProblem1 = resolve(() -> cultivateTable14.getRow(1).getCell(3).getText()).orElse("");

            String typeOfProblem2 = resolve(() -> cultivateTable14.getRow(2).getCell(1).getText()).orElse("");
            String specifyProblem2 = resolve(() -> cultivateTable14.getRow(2).getCell(2).getText()).orElse("");
            String increaseProblem2 = resolve(() -> cultivateTable14.getRow(2).getCell(3).getText()).orElse("");

            String typeOfProblem3 = resolve(() -> cultivateTable14.getRow(3).getCell(1).getText()).orElse("");
            String specifyProblem3 = resolve(() -> cultivateTable14.getRow(3).getCell(2).getText()).orElse("");
            String increaseProblem3 = resolve(() -> cultivateTable14.getRow(3).getCell(3).getText()).orElse("");

            String typeOfProblem4 = resolve(() -> cultivateTable14.getRow(4).getCell(1).getText()).orElse("");
            String specifyProblem4 = resolve(() -> cultivateTable14.getRow(4).getCell(2).getText()).orElse("");
            String increaseProblem4 = resolve(() -> cultivateTable14.getRow(4).getCell(3).getText()).orElse("");

            String typeOfProblem5 = resolve(() -> cultivateTable14.getRow(5).getCell(1).getText()).orElse("");
            String specifyProblem5 = resolve(() -> cultivateTable14.getRow(5).getCell(2).getText()).orElse("");
            String increaseProblem5 = resolve(() -> cultivateTable14.getRow(5).getCell(3).getText()).orElse("");

            String typeOfProblem6 = resolve(() -> cultivateTable14.getRow(6).getCell(1).getText()).orElse("");
            String specifyProblem6 = resolve(() -> cultivateTable14.getRow(6).getCell(2).getText()).orElse("");
            String increaseProblem6 = resolve(() -> cultivateTable14.getRow(6).getCell(3).getText()).orElse("");

            String typeOfProblem7 = resolve(() -> cultivateTable14.getRow(7).getCell(1).getText()).orElse("");
            String specifyProblem7 = resolve(() -> cultivateTable14.getRow(7).getCell(2).getText()).orElse("");
            String increaseProblem7 = resolve(() -> cultivateTable14.getRow(7).getCell(3).getText()).orElse("");

            String typeOfProblem8 = resolve(() -> cultivateTable14.getRow(8).getCell(1).getText()).orElse("");
            String specifyProblem8 = resolve(() -> cultivateTable14.getRow(8).getCell(2).getText()).orElse("");
            String increaseProblem8 = resolve(() -> cultivateTable14.getRow(8).getCell(3).getText()).orElse("");

            Object[][] cultivateRowData = {
                    {
                            name,
                            cropName1, plantTime1, production1, yeild1, quality1, timeOfPlanting1, harvest1,
                            cropName2, plantTime2, production2, yeild2, quality2, timeOfPlanting2, harvest2,
                            cropName3, plantTime3, production3, yeild3, quality3, timeOfPlanting3, harvest3,
                            cropName4, plantTime4, production4, yeild4, quality4, timeOfPlanting4, harvest4,
                            cropName5, plantTime5, production5, yeild5, quality5, timeOfPlanting5, harvest5, " ",
                            crName1, salesDestination1, salesVolume1, totalSales1, unitPrice1,
                            crName2, salesDestination2, salesVolume2, totalSales2, unitPrice2,
                            crName3, salesDestination3, salesVolume3, totalSales3, unitPrice3,
                            crName4, salesDestination4, salesVolume4, totalSales4, unitPrice4,
                            crName5, salesDestination5, salesVolume5, totalSales5, unitPrice5,
                            crName6, salesDestination6, salesVolume6, totalSales6, unitPrice6, "", "", "", "", "", "", "",
                            cName1, cMarket1, cOther1, cTotal1,
                            cName2, cMarket2, cOther2, cTotal2,
                            cName3, cMarket3, cOther3, cTotal3,
                            cName4, cMarket4, cOther4, cTotal4,
                            cName5, cMarket5, cOther5, cTotal5, "",
                            season1_10years, season1_05years, season1CurrentYears,
                            season2_10years, season2_05years, season2CurrentYears,
                            season3_10years, season3_05years, season3CurrentYears,
                            season4_10years, season4_05years, season4CurrentYears, "",
                            specifyProblem1, increaseProblem1,
                            specifyProblem2, increaseProblem2,
                            specifyProblem3, increaseProblem3,
                            specifyProblem4, increaseProblem4,
                            specifyProblem5, increaseProblem5,
                            specifyProblem6, increaseProblem6,
                            specifyProblem7, increaseProblem7,
                            specifyProblem8, increaseProblem8,

                    }
            };

            updateInformation(cultivateRowData, 3, f, pathname);

            // Input materials
            XWPFTable inputMaterials = tables.get(15);
            XWPFTable inputMaterials16 = tables.get(16);
            XWPFTable inputMaterials17 = tables.get(17);



            Object[][] inputMaterialsRowData = {
                    {
                            resolve(() -> inputMaterials.getRow(1).getCell(1).getText()).orElse(""), resolve(() -> inputMaterials.getRow(1).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials.getRow(1).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials.getRow(1).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials.getRow(1).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials.getRow(2).getCell(1).getText()).orElse(""), resolve(() -> inputMaterials.getRow(2).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials.getRow(2).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials.getRow(2).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials.getRow(2).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials.getRow(3).getCell(1).getText()).orElse(""), resolve(() -> inputMaterials.getRow(3).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials.getRow(3).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials.getRow(3).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials.getRow(3).getCell(5).getText()).orElse(""),
                            "","","","",
                            "","","","","", "", resolve(() -> inputMaterials16.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> inputMaterials17.getRow(2).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(2).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(2).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(2).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(3).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(3).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(3).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(3).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(4).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(4).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(4).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(4).getCell(5).getText()).orElse(""),

                            resolve(() -> inputMaterials17.getRow(6).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(6).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(6).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(6).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(7).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(7).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(7).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(7).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(8).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(8).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(8).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(8).getCell(5).getText()).orElse(""),

                            resolve(() -> inputMaterials17.getRow(10).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(10).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(10).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(10).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(11).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(11).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(11).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(11).getCell(5).getText()).orElse(""),
                            resolve(() -> inputMaterials17.getRow(12).getCell(2).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(12).getCell(3).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(12).getCell(4).getText()).orElse(""), resolve(() -> inputMaterials17.getRow(12).getCell(5).getText()).orElse(""),

                    }
            };

            updateInformation(inputMaterialsRowData, 4, f, pathname);

            //Agriculture
            XWPFTable agriculture = tables.get(18);
            XWPFTable agriculture19 = tables.get(19);
            XWPFTable agriculture20 = tables.get(20);
            Object[][] agricultureRowData = {
                    {

                            "", resolve(() -> agriculture.getRow(1).getCell(1).getText()).orElse(""), resolve(() -> agriculture.getRow(1).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture.getRow(2).getCell(1).getText()).orElse(""), resolve(() -> agriculture.getRow(2).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture.getRow(3).getCell(1).getText()).orElse(""), resolve(() -> agriculture.getRow(3).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture.getRow(4).getCell(1).getText()).orElse(""), resolve(() -> agriculture.getRow(4).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture.getRow(5).getCell(1).getText()).orElse(""), resolve(() -> agriculture.getRow(5).getCell(2).getText()).orElse(""),"",

                            resolve(() -> agriculture19.getRow(1).getCell(0).getText()).orElse(""),resolve(() -> agriculture19.getRow(1).getCell(1).getText()).orElse(""),resolve(() -> agriculture19.getRow(1).getCell(2).getText()).orElse(""),resolve(() -> agriculture19.getRow(1).getCell(3).getText()).orElse(""),
                            resolve(() -> agriculture19.getRow(2).getCell(0).getText()).orElse(""),resolve(() -> agriculture19.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> agriculture19.getRow(2).getCell(2).getText()).orElse(""),resolve(() -> agriculture19.getRow(2).getCell(3).getText()).orElse(""),
                            resolve(() -> agriculture19.getRow(3).getCell(0).getText()).orElse(""),resolve(() -> agriculture19.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> agriculture19.getRow(3).getCell(2).getText()).orElse(""),resolve(() -> agriculture19.getRow(3).getCell(3).getText()).orElse(""),
                            "", "", "", "",
                            "", "", "", "", "",
                            resolve(() -> agriculture20.getRow(1).getCell(0).getText()).orElse(""),resolve(() -> agriculture20.getRow(1).getCell(1).getText()).orElse(""),resolve(() -> agriculture20.getRow(1).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture20.getRow(2).getCell(0).getText()).orElse(""),resolve(() -> agriculture20.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> agriculture20.getRow(2).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture20.getRow(3).getCell(0).getText()).orElse(""),resolve(() -> agriculture20.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> agriculture20.getRow(3).getCell(2).getText()).orElse(""),
                            resolve(() -> agriculture20.getRow(4).getCell(0).getText()).orElse(""),resolve(() -> agriculture20.getRow(4).getCell(1).getText()).orElse(""),resolve(() -> agriculture20.getRow(4).getCell(2).getText()).orElse(""),

                    }
            };

            updateInformation(agricultureRowData, 5, f, pathname);

            //Production cost
            XWPFTable production22 = tables.get(21);
            XWPFTable production23 = tables.get(22);
            Object[][] productionCost = {
                    {
                        "","","","",
                            resolve(() -> production22.getRow(1).getCell(1).getText()).orElse(""),resolve(() -> production22.getRow(1).getCell(2).getText()).orElse(""),resolve(() -> production22.getRow(1).getCell(3).getText()).orElse(""), resolve(() -> production22.getRow(1).getCell(4).getText()).orElse(""),
                            resolve(() -> production22.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> production22.getRow(2).getCell(2).getText()).orElse(""),resolve(() -> production22.getRow(2).getCell(3).getText()).orElse(""), resolve(() -> production22.getRow(2).getCell(4).getText()).orElse(""),
                            resolve(() -> production22.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> production22.getRow(3).getCell(2).getText()).orElse(""),resolve(() -> production22.getRow(3).getCell(3).getText()).orElse(""), resolve(() -> production22.getRow(3).getCell(4).getText()).orElse(""),
                            resolve(() -> production22.getRow(4).getCell(1).getText()).orElse(""),resolve(() -> production22.getRow(4).getCell(2).getText()).orElse(""),resolve(() -> production22.getRow(4).getCell(3).getText()).orElse(""), resolve(() -> production22.getRow(4).getCell(4).getText()).orElse(""),"",

                            resolve(() -> production23.getRow(1).getCell(1).getText()).orElse(""),
                            resolve(() -> production23.getRow(2).getCell(1).getText()).orElse(""),
                            resolve(() -> production23.getRow(3).getCell(1).getText()).orElse(""),
                            resolve(() -> production23.getRow(5).getCell(0).getText()).orElse(""),



                    }
            };

            updateInformation(productionCost, 6, f, pathname);

            // Price fluctuation
            XWPFTable priceFlactuation26 = tables.get(24);
            Object[][] priceFlatuationRowData = {
                    {
                        "",
                            resolve(() -> priceFlactuation26.getRow(1).getCell(1).getText()).orElse(""),
                            resolve(() -> priceFlactuation26.getRow(2).getCell(1).getText()).orElse(""),
                            resolve(() -> priceFlactuation26.getRow(3).getCell(1).getText()).orElse(""),
                            resolve(() -> priceFlactuation26.getRow(4).getCell(1).getText()).orElse("")
                    }
            };

            updateInformation(priceFlatuationRowData, 8, f, pathname);


            //Contact farming
            XWPFTable cultivationFarmingTable28 = tables.get(25);
            XWPFTable cultivationFarmingTable29 = tables.get(26);
            XWPFTable cultivationFarmingTable30 = tables.get(27);
            XWPFTable cultivationFarmingTable31 = tables.get(28);
            XWPFTable cultivationFarmingTable32 = tables.get(29);

            Object[][] contactFarmingRowData = {
                    {
                        "","","",
                            resolve(() -> cultivationFarmingTable28.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> cultivationFarmingTable29.getRow(1).getCell(0).getText()).orElse("") + resolve(() -> cultivationFarmingTable29.getRow(1).getCell(1).getText()).orElse(""), "", "", "", "","","","","","","",
                            resolve(() -> cultivationFarmingTable30.getRow(0).getCell(0).getText()).orElse(""), "", "", "",
                            resolve(() -> cultivationFarmingTable31.getRow(1).getCell(0).getText()).orElse("") + resolve(() -> cultivationFarmingTable31.getRow(1).getCell(1).getText()).orElse(""), "", "", "", "","","","","","","",
                            resolve(() -> cultivationFarmingTable32.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                    }
            };

            updateInformation(contactFarmingRowData, 9, f, pathname);



            //Cultivation planning
            XWPFTable cultivationTable34 = tables.get(30);
            XWPFTable cultivationTable35 = tables.get(31);
            XWPFTable cultivationTable36 = tables.get(32);
            XWPFTable cultivationTable37 = tables.get(33);
            XWPFTable cultivationTable38 = tables.get(34);
            XWPFTable cultivationTable39 = tables.get(35);
            XWPFTable cultivationTable40 = tables.get(36);

            Object[][] cultivationRowData = {
                    {
                        "","","",
                            resolve(() -> cultivationTable34.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "","",
                            resolve(() -> cultivationTable35.getRow(0).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> cultivationTable35.getRow(1).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> cultivationTable35.getRow(1).getCell(2).getText()).orElse(""), "", "", "", "",
                            resolve(() -> cultivationTable36.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> cultivationTable37.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> cultivationTable38.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> cultivationTable39.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> cultivationTable40.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                    }
            };
            updateInformation(cultivationRowData, 10, f, pathname);


            // Irrigation
            XWPFTable irrigation41 = tables.get(37);
            XWPFTable irrigation42 = tables.get(38);
            XWPFTable irrigation43 = tables.get(39);
            XWPFTable irrigation44 = tables.get(40);
            XWPFTable irrigation45 = tables.get(41);
            XWPFTable irrigation46 = tables.get(42);
            XWPFTable irrigation47 = tables.get(43);
            XWPFTable irrigation48 = tables.get(44);
            XWPFTable irrigation49 = tables.get(45);

            Object[][] irrigationInformation = {
                    {
                        "", resolve(() -> irrigation41.getRow(0).getCell(1).getText()).orElse(""), "", resolve(() -> irrigation41.getRow(1).getCell(1).getText()).orElse(""), resolve(() -> irrigation41.getRow(1).getCell(2).getText()).orElse(""), "",
                            resolve(() -> irrigation42.getRow(1).getCell(1).getText()).orElse(""),"", "",
                            resolve(() -> irrigation42.getRow(2).getCell(1).getText()).orElse(""),"", "",
                            resolve(() -> irrigation42.getRow(3).getCell(1).getText()).orElse(""),"", "",
                            resolve(() -> irrigation42.getRow(4).getCell(1).getText()).orElse(""),"", "",
                            resolve(() -> irrigation42.getRow(5).getCell(1).getText()).orElse(""),"", "",
                            resolve(() -> irrigation43.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation44.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation45.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation46.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation47.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation48.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> irrigation49.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                    }
            };
            updateInformation(irrigationInformation, 11, f, pathname);

            // Training program
            XWPFTable trainingProgram51 = tables.get(46);
            XWPFTable trainingProgram52 = tables.get(47);
            XWPFTable trainingProgram53 = tables.get(48);
            XWPFTable trainingProgram54 = tables.get(49);

            Object[][] trainingprogramInformation = {
                    {
                        "", "", "",
                            resolve(() -> trainingProgram51.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainingProgram52.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainingProgram53.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "","",

                            resolve(() -> trainingProgram54.getRow(1).getCell(1).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(1).getCell(2).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(1).getCell(3).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(1).getCell(4).getText()).orElse(""),
                            resolve(() -> trainingProgram54.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(2).getCell(2).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(2).getCell(3).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(2).getCell(4).getText()).orElse(""),
                            resolve(() -> trainingProgram54.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(3).getCell(2).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(3).getCell(3).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(3).getCell(4).getText()).orElse(""),
                            resolve(() -> trainingProgram54.getRow(4).getCell(1).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(4).getCell(2).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(4).getCell(3).getText()).orElse(""),resolve(() -> trainingProgram54.getRow(4).getCell(4).getText()).orElse(""),


                    }
            };
            updateInformation(trainingprogramInformation, 12, f, pathname);


            //For women value chain actors
            XWPFTable chainActor73 = tables.get(67);
            XWPFTable chainActor74 = tables.get(68);
            XWPFTable chainActor75 = tables.get(69);
            XWPFTable chainActor76 = tables.get(70);
            Object[][] chainActor = {
                    {
                        "","","","","","","",
                            resolve(() -> chainActor73.getRow(1).getCell(0).getText()).orElse(""),resolve(() -> chainActor73.getRow(1).getCell(1).getText()).orElse(""),
                            resolve(() -> chainActor73.getRow(2).getCell(0).getText()).orElse(""),resolve(() -> chainActor73.getRow(2).getCell(1).getText()).orElse(""),
                            resolve(() -> chainActor73.getRow(3).getCell(0).getText()).orElse(""),resolve(() -> chainActor73.getRow(3).getCell(1).getText()).orElse(""),
                            resolve(() -> chainActor73.getRow(4).getCell(0).getText()).orElse(""),resolve(() -> chainActor73.getRow(4).getCell(1).getText()).orElse(""),
                            resolve(() -> chainActor73.getRow(5).getCell(0).getText()).orElse(""),resolve(() -> chainActor73.getRow(5).getCell(1).getText()).orElse(""),
                            resolve(() -> chainActor74.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> chainActor75.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> chainActor76.getRow(1).getCell(0).getText()).orElse("")


                    }
            };

            updateInformation(chainActor, 16, f, pathname);




            XWPFTable question78 = tables.get(71);
            XWPFTable question79 = tables.get(72);
            XWPFTable question80 = tables.get(73);
            XWPFTable question81 = tables.get(74);
            XWPFTable question82 = tables.get(75);
            XWPFTable question83 = tables.get(76);
            XWPFTable question84 = tables.get(77);
            XWPFTable question85 = tables.get(78);
            XWPFTable question86 = tables.get(79);
            XWPFTable question87 = tables.get(80);
            XWPFTable question88 = tables.get(81);
            XWPFTable question90 = tables.get(83);
            XWPFTable question91 = tables.get(84);
            XWPFTable question92 = tables.get(85);
            XWPFTable question93 = tables.get(86);
            XWPFTable question94 = tables.get(87);
            Object[][] farmerData = {
                    {
                        "","","",
                            resolve(() -> question78.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question79.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question80.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question81.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question82.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question83.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question84.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question85.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question86.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question87.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question88.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            "", "", "", "", "", "",
                            resolve(() -> question90.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question91.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question92.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question93.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> question94.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                    }

            };

            updateInformation(farmerData, 17, f, pathname);

            // Training program 2
            XWPFTable trainProgram55 = tables.get(50);
            XWPFTable trainProgram56 = tables.get(51);
            XWPFTable trainProgram57 = tables.get(52);
            XWPFTable trainProgram58 = tables.get(53);
            XWPFTable trainProgram59 = tables.get(54);
            XWPFTable trainProgram60 = tables.get(55);
            XWPFTable trainProgram61 = tables.get(56);

            Object[][] trainig2 = {
                    {
                            resolve(() -> trainProgram55.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram56.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram57.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram58.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram59.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram60.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> trainProgram61.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",

                    }

            };

            updateInformation(trainig2, 13, f, pathname);

            // Market information

            XWPFTable marketInformation62 = tables.get(57);
            XWPFTable marketInformation63 = tables.get(58);
            XWPFTable marketInformation64 = tables.get(59);
            XWPFTable marketInformation65 = tables.get(60);
            XWPFTable marketInformation66 = tables.get(61);
            XWPFTable marketInformation67 = tables.get(62);

            Object[][] marketInformation = {
                    {
                            resolve(() -> marketInformation62.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> marketInformation63.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> marketInformation64.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> marketInformation65.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> marketInformation66.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> marketInformation67.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",

                    }

            };
            updateInformation(marketInformation, 14, f, pathname);

            // Marketing
            XWPFTable marketing68 = tables.get(63);
            XWPFTable marketing69 = tables.get(64);
            XWPFTable marketing70 = tables.get(65);
            XWPFTable marketing71 = tables.get(66);


            Object[][] marketing = {
                    {
                            "",
                            resolve(() -> marketing68.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> marketing68.getRow(2).getCell(2).getText()).orElse(""), resolve(() -> marketing68.getRow(2).getCell(3).getText()).orElse(""),resolve(() -> marketing68.getRow(2).getCell(4).getText()).orElse(""),"",
                            resolve(() -> marketing68.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> marketing68.getRow(3).getCell(2).getText()).orElse(""), resolve(() -> marketing68.getRow(3).getCell(3).getText()).orElse(""),resolve(() -> marketing68.getRow(3).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing68.getRow(4).getCell(1).getText()).orElse(""),resolve(() -> marketing68.getRow(4).getCell(2).getText()).orElse(""), resolve(() -> marketing68.getRow(4).getCell(3).getText()).orElse(""),resolve(() -> marketing68.getRow(4).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing68.getRow(5).getCell(1).getText()).orElse(""),resolve(() -> marketing68.getRow(5).getCell(2).getText()).orElse(""), resolve(() -> marketing68.getRow(5).getCell(3).getText()).orElse(""),resolve(() -> marketing68.getRow(5).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing69.getRow(0).getCell(0).getText()).orElse(""), "","","","","",
                            resolve(() -> marketing70.getRow(0).getCell(0).getText()).orElse(""), "","","","","","",

                            resolve(() -> marketing71.getRow(1).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(1).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(1).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(1).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing71.getRow(2).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(2).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(2).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(2).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing71.getRow(3).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(3).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(3).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(3).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing71.getRow(4).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(4).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(4).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(4).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing71.getRow(5).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(5).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(5).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(5).getCell(4).getText()).orElse(""),
                            resolve(() -> marketing71.getRow(6).getCell(1).getText()).orElse(""),resolve(() -> marketing71.getRow(6).getCell(2).getText()).orElse(""),resolve(() -> marketing71.getRow(6).getCell(3).getText()).orElse(""),resolve(() -> marketing71.getRow(6).getCell(4).getText()).orElse(""),


                    }

            };
            updateInformation(marketing, 15, f, pathname);

            // Women
            XWPFTable women95 = tables.get(88);
            XWPFTable women96 = tables.get(89);
            XWPFTable women97 = tables.get(90);
            XWPFTable women98 = tables.get(91);
            XWPFTable women99 = tables.get(92);
            XWPFTable women100 = tables.get(93);

            Object[][] womenInformation = {
                    {
                            resolve(() -> women95.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> women96.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> women97.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> women98.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> women99.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> women100.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",

                    }

            };

            updateInformation(womenInformation, 18, f, pathname);

            // Health
            XWPFTable health101 = tables.get(94);
            XWPFTable health102 = tables.get(95);
            XWPFTable health103 = tables.get(96);

            Object[][] healthInformation = {
                    {
                            resolve(() -> health101.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> health102.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> health103.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",

                    }

            };

            updateInformation(healthInformation, 19, f, pathname);

            // Environment
            XWPFTable environment104 = tables.get(97);
            XWPFTable environment105 = tables.get(98);
            XWPFTable environment106 = tables.get(99);
            XWPFTable environment107 = tables.get(100);
            XWPFTable environment108 = tables.get(101);

            Object[][] environmentInformation = {
                    {
                            resolve(() -> environment104.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> environment105.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> environment106.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> environment107.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",
                            resolve(() -> environment108.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "",

                    }

            };
            updateInformation(environmentInformation, 20, f, pathname);

            //Impact of covid19
            XWPFTable covid109 = tables.get(102);
            XWPFTable covid110 = tables.get(103);
            XWPFTable covid111 = tables.get(104);
            XWPFTable covid112 = tables.get(105);

            Object[][] ipmactOfCovid19 = {
                    {
                            resolve(() -> covid109.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "", "","",
                            resolve(() -> covid110.getRow(1).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> covid110.getRow(2).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> covid111.getRow(1).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> covid111.getRow(2).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> covid111.getRow(3).getCell(1).getText()).orElse(""), "", "", "", "",
                            resolve(() -> covid111.getRow(4).getCell(1).getText()).orElse(""), "", "", "", "","",
                            resolve(() -> covid112.getRow(0).getCell(0).getText()).orElse(""), "", "", "", "",

                    }

            };

            updateInformation(ipmactOfCovid19, 21, f, pathname);


            fis.close();
        } catch (Exception exep) {
            exep.printStackTrace();
        }
    }

    private void updateInformation(Object[][] bookData, int sheetNo, File f, String pathname) {
        try {
            FileInputStream inputStream = new FileInputStream(f);
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(sheetNo);

            int rowCount = sheet.getLastRowNum();

            for (Object[] aBook : bookData) {
                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                Cell cell;
//                    cell.setCellValue(rowCount);

                for (Object field : aBook) {
                    cell = row.createCell(columnCount++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }

            }

            inputStream.close();

            FileOutputStream outputStream = new FileOutputStream(pathname);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

    public static <T> Optional<T> resolve(Supplier<T> resolver) {
        try {
            T result = resolver.get();
            return Optional.ofNullable(result);
        } catch (NullPointerException e) {
            return Optional.empty();
        }
    }
}
