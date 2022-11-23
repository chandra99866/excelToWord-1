package readingExcel;

import model.Headsmodel;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.*;
import writingToWord.WritingToWord;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

public class ReadingExcel {



private Headsmodel headsmodel= new  Headsmodel();

    FileHandler handler = new FileHandler("default.log", true);

    Logger logger = Logger.getLogger("ReadingExcel.class");

    SimpleFormatter formatter = new SimpleFormatter();


    public ReadingExcel(String path) throws IOException, InvalidFormatException {
        getAllExcelsPath(path);
    }


    private void getAllExcelsPath(String path ) throws IOException, InvalidFormatException {
        logger.addHandler(handler);
        handler.setFormatter(formatter);
       File file = new File(path);
        if(file.isDirectory()){
            logger.info("Created Directory At : "+file.getAbsolutePath()+"\\outDocs");
           logger.info( "status : "+new File(file.getAbsolutePath()+"\\outDocs").mkdir());
            for(File file1 : file.listFiles()){
                if(FilenameUtils.getExtension(file1.getAbsolutePath()).equals("xlsx")){
                    //here call the readfile
                    System.out.println(file1.getAbsolutePath());
                    readFile(file1);

                }
                else{
                    logger.info( " This file cannot be processed only xlsx file are allowed : "+file1.getName());
                }

            }
        }
        else{
            if(FilenameUtils.getExtension(path).equals("xlsx")){
                //here call the readfile
                logger.info("Created Directory At : "+file.getParent()+"\\outDocs");
                logger.info( "Status : "+new File(file.getParent()+"\\outDocs").mkdir());
                readFile(file);
            }
            else{
                logger.info( "This file cannot be processed only xlsx file are allowed");
            }
        }

    }


    private  void readFile(File file) throws IOException, InvalidFormatException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new File(file.getAbsolutePath()));

        String directory= file.getParent()+"\\outDocs\\"+FilenameUtils.getBaseName(file.getName());
        logger.info("*************************************************************Started of New File*****************************************************");
        logger.info("File name : "+FilenameUtils.getBaseName(file.getName()));
        logger.info("*************************************************************************************************************************************");
        logger.info("Created Directory At : "+directory);
        logger.info("Status "+ new File(directory).mkdir());

        for(int sheetIndex=0;sheetIndex<xssfWorkbook.getNumberOfSheets();sheetIndex++) {
            try {
                XSSFSheet sheet = xssfWorkbook.getSheetAt(sheetIndex);


                Map<Pair<Integer, Integer>, ArrayList<PictureData>> pictureDataMap = getImageLocations(sheet);

                String path = directory + "\\" + sheet.getSheetName().replace(" ", "_");
                logger.info("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Started of New sheet++++++++++++++++++++++++++++++++++++++++++++++++++++");
                logger.info("Sheet name : "+xssfWorkbook.getSheetName(sheetIndex));
                logger.info("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                logger.info("Created Directory At : "+path);
                logger.info("Status "+new File(path).mkdir());

                WritingToWord writingToWord = new WritingToWord(path + "\\" + sheet.getSheetName());
                LinkedList<String> relevantList = headsmodel.getReleventList(sheet.getSheetName());

                for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    try {

                        XSSFRow row = sheet.getRow(rowIndex);
                        LinkedHashMap<String, Object> data = new LinkedHashMap<>();


                        for (int cellIndex = 1; cellIndex <= relevantList.size(); cellIndex++) {

                            try {
                                XSSFCell cell = row.getCell(cellIndex);
                                String header = relevantList.get(cellIndex - 1);


                                if (header.equalsIgnoreCase("Publication Number")) {

                                    ArrayList<Object> arr = new ArrayList<>();
                                    arr.add(cell.getStringCellValue());

                                    if(cell.getStringCellValue().isBlank())
                                        logger.warning("NO patent found for"+" ==============>Error At cell " + cellIndex+",sheet: "+xssfWorkbook.getSheetName(sheetIndex)+",File name : "+FilenameUtils.getBaseName(file.getName()));


                                    if (cell.getHyperlink() == null){
                                        logger.warning("NO hyperlink is present for patent no : "+cell.getStringCellValue()+" ==============>Error At cell " + cellIndex+",sheet: "+xssfWorkbook.getSheetName(sheetIndex)+",File name : "+FilenameUtils.getBaseName(file.getName()));
                                        break;}


                                    arr.add(cell.getHyperlink().getAddress());
                                    data.put(header, arr);
                                    continue;
                                }


                                if (header.equalsIgnoreCase("patent focus")
                                        || header.equalsIgnoreCase("Claims")) continue;

                                Pair<Integer, Integer> pair = new Pair<>(cell.getColumnIndex(), cell.getRowIndex());

                                if (!pictureDataMap.isEmpty()) {
                                    if (pictureDataMap.containsKey(pair)) {
                                        ArrayList<Object> arr = new ArrayList<>();
                                        arr.add(header);
                                        arr.add(pictureDataMap.get(pair));
                                        data.put("pic", arr);
                                        continue;
                                    }
                                }


                                switch (cell.getCellType()) {
                                    case STRING:
                                        data.put(header, cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        if (header.toLowerCase().contains("date")) {
                                            Date itemDate = cell.getDateCellValue();
                                            String myDateStr = new SimpleDateFormat("dd-MMM-yyyy").format(itemDate);
                                            data.put(header, myDateStr);
                                        } else {
                                            data.put(header, (int) cell.getNumericCellValue());
                                        }
                                        break;

                                    case BLANK:
                                        break;
                                    case _NONE:
                                        break;
                                    default:
                                        data.put(header, cell.getNumericCellValue());


                                }


                            } catch (Exception e) {

                                logger.warning("==============>Error At cell " + cellIndex+",sheet: "+xssfWorkbook.getSheetName(sheetIndex)+",File name : "+FilenameUtils.getBaseName(file.getName()));
                                logger.warning( e.toString());
                                e.printStackTrace();

                            }
                        }
                        if (((ArrayList<Object>) data.get("Publication Number")) != null) {
                            //System.out.println(data);
                            writingToWord.writeData(data);
                        }


                    } catch (Exception e) {

                        logger.warning("==============>Error At row " + rowIndex+",sheet: "+xssfWorkbook.getSheetName(sheetIndex)+",File name : "+FilenameUtils.getBaseName(file.getName()));
                        logger.warning( e.toString());
                        e.printStackTrace();
                    }
                }
                writingToWord.close();

            }catch (Exception e){

                logger.warning("==============>Error At sheet " + xssfWorkbook.getSheetName(sheetIndex)+",File name : "+FilenameUtils.getBaseName(file.getName()));
                logger.warning( e.toString());
                e.printStackTrace();
            }
        }


    }

    private Map<Pair<Integer,Integer>, ArrayList<PictureData>> getImageLocations(XSSFSheet xssfSheet) throws FileNotFoundException {
        XSSFDrawing dp = xssfSheet.createDrawingPatriarch();
       Map<Pair<Integer,Integer>,ArrayList<PictureData>> pitcherData = new HashMap<>();
        List<XSSFShape> pics = dp.getShapes();
        for(XSSFShape pic : pics){
            XSSFPicture inpPic = (XSSFPicture)pic;
            XSSFClientAnchor clientAnchor = (XSSFClientAnchor) inpPic.getAnchor();
            inpPic.getPictureData(); // узнаю название картинки
            Pair<Integer,Integer> pair = new Pair<>((int)clientAnchor.getCol1(),(int)clientAnchor.getRow1());


            if(pitcherData.containsKey(pair)){
                pitcherData.get(pair).add(inpPic.getPictureData());
            }
            else {
                pitcherData.put(pair,new ArrayList<PictureData>());
                pitcherData.get(pair).add(inpPic.getPictureData());

            }

           // System.out.println("col: " + clientAnchor.getCol1()  + ", row: " + clientAnchor.getRow1()  );
        }

    return pitcherData;

    }










}
