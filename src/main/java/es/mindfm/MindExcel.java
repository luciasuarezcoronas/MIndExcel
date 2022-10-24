package es.mindfm;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import es.mindfm.dto.AssetDTO;
import es.mindfm.excel.AssetCell;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Slf4j
public class MindExcel {

    private static final String FILE_NAME = "src/main/resources/Assets.xlsx";




    public static void main(String[] args) {

        try {
            List<AssetDTO> assetDTOList = new ArrayList<AssetDTO>();
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            for(Row row : sheet){
                if(row.getRowNum() == 0){
                    if (!row.getCell(AssetCell.SITE_SPACE_ID.getPosition()).getStringCellValue().equals(AssetCell.SITE_SPACE_ID.name())
                            ||!row.getCell(AssetCell.DESCRIPTION.getPosition()).getStringCellValue().equals(AssetCell.DESCRIPTION.name())
                            ||!row.getCell(AssetCell.ELEMENTS .getPosition()).getStringCellValue().equals(AssetCell.ELEMENTS.name())
                            ||!row.getCell(AssetCell.TYPE.getPosition()).getStringCellValue().equals(AssetCell.TYPE.name())
                            ||!row.getCell(AssetCell.STATUS.getPosition()).getStringCellValue().equals(AssetCell.STATUS.name())
                    ) {
                        log.error("Header format file not supported for file "+row.getCell(AssetCell.SITE_SPACE_ID.getPosition()).getStringCellValue()+" --> "+AssetCell.SITE_SPACE_ID.name());
                    }
                }
                if(row.getRowNum()> 0) {
                    if(!(row.getCell(AssetCell.SITE_SPACE_ID.getPosition())== null ||
                            row.getCell(AssetCell.SITE_SPACE_ID.getPosition()).getNumericCellValue() == Double.valueOf(0).doubleValue())
                      && !(row.getCell(AssetCell.DESCRIPTION.getPosition())== null ||
                            row.getCell(AssetCell.DESCRIPTION.getPosition()).getStringCellValue().isEmpty())
                    ){
                        try {


                        AssetDTO dto = AssetDTO.builder()
                                .siteSpaceId( Double.valueOf(row.getCell(AssetCell.SITE_SPACE_ID.getPosition()).getNumericCellValue()).longValue())
                                .description(row.getCell(AssetCell.DESCRIPTION.getPosition()).getStringCellValue())
                                .build();
                        assetDTOList.add(dto);
                        } catch (NumberFormatException nfe) {
                            System.out.println("NumberFormatException: " + nfe.getMessage());
                        }


                    }
                    else {
                        log.info("Row  {} not processed, because missed data",row.getRowNum());
                        break;
                    }
                }
            }

            log.info(" ASSETS: "+assetDTOList.size());

        } catch (Exception e) {
            log.error("Error reading and decorating data from Assets ", e.getMessage());

        }

    }

}
