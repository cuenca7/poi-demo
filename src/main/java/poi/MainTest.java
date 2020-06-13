package poi;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;

public class MainTest {

    public static void main(String[] args) {
        List<ExcelDataVO> data = mockData();
        writerData(data);
    }

    private static void writerData(List<ExcelDataVO> data) {
        Workbook workbook = ExcelWriter.exportData(data);
        FileOutputStream fos = null;
        try {
            String exportFilePath = "/Users/cuenca/Desktop/demo.xls";
            File file = new File(exportFilePath);
            if (!file.exists()) {
                file.createNewFile();
            }
            fos = new FileOutputStream(exportFilePath);
            workbook.write(fos);
            fos.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }
    }

    private static List<ExcelDataVO> mockData() {
        List<ExcelDataVO> data = new ArrayList<ExcelDataVO>(100);
        int index = 1;
        int securityCode = 60000;
        for (int i = 0; i < 100; i++) {
            ExcelDataVO vo = new ExcelDataVO();
            vo.setFundCode(String.format("%05d", (index++)));
            vo.setRecordCreateDate(Date.valueOf("2020-02-02"));
            vo.setSecurityCode(String.valueOf(securityCode++));
            vo.setPurchRedmDate(Date.valueOf("2020-02-03"));
            vo.setPurchRedmQuantity(BigDecimal.TEN);
            data.add(vo);
        }
        return data;
    }

}
