package poi;

import lombok.Data;

import java.math.BigDecimal;
import java.sql.Date;

@Data
public class ExcelDataVO {

    private String fundCode;

    private Date recordCreateDate;

    private String securityCode;

    private Date purchRedmDate;

    private BigDecimal purchRedmQuantity;

}
