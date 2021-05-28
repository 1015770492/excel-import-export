package top.yumbo.excel.util;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/24 18:51
 */
public class BigDecimalUtils {
    // 一万
    public static final String TEN_THOUSAND_STRING="10000";
    // 一亿
    public static final String ONE_HUNDRED_MILLION_STRING="1"+"0000"+"0000";
    public static final BigDecimal TEN_THOUSAND_BIG_DECIMAL = new BigDecimal(TEN_THOUSAND_STRING);
    public static final BigDecimal ONE_HUNDRED_MILLION_BIG_DECIMAL = new BigDecimal(ONE_HUNDRED_MILLION_STRING);


    /**
     * @param dived  被除数
     * @param div    除数
     * @param scales 保留位数 不限制传null
     * @return
     */
    public static BigDecimal bigDecimalDivBigDecimalFormatScale(BigDecimal dived, BigDecimal div, Integer scales) {
        if (dived == null || div == null) {
            return null;
        }
        return dived.divide(div, scales, BigDecimal.ROUND_HALF_UP);
    }

    /**
     * 保留2位小数
     *
     * @param dived 被除数
     * @param div   除数
     */
    public static BigDecimal bigDecimalDivBigDecimalFormatTwo(BigDecimal dived, BigDecimal div) {

        return bigDecimalDivBigDecimalFormatScale(dived, div, 2);
    }
}