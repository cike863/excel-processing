/**
 *
 */
package util;

import org.apache.commons.lang3.StringUtils;

import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;

/**
 * 使用 Duration.between(startTime, endTime).getSeconds() 去计算时间差
 * Period.between(startTime, endTime) 计算天数差
 *
 * @version: v1.0
 * @author: Haixiang.Dai
 * project: xfxb-base-component
 * copyright:
 * createTime:
 * modifyTime:
 * modifyBy:
 */
public final class DateTimeUtil {

    /**
     * 返回销毁日期，暂时使用为2999年到期。
     *
     * @return
     */
    public static LocalDateTime buildAbolishTime() {
        return LocalDateTime.of(2999, 12, 31, 23, 59, 59);
    }

    /**
     * 把当前Date型时间转换成LocalDatetime类型的时间
     *
     * @param currentDate
     * @return
     */
    public static LocalDateTime buildDateToLocalDateTime(Date currentDate) {
        Instant expiredTimeInstant = Instant.ofEpochMilli(currentDate.getTime());
        LocalDateTime currentLocalDateTime = LocalDateTime.ofInstant(expiredTimeInstant, ZoneId.systemDefault());
        return currentLocalDateTime;
    }

    /**
     * 转换long型时间为LocalDateTime
     *
     * @param dateTime
     * @return
     */
    public static LocalDateTime buildLongToLocalDateTime(long dateTime) {
        Calendar calendar = Calendar.getInstance();
        Date createTime = new Date(dateTime);
        calendar.setTime(createTime);

        int year = calendar.get(Calendar.YEAR);// 获取年份
        int month = calendar.get(Calendar.MONTH) + 1;// 获取月份
        int day = calendar.get(Calendar.DATE);// 获取日
        int hour = calendar.get(Calendar.HOUR);// 小时
        int minute = calendar.get(Calendar.MINUTE);// 分
        int second = calendar.get(Calendar.SECOND);// 秒

        LocalDateTime createTimeNew = LocalDateTime.of(year, month, day, hour, minute, second);
        return createTimeNew;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_YYYYMMDDHHMMSS(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_YYYYMMDDHHMMSSFFF(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmssfff");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_YYMMDDHHMMSS(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMMddHHmmss");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_YYMMDDHHMMSSFFF(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMMddHHmmssfff");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_YYMMDDHHMM(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMMddHHmm");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据LocalDateTime转成String型日期。
     *
     * @param dateTime
     * @return
     */
    public static String buildLocalDateTimeToString_yyyy_MM_dd_HH_mm(LocalDateTime dateTime) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
        String formatted = dateTime.format(dateFormatter);
        return formatted;
    }

    /**
     * 根据String型时间日期转成LocalDateTime
     *
     * @param dateTimeStr
     * @return
     */
    public static LocalDateTime buildStringToLocalDateTime(String dateTimeStr) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        TemporalAccessor temporalAccessor = dateFormatter.parse(dateTimeStr);
        LocalDateTime value = LocalDateTime.from(temporalAccessor);
        return value;
    }

    /**
     * 根据String型时间日期转成LocalDateTime
     *
     * @param dateTimeStr
     * @return
     */
    public static LocalDateTime buildStringToLocalDateTimeWithHHMM(String dateTimeStr) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
        TemporalAccessor temporalAccessor = dateFormatter.parse(dateTimeStr);
        LocalDateTime value = LocalDateTime.from(temporalAccessor);
        return value;
    }

    /**
     * 根据LocalDateTime转成20150613这种格式的String型日期。
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString_yyyyMMdd(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换Sting型日期为LocalDate格式为：yyyyMMdd
     *
     * @param dateString
     * @return
     */
    public static LocalDate buildStringToLocalDate_yyyyMMdd(String dateString) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        TemporalAccessor temporalAccessor = dateFormatter.parse(dateString);
        LocalDate value = LocalDate.from(temporalAccessor);
        return value;
    }

    /**
     * 转换Sting型日期为LocalDate
     *
     * @param dateString
     * @return
     */
    public static LocalDate buildStringToLocalDate(String dateString) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        TemporalAccessor temporalAccessor = dateFormatter.parse(dateString);
        LocalDate value = LocalDate.from(temporalAccessor);
        return value;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString_yyyy_MM(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString_dd(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("dd");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString_yyMM(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMM");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildLocalDateToString_yyMMdd(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMMdd");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 转换LocalDate为String型日期
     *
     * @param date
     * @return
     */
    public static String buildTheDayMin(LocalDate date) {
        DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyMMdd");
        String formatted = date.format(dateFormatter);
        return formatted;
    }

    /**
     * 获取某个日期的当天最小时间
     *
     * @param startTime
     * @return
     */
    public static String setStartTimeMin(String startTime) {
        if (StringUtils.isBlank(startTime)) {
            return startTime;
        }
        if (startTime.length() > 11) {
            return startTime;
        }
        return startTime + " " + LocalTime.MIN;
    }

    /**
     * 获取某个日期的当天最大时间
     *
     * @param endTime
     * @return
     */
    public static String setEndTimeMax(String endTime) {
        if (StringUtils.isBlank(endTime)) {
            return endTime;
        }
        if (endTime.length() > 11) {
            return endTime;
        }
        return endTime + " " + LocalTime.MAX;
    }
}
