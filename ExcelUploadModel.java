package jp.co.wisdom_technology.employee.model;

import java.util.Date;

import lombok.Data;

/**
 * プロセス拡張モデル
 */
@Data
public class ExcelUploadModel {
	
    /**
     * 氏名
     */
    private String name;
    /**
     * 所属
     */
    private String department;
    /**
     * 年
     */
    private String year;
    /**
     * 月
     */
    private String month;    
    /**
     * 出勤日数
     */
    private double attendanceDays;
    /**
     * 休出日数 
     */
    private double daysOff;
    /**
     * 就業
     */
    private String employment;
    /**
     * 早出残
     */
    private String overtimeMorning;
    /**
     * 普通残
     */
    private double overtimeDay;
    /**
     * 深夜残
     */
    private double overtimeNight;
    /**
     * 休出
     */
    private double rest;
    /**
     * 遅刻
     */
    private double late;
    /**
     * 早退
     */
    private double advanceLeave;
    /**
     * 有休
     */
    private double paidLeave;
    /**
     * 特休
     */
    private double specialLeave;
    /**
     * 代休
     */
    private double replaceLeave;
    /**
     * 欠勤
     */
    private double unrecognizedLeave;
}