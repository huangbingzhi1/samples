package com.hisense.entity;

import lombok.*;

import java.util.*;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/12 8:44
 * @Version 1.0
 */
@Data
@RequiredArgsConstructor
@AllArgsConstructor
@NoArgsConstructor
public class People{
    /**
     * 分公司名
     */
    private String company;
    /**
     * 员工编码
     */
    @NonNull
    private String pCode;
    /**
     * 姓名
     */
    @NonNull
    private String pName;
    /**
     * 二级部门
     */
    private String secondPart;
    /**
     * 职位名称
     */
    private String position;
    /**
     * 成功样本数
     */
    private int successSample;
    /*
    还需捕获样本数量
     */
    private int needSample;
}
