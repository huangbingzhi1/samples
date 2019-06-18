package com.hisense.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NonNull;
import lombok.RequiredArgsConstructor;
import java.util.Map;
import java.util.Set;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/12 8:43
 * @Version 1.0
 */
@Data
@AllArgsConstructor
@RequiredArgsConstructor
public class Enterprise{
    /**
     * 编码
     **/
    @NonNull
    private String eCode;
    /**
     * 分公司
     **/
    @NonNull
    private String company;
    /**
     * 名称
     **/
    @NonNull
    private String eName;
    /**
     * cis编码
     **/
    @NonNull
    private String cisCode;
    /**
     * 省
     **/
    @NonNull
    private String province;
    /**
     * 市
     **/
    @NonNull
    private String city;
    /**
     * 区
     **/
    @NonNull
    private String district;
    /**
     * 市场级别
     **/
    @NonNull
    private String marketLevel;
    /**
     * 营销模式
     **/
    @NonNull
    private String marketModel;
    /**
     * 商家经理联系人
     **/
    @NonNull
    private String manager;
    /**
     * 商家经理联系人电话
     **/
    @NonNull
    private String managerPhone;
    /**
     * 门店名称
     **/
    @NonNull
    private String storeName;
    /**
     * 门店编码
     **/
    @NonNull
    private String storeCode;
    /**
     * 办事处名称
     **/
    @NonNull
    private String officeName;
    /**
     * 办事处经理编号
     **/

    @NonNull
    private String officeManagerCode;
    /**
     * 办事处经理姓名
     **/
    @NonNull
    private String officeManagerName;
    /**
     * 产品总监,key包括
     */
    @NonNull
    private Map<String, People> inspector;
    /**
     * 产品/客户经理
     */
    @NonNull
    private Map<String, People> saler;

    /**
     * 所有产品/客户经理所需要的样本统计
     **/
    private int needSampleCount;

    /**
     * 产品/客户经理临时值（随着样本提取的进行，会变化）
     **/
    private Set<String> salerSelected;

}



















